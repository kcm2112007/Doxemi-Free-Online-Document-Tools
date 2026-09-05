/* ==========================================================================
   Doxemi — script.js
   Single-file SPA: router, search, theme, and every document tool.
   No build step. No server. Everything below runs in the browser.
   ========================================================================== */
"use strict";

/* --------------------------------------------------------------------
   0. Editable configuration
   -------------------------------------------------------------------- */
const SOCIAL_LINKS = {
  instagram: "https://www.instagram.com/goodx_official?igsh=MjViOTB0M3o5OHN6",
  youtube: "YOUR_YOUTUBE_URL",
  x: "YOUR_X_URL",
  facebook: "https://www.facebook.com/share/19CAzUUMcJ/",
  linkedin: "YOUR_LINKEDIN_URL",
  github: "YOUR_GITHUB_URL",
  whatsapp: "https://whatsapp.com/channel/0029VbDU9dmEQIameOnw2R47",
  website: "https://kcm2112007.github.io/kalicharanmurmu-/"
};
const SITE_URL = "./"; // replace with your full https://USERNAME.github.io/REPOSITORY/ once deployed

/* --------------------------------------------------------------------
   1. Small helpers
   -------------------------------------------------------------------- */
const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));

function escapeHtml(str) {
  return String(str ?? "").replace(/[&<>"']/g, (c) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
  }[c]));
}

function formatBytes(bytes) {
  if (bytes === 0) return "0 B";
  if (!bytes && bytes !== 0) return "—";
  const units = ["B", "KB", "MB", "GB"];
  const i = Math.min(units.length - 1, Math.floor(Math.log(bytes) / Math.log(1024)));
  return `${(bytes / Math.pow(1024, i)).toFixed(i === 0 ? 0 : 1)} ${units[i]}`;
}

function extOf(name) {
  const m = /\.([a-z0-9]+)$/i.exec(name || "");
  return m ? m[1].toLowerCase() : "";
}

function downloadBlob(filename, blob) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 4000);
}
function downloadText(filename, text, mime = "text/plain") {
  downloadBlob(filename, new Blob([text], { type: mime + ";charset=utf-8" }));
}

function readFileAsText(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result);
    r.onerror = () => rej(new Error("read failed"));
    r.readAsText(file);
  });
}
function readFileAsArrayBuffer(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result);
    r.onerror = () => rej(new Error("read failed"));
    r.readAsArrayBuffer(file);
  });
}

const countWords = (t) => (t.trim() ? (t.trim().match(/\S+/g) || []).length : 0);
const countChars = (t) => t.length;
const countCharsNoSpaces = (t) => t.replace(/\s/g, "").length;
const countLines = (t) => (t === "" ? 0 : t.split(/\r\n|\r|\n/).length);
const countParagraphs = (t) => t.split(/\r?\n\s*\r?\n/).map(s => s.trim()).filter(Boolean).length;

/* lazy CDN script loader, caches promises so a library loads once */
const _libCache = {};
function loadScript(src) {
  if (_libCache[src]) return _libCache[src];
  _libCache[src] = new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = src; s.async = true;
    s.onload = () => resolve();
    s.onerror = () => reject(new Error("Could not load a required library from the CDN. Check your connection and try again."));
    document.head.appendChild(s);
  });
  return _libCache[src];
}
const loadMammoth = () => loadScript("https://cdn.jsdelivr.net/npm/mammoth@1.7.0/mammoth.browser.min.js").then(() => window.mammoth);
const loadXLSX = () => loadScript("https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js").then(() => window.XLSX);
const loadJSZip = () => loadScript("https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js").then(() => window.JSZip);

/* --------------------------------------------------------------------
   2. CSV parsing / serialising (RFC 4180-ish, handles quotes + commas)
   -------------------------------------------------------------------- */
function parseCSV(text, delim = ",") {
  const rows = [];
  let row = [], field = "", inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i], next = text[i + 1];
    if (inQuotes) {
      if (c === '"' && next === '"') { field += '"'; i++; }
      else if (c === '"') { inQuotes = false; }
      else field += c;
    } else {
      if (c === '"') inQuotes = true;
      else if (c === delim) { row.push(field); field = ""; }
      else if (c === "\n") { row.push(field); rows.push(row); row = []; field = ""; }
      else if (c === "\r") { /* skip, \n handles the break */ }
      else field += c;
    }
  }
  if (field.length || row.length) { row.push(field); rows.push(row); }
  return rows.filter((r, idx) => !(idx === rows.length - 1 && r.length === 1 && r[0] === ""));
}
function toCSV(rows, delim = ",", forceQuote = false) {
  return rows.map(r => r.map(cellRaw => {
    const cell = cellRaw === undefined || cellRaw === null ? "" : String(cellRaw);
    const needsQuote = forceQuote || cell.includes(delim) || cell.includes('"') || cell.includes("\n") || cell.includes("\r");
    return needsQuote ? `"${cell.replace(/"/g, '""')}"` : cell;
  }).join(delim)).join("\r\n");
}

/* --------------------------------------------------------------------
   3. Minimal RTF -> text converter
   Strips control words/groups. Not full-fidelity, but reads real content.
   -------------------------------------------------------------------- */
function rtfToText(rtf) {
  let s = rtf;
  // drop destination groups we never want the text of
  const skipGroups = ["fonttbl", "colortbl", "stylesheet", "info", "pict", "object", "themedata", "colorschememapping", "generator", "*"];
  let out = "";
  let i = 0, depth = 0;
  const skipDepth = [];
  while (i < s.length) {
    const c = s[i];
    if (c === "\\") {
      // check for hex escape \'xx
      if (s[i + 1] === "'") {
        const hex = s.substr(i + 2, 2);
        const code = parseInt(hex, 16);
        if (!isCurrentlySkipped()) out += isNaN(code) ? "" : String.fromCharCode(code);
        i += 4; continue;
      }
      // control word
      const m = /^\\([a-zA-Z]+)(-?\d+)?\s?/.exec(s.slice(i));
      if (m) {
        const word = m[1];
        if (word === "par" || word === "line") { if (!isCurrentlySkipped()) out += "\n"; }
        if (word === "tab") { if (!isCurrentlySkipped()) out += "\t"; }
        if (["*"].includes(word)) skipDepth.push(depth + 1);
        if (skipGroups.includes(word)) skipDepth.push(depth + 1);
        i += m[0].length; continue;
      }
      // escaped brace/backslash
      if (s[i + 1] === "{" || s[i + 1] === "}" || s[i + 1] === "\\") {
        if (!isCurrentlySkipped()) out += s[i + 1];
        i += 2; continue;
      }
      i++; continue;
    }
    if (c === "{") { depth++; i++; continue; }
    if (c === "}") {
      while (skipDepth.length && skipDepth[skipDepth.length - 1] >= depth) skipDepth.pop();
      depth--; i++; continue;
    }
    if (!isCurrentlySkipped()) out += c;
    i++;
  }
  function isCurrentlySkipped() { return skipDepth.length > 0; }
  return out.replace(/\n{3,}/g, "\n\n").replace(/[ \t]{2,}/g, " ").trim();
}

/* --------------------------------------------------------------------
   4. Shared icons + UI mounters
   -------------------------------------------------------------------- */
const ICON_UPLOAD = `<svg width="30" height="30" viewBox="0 0 30 30" fill="none" aria-hidden="true"><path d="M7 4h11l5 5v17H7V4Z" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/><path d="M18 4v5h5" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/><line x1="15" y1="14" x2="15" y2="22" stroke="var(--accent)" stroke-width="1.7" stroke-linecap="round"/><path d="M11.5 18l3.5-3.5 3.5 3.5" stroke="var(--accent)" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"/></svg>`;

function iconFor(category) {
  const common = 'width="22" height="22" viewBox="0 0 22 22" fill="none" aria-hidden="true"';
  switch (category) {
    case "docx": return `<svg ${common}><path d="M5 2h8l4 4v14H5V2Z" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/><path d="M13 2v4h4" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/><path d="M7.5 12l1.2 5 1.3-4 1.3 4 1.2-5" stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round"/></svg>`;
    case "txt": return `<svg ${common}><path d="M5 2h8l4 4v14H5V2Z" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/><path d="M13 2v4h4" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/><line x1="7.5" y1="12" x2="14.5" y2="12" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"/><line x1="7.5" y1="15.5" x2="12.5" y2="15.5" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"/></svg>`;
    case "rtf": return `<svg ${common}><path d="M5 2h8l4 4v14H5V2Z" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/><path d="M13 2v4h4" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/><path d="M8 11.5h4.2a1.6 1.6 0 0 1 0 3.2H8V11.5Zm0 3.2h4.6a1.6 1.6 0 0 1 0 3.2H8" stroke="currentColor" stroke-width="1.2" stroke-linejoin="round"/></svg>`;
    case "csv": return `<svg ${common}><rect x="3" y="4" width="16" height="14" rx="1.2" stroke="currentColor" stroke-width="1.5"/><line x1="3" y1="9" x2="19" y2="9" stroke="currentColor" stroke-width="1.3"/><line x1="8.5" y1="4" x2="8.5" y2="18" stroke="currentColor" stroke-width="1.3"/><line x1="13.5" y1="4" x2="13.5" y2="18" stroke="currentColor" stroke-width="1.3"/></svg>`;
    case "xlsx": return `<svg ${common}><rect x="3" y="4" width="16" height="14" rx="1.2" stroke="currentColor" stroke-width="1.5"/><path d="M8 8l6 8M14 8l-6 8" stroke="var(--accent)" stroke-width="1.4" stroke-linecap="round"/></svg>`;
    default: return `<svg ${common}><rect x="3" y="3" width="16" height="16" rx="3" stroke="currentColor" stroke-width="1.5"/><path d="M7 8h8M7 11h8M7 14h5" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"/></svg>`;
  }
}

/** Mounts a drag/drop + click-to-browse file uploader into `el`.
 *  onFile(file, resultArea) receives the chosen File and the container
 *  to render into; it may return a Promise. */
function mountFileTool(el, { accept = "*", note = "", onFile }) {
  el.innerHTML = `
    <div class="upload-zone" id="uz" tabindex="0" role="button" aria-label="Upload a file">
      ${ICON_UPLOAD}
      <p><strong>Drop your file here</strong></p>
      <p class="muted">or</p>
      <label class="btn btn-primary btn-small" for="uz-input">Choose file</label>
      <input type="file" id="uz-input" accept="${accept}" hidden>
      ${note ? `<p class="fine-print">${note}</p>` : ""}
    </div>
    <div id="tool-body"></div>`;
  const zone = $("#uz", el), input = $("#uz-input", el), body = $("#tool-body", el);
  const openPicker = () => input.click();
  zone.addEventListener("click", openPicker);
  zone.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); openPicker(); } });
  ["dragenter", "dragover"].forEach(ev => zone.addEventListener(ev, (e) => { e.preventDefault(); zone.classList.add("drag"); }));
  ["dragleave", "drop"].forEach(ev => zone.addEventListener(ev, (e) => { e.preventDefault(); zone.classList.remove("drag"); }));
  zone.addEventListener("drop", (e) => { const f = e.dataTransfer.files[0]; if (f) handleFile(f); });
  input.addEventListener("change", (e) => { const f = e.target.files[0]; if (f) handleFile(f); });

  function handleFile(file) {
    body.innerHTML = `
      <div class="file-summary">
        <div><div class="f-name">${escapeHtml(file.name)}</div><div class="f-meta">${formatBytes(file.size)} • ${escapeHtml(file.type || "unknown type")}</div></div>
        <button class="text-btn" id="reset-btn" type="button">Reset</button>
      </div>
      <div id="result-area"><div class="state-box"><div class="spinner"></div><p>Processing…</p></div></div>`;
    $("#reset-btn", body).addEventListener("click", () => { body.innerHTML = ""; input.value = ""; });
    const resultArea = $("#result-area", body);
    Promise.resolve().then(() => onFile(file, resultArea)).catch((err) => {
      console.error(err);
      resultArea.innerHTML = `<div class="alert alert-error">We couldn't process this document. Please check the file and try again.</div>`;
    });
  }
}

/** Mounts a paste-or-upload plain text workspace. */
function mountTextTool(el, { placeholder = "Paste or type text here…", accept = ".txt,text/plain", actions }) {
  el.innerHTML = `
    <div class="field-row">
      <label class="btn btn-secondary btn-small" for="tt-file">Load .txt file</label>
      <input type="file" id="tt-file" accept="${accept}" hidden>
      <button class="text-btn" id="tt-clear" type="button">Clear</button>
    </div>
    <textarea class="tool-textarea" id="tt-input" placeholder="${escapeHtml(placeholder)}"></textarea>
    <div id="tt-actions" class="action-row"></div>
    <div id="tt-results"></div>`;
  const ta = $("#tt-input", el), fileInput = $("#tt-file", el), actionsEl = $("#tt-actions", el), results = $("#tt-results", el);
  fileInput.addEventListener("change", async (e) => {
    const f = e.target.files[0]; if (!f) return;
    ta.value = await readFileAsText(f);
  });
  $("#tt-clear", el).addEventListener("click", () => { ta.value = ""; results.innerHTML = ""; });
  (actions || []).forEach(({ label, run, primary }) => {
    const b = document.createElement("button");
    b.type = "button"; b.className = `btn ${primary ? "btn-primary" : "btn-secondary"} btn-small`;
    b.textContent = label;
    b.addEventListener("click", () => run(ta.value, results));
    actionsEl.appendChild(b);
  });
  return { textarea: ta, results };
}

function statCards(items) {
  return `<div class="stat-grid">${items.map(([num, lbl]) => `<div class="stat-card"><span class="num">${escapeHtml(String(num))}</span><span class="lbl">${escapeHtml(lbl)}</span></div>`).join("")}</div>`;
}
function dataTable(rows, { maxRows = 200 } = {}) {
  if (!rows.length) return `<p class="muted">No rows to display.</p>`;
  const head = rows[0], body = rows.slice(1, maxRows + 1);
  const truncated = rows.length - 1 > maxRows;
  return `<div class="table-scroll"><table class="data-table"><thead><tr>${head.map(h => `<th>${escapeHtml(h)}</th>`).join("")}</tr></thead>
    <tbody>${body.map(r => `<tr>${head.map((_, i) => `<td>${escapeHtml(r[i] ?? "")}</td>`).join("")}</tr>`).join("")}</tbody></table></div>
    ${truncated ? `<p class="fine-print">Showing ${maxRows} of ${rows.length - 1} rows.</p>` : ""}`;
}

/* Minimal magic-number sniffing for common document/container formats */
function sniffSignature(bytes) {
  const b = new Uint8Array(bytes.slice(0, 8));
  const hex = Array.from(b).map(x => x.toString(16).padStart(2, "0")).join(" ");
  if (b[0] === 0x25 && b[1] === 0x50 && b[2] === 0x44 && b[3] === 0x46) return { kind: "PDF document", hex };
  if (b[0] === 0x50 && b[1] === 0x4b && b[2] === 0x03 && b[3] === 0x04) return { kind: "ZIP-based file (DOCX, XLSX, PPTX or ZIP)", hex };
  if (b[0] === 0xd0 && b[1] === 0xcf && b[2] === 0x11 && b[3] === 0xe0) return { kind: "Legacy Microsoft Office file (DOC/XLS/PPT)", hex };
  if (b[0] === 0x7b && b[1] === 0x5c && b[2] === 0x72 && b[3] === 0x74) return { kind: "RTF document", hex };
  if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return { kind: "PNG image", hex };
  if (b[0] === 0xff && b[1] === 0xd8) return { kind: "JPEG image", hex };
  return { kind: null, hex };
}
function looksLikeText(bytes) {
  const sample = new Uint8Array(bytes.slice(0, 1000));
  let printable = 0;
  for (const byte of sample) if ((byte >= 32 && byte < 127) || byte === 9 || byte === 10 || byte === 13) printable++;
  return sample.length === 0 || printable / sample.length > 0.85;
}

/* --------------------------------------------------------------------
   5. Categories + Tools registry
   -------------------------------------------------------------------- */
const CATEGORIES = {
  docx: { name: "DOCX Tools", ext: ".DOCX", desc: "Work with Microsoft Word documents." },
  txt: { name: "TXT Tools", ext: ".TXT", desc: "Analyze, clean and transform plain text." },
  rtf: { name: "RTF Tools", ext: ".RTF", desc: "View and convert rich-text documents." },
  csv: { name: "CSV Tools", ext: ".CSV", desc: "Inspect, clean, convert and analyze CSV data." },
  xlsx: { name: "XLSX Tools", ext: ".XLSX", desc: "Work with Excel spreadsheets." },
  general: { name: "General Tools", ext: "ANY", desc: "Cross-format checks that work on any document." }
};

const TOOLS = [];
function reg(t) { TOOLS.push(t); return t; }

/* ---------------- DOCX ---------------- */
reg({
  slug: "docx-viewer", name: "DOCX Viewer", category: "docx", formats: ["DOCX"],
  desc: "Preview a Word document's content directly in your browser.",
  seoDesc: "View a DOCX file online in your browser with Doxemi's free DOCX Viewer. No upload to a server required.",
  render(el) {
    mountFileTool(el, { accept: ".docx", note: "Works fully in your browser — the file is not uploaded anywhere.", onFile: async (file, area) => {
      const mammoth = await loadMammoth();
      const buf = await readFileAsArrayBuffer(file);
      const { value, messages } = await mammoth.convertToHtml({ arrayBuffer: buf });
      area.innerHTML = `${messages.length ? `<div class="alert alert-info">${messages.length} formatting note(s) were simplified during preview.</div>` : ""}<div class="docx-preview">${value}</div>`;
    }});
  }
});
reg({
  slug: "docx-to-pdf", name: "DOCX to PDF", category: "docx", formats: ["DOCX", "PDF"],
  desc: "Convert a Word document to PDF using your browser's print dialog.",
  seoDesc: "Convert DOCX to PDF online for free with Doxemi. Opens your browser's print dialog so you can save as PDF.",
  render(el) {
    mountFileTool(el, { accept: ".docx", note: "Doxemi renders the document, then opens your browser's print dialog — choose “Save as PDF” there.", onFile: async (file, area) => {
      const mammoth = await loadMammoth();
      const buf = await readFileAsArrayBuffer(file);
      const { value } = await mammoth.convertToHtml({ arrayBuffer: buf });
      area.innerHTML = `<div class="alert alert-success">Your document is ready.</div><div class="docx-preview" id="pdf-preview" style="max-height:320px">${value}</div><div class="action-row"><button class="btn btn-primary" id="print-btn" type="button">Open print dialog → Save as PDF</button></div>`;
      $("#print-btn", area).addEventListener("click", () => {
        const w = window.open("", "_blank");
        w.document.write(`<!doctype html><html><head><title>${escapeHtml(file.name)}</title><meta charset="utf-8"><style>body{font-family:Georgia,serif;max-width:800px;margin:40px auto;padding:0 20px;line-height:1.5;color:#111} img{max-width:100%}</style></head><body>${value}</body></html>`);
        w.document.close();
        w.onload = () => w.print();
      });
    }});
  }
});
reg({
  slug: "docx-to-txt", name: "DOCX to TXT", category: "docx", formats: ["DOCX", "TXT"],
  desc: "Extract the plain text content from a Word document.",
  seoDesc: "Convert DOCX to TXT online for free. Extract plain text from a Word document with Doxemi's DOCX to TXT tool.",
  render(el) {
    mountFileTool(el, { accept: ".docx", onFile: async (file, area) => {
      const mammoth = await loadMammoth();
      const buf = await readFileAsArrayBuffer(file);
      const { value } = await mammoth.extractRawText({ arrayBuffer: buf });
      area.innerHTML = `<div class="alert alert-success">Your document is ready.</div>${statCards([[countWords(value), "words"], [countChars(value), "characters"]])}
        <textarea class="tool-textarea" readonly>${escapeHtml(value)}</textarea>
        <div class="action-row"><button class="btn btn-primary" id="dl">Download .txt</button></div>`;
      $("#dl", area).addEventListener("click", () => downloadText(file.name.replace(/\.docx$/i, "") + ".txt", value));
    }});
  }
});
reg({
  slug: "docx-to-html", name: "DOCX to HTML", category: "docx", formats: ["DOCX", "HTML"],
  desc: "Convert a Word document into clean HTML markup.",
  seoDesc: "Convert DOCX to HTML online for free with Doxemi. Get clean, downloadable HTML from a Word document.",
  render(el) {
    mountFileTool(el, { accept: ".docx", onFile: async (file, area) => {
      const mammoth = await loadMammoth();
      const buf = await readFileAsArrayBuffer(file);
      const { value } = await mammoth.convertToHtml({ arrayBuffer: buf });
      area.innerHTML = `<div class="alert alert-success">Your document is ready.</div>
        <textarea class="tool-textarea" readonly>${escapeHtml(value)}</textarea>
        <div class="action-row"><button class="btn btn-primary" id="dl">Download .html</button></div>
        <h3>Preview</h3><div class="docx-preview">${value}</div>`;
      $("#dl", area).addEventListener("click", () => downloadText(file.name.replace(/\.docx$/i, "") + ".html", value, "text/html"));
    }});
  }
});
reg({
  slug: "docx-word-counter", name: "DOCX Word Counter", category: "docx", formats: ["DOCX"],
  desc: "Count words, characters, paragraphs and lines in a DOCX document.",
  seoDesc: "Count words, characters and paragraphs in DOCX files online with Doxemi's free DOCX Word Counter.",
  render(el) {
    mountFileTool(el, { accept: ".docx", onFile: async (file, area) => {
      const mammoth = await loadMammoth();
      const buf = await readFileAsArrayBuffer(file);
      const { value } = await mammoth.extractRawText({ arrayBuffer: buf });
      area.innerHTML = statCards([[countWords(value), "words"], [countChars(value), "characters"], [countParagraphs(value), "paragraphs"], [countLines(value), "lines"]]);
    }});
  }
});
reg({
  slug: "docx-metadata-viewer", name: "DOCX Metadata Viewer", category: "docx", formats: ["DOCX"],
  desc: "Inspect the document properties stored inside a DOCX file.",
  seoDesc: "View DOCX metadata online free — title, author and dates stored inside a Word document, with Doxemi.",
  render(el) {
    mountFileTool(el, { accept: ".docx", onFile: async (file, area) => {
      const meta = await readDocxMetadata(file);
      area.innerHTML = `<div class="table-scroll"><table class="data-table"><tbody>${
        Object.entries(meta).map(([k, v]) => `<tr><th>${escapeHtml(k)}</th><td>${escapeHtml(v || "—")}</td></tr>`).join("")
      }</tbody></table></div>`;
    }});
  }
});

async function readDocxMetadata(file) {
  const JSZip = await loadJSZip();
  const buf = await readFileAsArrayBuffer(file);
  const zip = await JSZip.loadAsync(buf);
  const meta = { "File name": file.name, "File size": formatBytes(file.size) };
  const parse = async (path, map) => {
    const entry = zip.file(path);
    if (!entry) return;
    const xml = await entry.async("string");
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    Object.entries(map).forEach(([tag, label]) => {
      const node = doc.getElementsByTagName(tag)[0];
      if (node && node.textContent) meta[label] = node.textContent;
    });
  };
  await parse("docProps/core.xml", { "dc:title": "Title", "dc:creator": "Author", "cp:lastModifiedBy": "Last modified by", "dcterms:created": "Created", "dcterms:modified": "Modified", "cp:keywords": "Keywords", "dc:subject": "Subject" });
  await parse("docProps/app.xml", { "Pages": "Pages", "Words": "Words", "Characters": "Characters", "Lines": "Lines", "Paragraphs": "Paragraphs", "Application": "Application" });
  return meta;
}

/* ---------------- TXT ---------------- */
reg({
  slug: "txt-viewer", name: "TXT Viewer", category: "txt", formats: ["TXT"],
  desc: "Open and read a plain text file in your browser.",
  seoDesc: "View TXT files online for free with Doxemi's TXT Viewer. Opens plain text files directly in your browser.",
  render(el) {
    mountFileTool(el, { accept: ".txt,text/plain", onFile: async (file, area) => {
      const text = await readFileAsText(file);
      area.innerHTML = `${statCards([[countWords(text), "words"], [countLines(text), "lines"]])}<textarea class="tool-textarea" readonly>${escapeHtml(text)}</textarea>`;
    }});
  }
});
reg({
  slug: "txt-to-pdf", name: "TXT to PDF", category: "txt", formats: ["TXT", "PDF"],
  desc: "Convert a plain text file to PDF using your browser's print dialog.",
  seoDesc: "Convert TXT to PDF online for free with Doxemi. Opens your browser's print dialog so you can save as PDF.",
  render(el) {
    mountFileTool(el, { accept: ".txt,text/plain", onFile: async (file, area) => {
      const text = await readFileAsText(file);
      area.innerHTML = `<div class="alert alert-success">Your document is ready.</div><div class="action-row"><button class="btn btn-primary" id="print-btn" type="button">Open print dialog → Save as PDF</button></div>`;
      $("#print-btn", area).addEventListener("click", () => {
        const w = window.open("", "_blank");
        w.document.write(`<!doctype html><html><head><title>${escapeHtml(file.name)}</title><meta charset="utf-8"><style>body{font-family:'Courier New',monospace;white-space:pre-wrap;max-width:800px;margin:40px auto;padding:0 20px;line-height:1.5;color:#111}</style></head><body>${escapeHtml(text)}</body></html>`);
        w.document.close();
        w.onload = () => w.print();
      });
    }});
  }
});
reg({
  slug: "txt-word-counter", name: "TXT Word Counter", category: "txt", formats: ["TXT"],
  desc: "Count words in pasted text or an uploaded .txt file.",
  seoDesc: "Free online word counter for TXT files and pasted text. Count words instantly with Doxemi.",
  render(el) {
    const { results } = mountTextTool(el, { actions: [{ label: "Count words", primary: true, run: (text, out) => {
      out.innerHTML = statCards([[countWords(text), "words"], [countParagraphs(text), "paragraphs"]]);
    }}]});
  }
});
reg({
  slug: "txt-character-counter", name: "TXT Character Counter", category: "txt", formats: ["TXT"],
  desc: "Count characters, with and without spaces, in text or a .txt file.",
  seoDesc: "Free online character counter for TXT files and pasted text, with and without spaces, from Doxemi.",
  render(el) {
    mountTextTool(el, { actions: [{ label: "Count characters", primary: true, run: (text, out) => {
      out.innerHTML = statCards([[countChars(text), "characters"], [countCharsNoSpaces(text), "characters (no spaces)"]]);
    }}]});
  }
});
reg({
  slug: "txt-line-counter", name: "TXT Line Counter", category: "txt", formats: ["TXT"],
  desc: "Count the number of lines in text or a .txt file.",
  seoDesc: "Free online line counter for TXT files and pasted text with Doxemi.",
  render(el) {
    mountTextTool(el, { actions: [{ label: "Count lines", primary: true, run: (text, out) => {
      const nonEmpty = text.split(/\r\n|\r|\n/).filter(l => l.trim() !== "").length;
      out.innerHTML = statCards([[countLines(text), "total lines"], [nonEmpty, "non-empty lines"]]);
    }}]});
  }
});
reg({
  slug: "txt-case-converter", name: "TXT Case Converter", category: "txt", formats: ["TXT"],
  desc: "Change text to UPPERCASE, lowercase, Title Case or Sentence case.",
  seoDesc: "Free online case converter — UPPERCASE, lowercase, Title Case and Sentence case — with Doxemi.",
  render(el) {
    const toTitle = (t) => t.toLowerCase().replace(/(^|\s|["'(\-])([a-z])/g, (m, p1, p2) => p1 + p2.toUpperCase());
    const toSentence = (t) => t.toLowerCase().replace(/(^\s*[a-z])|([.!?]\s*[a-z])/g, (m) => m.toUpperCase());
    const { results } = mountTextTool(el, { actions: [
      { label: "UPPERCASE", run: (t, out) => showCaseResult(out, t.toUpperCase()) },
      { label: "lowercase", run: (t, out) => showCaseResult(out, t.toLowerCase()) },
      { label: "Title Case", run: (t, out) => showCaseResult(out, toTitle(t)) },
      { label: "Sentence case", primary: true, run: (t, out) => showCaseResult(out, toSentence(t)) }
    ]});
    function showCaseResult(out, converted) {
      out.innerHTML = `<textarea class="tool-textarea" readonly>${escapeHtml(converted)}</textarea><div class="action-row"><button class="btn btn-primary" id="dl">Download .txt</button></div>`;
      $("#dl", out).addEventListener("click", () => downloadText("converted.txt", converted));
    }
  }
});
reg({
  slug: "txt-cleaner", name: "TXT Cleaner", category: "txt", formats: ["TXT"],
  desc: "Remove extra blank lines, trailing spaces and repeated whitespace.",
  seoDesc: "Clean up messy text files online for free — trim whitespace and extra blank lines with Doxemi's TXT Cleaner.",
  render(el) {
    mountTextTool(el, { actions: [{ label: "Clean text", primary: true, run: (t, out) => {
      const cleaned = t.split(/\r\n|\r|\n/).map(l => l.replace(/[ \t]+$/g, "").replace(/[ \t]{2,}/g, " ")).join("\n").replace(/\n{3,}/g, "\n\n").trim();
      out.innerHTML = `<textarea class="tool-textarea" readonly>${escapeHtml(cleaned)}</textarea><div class="action-row"><button class="btn btn-primary" id="dl">Download .txt</button></div>`;
      $("#dl", out).addEventListener("click", () => downloadText("cleaned.txt", cleaned));
    }}]});
  }
});
reg({
  slug: "txt-formatter", name: "TXT Formatter", category: "txt", formats: ["TXT"],
  desc: "Wrap long lines to a set width or collapse text into single paragraphs.",
  seoDesc: "Reformat plain text online for free — wrap lines to a width or collapse paragraphs — with Doxemi.",
  render(el) {
    el.innerHTML = `<div class="field-row"><label for="wrap-width">Wrap width</label><input type="text" id="wrap-width" value="80" style="width:70px" inputmode="numeric"></div>`;
    const holder = document.createElement("div");
    el.appendChild(holder);
    mountTextTool(holder, { actions: [
      { label: "Wrap lines", primary: true, run: (t, out) => {
        const width = Math.max(20, parseInt($("#wrap-width", el).value, 10) || 80);
        const wrapped = t.split(/\r?\n/).map(line => wrapLine(line, width)).join("\n");
        finish(out, wrapped);
      }},
      { label: "Collapse to one paragraph", run: (t, out) => finish(out, t.replace(/\s+/g, " ").trim()) }
    ]});
    function wrapLine(line, width) {
      if (line.length <= width) return line;
      const words = line.split(" "); let out = "", cur = "";
      words.forEach(w => { if ((cur + " " + w).trim().length > width) { out += cur.trim() + "\n"; cur = w; } else cur += " " + w; });
      return (out + cur).trim();
    }
    function finish(out, text) {
      out.innerHTML = `<textarea class="tool-textarea" readonly>${escapeHtml(text)}</textarea><div class="action-row"><button class="btn btn-primary" id="dl">Download .txt</button></div>`;
      $("#dl", out).addEventListener("click", () => downloadText("formatted.txt", text));
    }
  }
});

/* ---------------- RTF ---------------- */
reg({
  slug: "rtf-viewer", name: "RTF Viewer", category: "rtf", formats: ["RTF"],
  desc: "Read the text content of a rich text (.rtf) file.",
  seoDesc: "View RTF files online for free with Doxemi's RTF Viewer. Reads the text content of rich text files in your browser.",
  render(el) {
    mountFileTool(el, { accept: ".rtf", note: "Shows the document's text content. Formatting such as fonts and colors is not rendered.", onFile: async (file, area) => {
      const raw = await readFileAsText(file);
      const text = rtfToText(raw);
      area.innerHTML = `${statCards([[countWords(text), "words"], [countLines(text), "lines"]])}<textarea class="tool-textarea" readonly>${escapeHtml(text)}</textarea>`;
    }});
  }
});
reg({
  slug: "rtf-to-txt", name: "RTF to TXT", category: "rtf", formats: ["RTF", "TXT"],
  desc: "Extract plain text from a rich text (.rtf) file.",
  seoDesc: "Convert RTF to TXT online for free with Doxemi. Extract the plain text content from a rich text file.",
  render(el) {
    mountFileTool(el, { accept: ".rtf", onFile: async (file, area) => {
      const raw = await readFileAsText(file);
      const text = rtfToText(raw);
      area.innerHTML = `<div class="alert alert-success">Your document is ready.</div><textarea class="tool-textarea" readonly>${escapeHtml(text)}</textarea><div class="action-row"><button class="btn btn-primary" id="dl">Download .txt</button></div>`;
      $("#dl", area).addEventListener("click", () => downloadText(file.name.replace(/\.rtf$/i, "") + ".txt", text));
    }});
  }
});
reg({
  slug: "rtf-to-html", name: "RTF to HTML", category: "rtf", formats: ["RTF", "HTML"],
  desc: "Convert a rich text file's text content into basic HTML paragraphs.",
  seoDesc: "Convert RTF to HTML online for free with Doxemi — basic, text-only HTML conversion for rich text files.",
  render(el) {
    mountFileTool(el, { accept: ".rtf", note: "Basic conversion: preserves text and paragraphs only, not fonts, colors or images.", onFile: async (file, area) => {
      const raw = await readFileAsText(file);
      const text = rtfToText(raw);
      const html = text.split(/\n{2,}/).map(p => `<p>${escapeHtml(p).replace(/\n/g, "<br>")}</p>`).join("\n");
      area.innerHTML = `<div class="alert alert-success">Your document is ready.</div><textarea class="tool-textarea" readonly>${escapeHtml(html)}</textarea><div class="action-row"><button class="btn btn-primary" id="dl">Download .html</button></div><h3>Preview</h3><div class="docx-preview">${html}</div>`;
      $("#dl", area).addEventListener("click", () => downloadText(file.name.replace(/\.rtf$/i, "") + ".html", html, "text/html"));
    }});
  }
});
reg({
  slug: "rtf-word-counter", name: "RTF Word Counter", category: "rtf", formats: ["RTF"],
  desc: "Count words, characters and lines in a rich text (.rtf) file.",
  seoDesc: "Count words and characters in RTF files online for free with Doxemi's RTF Word Counter.",
  render(el) {
    mountFileTool(el, { accept: ".rtf", onFile: async (file, area) => {
      const text = rtfToText(await readFileAsText(file));
      area.innerHTML = statCards([[countWords(text), "words"], [countChars(text), "characters"], [countLines(text), "lines"]]);
    }});
  }
});
reg({
  slug: "rtf-cleaner", name: "RTF Cleaner", category: "rtf", formats: ["RTF", "TXT"],
  desc: "Strip formatting codes from an RTF file and download clean plain text.",
  seoDesc: "Clean RTF files online for free — strip formatting codes and get plain text with Doxemi's RTF Cleaner.",
  render(el) {
    mountFileTool(el, { accept: ".rtf", onFile: async (file, area) => {
      const raw = await readFileAsText(file);
      const cleaned = rtfToText(raw).replace(/\n{3,}/g, "\n\n").trim();
      area.innerHTML = `<div class="alert alert-success">Your document is ready.</div><textarea class="tool-textarea" readonly>${escapeHtml(cleaned)}</textarea><div class="action-row"><button class="btn btn-primary" id="dl">Download .txt</button></div>`;
      $("#dl", area).addEventListener("click", () => downloadText(file.name.replace(/\.rtf$/i, "") + "-clean.txt", cleaned));
    }});
  }
});

/* ---------------- CSV ---------------- */
reg({
  slug: "csv-viewer", name: "CSV Viewer", category: "csv", formats: ["CSV"],
  desc: "Open a CSV file and view it as a table.",
  seoDesc: "View CSV files online for free as a table with Doxemi's CSV Viewer.",
  render(el) {
    mountFileTool(el, { accept: ".csv,text/csv", onFile: async (file, area) => {
      const rows = parseCSV(await readFileAsText(file));
      area.innerHTML = statCards([[rows.length - 1 >= 0 ? rows.length - 1 : 0, "rows"], [(rows[0] || []).length, "columns"]]) + dataTable(rows);
    }});
  }
});
reg({
  slug: "csv-to-json", name: "CSV to JSON", category: "csv", formats: ["CSV", "JSON"],
  desc: "Convert CSV data into a JSON array of objects.",
  seoDesc: "Convert CSV to JSON online for free with Doxemi. Uses the first row as JSON object keys.",
  render(el) {
    mountFileTool(el, { accept: ".csv,text/csv", onFile: async (file, area) => {
      const rows = parseCSV(await readFileAsText(file));
      const [head, ...body] = rows;
      const json = body.map(r => Object.fromEntries(head.map((h, i) => [h, r[i] ?? ""])));
      const text = JSON.stringify(json, null, 2);
      area.innerHTML = `<div class="alert alert-success">Your document is ready.</div>${statCards([[json.length, "records"]])}<textarea class="tool-textarea" readonly>${escapeHtml(text)}</textarea><div class="action-row"><button class="btn btn-primary" id="dl">Download .json</button></div>`;
      $("#dl", area).addEventListener("click", () => downloadText(file.name.replace(/\.csv$/i, "") + ".json", text, "application/json"));
    }});
  }
});
reg({
  slug: "json-to-csv", name: "JSON to CSV", category: "csv", formats: ["JSON", "CSV"],
  desc: "Convert a JSON array of objects into CSV rows.",
  seoDesc: "Convert JSON to CSV online for free with Doxemi. Turns a JSON array of objects into a downloadable CSV file.",
  render(el) {
    mountFileTool(el, { accept: ".json,application/json", onFile: async (file, area) => {
      const text = await readFileAsText(file);
      let json;
      try { json = JSON.parse(text); } catch { throw new Error("invalid json"); }
      const arr = Array.isArray(json) ? json : [json];
      const keys = Array.from(arr.reduce((set, o) => { Object.keys(o || {}).forEach(k => set.add(k)); return set; }, new Set()));
      const rows = [keys, ...arr.map(o => keys.map(k => o?.[k] ?? ""))];
      const csv = toCSV(rows);
      area.innerHTML = `<div class="alert alert-success">Your document is ready.</div>${statCards([[arr.length, "rows"], [keys.length, "columns"]])}<textarea class="tool-textarea" readonly>${escapeHtml(csv)}</textarea><div class="action-row"><button class="btn btn-primary" id="dl">Download .csv</button></div>`;
      $("#dl", area).addEventListener("click", () => downloadText(file.name.replace(/\.json$/i, "") + ".csv", csv, "text/csv"));
    }});
  }
});
reg({
  slug: "csv-to-txt", name: "CSV to TXT", category: "csv", formats: ["CSV", "TXT"],
  desc: "Convert CSV rows into plain, tab-separated text.",
  seoDesc: "Convert CSV to TXT online for free with Doxemi — tab-separated plain text output.",
  render(el) {
    mountFileTool(el, { accept: ".csv,text/csv", onFile: async (file, area) => {
      const rows = parseCSV(await readFileAsText(file));
      const text = rows.map(r => r.join("\t")).join("\n");
      area.innerHTML = `<div class="alert alert-success">Your document is ready.</div><textarea class="tool-textarea" readonly>${escapeHtml(text)}</textarea><div class="action-row"><button class="btn btn-primary" id="dl">Download .txt</button></div>`;
      $("#dl", area).addEventListener("click", () => downloadText(file.name.replace(/\.csv$/i, "") + ".txt", text));
    }});
  }
});
reg({
  slug: "csv-column-extractor", name: "CSV Column Extractor", category: "csv", formats: ["CSV"],
  desc: "Pick specific columns from a CSV file and export just those.",
  seoDesc: "Extract specific columns from a CSV file online for free with Doxemi's CSV Column Extractor.",
  render(el) {
    mountFileTool(el, { accept: ".csv,text/csv", onFile: async (file, area) => {
      const rows = parseCSV(await readFileAsText(file));
      const head = rows[0] || [];
      area.innerHTML = `<div class="field-row">${head.map((h, i) => `<label><input type="checkbox" class="col-check" value="${i}" checked> ${escapeHtml(h || "Column " + (i + 1))}</label>`).join("")}</div>
        <div class="action-row"><button class="btn btn-primary" id="extract">Extract selected columns</button></div><div id="extract-out"></div>`;
      $("#extract", area).addEventListener("click", () => {
        const idx = $$(".col-check", area).filter(c => c.checked).map(c => parseInt(c.value, 10));
        const out = rows.map(r => idx.map(i => r[i] ?? ""));
        const csv = toCSV(out);
        const target = $("#extract-out", area);
        target.innerHTML = `<textarea class="tool-textarea" readonly>${escapeHtml(csv)}</textarea><div class="action-row"><button class="btn btn-primary" id="dl">Download .csv</button></div>`;
        $("#dl", target).addEventListener("click", () => downloadText(file.name.replace(/\.csv$/i, "") + "-columns.csv", csv, "text/csv"));
      });
    }});
  }
});
reg({
  slug: "csv-row-counter", name: "CSV Row Counter", category: "csv", formats: ["CSV"],
  desc: "Count the number of data rows and columns in a CSV file.",
  seoDesc: "Count rows and columns in a CSV file online for free with Doxemi's CSV Row Counter.",
  render(el) {
    mountFileTool(el, { accept: ".csv,text/csv", onFile: async (file, area) => {
      const rows = parseCSV(await readFileAsText(file));
      area.innerHTML = statCards([[Math.max(0, rows.length - 1), "data rows"], [(rows[0] || []).length, "columns"], [rows.length, "total rows incl. header"]]);
    }});
  }
});
reg({
  slug: "csv-cleaner", name: "CSV Cleaner", category: "csv", formats: ["CSV"],
  desc: "Remove blank rows, duplicate rows and stray whitespace from a CSV file.",
  seoDesc: "Clean up a CSV file online for free — remove blank and duplicate rows with Doxemi's CSV Cleaner.",
  render(el) {
    mountFileTool(el, { accept: ".csv,text/csv", onFile: async (file, area) => {
      const rows = parseCSV(await readFileAsText(file)).map(r => r.map(c => c.trim()));
      const nonBlank = rows.filter(r => r.some(c => c !== ""));
      const seen = new Set(); const deduped = [];
      nonBlank.forEach(r => { const key = r.join("\u0001"); if (!seen.has(key)) { seen.add(key); deduped.push(r); } });
      const csv = toCSV(deduped);
      area.innerHTML = `<div class="alert alert-success">Your document is ready.</div>${statCards([[rows.length - deduped.length, "rows removed"], [deduped.length, "rows kept"]])}<textarea class="tool-textarea" readonly>${escapeHtml(csv)}</textarea><div class="action-row"><button class="btn btn-primary" id="dl">Download .csv</button></div>`;
      $("#dl", area).addEventListener("click", () => downloadText(file.name.replace(/\.csv$/i, "") + "-cleaned.csv", csv, "text/csv"));
    }});
  }
});
reg({
  slug: "csv-formatter", name: "CSV Formatter", category: "csv", formats: ["CSV"],
  desc: "Re-serialize a CSV file with a chosen delimiter and consistent quoting.",
  seoDesc: "Reformat a CSV file online for free — change the delimiter or quote every field with Doxemi's CSV Formatter.",
  render(el) {
    mountFileTool(el, { accept: ".csv,text/csv", onFile: async (file, area) => {
      const rows = parseCSV(await readFileAsText(file));
      area.innerHTML = `<div class="field-row">
          <label>Delimiter <select id="delim"><option value=",">Comma (,)</option><option value=";">Semicolon (;)</option><option value="\t">Tab</option></select></label>
          <label><input type="checkbox" id="force-quote"> Quote every field</label>
        </div><div class="action-row"><button class="btn btn-primary" id="format-btn">Format</button></div><div id="fmt-out"></div>`;
      $("#format-btn", area).addEventListener("click", () => {
        const delim = $("#delim", area).value;
        const csv = toCSV(rows, delim, $("#force-quote", area).checked);
        const target = $("#fmt-out", area);
        target.innerHTML = `<textarea class="tool-textarea" readonly>${escapeHtml(csv)}</textarea><div class="action-row"><button class="btn btn-primary" id="dl">Download .csv</button></div>`;
        $("#dl", target).addEventListener("click", () => downloadText(file.name.replace(/\.csv$/i, "") + "-formatted.csv", csv, "text/csv"));
      });
    }});
  }
});

/* ---------------- XLSX ---------------- */
reg({
  slug: "xlsx-viewer", name: "XLSX Viewer", category: "xlsx", formats: ["XLSX"],
  desc: "Open an Excel workbook and browse its sheets as tables.",
  seoDesc: "View XLSX files online for free with Doxemi's XLSX Viewer. Browse Excel sheets as tables in your browser.",
  render(el) {
    mountFileTool(el, { accept: ".xlsx", onFile: async (file, area) => {
      const XLSX = await loadXLSX();
      const wb = XLSX.read(await readFileAsArrayBuffer(file), { type: "array" });
      area.innerHTML = `<div class="field-row"><label for="sheet-select">Sheet</label><select id="sheet-select">${wb.SheetNames.map(n => `<option value="${escapeHtml(n)}">${escapeHtml(n)}</option>`).join("")}</select></div><div id="sheet-out"></div>`;
      const renderSheet = (name) => {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: "" });
        $("#sheet-out", area).innerHTML = statCards([[Math.max(0, rows.length - 1), "rows"], [(rows[0] || []).length, "columns"]]) + dataTable(rows);
      };
      $("#sheet-select", area).addEventListener("change", (e) => renderSheet(e.target.value));
      renderSheet(wb.SheetNames[0]);
    }});
  }
});
reg({
  slug: "xlsx-to-csv", name: "XLSX to CSV", category: "xlsx", formats: ["XLSX", "CSV"],
  desc: "Convert one sheet of an Excel workbook to a CSV file.",
  seoDesc: "Convert XLSX to CSV online for free with Doxemi. Choose a sheet and download it as CSV.",
  render(el) {
    mountFileTool(el, { accept: ".xlsx", onFile: async (file, area) => {
      const XLSX = await loadXLSX();
      const wb = XLSX.read(await readFileAsArrayBuffer(file), { type: "array" });
      area.innerHTML = `<div class="field-row"><label for="sheet-select">Sheet</label><select id="sheet-select">${wb.SheetNames.map(n => `<option value="${escapeHtml(n)}">${escapeHtml(n)}</option>`).join("")}</select>
        <button class="btn btn-primary btn-small" id="convert-btn">Convert &amp; download</button></div><div id="csv-out"></div>`;
      $("#convert-btn", area).addEventListener("click", () => {
        const name = $("#sheet-select", area).value;
        const csv = XLSX.utils.sheet_to_csv(wb.Sheets[name]);
        $("#csv-out", area).innerHTML = `<div class="alert alert-success">Your document is ready.</div><textarea class="tool-textarea" readonly>${escapeHtml(csv)}</textarea>`;
        downloadText(`${file.name.replace(/\.xlsx$/i, "")}-${name}.csv`, csv, "text/csv");
      });
    }});
  }
});
reg({
  slug: "csv-to-xlsx", name: "CSV to XLSX", category: "xlsx", formats: ["CSV", "XLSX"],
  desc: "Convert a CSV file into a downloadable Excel workbook.",
  seoDesc: "Convert CSV to XLSX online for free with Doxemi. Turns a CSV file into a real Excel workbook.",
  render(el) {
    mountFileTool(el, { accept: ".csv,text/csv", onFile: async (file, area) => {
      const XLSX = await loadXLSX();
      const rows = parseCSV(await readFileAsText(file));
      const ws = XLSX.utils.aoa_to_sheet(rows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
      const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      area.innerHTML = `<div class="alert alert-success">Your document is ready.</div>${statCards([[Math.max(0, rows.length - 1), "rows"], [(rows[0] || []).length, "columns"]])}<div class="action-row"><button class="btn btn-primary" id="dl">Download .xlsx</button></div>`;
      $("#dl", area).addEventListener("click", () => downloadBlob(file.name.replace(/\.csv$/i, "") + ".xlsx", new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" })));
    }});
  }
});
reg({
  slug: "xlsx-sheet-extractor", name: "XLSX Sheet Extractor", category: "xlsx", formats: ["XLSX"],
  desc: "Pull a single sheet out of a workbook and save it as its own XLSX file.",
  seoDesc: "Extract a single sheet from an Excel workbook online for free with Doxemi's XLSX Sheet Extractor.",
  render(el) {
    mountFileTool(el, { accept: ".xlsx", onFile: async (file, area) => {
      const XLSX = await loadXLSX();
      const wb = XLSX.read(await readFileAsArrayBuffer(file), { type: "array" });
      area.innerHTML = `<div class="field-row"><label for="sheet-select">Sheet to extract</label><select id="sheet-select">${wb.SheetNames.map(n => `<option value="${escapeHtml(n)}">${escapeHtml(n)}</option>`).join("")}</select>
        <button class="btn btn-primary btn-small" id="extract-btn">Extract &amp; download</button></div>`;
      $("#extract-btn", area).addEventListener("click", () => {
        const name = $("#sheet-select", area).value;
        const newWb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWb, wb.Sheets[name], name);
        const out = XLSX.write(newWb, { bookType: "xlsx", type: "array" });
        downloadBlob(`${file.name.replace(/\.xlsx$/i, "")}-${name}.xlsx`, new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
      });
    }});
  }
});
reg({
  slug: "xlsx-metadata-viewer", name: "XLSX Metadata Viewer", category: "xlsx", formats: ["XLSX"],
  desc: "View sheet names, dimensions and document properties in an Excel workbook.",
  seoDesc: "View XLSX metadata online for free — sheet names, size and document properties — with Doxemi.",
  render(el) {
    mountFileTool(el, { accept: ".xlsx", onFile: async (file, area) => {
      const XLSX = await loadXLSX();
      const wb = XLSX.read(await readFileAsArrayBuffer(file), { type: "array", bookProps: true });
      const props = wb.Props || {};
      const sheetRows = wb.SheetNames.map(n => {
        const ref = wb.Sheets[n]["!ref"];
        return `<tr><td>${escapeHtml(n)}</td><td>${escapeHtml(ref || "empty")}</td></tr>`;
      }).join("");
      area.innerHTML = `<div class="table-scroll"><table class="data-table"><tbody>
          <tr><th>File name</th><td>${escapeHtml(file.name)}</td></tr>
          <tr><th>File size</th><td>${formatBytes(file.size)}</td></tr>
          <tr><th>Sheets</th><td>${wb.SheetNames.length}</td></tr>
          <tr><th>Title</th><td>${escapeHtml(props.Title || "—")}</td></tr>
          <tr><th>Author</th><td>${escapeHtml(props.Author || "—")}</td></tr>
          <tr><th>Created</th><td>${escapeHtml(props.CreatedDate ? new Date(props.CreatedDate).toString() : "—")}</td></tr>
        </tbody></table></div>
        <h3>Sheets</h3>
        <div class="table-scroll"><table class="data-table"><thead><tr><th>Name</th><th>Used range</th></tr></thead><tbody>${sheetRows}</tbody></table></div>`;
    }});
  }
});
reg({
  slug: "xlsx-data-cleaner", name: "XLSX Data Cleaner", category: "xlsx", formats: ["XLSX"],
  desc: "Remove empty rows and columns from a sheet, then download the result.",
  seoDesc: "Clean an Excel workbook online for free — remove empty rows and columns with Doxemi's XLSX Data Cleaner.",
  render(el) {
    mountFileTool(el, { accept: ".xlsx", onFile: async (file, area) => {
      const XLSX = await loadXLSX();
      const wb = XLSX.read(await readFileAsArrayBuffer(file), { type: "array" });
      area.innerHTML = `<div class="field-row"><label for="sheet-select">Sheet</label><select id="sheet-select">${wb.SheetNames.map(n => `<option value="${escapeHtml(n)}">${escapeHtml(n)}</option>`).join("")}</select>
        <button class="btn btn-primary btn-small" id="clean-btn">Clean &amp; download</button></div><div id="clean-out"></div>`;
      $("#clean-btn", area).addEventListener("click", () => {
        const name = $("#sheet-select", area).value;
        let rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: "" });
        rows = rows.filter(r => r.some(c => String(c).trim() !== ""));
        const colHasData = [];
        rows.forEach(r => r.forEach((c, i) => { if (String(c).trim() !== "") colHasData[i] = true; }));
        rows = rows.map(r => r.filter((_, i) => colHasData[i]));
        const ws = XLSX.utils.aoa_to_sheet(rows);
        const newWb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWb, ws, name);
        const out = XLSX.write(newWb, { bookType: "xlsx", type: "array" });
        $("#clean-out", area).innerHTML = statCards([[rows.length, "rows kept"], [colHasData.filter(Boolean).length, "columns kept"]]);
        downloadBlob(`${file.name.replace(/\.xlsx$/i, "")}-cleaned.xlsx`, new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
      });
    }});
  }
});

/* ---------------- General ---------------- */
async function extractTextGeneric(file) {
  const ext = extOf(file.name);
  if (ext === "docx") {
    const mammoth = await loadMammoth();
    const { value } = await mammoth.extractRawText({ arrayBuffer: await readFileAsArrayBuffer(file) });
    return { text: value, supported: true };
  }
  if (ext === "txt" || ext === "csv" || (file.type && file.type.startsWith("text/"))) {
    return { text: await readFileAsText(file), supported: true };
  }
  if (ext === "rtf") return { text: rtfToText(await readFileAsText(file)), supported: true };
  return { text: "", supported: false };
}
reg({
  slug: "document-word-counter", name: "Document Word Counter", category: "general", formats: ["DOCX", "TXT", "RTF", "CSV"],
  desc: "Count words and paragraphs in a DOCX, TXT, RTF or CSV file.",
  seoDesc: "Free online word counter that works across DOCX, TXT, RTF and CSV files — from Doxemi.",
  render(el) {
    mountFileTool(el, { onFile: async (file, area) => {
      const { text, supported } = await extractTextGeneric(file);
      if (!supported) { area.innerHTML = `<div class="alert alert-error">This document format isn't supported by this tool.</div>`; return; }
      area.innerHTML = statCards([[countWords(text), "words"], [countParagraphs(text), "paragraphs"], [countChars(text), "characters"]]);
    }});
  }
});
reg({
  slug: "character-counter", name: "Character Counter", category: "general", formats: ["DOCX", "TXT", "RTF", "CSV"],
  desc: "Count characters in a DOCX, TXT, RTF or CSV file.",
  seoDesc: "Free online character counter for DOCX, TXT, RTF and CSV files, with and without spaces — Doxemi.",
  render(el) {
    mountFileTool(el, { onFile: async (file, area) => {
      const { text, supported } = await extractTextGeneric(file);
      if (!supported) { area.innerHTML = `<div class="alert alert-error">This document format isn't supported by this tool.</div>`; return; }
      area.innerHTML = statCards([[countChars(text), "characters"], [countCharsNoSpaces(text), "characters (no spaces)"]]);
    }});
  }
});
reg({
  slug: "line-counter", name: "Line Counter", category: "general", formats: ["DOCX", "TXT", "RTF", "CSV"],
  desc: "Count lines in a DOCX, TXT, RTF or CSV file.",
  seoDesc: "Free online line counter for DOCX, TXT, RTF and CSV files — Doxemi.",
  render(el) {
    mountFileTool(el, { onFile: async (file, area) => {
      const { text, supported } = await extractTextGeneric(file);
      if (!supported) { area.innerHTML = `<div class="alert alert-error">This document format isn't supported by this tool.</div>`; return; }
      area.innerHTML = statCards([[countLines(text), "lines"]]);
    }});
  }
});
reg({
  slug: "file-size-checker", name: "File Size Checker", category: "general", formats: ["ANY"],
  desc: "Check any file's exact size, in bytes and human-readable units.",
  seoDesc: "Check a file's exact size online for free with Doxemi's File Size Checker — any file type.",
  render(el) {
    mountFileTool(el, { onFile: async (file, area) => {
      area.innerHTML = statCards([[file.size.toLocaleString(), "bytes"], [formatBytes(file.size), "size"]]) +
        `<p class="fine-print">For reference: many email providers cap attachments around 25 MB, and WhatsApp allows files up to 100 MB.</p>`;
    }});
  }
});
reg({
  slug: "file-type-checker", name: "File Type Checker", category: "general", formats: ["ANY"],
  desc: "Check a file's reported MIME type, extension, and its actual file signature.",
  seoDesc: "Check a file's real type online for free with Doxemi's File Type Checker — MIME type, extension and file signature.",
  render(el) {
    mountFileTool(el, { onFile: async (file, area) => {
      const buf = await readFileAsArrayBuffer(file);
      const sig = sniffSignature(buf);
      const asText = sig.kind ? false : looksLikeText(buf);
      area.innerHTML = `<div class="table-scroll"><table class="data-table"><tbody>
        <tr><th>File name</th><td>${escapeHtml(file.name)}</td></tr>
        <tr><th>Extension</th><td>.${escapeHtml(extOf(file.name) || "none")}</td></tr>
        <tr><th>Reported MIME type</th><td>${escapeHtml(file.type || "not reported by browser")}</td></tr>
        <tr><th>Detected from content</th><td>${escapeHtml(sig.kind || (asText ? "Plain text-like content" : "Unrecognized binary content"))}</td></tr>
        <tr><th>First bytes (hex)</th><td>${escapeHtml(sig.hex)}</td></tr>
      </tbody></table></div>`;
    }});
  }
});
reg({
  slug: "mime-type-checker", name: "MIME Type Checker", category: "general", formats: ["ANY"],
  desc: "See the MIME type your browser reports for any file.",
  seoDesc: "Check a file's MIME type online for free with Doxemi's MIME Type Checker.",
  render(el) {
    mountFileTool(el, { onFile: async (file, area) => {
      area.innerHTML = `<div class="table-scroll"><table class="data-table"><tbody>
        <tr><th>File name</th><td>${escapeHtml(file.name)}</td></tr>
        <tr><th>MIME type</th><td>${escapeHtml(file.type || "Not reported — browsers can't always detect a type for every extension")}</td></tr>
        <tr><th>Extension</th><td>.${escapeHtml(extOf(file.name) || "none")}</td></tr>
      </tbody></table></div>`;
    }});
  }
});
reg({
  slug: "document-metadata-viewer", name: "Document Metadata Viewer", category: "general", formats: ["DOCX", "XLSX", "ANY"],
  desc: "View file properties for any file, plus document properties for DOCX and XLSX.",
  seoDesc: "View file and document metadata online for free with Doxemi — works with any file, plus deep DOCX/XLSX properties.",
  render(el) {
    mountFileTool(el, { onFile: async (file, area) => {
      const ext = extOf(file.name);
      const base = { "File name": file.name, "File size": formatBytes(file.size), "MIME type": file.type || "—", "Last modified": file.lastModified ? new Date(file.lastModified).toString() : "—" };
      let extra = {};
      if (ext === "docx") extra = await readDocxMetadata(file);
      if (ext === "xlsx") {
        const XLSX = await loadXLSX();
        const wb = XLSX.read(await readFileAsArrayBuffer(file), { type: "array", bookProps: true });
        extra = { Sheets: wb.SheetNames.join(", "), Title: (wb.Props || {}).Title || "—", Author: (wb.Props || {}).Author || "—" };
      }
      const merged = { ...base, ...extra };
      area.innerHTML = `<div class="table-scroll"><table class="data-table"><tbody>${Object.entries(merged).map(([k, v]) => `<tr><th>${escapeHtml(k)}</th><td>${escapeHtml(v || "—")}</td></tr>`).join("")}</tbody></table></div>`;
    }});
  }
});
reg({
  slug: "document-format-checker", name: "Document Format Checker", category: "general", formats: ["ANY"],
  desc: "Check whether a file's extension matches its actual content signature.",
  seoDesc: "Check whether a file's extension matches its real format online for free with Doxemi's Document Format Checker.",
  render(el) {
    mountFileTool(el, { onFile: async (file, area) => {
      const ext = extOf(file.name);
      const buf = await readFileAsArrayBuffer(file);
      const sig = sniffSignature(buf);
      const textLike = looksLikeText(buf);
      const expectations = { pdf: "PDF document", docx: "ZIP-based file (DOCX, XLSX, PPTX or ZIP)", xlsx: "ZIP-based file (DOCX, XLSX, PPTX or ZIP)", doc: "Legacy Microsoft Office file (DOC/XLS/PPT)", xls: "Legacy Microsoft Office file (DOC/XLS/PPT)", rtf: "RTF document" };
      let verdict, tone;
      if (expectations[ext]) {
        const match = sig.kind === expectations[ext];
        verdict = match ? "Matches — the file's content signature is consistent with its .{ext} extension.".replace("{ext}", ext) : `Mismatch — the extension suggests ${expectations[ext]}, but the content looks like ${sig.kind || "something else"}.`;
        tone = match ? "alert-success" : "alert-error";
      } else if (["txt", "csv", "md", "json", "html", "xml"].includes(ext)) {
        verdict = textLike ? "Matches — the content looks like plain text, consistent with this extension." : "Mismatch — the content doesn't look like plain text.";
        tone = textLike ? "alert-success" : "alert-error";
      } else {
        verdict = "No signature check is available for this extension yet.";
        tone = "alert-info";
      }
      area.innerHTML = `<div class="alert ${tone}">${verdict}</div><p class="fine-print">Detected content type: ${escapeHtml(sig.kind || (textLike ? "plain text" : "unrecognized binary"))}</p>`;
    }});
  }
});
reg({
  slug: "document-extension-checker", name: "Document Extension Checker", category: "general", formats: ["ANY"],
  desc: "Check a file's extension against a list of common document formats.",
  seoDesc: "Check a file's extension online for free with Doxemi's Document Extension Checker.",
  render(el) {
    const known = { docx: "Word document", doc: "Legacy Word document", txt: "Plain text", rtf: "Rich text", csv: "Comma-separated values", xlsx: "Excel workbook", xls: "Legacy Excel workbook", pdf: "PDF document", json: "JSON data", html: "HTML document" };
    mountFileTool(el, { onFile: async (file, area) => {
      const ext = extOf(file.name);
      area.innerHTML = `<div class="table-scroll"><table class="data-table"><tbody>
        <tr><th>File name</th><td>${escapeHtml(file.name)}</td></tr>
        <tr><th>Extension</th><td>${ext ? "." + escapeHtml(ext) : "No extension found"}</td></tr>
        <tr><th>Recognized as</th><td>${escapeHtml(known[ext] || "Not a document format Doxemi recognizes")}</td></tr>
      </tbody></table></div>`;
    }});
  }
});

/* --------------------------------------------------------------------
   6. Featured / popular tools shown on the homepage
   -------------------------------------------------------------------- */
const POPULAR_SLUGS = ["docx-to-pdf", "docx-word-counter", "csv-to-json", "xlsx-to-csv", "txt-case-converter", "rtf-to-txt", "file-type-checker", "csv-cleaner"];

function toolBySlug(slug) { return TOOLS.find(t => t.slug === slug); }
function toolsInCategory(cat) { return TOOLS.filter(t => t.category === cat); }

function toolCardHTML(t) {
  return `<a class="tool-card" href="#/tool/${t.slug}">
    <div class="t-top"><span class="t-icon">${iconFor(t.category)}</span><span class="status-pill live">Live</span></div>
    <h3>${escapeHtml(t.name)}</h3>
    <p>${escapeHtml(t.desc)}</p>
    <span class="t-formats">${t.formats.join(" · ")}</span>
  </a>`;
}

function renderHomeExtras() {
  $("#category-grid").innerHTML = Object.entries(CATEGORIES).filter(([id]) => id !== "general").map(([id, c]) => `
    <a class="category-card" href="#/category/${id}">
      <span class="ext-tag">${c.ext}</span>
      <h3>${c.name}</h3>
      <p>${c.desc}</p>
      <span class="count">${toolsInCategory(id).length} tools · View tools</span>
    </a>`).join("");
  $("#popular-grid").innerHTML = POPULAR_SLUGS.map(s => toolBySlug(s)).filter(Boolean).map(toolCardHTML).join("");
}

/* --------------------------------------------------------------------
   7. Router
   -------------------------------------------------------------------- */
const homeView = () => $("#home-view");
const dynView = () => $("#dynamic-view");
const dynContent = () => $("#dynamic-content");
const crumbsEl = () => $("#breadcrumbs");

function setMeta({ title, description }) {
  document.title = title;
  const desc = description || "";
  $("head meta[name='description']").setAttribute("content", desc);
  $("#og-title").setAttribute("content", title);
  $("#og-description").setAttribute("content", desc);
  $("#twitter-title").setAttribute("content", title);
  $("#twitter-description").setAttribute("content", desc);
}

function setBreadcrumbs(items) {
  crumbsEl().innerHTML = items.map((it, i) => {
    const isLast = i === items.length - 1;
    return `${i > 0 ? '<span class="sep">/</span>' : ""}${isLast ? `<span aria-current="page">${escapeHtml(it.label)}</span>` : `<a href="${it.href}">${escapeHtml(it.label)}</a>`}`;
  }).join("");
}

function showHome() { homeView().hidden = false; dynView().hidden = true; }
function showDynamic() { homeView().hidden = true; dynView().hidden = false; }

function renderAllTools(initialFilter) {
  showDynamic();
  setBreadcrumbs([{ label: "Home", href: "#/" }, { label: "All Tools" }]);
  setMeta({ title: "All Document Tools | Doxemi", description: "Explore free tools for DOCX, TXT, RTF, CSV, XLSX and other document formats." });
  dynContent().innerHTML = `
    <div class="tool-page">
      <div class="tool-header">
        <h1>All document tools</h1>
        <p>Explore free tools for DOCX, TXT, RTF, CSV, XLSX and other document formats.</p>
      </div>
      <div class="tools-toolbar">
        <div class="page-search">
          <svg width="16" height="16" viewBox="0 0 18 18" fill="none" aria-hidden="true"><circle cx="8" cy="8" r="5.5" stroke="currentColor" stroke-width="1.6"/><line x1="12.2" y1="12.2" x2="16" y2="16" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/></svg>
          <input type="text" id="at-search" placeholder="Filter tools…">
        </div>
        <button class="filter-chip active" data-cat="all" type="button">All</button>
        ${Object.entries(CATEGORIES).map(([id, c]) => `<button class="filter-chip" data-cat="${id}" type="button">${c.name}</button>`).join("")}
      </div>
      <div class="tool-grid" id="at-grid"></div>
      <p class="no-results" id="at-empty" hidden>No tools match your filters.</p>
    </div>`;
  let activeCat = initialFilter || "all";
  const grid = $("#at-grid"), empty = $("#at-empty"), search = $("#at-search");
  function draw() {
    const q = search.value.trim().toLowerCase();
    const list = TOOLS.filter(t => (activeCat === "all" || t.category === activeCat) &&
      (!q || t.name.toLowerCase().includes(q) || t.desc.toLowerCase().includes(q) || t.formats.join(" ").toLowerCase().includes(q)));
    grid.innerHTML = list.map(toolCardHTML).join("");
    empty.hidden = list.length > 0;
  }
  $$(".filter-chip", dynContent()).forEach(chip => chip.addEventListener("click", () => {
    $$(".filter-chip", dynContent()).forEach(c => c.classList.remove("active"));
    chip.classList.add("active"); activeCat = chip.dataset.cat; draw();
  }));
  search.addEventListener("input", draw);
  if (initialFilter) $$(".filter-chip", dynContent()).forEach(c => c.classList.toggle("active", c.dataset.cat === initialFilter));
  draw();
}

function renderCategory(catId) {
  const cat = CATEGORIES[catId];
  if (!cat) return renderNotFound();
  showDynamic();
  setBreadcrumbs([{ label: "Home", href: "#/" }, { label: "Tools", href: "#/tools" }, { label: cat.name }]);
  setMeta({ title: `${cat.name} Online | Doxemi`, description: `${cat.desc} Free ${cat.name.toLowerCase()} from Doxemi — ${toolsInCategory(catId).length} tools, no sign-up required.` });
  dynContent().innerHTML = `
    <div class="tool-page">
      <div class="tool-header"><h1>${escapeHtml(cat.name)}</h1><p>${escapeHtml(cat.desc)}</p></div>
      <div class="tool-grid">${toolsInCategory(catId).map(toolCardHTML).join("")}</div>
    </div>`;
}

function renderToolPage(slug) {
  const tool = toolBySlug(slug);
  if (!tool) return renderNotFound();
  showDynamic();
  const cat = CATEGORIES[tool.category];
  setBreadcrumbs([{ label: "Home", href: "#/" }, { label: cat.name, href: `#/category/${tool.category}` }, { label: tool.name }]);
  setMeta({ title: `${tool.name} Online | Doxemi`, description: tool.seoDesc });
  const related = TOOLS.filter(t => t.category === tool.category && t.slug !== tool.slug).slice(0, 3);
  dynContent().innerHTML = `
    <article class="tool-page">
      <div class="tool-header">
        <div class="tool-meta-row"><span class="t-formats">${tool.formats.join(" · ")}</span><span class="status-pill live">Live tool</span></div>
        <h1>${escapeHtml(tool.name)}</h1>
        <p>${escapeHtml(tool.desc)}</p>
      </div>
      <div class="tool-panel" id="tool-mount"></div>
      <ol class="how-to">
        <li>Choose or drop your file into the box above.</li>
        <li>Doxemi processes it right in your browser.</li>
        <li>Review the result and download it if the tool produces a file.</li>
      </ol>
      ${related.length ? `<div class="related-tools"><h2>Related document tools</h2><div class="tool-grid">${related.map(toolCardHTML).join("")}</div></div>` : ""}
    </article>`;
  try { tool.render($("#tool-mount")); }
  catch (err) {
    console.error(err);
    $("#tool-mount").innerHTML = `<div class="alert alert-error">This tool couldn't load. Please refresh and try again.</div>`;
  }
}

const STATIC_PAGES = {
  about: { title: "About Doxemi", body: `<h1>About Doxemi</h1><p>Doxemi is a free collection of document tools for DOCX, TXT, RTF, CSV and XLSX files. Every tool is built to do one job well, and most run entirely in your browser rather than on a server.</p><p>Doxemi is an independent project, built and maintained by one developer.</p>` },
  contact: { title: "Contact Doxemi", body: `<h1>Contact</h1><p>Doxemi doesn't run a support inbox yet. For now, the best way to reach the person behind it is through the social links in the footer.</p>` },
  privacy: { title: "Privacy Policy", body: `<h1>Privacy policy</h1><p>Most Doxemi tools process files entirely in your browser using JavaScript. Those files are read locally and are not uploaded to a server as part of running the tool.</p><p>Some tools load a third-party JavaScript library (for example, for reading DOCX or XLSX files) from a public content delivery network the first time you use them. Loading that script does not send your document anywhere — it only fetches the code that runs the conversion locally.</p><p>Doxemi does not use tracking cookies or sell personal data. If that changes, this page will be updated.</p>` },
  terms: { title: "Terms of Use", body: `<h1>Terms of use</h1><p>Doxemi is provided free of charge, as-is, with no warranty of any kind. You're responsible for the files you process with it and for keeping your own backups.</p><p>Don't use Doxemi to process content you don't have the right to handle, or in a way that disrupts the service for other people.</p>` },
  guides: { title: "Guides", body: `<h1>Guides</h1><p>Longer how-to guides for common document workflows are on the way. In the meantime, every tool page includes short step-by-step instructions and links to related tools.</p>` },
  blog: { title: "Blog", body: `<h1>Blog</h1><p>The Doxemi blog doesn't have any posts yet. Check back soon.</p>` },
  faq: null // handled by scrolling to the home FAQ
};

function renderStatic(id) {
  const page = STATIC_PAGES[id];
  if (!page) return renderNotFound();
  showDynamic();
  setBreadcrumbs([{ label: "Home", href: "#/" }, { label: page.title }]);
  setMeta({ title: `${page.title} | Doxemi`, description: page.title + " — Doxemi, free online document tools." });
  dynContent().innerHTML = `<div class="static-page">${page.body}</div>`;
}

function renderNotFound() {
  showDynamic();
  setBreadcrumbs([{ label: "Home", href: "#/" }, { label: "Not found" }]);
  setMeta({ title: "Page not found | Doxemi", description: "This page doesn't exist on Doxemi." });
  dynContent().innerHTML = `<div class="static-page">
    <h1>Oops! Page not found.</h1>
    <p>That page doesn't exist. Try one of these instead.</p>
    <div class="action-row"><a class="btn btn-primary" href="#/">Back home</a><a class="btn btn-ghost" href="#/tools">Explore tools</a></div>
  </div>`;
}

function router() {
  const hash = location.hash.replace(/^#/, "") || "/";
  const parts = hash.split("/").filter(Boolean);
  window.scrollTo({ top: 0, behavior: "instant" in window ? "instant" : "auto" });

  $$(".main-nav a").forEach(a => a.classList.toggle("active", a.getAttribute("href") === "#" + hash));
  $("#mobile-nav").classList.remove("open");
  $("#nav-toggle").setAttribute("aria-expanded", "false");

  if (parts.length === 0) {
    showHome();
    setMeta({ title: "Doxemi – Free Online Document Tools", description: "Doxemi provides free online document tools for DOCX, TXT, RTF, CSV and XLSX files. View, convert, analyze, clean and manage documents quickly and easily." });
    return;
  }
  if (parts[0] === "tools") return renderAllTools();
  if (parts[0] === "category" && parts[1]) return renderCategory(parts[1]);
  if (parts[0] === "tool" && parts[1]) return renderToolPage(parts[1]);
  if (parts[0] === "faq") { showHome(); requestAnimationFrame(() => $("#faq").scrollIntoView({ behavior: "smooth" })); return; }
  if (STATIC_PAGES[parts[0]] !== undefined || parts[0] in STATIC_PAGES) return renderStatic(parts[0]);
  return renderNotFound();
}
window.addEventListener("hashchange", router);

/* --------------------------------------------------------------------
   8. Search overlay
   -------------------------------------------------------------------- */
function initSearch() {
  const overlay = $("#search-overlay"), input = $("#search-input"), list = $("#search-results"), empty = $("#search-empty"), clearBtn = $("#search-clear");
  let activeIndex = -1;

  function open() {
    overlay.hidden = false; input.value = ""; input.focus(); renderResults("");
    document.body.style.overflow = "hidden";
  }
  function close() {
    overlay.hidden = true; document.body.style.overflow = "";
  }
  function renderResults(q) {
    const query = q.trim().toLowerCase();
    const matches = !query ? TOOLS.slice(0, 8) : TOOLS.filter(t =>
      t.name.toLowerCase().includes(query) || t.desc.toLowerCase().includes(query) ||
      t.category.includes(query) || t.formats.join(" ").toLowerCase().includes(query)
    ).slice(0, 20);
    activeIndex = -1;
    clearBtn.hidden = !q;
    empty.hidden = matches.length > 0;
    list.innerHTML = matches.map((t, i) => `
      <li role="option"><button type="button" data-slug="${t.slug}" data-i="${i}">
        <span class="r-name">${escapeHtml(t.name)}</span>
        <span class="r-meta">${CATEGORIES[t.category].name} · ${t.formats.join(", ")}</span>
      </button></li>`).join("");
    $$("button", list).forEach(b => b.addEventListener("click", () => go(b.dataset.slug)));
  }
  function go(slug) { close(); location.hash = `#/tool/${slug}`; }
  function move(delta) {
    const buttons = $$("button", list);
    if (!buttons.length) return;
    activeIndex = (activeIndex + delta + buttons.length) % buttons.length;
    buttons.forEach(b => b.classList.remove("active"));
    buttons[activeIndex].classList.add("active");
    buttons[activeIndex].scrollIntoView({ block: "nearest" });
  }

  $("#search-open").addEventListener("click", open);
  $("#search-close").addEventListener("click", close);
  clearBtn.addEventListener("click", () => { input.value = ""; renderResults(""); input.focus(); });
  overlay.addEventListener("click", (e) => { if (e.target === overlay) close(); });
  input.addEventListener("input", (e) => renderResults(e.target.value));
  input.addEventListener("keydown", (e) => {
    if (e.key === "ArrowDown") { e.preventDefault(); move(1); }
    else if (e.key === "ArrowUp") { e.preventDefault(); move(-1); }
    else if (e.key === "Enter") { const b = $$("button", list)[activeIndex] || $$("button", list)[0]; if (b) go(b.dataset.slug); }
    else if (e.key === "Escape") close();
  });
  document.addEventListener("keydown", (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "k") { e.preventDefault(); overlay.hidden ? open() : close(); }
    if (e.key === "Escape" && !overlay.hidden) close();
  });
}

/* --------------------------------------------------------------------
   9. Theme (light / dark / system) with localStorage persistence
   -------------------------------------------------------------------- */
function initTheme() {
  const sunIcon = $("#theme-icon-sun"), moonIcon = $("#theme-icon-moon");
  const stored = localStorage.getItem("doxemi-theme");
  const prefersDark = window.matchMedia("(prefers-color-scheme: dark)").matches;
  applyTheme(stored || (prefersDark ? "dark" : "light"));

  function applyTheme(theme) {
    document.documentElement.setAttribute("data-theme", theme);
    sunIcon.hidden = theme === "dark";
    moonIcon.hidden = theme !== "dark";
  }
  $("#theme-toggle").addEventListener("click", () => {
    const next = document.documentElement.getAttribute("data-theme") === "dark" ? "light" : "dark";
    applyTheme(next);
    localStorage.setItem("doxemi-theme", next);
  });
}

/* --------------------------------------------------------------------
   10. Mobile nav
   -------------------------------------------------------------------- */
function initMobileNav() {
  const toggle = $("#nav-toggle"), nav = $("#mobile-nav");
  toggle.addEventListener("click", () => {
    const open = nav.classList.toggle("open");
    toggle.setAttribute("aria-expanded", String(open));
  });
  $$("#mobile-nav a").forEach(a => a.addEventListener("click", () => { nav.classList.remove("open"); toggle.setAttribute("aria-expanded", "false"); }));
}

/* --------------------------------------------------------------------
   11. Footer social links (only show configured ones)
   -------------------------------------------------------------------- */
const SOCIAL_ICONS = {
  instagram: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none"><rect x="3" y="3" width="18" height="18" rx="5" stroke="currentColor" stroke-width="1.6"/><circle cx="12" cy="12" r="4" stroke="currentColor" stroke-width="1.6"/><circle cx="17.2" cy="6.8" r="1.1" fill="currentColor"/></svg>`,
  youtube: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none"><rect x="2.5" y="5.5" width="19" height="13" rx="3.5" stroke="currentColor" stroke-width="1.6"/><path d="M10.5 9.5l5 2.5-5 2.5v-5Z" fill="currentColor"/></svg>`,
  x: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none"><path d="M4 4l16 16M20 4L4 20" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/></svg>`,
  facebook: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none"><path d="M14 21v-7h2.5l.5-3H14V9c0-.9.2-1.5 1.6-1.5H17V5c-.3 0-1.3-.1-2.4-.1-2.4 0-4 1.5-4 4.1V11H8v3h2.6v7H14Z" stroke="currentColor" stroke-width="1.4" stroke-linejoin="round"/></svg>`,
  linkedin: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none"><rect x="3" y="3" width="18" height="18" rx="3" stroke="currentColor" stroke-width="1.6"/><line x1="7.5" y1="10" x2="7.5" y2="17" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/><circle cx="7.5" cy="7" r="1" fill="currentColor"/><path d="M11.5 17v-4.2c0-1.6 1-2.6 2.4-2.6 1.3 0 2.1 1 2.1 2.6V17" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/></svg>`,
  github: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none"><path d="M12 2a10 10 0 0 0-3.16 19.5c.5.1.68-.22.68-.48v-1.7c-2.78.6-3.37-1.34-3.37-1.34-.46-1.15-1.1-1.46-1.1-1.46-.9-.62.07-.6.07-.6 1 .07 1.53 1.02 1.53 1.02.9 1.52 2.34 1.08 2.91.83.09-.65.35-1.08.63-1.33-2.22-.25-4.56-1.11-4.56-4.94 0-1.1.39-1.99 1.03-2.69-.1-.25-.45-1.27.1-2.65 0 0 .84-.27 2.75 1.02a9.4 9.4 0 0 1 5 0c1.9-1.3 2.75-1.02 2.75-1.02.55 1.38.2 2.4.1 2.65.64.7 1.03 1.6 1.03 2.69 0 3.84-2.35 4.68-4.58 4.93.36.31.68.92.68 1.85v2.75c0 .27.18.58.69.48A10 10 0 0 0 12 2Z" fill="currentColor"/></svg>`,
  whatsapp: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none"><path d="M12 2a10 10 0 0 0-8.6 15.1L2 22l5.05-1.36A10 10 0 1 0 12 2Z" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/><path d="M8.3 8.3c-.2.9 0 2.3 1.5 4 1.5 1.7 3 2.2 4 2.1.5-.05 1.1-.5 1.3-1l.2-.5-1.8-1.1-.5.6c-.15.15-.35.15-.5.05-.5-.3-1.2-.9-1.7-1.6-.1-.15-.1-.35.05-.5l.5-.55-1-1.9-.6.15c-.4.1-.75.4-.9.75Z" fill="currentColor"/></svg>`,
  website: `<svg width="16" height="16" viewBox="0 0 24 24" fill="none"><circle cx="12" cy="12" r="9" stroke="currentColor" stroke-width="1.6"/><path d="M3 12h18M12 3c2.5 2.6 3.8 5.7 3.8 9s-1.3 6.4-3.8 9c-2.5-2.6-3.8-5.7-3.8-9s1.3-6.4 3.8-9Z" stroke="currentColor" stroke-width="1.4"/></svg>`
};
function initFooterSocial() {
  const active = Object.entries(SOCIAL_LINKS).filter(([, url]) => url && !/^YOUR_.*_URL$/.test(url));
  const wrap = $("#social-links"), col = $("#footer-social");
  if (!active.length) { col.hidden = true; return; }
  wrap.innerHTML = active.map(([key, url]) => `<a href="${escapeHtml(url)}" target="_blank" rel="noopener noreferrer" aria-label="Doxemi on ${key}">${SOCIAL_ICONS[key] || ""}</a>`).join("");
}

/* --------------------------------------------------------------------
   12. Homepage smart drop zone — suggests a matching tool for the file
   -------------------------------------------------------------------- */
function initHeroDrop() {
  const zone = $("#hero-dropzone"), input = $("#hero-file-input"), out = $("#hero-drop-result");
  const openPicker = () => input.click();
  zone.addEventListener("click", openPicker);
  zone.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); openPicker(); } });
  ["dragenter", "dragover"].forEach(ev => zone.addEventListener(ev, (e) => { e.preventDefault(); zone.classList.add("drag"); }));
  ["dragleave", "drop"].forEach(ev => zone.addEventListener(ev, (e) => { e.preventDefault(); zone.classList.remove("drag"); }));
  zone.addEventListener("drop", (e) => { const f = e.dataTransfer.files[0]; if (f) suggest(f); });
  input.addEventListener("change", (e) => { const f = e.target.files[0]; if (f) suggest(f); });

  const suggestions = { docx: "docx-viewer", txt: "txt-viewer", rtf: "rtf-viewer", csv: "csv-viewer", xlsx: "xlsx-viewer" };
  function suggest(file) {
    const ext = extOf(file.name);
    const slug = suggestions[ext];
    out.hidden = false;
    if (slug) {
      const t = toolBySlug(slug);
      out.innerHTML = `<p><strong>${escapeHtml(file.name)}</strong> looks like a ${ext.toUpperCase()} file.</p><p><a href="#/tool/${slug}">Open it in ${escapeHtml(t.name)} →</a></p><p class="fine-print">Or browse all <a href="#/category/${ext}">${ext.toUpperCase()} tools</a>.</p>`;
    } else {
      out.innerHTML = `<p>Doxemi doesn't have a dedicated viewer for .${escapeHtml(ext || "this")} files yet. Try the <a href="#/tool/file-type-checker">File Type Checker</a> or browse <a href="#/tools">all tools</a>.</p>`;
    }
  }
}

/* --------------------------------------------------------------------
   13. Init
   -------------------------------------------------------------------- */
document.addEventListener("DOMContentLoaded", () => {
  $("#footer-year").textContent = new Date().getFullYear();
  renderHomeExtras();
  initSearch();
  initTheme();
  initMobileNav();
  initFooterSocial();
  initHeroDrop();
  router();
});
