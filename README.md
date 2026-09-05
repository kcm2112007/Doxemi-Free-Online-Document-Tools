# Doxemi

Simple tools for your documents.

Doxemi is a free, static, single-page website with real, working tools for **DOCX, TXT, RTF, CSV and XLSX** files — viewers, converters, word/character/line counters, cleaners and metadata inspectors. Most tools run entirely in your browser: your file is read locally and is not uploaded anywhere.

## What's in this repository

```
Doxemi/
├── index.html      the entire site (header, homepage, and the app shell)
├── style.css        all styling
├── script.js        all app logic — routing, search, and every tool
├── favicon.svg
├── manifest.json
├── robots.txt
├── sitemap.xml
├── 404.html
└── README.md
```

There is no build step, no `npm`, no framework. It's plain HTML/CSS/JS and works as a static site straight out of the folder.

## Supported formats

DOCX, TXT, RTF, CSV, XLSX — 42 tools in total across those five formats plus a handful of general, format-agnostic checks (file size, MIME type, extension). A few tools (DOCX/RTF → PDF) open your browser's print dialog rather than generating a PDF directly, since that's the reliable way to produce a real PDF entirely client-side; this is explained on each of those tool pages.

## Running it locally

Because the site uses `fetch`-free, same-origin JavaScript, you can usually just open `index.html` directly in a browser. If a tool's CDN library (Mammoth.js or SheetJS) doesn't load when opened as a plain `file://` page in your browser, serve the folder instead:

```bash
# any static server works, for example:
npx serve .
# or
python3 -m http.server 8080
```

Then visit `http://localhost:8080` (or whatever port your server prints).

## Uploading to GitHub

1. Create a new repository on GitHub (public, for GitHub Pages on the free plan).
2. Upload every file in this folder to the repository root (drag-and-drop through GitHub's web UI works fine — no command line needed).
3. Commit to the `main` branch.

## Enabling GitHub Pages

1. In your repository, go to **Settings → Pages**.
2. Under "Build and deployment", set **Source** to "Deploy from a branch".
3. Choose the `main` branch and the `/ (root)` folder, then save.
4. Your site will be live at:
   ```
   https://USERNAME.github.io/REPOSITORY/
   ```
   GitHub Pages usually takes a minute or two to publish after the first save.

Because every internal link in `index.html`/`script.js` is a relative or hash link, the site works correctly at that `/REPOSITORY/` sub-path — you don't need to change anything for it to work.

## Configuring your social links

Open `script.js` and find the `SOCIAL_LINKS` object near the top:

```js
const SOCIAL_LINKS = {
  instagram: "YOUR_INSTAGRAM_URL",
  youtube: "YOUR_YOUTUBE_URL",
  x: "YOUR_X_URL",
  facebook: "YOUR_FACEBOOK_URL",
  linkedin: "YOUR_LINKEDIN_URL",
  github: "YOUR_GITHUB_URL"
};
```

Replace the placeholder strings with your real profile URLs. Any entry left as `YOUR_..._URL` is automatically hidden from the footer, so you only need to fill in the ones you actually use.

## Filling in your real domain

Two files use a `YOUR_DOMAIN_HERE` placeholder instead of a guessed URL:

- `robots.txt` — the `Sitemap:` line
- `sitemap.xml` — the `<loc>` value

Once your GitHub Pages URL (or custom domain) is live, replace `YOUR_DOMAIN_HERE` with it, e.g. `https://USERNAME.github.io/REPOSITORY`.

A note on the sitemap: Doxemi is a single-page app, and every tool/category page lives at a `#/...` hash route rather than its own server-rendered URL. Search engines generally don't index URL fragments, so the sitemap intentionally lists only the one real document (`index.html`) rather than fabricated tool URLs that wouldn't resolve on their own.

## Customizing branding

- **Colors and type** live as CSS custom properties at the top of `style.css` (`:root` and `[data-theme="dark"]`) — change `--accent`, `--paper`, fonts, etc. there.
- **Logo** is inline SVG in the `<header>` and `<footer>` of `index.html`, and mirrored (simplified) in `favicon.svg`.
- **Copy** (hero text, FAQ, About/Privacy/Terms pages) lives directly in `index.html` and in the `STATIC_PAGES` object in `script.js`.

## Adding a new tool

Every tool is one object pushed into the `TOOLS` array in `script.js`, for example:

```js
reg({
  slug: "my-new-tool",
  name: "My New Tool",
  category: "txt",              // one of: docx, txt, rtf, csv, xlsx, general
  formats: ["TXT"],
  desc: "One-sentence description shown on cards.",
  seoDesc: "Longer sentence used as the page's meta description.",
  render(el) {
    // build your tool's UI inside `el`, using the mountFileTool()
    // or mountTextTool() helpers already used by the other tools
  }
});
```

It will automatically show up in its category page, the All Tools page, and search — no other file needs to change.

## Known limitations (disclosed on purpose)

- **DOCX/RTF/TXT → PDF** use the browser's native print dialog ("Save as PDF") rather than generating a PDF file directly, since that's the most reliable way to do it without a server.
- **RTF tools** use a lightweight, hand-written RTF-to-text parser. It correctly extracts text from typical RTF files but does not reproduce fonts, colors or embedded images — this is stated on the relevant tool pages.
- A few library-backed tools (DOCX via Mammoth.js, XLSX via SheetJS) load their library from a public CDN the first time you use them. If you're offline or the CDN is blocked, those specific tools won't load until you're back online; everything else keeps working.
