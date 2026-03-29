# PDF Tools — Chrome Extension

A privacy-first Chrome extension for working with PDF files. All processing happens locally in your browser — no uploads, no servers.

## Features

- **Merge PDF** — combine multiple PDFs in any order (drag to reorder)
- **PDF to Word** — extract text from PDFs and export as `.docx`

## Installation

This extension is not on the Chrome Web Store. Load it manually:

1. Clone or download this repo
2. Run `npm install && npm run build`
3. Open `chrome://extensions` in Chrome
4. Enable **Developer mode** (top-right toggle)
5. Click **Load unpacked** and select this folder

## Development

```bash
npm install       # install dependencies
npm run build     # build bundle.js and copy pdf.worker.min.mjs
```

Reload the extension in `chrome://extensions` after each build.

## Project Structure

```
├── src/
│   ├── popup.js          # Main UI logic (views, drag-drop, merge, convert)
│   └── pdf-to-word.js    # PDF text extraction + .docx generation
├── scripts/
│   └── copy-worker.js    # Build script: copies pdfjs worker to project root
├── popup.html            # Extension popup UI
├── tab.html              # Full browser tab UI (same features, larger canvas)
├── open-tab.js           # "Open in full tab" button handler
├── popup.css             # Shared styles
└── manifest.json         # Chrome extension manifest (MV3)
```

## How It Works

**Merge**: Uses [`pdf-lib`](https://pdf-lib.js.org/) to load each PDF and copy all pages into a new document, then triggers a download.

**PDF to Word**: Uses [`pdfjs-dist`](https://mozilla.github.io/pdf.js/) to extract text items from each page (sorted by position to approximate reading order), then builds a `.docx` with [`docx`](https://docx.js.org/). Layout, images, and tables are not preserved — text content only.

## Limitations

- Password-protected PDFs are not supported
- PDF to Word is text-only; formatting and images are lost
