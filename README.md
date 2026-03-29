# PDF Tools — Chrome Extension

> A privacy-first Chrome extension for working with PDF files.
> All processing happens on your device — nothing is uploaded unless you explicitly opt in to a cloud API.

![Manifest V3](https://img.shields.io/badge/Manifest-V3-blue?style=flat-square)
![License](https://img.shields.io/badge/license-MIT-green?style=flat-square)
![Built with](https://img.shields.io/badge/built%20with-esbuild-yellow?style=flat-square)

---

## Features

| Feature | Description |
|---|---|
| **Merge PDF** | Combine multiple PDFs in any order. Drag to reorder, click Merge. |
| **PDF to Word** | Convert PDF to `.docx` with font size, bold/italic, and paragraph detection. |

---

## PDF to Word — Conversion Modes

The extension supports three conversion modes in priority order:

```
┌─────────────────────────────────────────────────────────┐
│                   PDF to Word Request                   │
└──────────────────────────┬──────────────────────────────┘
                           │
              ┌────────────▼────────────┐
              │  Local server URL set?  │
              └────────────┬────────────┘
               YES         │          NO
    ┌──────────▼──────┐    │    ┌──────▼──────────────────┐
    │  local-server   │    │    │  ConvertAPI key set?     │
    │  (Word or LO)   │    │    └──────┬──────────────────┘
    │  Perfect layout │    │     YES   │        NO
    └─────────────────┘    │  ┌────────▼──────┐  ┌────────────────┐
                           │  │  ConvertAPI   │  │  Built-in      │
                           │  │  Cloud, free  │  │  Text-only     │
                           │  │  250/month    │  │  Always works  │
                           │  └───────────────┘  └────────────────┘
```

| Mode | Quality | Privacy | Requires |
|---|---|---|---|
| **Built-in** | Text only | 100% local | Nothing |
| **ConvertAPI** | Perfect | Files sent to their servers | Free API key |
| **Local server** | Perfect | 100% local | Python + Word or LibreOffice |

---

## Installation

> This extension is not on the Chrome Web Store. Load it manually:

```bash
git clone <this-repo>
cd chrome_pdf_extension
npm install
npm run build
```

1. Open `chrome://extensions` in Chrome
2. Enable **Developer mode** (top-right toggle)
3. Click **Load unpacked** → select this folder

---

## Enabling Better Conversion Quality

### Option A — ConvertAPI (free, no install)

1. Sign up free at **[convertapi.com](https://www.convertapi.com/a/auth/register)** → get 250 conversions instantly
2. Copy your secret key
3. Open the extension → **PDF to Word** → paste key into **ConvertAPI key** field

### Option B — Local server (no files leave your machine)

Requires Python 3.8+ and either **Microsoft Word** or **LibreOffice** already installed.

```bash
python local-server.py
```

The server auto-detects whichever office engine is available:

```
Engine detection:
  Microsoft Word : C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE
  LibreOffice    : not found

Active engine  : Microsoft Word
Listening      : http://127.0.0.1:8765
```

Then in the extension → **PDF to Word** → **Server URL** → `http://localhost:8765`

---

## Project Structure

```
chrome_pdf_extension/
│
├── src/
│   ├── features/
│   │   ├── merge/
│   │   │   └── merge.js              # PDF merge logic (pdf-lib)
│   │   └── pdf-to-word/
│   │       ├── pdf-to-word.js        # Built-in converter (pdfjs + docx)
│   │       ├── stirling-pdf.js       # Stirling PDF / local server client
│   │       └── convertapi.js         # ConvertAPI cloud client
│   ├── popup.js                      # All UI logic and event wiring
│   └── open-tab.js                   # "Open in full tab" button
│
├── assets/
│   ├── css/
│   │   └── popup.css                 # Shared styles
│   ├── images/
│   │   ├── merge.svg                 # Merge feature icon
│   │   └── pdf-to-word.svg           # PDF to Word feature icon
│   └── icons/
│       └── icon.svg                  # Extension toolbar icon (source)
│
├── scripts/
│   └── copy-worker.js                # Build helper: copies pdfjs worker
│
├── local-server.py                   # Local conversion server (Word / LibreOffice)
├── popup.html                        # Extension popup
├── tab.html                          # Full browser tab view
├── manifest.json                     # Chrome MV3 manifest
└── package.json
```

---

## How It Works

### Merge PDF

```
Files selected
     │
     ▼
pdf-lib: PDFDocument.create()
     │
     ├── for each file:
     │     load → copyPages → addPage
     │
     ▼
merged.save() → Blob → download
```

Uses [`pdf-lib`](https://pdf-lib.js.org/) entirely in-memory. No page size limits.

### PDF to Word (built-in)

```
pdfjs: getTextContent()
     │
     ▼
Annotate each item:
  fontSize  ← transform matrix
  isBold    ← font name contains "bold"
  isItalic  ← font name contains "italic"
     │
     ▼
Group items → lines (same Y ± 3pt)
Group lines → paragraphs (Y-gap > 1.4× avg line height)
     │
     ▼
Detect headings by font size ratio:
  ≥ 1.8× body → H1
  ≥ 1.4× body → H2
  ≥ 1.15× body → H3
     │
     ▼
docx: Paragraph + TextRun (bold, italic, size)
     │
     ▼
Packer.toBlob() → download
```

---

## Known Limitations

- Password-protected PDFs are not supported
- Built-in PDF to Word is text only — scanned PDFs, images, and complex tables require a local server or ConvertAPI
- Multi-column layouts may mix column order in built-in mode
- `local-server.py` is single-threaded — one conversion at a time

---

## Development

```bash
npm run build     # bundles src/popup.js → bundle.js, copies pdf.worker.min.mjs
```

Reload the extension in `chrome://extensions` after each build.

**Build artifacts** (not committed):
- `bundle.js`
- `pdf.worker.min.mjs`
