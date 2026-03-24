# PDF Editor

A client-side PDF editor built with vanilla JavaScript. All processing happens in the browser — no files are uploaded to any server.

## Features

**Annotation Tools**
- Shapes: rectangle, ellipse, line, arrow
- Freehand pen drawing
- Text with font family/size controls (Helvetica, Times New Roman, Courier)
- Highlighter with multiple colors
- Eraser
- Sticky notes
- Signature pad (draw and save for reuse)
- Stamps (DRAFT, CONFIDENTIAL, APPROVED, REVIEWED, FINAL, or custom text)

**Editing Tools**
- Redact sensitive content
- Crop pages
- Select, move, and resize annotations
- Adjustable thickness, opacity, color (presets + hex picker)
- Fill/stroke toggle

**Document Management**
- Merge multiple PDFs
- Split PDF by page ranges
- Page thumbnails with drag-to-reorder
- Find text in document
- Zoom in/out (0.5x–3x)

**UX**
- Undo/redo
- Keyboard shortcuts for every tool
- Dark mode
- Drag-and-drop file upload
- Download annotated PDF
- Fully accessible (ARIA labels, screen reader announcements)

## Getting Started

No build step required. Serve the files with any static HTTP server:

```bash
# Python
python3 -m http.server 8000

# Node.js (npx)
npx serve .

# PHP
php -S localhost:8000
```

Then open `http://localhost:8000` in your browser.

> **Note:** Opening `index.html` directly via `file://` won't work due to ES module imports.

## Project Structure

```
index.html          — Main HTML shell
pdf-editor.js       — Application logic (~3,800 lines)
pdf-editor.css      — Styles and design tokens
vendor/
  pdf-lib.esm.js    — pdf-lib (PDF creation/modification)
  pdf.min.mjs       — PDF.js (PDF rendering)
  pdf.worker.min.mjs — PDF.js web worker
```

## Libraries

- [pdf-lib](https://pdf-lib.js.org/) — Create and modify PDFs
- [PDF.js](https://mozilla.github.io/pdf.js/) — Render PDFs in the browser

## License

MIT
