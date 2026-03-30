# PDF-Flattener

A collection of drag-and-drop Windows batch tools for processing PDFs locally. No admin rights required — everything runs in your browser and files never leave your computer.

## Tools

### flattener.bat — Flatten PDFs

Double-click to open a browser-based drag-and-drop tool that flattens PDFs using [qpdf](https://github.com/qpdf/qpdf). Flattening bakes form fields, annotations, stamps, and signatures into the page content so they can no longer be edited.

- Handles forms, annotations, stamps, and digital signatures
- Detects oversized pages and automatically resizes them to US Letter (8.5 x 11") via Ghostscript
- Outputs a downloadable flattened PDF and saves a copy to a persistent output folder

### page-shrink.bat — Shrink Oversized Pages

Double-click to open a browser-based drag-and-drop tool that detects oversized PDF pages and shrinks them to US Letter (8.5 x 11") using [Ghostscript](https://www.ghostscript.com/), similar to Microsoft's "Print to PDF" shrink-to-fit behavior.

- Checks every page's mediabox against US Letter dimensions (612 x 792 points)
- Oversized pages are scaled to fit Letter using Ghostscript's `-dFIXEDMEDIA -dPDFFitPage` options
- PDFs already at Letter size are detected and skipped with a status message
- Does **not** flatten annotations or form fields (use `flattener.bat` for that)

## How It Works

Both tools use the same architecture:

1. A `.bat` file bootstraps an embedded PowerShell script
2. The PowerShell script starts a local HTTP server (default port 8739)
3. A browser-based drag-and-drop UI is served from memory
4. Uploaded PDFs are processed locally and served back for download
5. A persistent output folder is created in `%TEMP%` for easy access

## Dependencies

Downloaded automatically on first run — no manual setup needed:

| Tool | Purpose | Size |
|------|---------|------|
| [qpdf](https://github.com/qpdf/qpdf) v12.3.2 | PDF page inspection, flattening | ~23 MB |
| [Ghostscript](https://www.ghostscript.com/) v10.07.0 | Page resizing to Letter | ~62 MB |

Both are installed to `%LOCALAPPDATA%` (user-writable, no admin needed).

- `flattener.bat` always downloads qpdf; Ghostscript is downloaded only if oversized pages are detected.
- `page-shrink.bat` always downloads both qpdf (for page size detection) and Ghostscript (for resizing).

## Requirements

- Windows 10 or later
- PowerShell 5.1+ (included with Windows)
- Internet connection for first-run dependency download
