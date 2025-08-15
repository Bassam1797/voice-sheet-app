# Voice Sheet App – V2

An Excel-like grid with voice entry, import/export, and power-user features.

## Features
- Inline editing, multi-cell selection (click + drag).
- Copy/Cut/Paste/Delete; clipboard uses TSV so it round-trips with Excel and Google Sheets.
- Drag-fill horizontally/vertically, including simple numeric series.
- Merge / Unmerge cells.
- Insert/Delete rows & columns; auto-fit column width (double-click header).
- Column resize with small minimum widths; widths persist in `localStorage`.
- Undo / Redo.
- Blue active-cell dot only on last active cell.
- Import/Export: XLSX, CSV, TSV (SheetJS).
- New Grid + Clear Grid buttons.
- Voice recognition (Web Speech API) with: numeric-only acceptance, no undo spam, auto-advance (configurable dir), silence timeout control.

## Run
Open `index.html` in a modern Chromium-based browser for best results (Chrome/Edge).

## Notes
- Browser clipboard permissions may be required for programmatic paste on some platforms.
- Web Speech API is experimental and not available in all browsers.
