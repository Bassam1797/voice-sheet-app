# Voice Sheet App (V2+, Single-File)

Patched features:
- Merge-safe row/col edits and moves
- Column-width persistence with insert/delete/move sync
- Undo **batching** for paste/import/delete/fill
- Scroll container + `scrollIntoView` + basic ARIA roles
- Safer clipboard fallback prompt
- Mic input: numeric-only parsing, undoable; manual start/stop; auto-advance
- Auto-fit uses computed input font

## Files
- `voice-sheet-app.html` — app
- `manifest.json` — PWA metadata
- `icon-192.png`, `icon-512.png`, `apple-touch-icon.png`, `favicon.ico` — icons
- `README.md`

## Deploy
1. Upload all files to a GitHub Pages repo (e.g., `docs/` or root).
2. Open `voice-sheet-app.html`.
3. For PWA install, serve over HTTPS (GitHub Pages does). Add to Home Screen.
