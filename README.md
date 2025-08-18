# Voice Sheet App (Single-File)

Excel‑style voice-enabled sheet in a single HTML file. No build step, no dependencies.

## What’s in this repo
- `index.html` — the entire app (UI + logic).  
- `.gitignore` — ignores OS cruft.

## Quick start
Open `index.html` directly in a browser **or** serve locally for best results:
```bash
# Option A: VS Code Live Server (recommended)
# Right-click index.html → "Open with Live Server"

# Option B: Node http-server
npx http-server -p 5173
# then open http://127.0.0.1:5173/
```

## Configuration knobs
Edit these in `index.html`:

- **Default rows / cols** (boot line at the bottom):
  ```js
  createGrid(document.getElementById('grid-container'), 500, 26);
  ```
- **Silence limit** input width (layout only):
  ```css
  #silenceMs{ width: 10ch !important; }
  ```
- **Advance limit** input width (layout only):
  ```css
  #advanceLimit{ width: 7ch !important; }
  ```
- **Toolbar density** — tweak inside `<style id="toolbar-compact-0818">`.

## Features (high level)
- Spreadsheet‑like grid: edit, multi‑select, copy/cut/paste, drag‑fill, merge/unmerge, insert/delete rows/cols, resize, auto‑fit.
- Undo/Redo (whole‑cell operations).
- Voice dictation with mic toggle; auto‑advance directions (Right/Down and Left/Up) with Advance Limit.
- Import/Export: CSV/TSV and Excel (via SheetJS if included in your build).
- Mobile‑friendly gestures and compact ribbon layout.

## GitHub Pages (optional)
1. Push this repo to GitHub.
2. Settings → Pages → **Deploy from a branch** → Branch: `main` (root).  
3. Your app will be served from `https://<your-user>.github.io/<repo>/`.

## Notes
- This is a static app. No backend.
- Tested on latest Edge/Chrome. Clear cache (Ctrl/Cmd+Shift+R) after CSS tweaks.

_Last updated: 2025-08-18_


## License
MIT — see [LICENSE](LICENSE) for details.
