# Voice Sheet App

A tiny, client‑side web app that lets you **dictate data into a spreadsheet‑like grid** and **export to Excel or CSV**.

- Works in modern Chromium browsers (Chrome, Edge). Safari/iOS support may vary.
- No server required; just open `index.html`.
- Data is auto‑saved to your browser (IndexedDB/localStorage).

## Features
- Start/stop voice input (Web Speech API).
- Command grammar for structured input:
  - `cell A1 42` → put `42` in A1
  - `row 3 values 10 11 12`
  - `column B from 2 values 5 6 7`
  - `next 8 9 10` (fills across the active row from the active cell)
  - `select A5` (moves cursor without writing)
  - `undo`, `redo`, `new sheet`, `clear sheet`
  - `save sheet "Week 1"` and `load sheet "Week 1"`
- Manual typing and arrow‑key navigation supported.
- Export to **Excel (.xlsx)** or **CSV**.

## Quick Start
1. Download and unzip the project.
2. Open `index.html` in Chrome/Edge (desktop is best).
3. Click **Mic ▶** to start dictating (allow microphone permissions).
4. Export via **Export → Excel** or **Export → CSV**.

## GitHub: How to publish
```bash
# 1) Create a new repo on GitHub (empty, no README)
# 2) Locally:
git clone <YOUR-NEW-REPO-URL> voice-sheet-app
cd voice-sheet-app
# 3) Copy these files into the repo folder, then:
git add .
git commit -m "feat: initial voice-to-spreadsheet app"
git push origin main
```
To enable GitHub Pages (optional): Settings → Pages → Deploy from branch → `main` / `/root`. Then visit the provided URL.

## License
MIT — do whatever you want, no warranty.
