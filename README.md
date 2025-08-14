# Voice Sheet App v2

A browser-based, speech-powered spreadsheet application.  
Enter data into a grid by **speaking**, navigate hands-free, and export/import Excel or CSV files — all without any backend server.

## ✨ Features

### Voice Input
- **Speech-to-cell** entry with [Web Speech API](https://developer.mozilla.org/en-US/docs/Web/API/Web_Speech_API/Using_the_Web_Speech_API)  
- Auto-advance after a configurable silence delay  
- Direction control: Right / Left / Down / Up  
- Number-only mode with word-to-digit conversion  
- **SpeechGrammar** hints for improved number recognition

### Spreadsheet Grid
- 100 rows × 26 columns (A–Z)
- Keyboard navigation + manual editing
- Undo / Redo
- Active cell highlighting

### Import & Export
- **Export** full sheet to `.xlsx` or `.csv` with sheet name + timestamp in filename
- **Import** `.xlsx` or `.csv` directly into the grid (first sheet only)

### Formulas (Basic)
- Supports `=A1`, `=SUM(A1:A5)`, `=SUM(A1,B3,C7)`
- Supports `=AVG(...)`
- Auto recalculates when referenced cells change

### Auto-Wrap Logic
- Wrap rows automatically after a set number of entries
- Configurable block width

### Visual Indicators
- Active cell pulses while listening
- Confirmation prompts for **Clear**, **New**, and **Load**

### Snapshots & History
- Save up to 10 local snapshots per sheet
- Restore previous versions instantly
- Snapshots are stored in browser localStorage (per device/browser)

---

## 🚀 Getting Started

### 1. Run locally
Clone or download this repo, then open `index.html` in Chrome or Edge (desktop).

```bash
git clone <your-repo-url>
cd voice-sheet-app
