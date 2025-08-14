import { addrToRC, setActive, setCell, setRowValues, setColumnValues, getActive } from './grid.js';

let recog = null;
let heardCb = null;

export function setHeardCallback(cb){ heardCb = cb; }

export async function startListening(onPhrase) {
  const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
  if (!SR) { alert('Web Speech API not supported in this browser. Use Chrome/Edge desktop.'); return false; }
  recog = new SR();
  recog.lang = 'en-GB';
  recog.interimResults = true;
  recog.continuous = true;

  let transcript = '';
  recog.onresult = (e) => {
    transcript = '';
    for (let i= e.resultIndex; i<e.results.length; i++) {
      transcript += e.results[i][0].transcript;
      if (e.results[i].isFinal) {
        const phrase = transcript.trim();
        heardCb?.(phrase);
        onPhrase(phrase);
        transcript='';
      } else {
        heardCb?.(transcript.trim());
      }
    }
  };
  recog.onerror = (e) => console.warn('Speech error', e);
  recog.onend = () => {/* auto-stop */};

  try {
    await recog.start();
    return true;
  } catch (e) {
    alert('Microphone permission denied or unavailable.');
    return false;
  }
}

export function stopListening() {
  try { recog?.stop(); } catch {}
}

// --- Parser ---
export function parseAndExecute(text) {
  const t = text.trim().toLowerCase();

  // undo/redo/new/clear
  if (/^undo\b/.test(t)) { document.getElementById('undoBtn').click(); return; }
  if (/^redo\b/.test(t)) { document.getElementById('redoBtn').click(); return; }
  if (/^new( sheet)?\b/.test(t)) { document.getElementById('newSheetBtn').click(); return; }
  if (/^clear( sheet)?\b/.test(t)) { document.getElementById('clearSheetBtn').click(); return; }

  // save/load
  let m = t.match(/^save sheet\s+"(.+)"$/);
  if (m) { document.getElementById('sheetName').value = m[1]; document.getElementById('saveSheetBtn').click(); return; }
  m = t.match(/^load sheet\s+"(.+)"$/);
  if (m) { document.getElementById('sheetName').value = m[1]; document.getElementById('loadSheetBtn').click(); return; }

  // select cell
  m = t.match(/^select\s+([a-z]+\d{1,4})$/);
  if (m) { const rc = addrToRC(m[1]); if (rc) setActive(rc.r, rc.c); return; }

  // cell A1 42 (value may have spaces; capture everything after address)
  m = t.match(/^cell\s+([a-z]+\d{1,4})\s+(.+)$/);
  if (m) {
    const rc = addrToRC(m[1]);
    if (rc) setCell(rc.r, rc.c, m[2]);
    return;
  }

  // row 3 values 10 11 12
  m = t.match(/^row\s+(\d{1,4})\s+values\s+(.+)$/);
  if (m) {
    const r = parseInt(m[1],10);
    const values = splitValues(m[2]);
    setRowValues(r, values, 1);
    return;
  }

  // column C from 2 values 5 6 7
  m = t.match(/^column\s+([a-z]+)\s+from\s+(\d{1,4})\s+values\s+(.+)$/);
  if (m) {
    const cLabel = m[1].toUpperCase();
    const c = [...cLabel].reduce((n,ch)=>n*26+(ch.charCodeAt(0)-64),0);
    const startRow = parseInt(m[2],10);
    const values = splitValues(m[3]);
    setColumnValues(c, values, startRow);
    return;
  }

  // next 8 9 10 (fill across from active cell)
  m = t.match(/^next\s+(.+)$/);
  if (m) {
    const cur = getActive();
    const rc = addrToRC(cur);
    if (!rc) return;
    const values = splitValues(m[1]);
    for (let i=0;i<values.length;i++) {
      setCell(rc.r, rc.c + i, values[i]);
    }
    return;
  }

  // Fallback: if phrase looks like "A1 42"
  m = t.match(/^([a-z]+\d{1,4})\s+(.+)$/);
  if (m) {
    const rc = addrToRC(m[1]);
    if (rc) setCell(rc.r, rc.c, m[2]);
    return;
  }
}

function splitValues(s) {
  // split on spaces but keep quoted segments intact
  const out = [];
  let cur = '', inQ = false;
  for (let i=0;i<s.length;i++) {
    const ch = s[i];
    if (ch === '"') { inQ = !inQ; continue; }
    if (!inQ && ch === ' ') { if (cur) { out.push(cur); cur=''; } }
    else { cur += ch; }
  }
  if (cur) out.push(cur);
  return out;
}
