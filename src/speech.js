import { addrToRC, setActive, setCell, setRowValues, setColumnValues, getActive, moveActive } from './grid.js';

let recog = null;
let heardCb = null;

let advanceCfg = { direction: 'right', enabled: true, silenceMs: 1000, numbersOnly: true };
let wrapCfg = { wrapRows: true, blockWidth: 26 };
let silenceTimer = null;
let interimBuffer = '';
let listeningIndicator = true;

export function setHeardCallback(cb){ heardCb = cb; }
export function setAdvanceConfig(cfg){ advanceCfg = { ...advanceCfg, ...cfg }; }
export function setWrapConfig(cfg){ wrapCfg = { ...wrapCfg, ...cfg }; }
export function setListeningIndicator(v){ listeningIndicator = !!v; }

export async function startListening(onPhrase) {
  const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
  const GR = window.SpeechGrammarList || window.webkitSpeechGrammarList;
  if (!SR) { alert('Web Speech API not supported in this browser. Use Chrome/Edge desktop.'); return false; }

  recog = new SR();
  recog.lang = 'en-GB';
  recog.interimResults = true;
  recog.continuous = true;

  // Simple grammar hinting
  if (GR) {
    const grammarList = new GR();
    const numbers = 'zero | one | two | three | four | five | six | seven | eight | nine | ten | eleven | twelve | thirteen | fourteen | fifteen | sixteen | seventeen | eighteen | nineteen | twenty | thirty | forty | fifty | sixty | seventy | eighty | ninety | hundred | thousand | point | negative | minus';
    const commands = 'undo | redo | new | clear | next | left | right | up | down | enter | select | row | column | save | load';
    const jsgf = `#JSGF V1.0; grammar vs; public <item> = (${numbers}) | (${commands}) | (A | B | C | D | E | F | G | H | I | J | K | L | M | N | O | P | Q | R | S | T | U | V | W | X | Y | Z)+;`;
    try {
      grammarList.addFromString(jsgf, 1);
      recog.grammars = grammarList;
    } catch {}
  }

  const resetSilenceTimer = () => {
    if (!advanceCfg.enabled) return;
    clearTimeout(silenceTimer);
    silenceTimer = setTimeout(commitInterimIfAny, advanceCfg.silenceMs);
    setListeningPulse(true);
  };

  recog.onresult = (e) => {
    let latest = '';
    for (let i = e.resultIndex; i < e.results.length; i++) {
      latest += e.results[i][0].transcript;
      if (e.results[i].isFinal) {
        const phrase = latest.trim();
        heardCb?.(phrase);
        if (isCommandLike(phrase)) {
          onPhrase(phrase);
          interimBuffer = '';
        } else {
          commitValue(phrase);
        }
        latest = '';
        setListeningPulse(false);
      } else {
        interimBuffer = (latest || '').trim();
        heardCb?.(interimBuffer);
        resetSilenceTimer();
      }
    }
  };

  recog.onerror = (e) => { console.warn('Speech error', e); setListeningPulse(false); };
  recog.onend = () => { setListeningPulse(false); };

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
  setListeningPulse(false);
}

export function parseAndExecute(text) {
  const t = text.trim().toLowerCase();

  if (/^undo\b/.test(t)) { document.getElementById('undoBtn').click(); return; }
  if (/^redo\b/.test(t)) { document.getElementById('redoBtn').click(); return; }
  if (/^new( sheet)?\b/.test(t)) { document.getElementById('newSheetBtn').click(); return; }
  if (/^clear( sheet)?\b/.test(t)) { document.getElementById('clearSheetBtn').click(); return; }

  let m = t.match(/^save sheet\s+"(.+)"$/);
  if (m) { document.getElementById('sheetName').value = m[1]; document.getElementById('saveSheetBtn').click(); return; }
  m = t.match(/^load sheet\s+"(.+)"$/);
  if (m) { document.getElementById('sheetName').value = m[1]; document.getElementById('loadSheetBtn').click(); return; }

  m = t.match(/^select\s+([a-z]+\d{1,4})$/);
  if (m) { const rc = addrToRC(m[1]); if (rc) setActive(rc.r, rc.c); return; }

  m = t.match(/^(next|right)(?:\s+(\d+))?$/);
  if (m) { moveActive('right', parseInt(m[2] || '1',10), wrapCfg); return; }
  m = t.match(/^left(?:\s+(\d+))?$/);
  if (m) { moveActive('left', parseInt(m[1] || '1',10), wrapCfg); return; }
  m = t.match(/^(down(?:\s+(\d+))?|enter|next row)$/);
  if (m) { moveActive('down', parseInt(m[2] || '1',10), wrapCfg); return; }
  m = t.match(/^up(?:\s+(\d+))?$/);
  if (m) { moveActive('up', parseInt(m[1] || '1',10), wrapCfg); return; }

  m = t.match(/^cell\s+([a-z]+\d{1,4})\s+(.+)$/);
  if (m) {
    const rc = addrToRC(m[1]);
    if (rc) { setCell(rc.r, rc.c, m[2]); setActive(rc.r, rc.c); }
    return;
  }

  m = t.match(/^row\s+(\d{1,4})\s+values\s+(.+)$/);
  if (m) {
    const r = parseInt(m[1],10);
    const values = splitValues(m[2]);
    setRowValues(r, values, 1);
    return;
  }

  m = t.match(/^column\s+([a-z]+)\s+from\s+(\d{1,4})\s+values\s+(.+)$/);
  if (m) {
    const cLabel = m[1].toUpperCase();
    const c = [...cLabel].reduce((n,ch)=>n*26+(ch.charCodeAt(0)-64),0);
    const startRow = parseInt(m[2],10);
    const values = splitValues(m[3]);
    setColumnValues(c, values, startRow);
    return;
  }

  m = t.match(/^next\s+(.+)$/);
  if (m) {
    const values = splitValues(m[1]);
    moveActive('right', 1, wrapCfg);
    const cur = addrToRC(getActive());
    for (let i=0;i<values.length;i++) setCell(cur.r, cur.c + i, values[i]);
    setActive(cur.r, cur.c + Math.max(0, values.length-1));
    return;
  }

  m = t.match(/^([a-z]+\d{1,4})\s+(.+)$/);
  if (m) {
    const rc = addrToRC(m[1]);
    if (rc) { setCell(rc.r, rc.c, m[2]); setActive(rc.r, rc.c); }
    return;
  }
}

// Helpers
function isCommandLike(t){
  const x = t.toLowerCase();
  return /^(undo|redo|new|clear|save sheet|load sheet|select|row|column|next|left|right|up|down|enter|next row)\b/.test(x)
      || /^[a-z]+\d{1,4}\s+.+$/.test(x)
      || /^cell\s+[a-z]+\d{1,4}\s+.+$/.test(x);
}

function wordsToNumber(s){
  s = s.toLowerCase().trim();
  if (!s) return null;
  const small = {
    'zero':0,'one':1,'two':2,'three':3,'four':4,'five':5,'six':6,'seven':7,'eight':8,'nine':9,'ten':10,
    'eleven':11,'twelve':12,'thirteen':13,'fourteen':14,'fifteen':15,'sixteen':16,'seventeen':17,'eighteen':18,'nineteen':19
  };
  const tens = {'twenty':20,'thirty':30,'forty':40,'fifty':50,'sixty':60,'seventy':70,'eighty':80,'ninety':90};
  let total = 0, current = 0, hadWord=false;

  const tokens = s.replace(/-/g,' ').split(/\s+/);
  let i=0;
  while (i<tokens.length) {
    const w = tokens[i];
    if (w === 'minus' || w === 'negative') {
      const rest = wordsToNumber(tokens.slice(i+1).join(' '));
      return (rest==null) ? null : -rest;
    }
    if (w === 'point') {
      const intPart = hadWord ? total+current : null;
      if (intPart==null) return null;
      let fracStr = '';
      i++;
      for (; i<tokens.length; i++){
        if (tokens[i] in small) { fracStr += small[tokens[i]]; }
        else if (/^\d$/.test(tokens[i])) { fracStr += tokens[i]; }
        else break;
      }
      const frac = fracStr ? parseFloat('0.'+fracStr) : 0;
      return intPart + frac;
    }
    if (w in small) { current += small[w]; hadWord=true; i++; continue; }
    if (w in tens) { current += tens[w]; hadWord=true; i++; continue; }
    if (w === 'hundred') { current *= 100; hadWord=true; i++; continue; }
    if (w === 'thousand') { total += current * 1000; current = 0; hadWord=true; i++; continue; }
    if (/^\d+(?:\.\d+)?$/.test(w)) { return parseFloat(w); }
    break;
  }
  if (!hadWord) return null;
  return total + current;
}

function extractNumber(text){
  const trimmed = (text||'').trim();
  if (!trimmed) return null;
  const m = trimmed.match(/^-?\d+(?:\.\d+)?$/);
  if (m) return parseFloat(m[0]);
  const wnum = wordsToNumber(trimmed);
  if (wnum != null) return wnum;
  return null;
}

function normalizeValue(s){
  if (!s) return '';
  const num = extractNumber(s);
  if (advanceCfg.numbersOnly) {
    return (num == null) ? '' : String(num);
  }
  return (num == null) ? s.trim() : String(num);
}

function commitValue(phrase){
  const val = normalizeValue(phrase);
  if (!val) return;
  const cur = addrToRC(getActive()); if (!cur) return;
  setCell(cur.r, cur.c, val);
  if (advanceCfg.enabled) moveActive(advanceCfg.direction, 1, wrapCfg);
  interimBuffer = '';
}

function commitInterimIfAny(){
  const val = normalizeValue(interimBuffer);
  if (val) commitValue(val);
}

function splitValues(s) {
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

function setListeningPulse(on){
  if (!listeningIndicator) return;
  const a = addrToRC(getActive());
  if (!a) return;
  const td = document.querySelector(`.cell[data-r="${a.r}"][data-c="${a.c}"]`);
  if (!td) return;
  if (on) td.classList.add('listening'); else td.classList.remove('listening');
}
