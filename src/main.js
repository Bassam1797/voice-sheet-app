// src/main.js — drives the table grid from grid.js, restores mic, and wires dropdown import/export

import {
  createGrid,
  newGrid,
  clearGrid,
  setActive,
  getActive,
  moveActive,
  getSelection,
  getData,
  setCell,
  insertRow, insertCol, deleteRow, deleteCol, moveRow, moveCol,
  mergeSelection, unmergeSelection,
  autoWidthFitSelected,
  copySelection, cutSelection, pasteFromClipboard, deleteSelection
} from './grid.js';

/* ---------- bootstrap grid ---------- */
const host = document.getElementById('grid-container') || document.getElementById('gridContainer');
createGrid(host, 100, 26);

/* ---------- helpers ---------- */
function activeCoordsFromDOM() {
  const td = document.querySelector('.cell.active');
  if (!td) return null;
  return { r: parseInt(td.dataset.r,10), c: parseInt(td.dataset.c,10) };
}
function filename(base, ext) {
  const d = new Date(), pad = n => String(n).padStart(2,'0');
  return `${base}_${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}_${pad(d.getHours())}-${pad(d.getMinutes())}.${ext}`;
}

/* ---------- status hook from grid ---------- */
window.__onActiveChange = () => {
  document.getElementById('status-cell')?.textContent = getActive();
};

/* ---------- row/col ops ---------- */
document.getElementById('btn-insert-row-above').onclick = () => insertRow('above');
document.getElementById('btn-insert-row-below').onclick = () => insertRow('below');
document.getElementById('btn-delete-row').onclick = () => deleteRow();
document.getElementById('btn-move-row-up').onclick = () => moveRow('up');
document.getElementById('btn-move-row-down').onclick = () => moveRow('down');

document.getElementById('btn-insert-col-left').onclick = () => insertCol('left');
document.getElementById('btn-insert-col-right').onclick = () => insertCol('right');
document.getElementById('btn-delete-col').onclick = () => deleteCol();
document.getElementById('btn-move-col-left').onclick = () => moveCol('left');
document.getElementById('btn-move-col-right').onclick = () => moveCol('right');

document.getElementById('btn-merge').onclick = () => mergeSelection();
document.getElementById('btn-unmerge').onclick = () => unmergeSelection();
document.getElementById('btn-autofit').onclick = () => autoWidthFitSelected();

/* ---------- grid lifecycle ---------- */
document.getElementById('btn-new-grid').onclick = () => newGrid(host, 100, 26);
document.getElementById('btn-clear-grid').onclick = () => {
  if (confirm('Clear ALL cells on this sheet?')) clearGrid();
};

/* ---------- context menu (copy/cut/paste/delete) ---------- */
const ctx = document.getElementById('context-menu');
document.addEventListener('contextmenu', (e)=>{
  const cell = e.target.closest('.cell'); if (!cell) return;
  e.preventDefault();
  ctx.style.display='block';
  const vw=innerWidth, vh=innerHeight, W=ctx.offsetWidth||180, H=ctx.offsetHeight||150, pad=6;
  ctx.style.left = Math.min(e.clientX, vw-W-pad) + 'px';
  ctx.style.top  = Math.min(e.clientY, vh-H-pad) + 'px';
});
document.addEventListener('click', (e)=>{ if (!ctx.contains(e.target)) ctx.style.display='none'; });

document.getElementById('ctx-copy').onclick   = async ()=>{ await copySelection(); ctx.style.display='none'; };
document.getElementById('ctx-cut').onclick    = async ()=>{ await cutSelection();  ctx.style.display='none'; };
document.getElementById('ctx-paste').onclick  = async ()=>{ await pasteFromClipboard(); ctx.style.display='none'; };
document.getElementById('ctx-delete').onclick = ()=>{ deleteSelection(); ctx.style.display='none'; };

/* ---------- Export (XLSX/CSV/TSV) using SheetJS ---------- */
function matrixForExport() {
  const sel = getSelection?.();
  const matrix = getData();
  if (!sel) return matrix;
  const r1=Math.min(sel.r1,sel.r2), r2=Math.max(sel.r1,sel.r2);
  const c1=Math.min(sel.c1,sel.c2), c2=Math.max(sel.c1,sel.c2);
  const single = r1===r2 && c1===c2;
  if (single) return matrix;
  const out=[]; for(let r=r1;r<=r2;r++) out.push(matrix[r-1].slice(c1-1,c2));
  return out;
}
function exportExcel(){
  const data = matrixForExport();
  const ws = XLSX.utils.aoa_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  XLSX.writeFile(wb, filename('VoiceSheet','xlsx'));
}
function exportCSV(){
  const data = matrixForExport();
  const ws = XLSX.utils.aoa_to_sheet(data);
  const csv = XLSX.utils.sheet_to_csv(ws);
  const blob = new Blob([csv], {type:'text/csv;charset=utf-8'});
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=filename('VoiceSheet','csv'); a.click();
}
function exportTSV(){
  const data = matrixForExport();
  const ws = XLSX.utils.aoa_to_sheet(data);
  const tsv = XLSX.utils.sheet_to_csv(ws, {FS:'\t'});
  const blob = new Blob([tsv], {type:'text/tab-separated-values;charset=utf-8'});
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=filename('VoiceSheet','tsv'); a.click();
}
document.getElementById('export-xlsx').onclick = (e)=>{ e.preventDefault(); exportExcel(); };
document.getElementById('export-csv').onclick  = (e)=>{ e.preventDefault(); exportCSV(); };
document.getElementById('export-tsv').onclick  = (e)=>{ e.preventDefault(); exportTSV(); };

/* ---------- Import (XLSX/CSV/TSV) — paste at active cell ---------- */
function fillFromMatrixAt(matrix, startR, startC){
  for(let i=0;i<matrix.length;i++){
    for(let j=0;j<matrix[i].length;j++){
      const r = startR + i, c = startC + j;
      setCell(r, c, matrix[i][j]);
    }
  }
}
async function pickFile(accept){
  return new Promise(res=>{
    const inp=document.createElement('input'); inp.type='file'; inp.accept=accept;
    inp.onchange=()=>res(inp.files && inp.files[0]); inp.click();
  });
}
document.getElementById('import-xlsx').onclick = async (e)=>{
  e.preventDefault();
  const f = await pickFile('.xlsx'); if(!f) return;
  const buf = await f.arrayBuffer();
  const wb = XLSX.read(buf, {type:'array'}); const ws = wb.Sheets[wb.SheetNames[0]];
  const matrix = XLSX.utils.sheet_to_json(ws, {header:1});
  const pos = activeCoordsFromDOM() || {r:1,c:1};
  fillFromMatrixAt(matrix, pos.r, pos.c);
};
document.getElementById('import-csv').onclick = async (e)=>{
  e.preventDefault();
  const f = await pickFile('.csv'); if(!f) return;
  const text = await f.text();
  const matrix = text.replace(/\r/g,'').split('\n').filter(l=>l.length).map(l=>l.split(','));
  const pos = activeCoordsFromDOM() || {r:1,c:1};
  fillFromMatrixAt(matrix, pos.r, pos.c);
};
document.getElementById('import-tsv').onclick = async (e)=>{
  e.preventDefault();
  const f = await pickFile('.tsv,.txt'); if(!f) return;
  const text = await f.text();
  const matrix = text.replace(/\r/g,'').split('\n').filter(l=>l.length).map(l=>l.split('\t'));
  const pos = activeCoordsFromDOM() || {r:1,c:1};
  fillFromMatrixAt(matrix, pos.r, pos.c);
};

/* ---------- Microphone (Web Speech) with 1s silence auto-advance ---------- */
const micBtn = document.getElementById('micBtn');
const dirSel = document.getElementById('direction');
const autoChk = document.getElementById('autoAdvance');
const silenceMsInp = document.getElementById('silenceMs');

let recognizing = false;
let recognition = null;

function browserHasSpeech(){
  return ('webkitSpeechRecognition' in window) || ('SpeechRecognition' in window);
}

if (browserHasSpeech()){
  const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
  recognition = new SR();
  recognition.continuous = true;
  recognition.interimResults = false;
  recognition.lang = 'en-US';

  recognition.onresult = (ev)=>{
    const phrase = ev.results[ev.results.length-1][0].transcript.trim();
    // write into active cell
    const pos = activeCoordsFromDOM();
    if (pos) setCell(pos.r, pos.c, phrase);
  };
  let lastResultAt = 0;
  recognition.onresultend = ()=>{ lastResultAt = Date.now(); };
  recognition.onend = ()=>{
    if (!recognizing) return;
    // auto-advance after silence
    const delay = Math.max(200, +silenceMsInp.value || 1000);
    setTimeout(()=> {
      if (!recognizing) return;
      const dir = dirSel.value || 'right';
      moveActive(dir, 1, { wrapRows: true, blockWidth: 26, blockHeight: 100 });
      recognition.start();
    }, delay);
  };
  micBtn.onclick = ()=>{
    if (!recognizing){
      recognizing = true; recognition.start(); micBtn.textContent = 'Mic ⏸';
    } else {
      recognizing = false; recognition.stop(); micBtn.textContent = 'Mic ▶';
    }
  };
} else {
  micBtn.textContent = 'Mic (unsupported)';
  micBtn.disabled = true;
}

/* ---------- optional: keyboard focus on first cell ---------- */
setActive(1,1);
