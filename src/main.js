// src/main.js — table grid driver + mic + import/export dropdowns

import {
  // grid lifecycle
  createGrid, newGrid, clearGrid,
  // cell navigation & access
  setActive, getActive, moveActive, getSelection, getData, setCell,
  // row/col ops
  insertRow, insertCol, deleteRow, deleteCol, moveRow, moveCol,
  // merge
  mergeSelection, unmergeSelection,
  // column width
  autoWidthFitSelected,
  // clipboard/context
  copySelection, cutSelection, pasteFromClipboard, deleteSelection
} from './grid.js';

/* -------------------- init grid -------------------- */
const host = document.getElementById('grid-container') || document.getElementById('gridContainer');
createGrid(host, 100, 26);         // renders the TABLE grid

// keep status bar in sync
window.__onActiveChange = () => {
  const el = document.getElementById('status-cell');
  if (el) el.textContent = getActive();
};
setActive(1,1);

/* -------------------- toolbar: row/col ops -------------------- */
document.getElementById('btn-insert-row-above').onclick = () => insertRow('above');
document.getElementById('btn-insert-row-below').onclick = () => insertRow('below');
document.getElementById('btn-delete-row').onclick       = () => deleteRow();
document.getElementById('btn-move-row-up').onclick      = () => moveRow('up');
document.getElementById('btn-move-row-down').onclick    = () => moveRow('down');

document.getElementById('btn-insert-col-left').onclick  = () => insertCol('left');
document.getElementById('btn-insert-col-right').onclick = () => insertCol('right');
document.getElementById('btn-delete-col').onclick       = () => deleteCol();
document.getElementById('btn-move-col-left').onclick    = () => moveCol('left');
document.getElementById('btn-move-col-right').onclick   = () => moveCol('right');

document.getElementById('btn-merge').onclick            = () => mergeSelection();
document.getElementById('btn-unmerge').onclick          = () => unmergeSelection();
document.getElementById('btn-autofit').onclick          = () => autoWidthFitSelected();

document.getElementById('btn-new-grid').onclick         = () => newGrid(host, 100, 26);
document.getElementById('btn-clear-grid').onclick       = () => { if (confirm('Clear ALL cells?')) clearGrid(); };

/* -------------------- context menu -------------------- */
const ctx = document.getElementById('context-menu');
document.addEventListener('contextmenu', (e)=>{
  const cell = e.target.closest?.('.cell');
  if (!cell) return;
  e.preventDefault();
  ctx.style.display = 'block';
  const pad=6, W=ctx.offsetWidth||180, H=ctx.offsetHeight||150;
  const x = Math.min(e.clientX, innerWidth - W - pad);
  const y = Math.min(e.clientY, innerHeight - H - pad);
  ctx.style.left = x+'px';
  ctx.style.top  = y+'px';
});
document.addEventListener('click', (e)=>{ if (!ctx.contains(e.target)) ctx.style.display='none'; });

document.getElementById('ctx-copy').onclick   = async ()=>{ await copySelection(); ctx.style.display='none'; };
document.getElementById('ctx-cut').onclick    = async ()=>{ await cutSelection();  ctx.style.display='none'; };
document.getElementById('ctx-paste').onclick  = async ()=>{ await pasteFromClipboard(); ctx.style.display='none'; };
document.getElementById('ctx-delete').onclick = ()=>{ deleteSelection(); ctx.style.display='none'; };

/* -------------------- export (xlsx/csv/tsv) -------------------- */
function filename(base, ext){
  const d=new Date(), p=n=>String(n).padStart(2,'0');
  return `${base}_${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())}_${p(d.getHours())}-${p(d.getMinutes())}.${ext}`;
}
function matrixForExport(){
  const all = getData();
  const sel = getSelection?.();
  if (!sel) return all;
  const r1=Math.min(sel.r1,sel.r2), r2=Math.max(sel.r1,sel.r2);
  const c1=Math.min(sel.c1,sel.c2), c2=Math.max(sel.c1,sel.c2);
  if (r1===r2 && c1===c2) return all; // single cell selected → export whole sheet
  const out=[]; for(let r=r1;r<=r2;r++) out.push(all[r-1].slice(c1-1, c2));
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
  const ws = XLSX.utils.aoa_to_sheet(matrixForExport());
  const csv = XLSX.utils.sheet_to_csv(ws);
  const blob = new Blob([csv], {type:'text/csv;charset=utf-8'});
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=filename('VoiceSheet','csv'); a.click();
}
function exportTSV(){
  const ws = XLSX.utils.aoa_to_sheet(matrixForExport());
  const tsv = XLSX.utils.sheet_to_csv(ws, {FS:'\t'});
  const blob = new Blob([tsv], {type:'text/tab-separated-values;charset=utf-8'});
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=filename('VoiceSheet','tsv'); a.click();
}
document.getElementById('export-xlsx').onclick = (e)=>{ e.preventDefault(); exportExcel(); };
document.getElementById('export-csv').onclick  = (e)=>{ e.preventDefault(); exportCSV(); };
document.getElementById('export-tsv').onclick  = (e)=>{ e.preventDefault(); exportTSV(); };

/* -------------------- import (xlsx/csv/tsv) at active cell -------------------- */
function activeRC(){
  const td = document.querySelector('.cell.active'); 
  return td ? { r:+td.dataset.r, c:+td.dataset.c } : { r:1, c:1 };
}
function fillFromMatrixAt(matrix, r0, c0){
  for(let i=0;i<matrix.length;i++){
    for(let j=0;j<matrix[i].length;j++){
      setCell(r0+i, c0+j, matrix[i][j]);
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
  const pos = activeRC();
  fillFromMatrixAt(matrix, pos.r, pos.c);
};
document.getElementById('import-csv').onclick = async (e)=>{
  e.preventDefault();
  const f = await pickFile('.csv'); if(!f) return;
  const text = await f.text();
  const matrix = text.replace(/\r/g,'').split('\n').filter(Boolean).map(l=>l.split(','));
  const pos = activeRC();
  fillFromMatrixAt(matrix, pos.r, pos.c);
};
document.getElementById('import-tsv').onclick = async (e)=>{
  e.preventDefault();
  const f = await pickFile('.tsv,.txt'); if(!f) return;
  const text = await f.text();
  const matrix = text.replace(/\r/g,'').split('\n').filter(Boolean).map(l=>l.split('\t'));
  const pos = activeRC();
  fillFromMatrixAt(matrix, pos.r, pos.c);
};

/* -------------------- microphone (Web Speech) -------------------- */
const micBtn     = document.getElementById('micBtn');
const dirSelect  = document.getElementById('direction');
const autoChk    = document.getElementById('autoAdvance');
const silenceInp = document.getElementById('silenceMs');

let recognizing = false;
let recognition = null;

function speechSupported(){
  return ('webkitSpeechRecognition' in window) || ('SpeechRecognition' in window);
}

if (speechSupported()){
  const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
  recognition = new SR();
  recognition.continuous = true;
  recognition.interimResults = false;
  recognition.lang = 'en-US';

  recognition.onresult = (ev)=>{
    const text = ev.results[ev.results.length-1][0].transcript.trim();
    const pos = activeRC();
    setCell(pos.r, pos.c, text);
  };
  recognition.onend = ()=>{
    if (!recognizing) return;
    const delay = Math.max(200, +silenceInp.value || 1000);
    setTimeout(()=>{
      if (!recognizing) return;
      if (autoChk.checked){
        const dir = dirSelect.value || 'right';
        moveActive(dir, 1, { wrapRows:true, blockWidth:26, blockHeight:100 });
      }
      recognition.start();
    }, delay);
  };

  micBtn.onclick = ()=>{
    if (!recognizing){
      recognizing = true; recognition.start(); micBtn.textContent='Mic ⏸';
    } else {
      recognizing = false; recognition.stop(); micBtn.textContent='Mic ▶';
    }
  };
} else {
  micBtn.textContent = 'Mic (unsupported)';
  micBtn.disabled = true;
}
