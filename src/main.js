// src/main.js — safe driver for table grid + mic + import/export dropdowns
import * as grid from './grid.js';

function $(id){ return document.getElementById(id); }

// Run after DOM is ready
window.addEventListener('DOMContentLoaded', () => {
  const host = $('grid-container') || $('gridContainer');
  if (!host) {
    console.error('Grid container not found.');
    return;
  }

  // Create grid (fallback size if API missing)
  grid.createGrid?.(host, 100, 26);
  grid.setActive?.(1,1);

  // Keep status bar in sync (grid calls this when active cell changes)
  window.__onActiveChange = () => {
    const txt = grid.getActive?.() || 'A1';
    $('status-cell') && ( $('status-cell').textContent = txt );
  };

  // ---- Toolbar wiring (guards with optional chaining) ----
  $('btn-insert-row-above').onclick = () => grid.insertRow?.('above');
  $('btn-insert-row-below').onclick = () => grid.insertRow?.('below');
  $('btn-delete-row').onclick       = () => grid.deleteRow?.();
  $('btn-move-row-up').onclick      = () => grid.moveRow?.('up');
  $('btn-move-row-down').onclick    = () => grid.moveRow?.('down');

  $('btn-insert-col-left').onclick  = () => grid.insertCol?.('left');
  $('btn-insert-col-right').onclick = () => grid.insertCol?.('right');
  $('btn-delete-col').onclick       = () => grid.deleteCol?.();
  $('btn-move-col-left').onclick    = () => grid.moveCol?.('left');
  $('btn-move-col-right').onclick   = () => grid.moveCol?.('right');

  $('btn-merge').onclick            = () => grid.mergeSelection?.();
  $('btn-unmerge').onclick          = () => grid.unmergeSelection?.();
  $('btn-autofit').onclick          = () => grid.autoWidthFitSelected?.();

  $('btn-new-grid').onclick         = () => grid.newGrid?.(host, 100, 26);
  $('btn-clear-grid').onclick       = () => { if (confirm('Clear ALL cells?')) grid.clearGrid?.(); };

  // ---- Context menu ----
  const ctx = $('context-menu');
  document.addEventListener('contextmenu', (e)=>{
    const cell = e.target.closest?.('.cell'); if (!cell) return;
    e.preventDefault();
    ctx.style.display='block';
    const pad=6, W=ctx.offsetWidth||180, H=ctx.offsetHeight||150;
    ctx.style.left = Math.min(e.clientX, innerWidth - W - pad) + 'px';
    ctx.style.top  = Math.min(e.clientY, innerHeight - H - pad) + 'px';
  });
  document.addEventListener('click', (e)=>{ if (!ctx.contains(e.target)) ctx.style.display='none'; });

  $('ctx-copy').onclick   = async ()=>{ await grid.copySelection?.(); ctx.style.display='none'; };
  $('ctx-cut').onclick    = async ()=>{ await grid.cutSelection?.();  ctx.style.display='none'; };
  $('ctx-paste').onclick  = async ()=>{ await grid.pasteFromClipboard?.(); ctx.style.display='none'; };
  $('ctx-delete').onclick = ()=>{ grid.deleteSelection?.(); ctx.style.display='none'; };

  // ---- Export helpers (XLSX/CSV/TSV) ----
  function filename(base, ext){
    const d=new Date(), p=n=>String(n).padStart(2,'0');
    return `${base}_${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())}_${p(d.getHours())}-${p(d.getMinutes())}.${ext}`;
  }
  function matrixForExport(){
    const all = grid.getData?.() || [];
    const sel = grid.getSelection?.();
    if (!sel) return all;
    const r1=Math.min(sel.r1,sel.r2), r2=Math.max(sel.r1,sel.r2);
    const c1=Math.min(sel.c1,sel.c2), c2=Math.max(sel.c1,sel.c2);
    if (r1===r2 && c1===c2) return all;
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
  $('export-xlsx').onclick = (e)=>{ e.preventDefault(); exportExcel(); };
  $('export-csv').onclick  = (e)=>{ e.preventDefault(); exportCSV(); };
  $('export-tsv').onclick  = (e)=>{ e.preventDefault(); exportTSV(); };

  // ---- Import (paste at active cell) ----
  function activeRC(){
    const td = document.querySelector('.cell.active');
    return td ? { r:+td.dataset.r, c:+td.dataset.c } : { r:1, c:1 };
  }
  function fillFromMatrixAt(matrix, r0, c0){
    for(let i=0;i<matrix.length;i++){
      for(let j=0;j<matrix[i].length;j++){
        grid.setCell?.(r0+i, c0+j, matrix[i][j]);
      }
    }
  }
  async function pickFile(accept){
    return new Promise(res=>{
      const inp=document.createElement('input'); inp.type='file'; inp.accept=accept;
      inp.onchange=()=>res(inp.files && inp.files[0]); inp.click();
    });
  }
  $('import-xlsx').onclick = async (e)=>{
    e.preventDefault();
    const f = await pickFile('.xlsx'); if(!f) return;
    const buf = await f.arrayBuffer();
    const wb = XLSX.read(buf, {type:'array'}); const ws = wb.Sheets[wb.SheetNames[0]];
    const matrix = XLSX.utils.sheet_to_json(ws, {header:1});
    const pos = activeRC(); fillFromMatrixAt(matrix, pos.r, pos.c);
  };
  $('import-csv').onclick = async (e)=>{
    e.preventDefault();
    const f = await pickFile('.csv'); if(!f) return;
    const text = await f.text();
    const matrix = text.replace(/\r/g,'').split('\n').filter(Boolean).map(l=>l.split(','));
    const pos = activeRC(); fillFromMatrixAt(matrix, pos.r, pos.c);
  };
  $('import-tsv').onclick = async (e)=>{
    e.preventDefault();
    const f = await pickFile('.tsv,.txt'); if(!f) return;
    const text = await f.text();
    const matrix = text.replace(/\r/g,'').split('\n').filter(Boolean).map(l=>l.split('\t'));
    const pos = activeRC(); fillFromMatrixAt(matrix, pos.r, pos.c);
  };

  // ---- Microphone (Web Speech) ----
  const micBtn     = $('micBtn');
  const dirSelect  = $('direction');
  const autoChk    = $('autoAdvance');
  const silenceInp = $('silenceMs');

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
      grid.setCell?.(pos.r, pos.c, text);
    };
    recognition.onend = ()=>{
      if (!recognizing) return;
      const delay = Math.max(200, +silenceInp.value || 1000);
      setTimeout(()=>{
        if (!recognizing) return;
        if (autoChk.checked){
          const dir = dirSelect.value || 'right';
          grid.moveActive?.(dir, 1, { wrapRows:true, blockWidth:26, blockHeight:100 });
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
});
