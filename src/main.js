// src/main.js — self-diagnosing bootstrap + mic + import/export
function $(id){ return document.getElementById(id); }
function showOverlay(msg){
  console.error(msg);
  let el = document.getElementById('__err');
  if (!el){
    el = document.createElement('pre');
    el.id='__err';
    el.style.cssText='position:fixed;left:10px;bottom:10px;max-width:90vw;max-height:40vh;overflow:auto;background:#fff3f3;border:1px solid #d33;color:#900;padding:10px;border-radius:8px;z-index:99999;white-space:pre-wrap;';
    document.body.appendChild(el);
  }
  el.textContent = String(msg);
}
function filename(base, ext){
  const d=new Date(), p=n=>String(n).padStart(2,'0');
  return `${base}_${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())}_${p(d.getHours())}-${p(d.getMinutes())}.${ext}`;
}

window.addEventListener('DOMContentLoaded', async () => {
  // Kill service workers (common GitHub Pages gotcha)
  if ('serviceWorker' in navigator) {
    try {
      const regs = await navigator.serviceWorker.getRegistrations();
      regs.forEach(r=>r.unregister());
    } catch {}
  }

  // DOM sanity
  const host = $('grid-container') || $('gridContainer');
  if (!host) {
    showOverlay('❌ Missing grid container element: #grid-container (or #gridContainer). Add:\n<main id="grid-container"></main>');
    return;
  }
  if (!$('status-cell')) {
    showOverlay('⚠️ Missing status element #status-cell. Add it to the Status bar (will still continue).');
  }

  // Try to import grid.js (case-sensitive path!)
  let grid;
  try {
    grid = await import('./grid.js?v=' + Date.now()); // cache-bust
  } catch (e) {
    showOverlay(`❌ Failed to import ./grid.js\n\n${e.message}\n\nCheck:\n• File exists at /src/grid.js (case-sensitive)\n• index.html uses <script type="module" src="src/main.js">\n• Served over HTTPS (not file://)`);
    // Render a tiny fallback so page isn’t blank
    host.innerHTML = '<table class="grid"><tbody>'
      + Array.from({length:3},(_,r)=>'<tr>'
        + Array.from({length:3},(_,c)=>`<td class="cell" data-r="${r+1}" data-c="${c+1}"><input value=""></td>`).join('')
      + '</tr>').join('')
      + '</tbody></table>';
    return;
  }

  // Create grid if the API exists; else fallback render
  if (typeof grid.createGrid === 'function') {
    grid.createGrid(host, 100, 26);
    grid.setActive?.(1,1);
  } else {
    showOverlay('❌ grid.createGrid not found in grid.js. Rendering tiny fallback.');
    host.innerHTML = '<table class="grid"><tbody>'
      + Array.from({length:3},(_,r)=>'<tr>'
        + Array.from({length:3},(_,c)=>`<td class="cell" data-r="${r+1}" data-c="${c+1}"><input value=""></td>`).join('')
      + '</tr>').join('')
      + '</tbody></table>';
  }

  // Keep status bar in sync
  window.__onActiveChange = () => {
    const txt = grid.getActive?.() || 'A1';
    $('status-cell') && ( $('status-cell').textContent = txt );
  };

  // ---- Toolbar wiring (guard everything) ----
  $('btn-insert-row-above')?.addEventListener('click', ()=>grid.insertRow?.('above'));
  $('btn-insert-row-below')?.addEventListener('click', ()=>grid.insertRow?.('below'));
  $('btn-delete-row')?.addEventListener('click', ()=>grid.deleteRow?.());
  $('btn-move-row-up')?.addEventListener('click', ()=>grid.moveRow?.('up'));
  $('btn-move-row-down')?.addEventListener('click', ()=>grid.moveRow?.('down'));

  $('btn-insert-col-left')?.addEventListener('click', ()=>grid.insertCol?.('left'));
  $('btn-insert-col-right')?.addEventListener('click', ()=>grid.insertCol?.('right'));
  $('btn-delete-col')?.addEventListener('click', ()=>grid.deleteCol?.());
  $('btn-move-col-left')?.addEventListener('click', ()=>grid.moveCol?.('left'));
  $('btn-move-col-right')?.addEventListener('click', ()=>grid.moveCol?.('right'));

  $('btn-merge')?.addEventListener('click', ()=>grid.mergeSelection?.());
  $('btn-unmerge')?.addEventListener('click', ()=>grid.unmergeSelection?.());
  $('btn-autofit')?.addEventListener('click', ()=>grid.autoWidthFitSelected?.());

  $('btn-new-grid')?.addEventListener('click', ()=>grid.newGrid?.(host, 100, 26));
  $('btn-clear-grid')?.addEventListener('click', ()=>{ if (confirm('Clear ALL cells?')) grid.clearGrid?.(); });

  // ---- Context menu ----
  const ctx = $('context-menu');
  document.addEventListener('contextmenu', (e)=>{
    const cell = e.target.closest?.('.cell'); if (!cell) return;
    e.preventDefault();
    if (!ctx) return;
    ctx.style.display='block';
    const pad=6, W=ctx.offsetWidth||180, H=ctx.offsetHeight||150;
    ctx.style.left = Math.min(e.clientX, innerWidth - W - pad) + 'px';
    ctx.style.top  = Math.min(e.clientY, innerHeight - H - pad) + 'px';
  });
  document.addEventListener('click', (e)=>{ if (ctx && !ctx.contains(e.target)) ctx.style.display='none'; });

  $('ctx-copy')?.addEventListener('click', async ()=>{ await grid.copySelection?.(); ctx && (ctx.style.display='none'); });
  $('ctx-cut')?.addEventListener('click',  async ()=>{ await grid.cutSelection?.();  ctx && (ctx.style.display='none'); });
  $('ctx-paste')?.addEventListener('click',async ()=>{ await grid.pasteFromClipboard?.(); ctx && (ctx.style.display='none'); });
  $('ctx-delete')?.addEventListener('click',     ()=>{ grid.deleteSelection?.(); ctx && (ctx.style.display='none'); });

  // ---- Export (XLSX/CSV/TSV) ----
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
    if (typeof XLSX==='undefined'){ showOverlay('❌ XLSX not loaded'); return; }
    const ws = XLSX.utils.aoa_to_sheet(matrixForExport());
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, filename('VoiceSheet','xlsx'));
  }
  function exportCSV(){
    if (typeof XLSX==='undefined'){ showOverlay('❌ XLSX not loaded'); return; }
    const ws = XLSX.utils.aoa_to_sheet(matrixForExport());
    const csv = XLSX.utils.sheet_to_csv(ws);
    const blob = new Blob([csv], {type:'text/csv;charset=utf-8'});
    const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=filename('VoiceSheet','csv'); a.click();
  }
  function exportTSV(){
    if (typeof XLSX==='undefined'){ showOverlay('❌ XLSX not loaded'); return; }
    const ws = XLSX.utils.aoa_to_sheet(matrixForExport());
    const tsv = XLSX.utils.sheet_to_csv(ws, {FS:'\t'});
    const blob = new Blob([tsv], {type:'text/tab-separated-values;charset=utf-8'});
    const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=filename('VoiceSheet','tsv'); a.click();
  }
  $('export-xlsx')?.addEventListener('click', e=>{ e.preventDefault(); exportExcel(); });
  $('export-csv')?.addEventListener('click',  e=>{ e.preventDefault(); exportCSV(); });
  $('export-tsv')?.addEventListener('click',  e=>{ e.preventDefault(); exportTSV(); });

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
  $('import-xlsx')?.addEventListener('click', async e=>{
    e.preventDefault();
    const f = await pickFile('.xlsx'); if(!f) return;
    const buf = await f.arrayBuffer();
    const wb = XLSX.read(buf, {type:'array'}); const ws = wb.Sheets[wb.SheetNames[0]];
    const matrix = XLSX.utils.sheet_to_json(ws, {header:1});
    const pos = activeRC(); fillFromMatrixAt(matrix, pos.r, pos.c);
  });
  $('import-csv')?.addEventListener('click', async e=>{
    e.preventDefault();
    const f = await pickFile('.csv'); if(!f) return;
    const text = await f.text();
    const matrix = text.replace(/\r/g,'').split('\n').filter(Boolean).map(l=>l.split(','));
    const pos = activeRC(); fillFromMatrixAt(matrix, pos.r, pos.c);
  });
  $('import-tsv')?.addEventListener('click', async e=>{
    e.preventDefault();
    const f = await pickFile('.tsv,.txt'); if(!f) return;
    const text = await f.text();
    const matrix = text.replace(/\r/g,'').split('\n').filter(Boolean).map(l=>l.split('\t'));
    const pos = activeRC(); fillFromMatrixAt(matrix, pos.r, pos.c);
  });

  // ---- Microphone (Web Speech) ----
  const micBtn     = $('micBtn');
  const dirSelect  = $('direction');
  const autoChk    = $('autoAdvance');
  const silenceInp = $('silenceMs');

  function speechSupported(){
    return ('webkitSpeechRecognition' in window) || ('SpeechRecognition' in window);
  }
  if (micBtn){
    if (speechSupported()){
      const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
      const recognition = new SR();
      recognition.continuous = true;
      recognition.interimResults = false;
      recognition.lang = 'en-US';

      let recognizing = false;
      recognition.onresult = (ev)=>{
        const text = ev.results[ev.results.length-1][0].transcript.trim();
        const pos = activeRC();
        grid.setCell?.(pos.r, pos.c, text);
      };
      recognition.onend = ()=>{
        if (!recognizing) return;
        const delay = Math.max(200, +(silenceInp?.value||1000));
        setTimeout(()=>{
          if (!recognizing) return;
          if (autoChk?.checked){
            const dir = dirSelect?.value || 'right';
            grid.moveActive?.(dir, 1, { wrapRows:true, blockWidth:26, blockHeight:100 });
          }
          recognition.start();
        }, delay);
      };
      micBtn.onclick = ()=>{
        if (!recognizing){ recognizing = true; recognition.start(); micBtn.textContent='Mic ⏸'; }
        else { recognizing = false; recognition.stop(); micBtn.textContent='Mic ▶'; }
      };
    } else {
      micBtn.textContent = 'Mic (unsupported)';
      micBtn.disabled = true;
    }
  }
});
