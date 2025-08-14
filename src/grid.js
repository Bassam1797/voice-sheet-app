// src/grid.js
// Core spreadsheet: render, selection, drag-fill (series), resize, merge, row/col ops, clipboard, undo/redo.

/////////////////////////// State ///////////////////////////

let ROWS = 100, COLS = 26;
let container = null;

// Matrix of { value: string, style?: {bold,size,color,bg,align,wrap} }
let cells = [];

// Active cell & selection
let active = { r: 1, c: 1 }; // 1-based
let sel = { r1: 1, c1: 1, r2: 1, c2: 1 };

// Undo/redo stacks
let undoStack = [];
let redoStack = [];

// Column widths (persisted)
let colWidths = JSON.parse(localStorage.getItem('colWidths') || '[]'); // index by 1..COLS

// Merges: array of rects { r1,c1,r2,c2 }
let merges = [];


/////////////////////////// Public API ///////////////////////////

export function createGrid(hostEl, rows = 100, cols = 26) {
  container = hostEl;
  ROWS = rows; COLS = cols;
  // init matrix
  cells = Array.from({ length: ROWS }, () =>
    Array.from({ length: COLS }, () => ({ value: '', style: {} }))
  );
  merges = [];
  undoStack = []; redoStack = [];

  render();
  attachGlobalShortcuts();
  setActive(1, 1);
  drawSelection();
  refreshStatsUI();
}

export function newGrid(hostEl, rows, cols) {
  createGrid(hostEl, rows, cols);
}

export function clearGrid() {
  for (let r = 1; r <= ROWS; r++) for (let c = 1; c <= COLS; c++) {
    cells[r-1][c-1].value = '';
    const inp = cellInput(r,c); if (inp) inp.value = '';
  }
  undoStack = []; redoStack = [];
  merges = [];
  drawSelection();
  refreshStatsUI();
}

export function getData() {
  const out = [];
  for (let r = 1; r <= ROWS; r++) {
    out.push(cells[r-1].map(o => o.value ?? ''));
  }
  return out;
}

export function setCell(r, c, val) {
  if (r < 1 || c < 1 || r > ROWS || c > COLS) return;
  const prev = cells[r-1][c-1].value ?? '';
  const next = val ?? '';
  if (prev === next) return;
  cells[r-1][c-1].value = next;
  const inp = cellInput(r,c);
  if (inp && inp.value !== next) inp.value = next;
  pushUndo({ type:'set', r, c, prev, next });
  redoStack = [];
}

export function onInputEdit(action){
  if (action === 'undo') undo();
  if (action === 'redo') redo();
}

export function getActive() { return `${colLabel(active.c)}${active.r}`; }
export function setActive(r, c, viaFocus=false) {
  active.r = clamp(r,1,ROWS); active.c = clamp(c,1,COLS);
  // single active dot (remove all, then add)
  container.querySelectorAll('.cell.active').forEach(el=>el.classList.remove('active'));
  const td = cellTd(active.r, active.c);
  td?.classList.add('active');
  if (!viaFocus) sel = { r1: active.r, c1: active.c, r2: active.r, c2: active.c };
  const inp = td?.querySelector('input'); if (inp) inp.focus({ preventScroll:true });
  window.__onActiveChange && window.__onActiveChange();
  drawSelection();
}

export function getSelection(){ return { ...sel }; }

export function moveActive(dir, steps=1, opts={wrapRows:true, blockWidth:26, blockHeight:100}){
  dir = String(dir||'').toLowerCase();
  let r = active.r, c = active.c;

  const BW = clamp(opts.blockWidth||COLS, 1, COLS);
  const startBlockCol = Math.floor((c-1)/BW)*BW + 1;
  const endBlockCol = Math.min(startBlockCol + BW - 1, COLS);

  const BH = clamp(opts.blockHeight||ROWS, 1, ROWS);
  const startBlockRow = Math.floor((r-1)/BH)*BH + 1;
  const endBlockRow = Math.min(startBlockRow + BH - 1, ROWS);

  if (dir === 'right') {
    c += steps;
    if (opts.wrapRows) while (c > endBlockCol) { c = startBlockCol + (c - endBlockCol - 1); r = Math.min(ROWS, r + 1); }
    else c = Math.min(COLS, c);
  } else if (dir === 'left') {
    c -= steps;
    if (opts.wrapRows) while (c < startBlockCol) { c = endBlockCol - (startBlockCol - c - 1); r = Math.max(1, r - 1); }
    else c = Math.max(1, c);
  } else if (dir === 'down') {
    r += steps;
    if (opts.wrapRows) while (r > endBlockRow) { r = startBlockRow + (r - endBlockRow - 1); c = Math.min(COLS, c + 1); }
    else r = Math.min(ROWS, r);
  } else if (dir === 'up') {
    r -= steps;
    if (opts.wrapRows) while (r < startBlockRow) { r = endBlockRow - (startBlockRow - r - 1); c = Math.max(1, c - 1); }
    else r = Math.max(1, r);
  }
  setActive(r, c);
}

export function autoWidthFitSelected(col){
  const c = col || active.c;
  const ctx = document.createElement('canvas').getContext('2d');
  // approximate current font:
  const body = getComputedStyle(document.body);
  ctx.font = `${body.fontSize} ${body.fontFamily}`;
  let max = 40;
  for (let r=1;r<=ROWS;r++){
    const v = (cells[r-1][c-1]?.value || '') + '  ';
    const w = ctx.measureText(v).width + 18;
    max = Math.max(max, w);
  }
  const th = container.querySelector(`thead th[data-c="${c}"]`);
  if (th) th.style.width = max + 'px';
  colWidths[c] = max;
  localStorage.setItem('colWidths', JSON.stringify(colWidths));
}

export function copySelection(){ return copySelectionTSV(); }
export async function cutSelection(){ await copySelectionTSV(); deleteSelection(); }
export async function pasteFromClipboard(){
  let text='';
  try { text = await navigator.clipboard.readText(); }
  catch { alert('Clipboard read blocked. Click a cell and try again.'); return; }
  if (!text) return;
  pasteTSVAtTopLeft(text);
}
export function deleteSelection(){
  const R = normSel(sel);
  for (let r=R.r1;r<=R.r2;r++) for (let c=R.c1;c<=R.c2;c++) setCell(r,c,'');
}

export function insertRow(where='below'){
  const r = active.r;
  const row = Array.from({length:COLS}, () => ({ value:'', style:{} }));
  const idx = (where==='above') ? (r-1) : r;
  cells.splice(idx, 0, row);
  ROWS = cells.length;
  rebuild();
  setActive((where==='above'? r : r+1), active.c);
}
export function insertCol(where='right'){
  const c = active.c;
  const idx = (where==='left') ? (c-1) : c;
  for (let r=0;r<ROWS;r++) cells[r].splice(idx, 0, { value:'', style:{} });
  COLS = cells[0].length;
  rebuild();
  setActive(active.r, (where==='left'? c : c+1));
}
export function deleteRow(){
  if (ROWS<=1) return;
  cells.splice(active.r-1,1);
  ROWS = cells.length;
  merges = merges.filter(m => !(m.r1>=active.r && m.r2<=active.r)); // naive cleanup
  rebuild();
  setActive(Math.min(active.r, ROWS), active.c);
}
export function deleteCol(){
  if (COLS<=1) return;
  for (let r=0;r<ROWS;r++) cells[r].splice(active.c-1,1);
  COLS = cells[0].length;
  merges = merges.filter(m => !(m.c1>=active.c && m.c2<=active.c));
  rebuild();
  setActive(active.r, Math.min(active.c, COLS));
}
export function moveRow(dir){
  const r = active.r;
  const to = (dir==='up') ? Math.max(1,r-1) : Math.min(ROWS, r+1);
  if (to===r) return;
  const [row] = cells.splice(r-1,1);
  cells.splice(to-1,0,row);
  rebuild();
  setActive(to, active.c);
}
export function moveCol(dir){
  const c = active.c;
  const to = (dir==='left') ? Math.max(1,c-1) : Math.min(COLS, c+1);
  if (to===c) return;
  for (let r=0;r<ROWS;r++){
    const [val] = cells[r].splice(c-1,1);
    cells[r].splice(to-1,0,val);
  }
  rebuild();
  setActive(active.r, to);
}

export function mergeSelection(){
  const R = normSel(sel);
  // Warn if multiple filled cells
  let filled=0;
  for (let r=R.r1;r<=R.r2;r++) for (let c=R.c1;c<=R.c2;c++) if (cells[r-1][c-1].value) filled++;
  if (filled>1 && !confirm('Merging will keep only the top-left value and clear the rest. Continue?')) return;
  // Clear all but top-left
  for (let r=R.r1;r<=R.r2;r++) for (let c=R.c1;c<=R.c2;c++) if (!(r===R.r1 && c===R.c1)) setCell(r,c,'');
  // Add merge rect
  merges.push(R);
  rebuild();
  setActive(R.r1, R.c1);
}
export function unmergeSelection(){
  const R = normSel(sel);
  merges = merges.filter(m => !(m.r1===R.r1 && m.c1===R.c1 && m.r2===R.r2 && m.c2===R.c2));
  rebuild();
}

export function refreshStatsUI(){ /* hook for index to compute stats, optional */ }

// expose needed bits for other modules
export { ROWS, COLS };


/////////////////////////// Rendering ///////////////////////////

function render(){
  const table = document.createElement('table');
  table.className = 'grid';

  // THEAD
  const thead = document.createElement('thead');
  const trh = document.createElement('tr');
  const corner = document.createElement('th');
  trh.appendChild(corner);
  for (let c=1;c<=COLS;c++){
    const th = document.createElement('th');
    th.textContent = colLabel(c);
    th.dataset.c = c;
    if (colWidths[c]) th.style.width = colWidths[c]+'px';
    const rez = document.createElement('div');
    rez.className = 'col-resizer';
    rez.addEventListener('mousedown', e => startColResize(e, c, th));
    th.appendChild(rez);
    trh.appendChild(th);
  }
  thead.appendChild(trh);
  table.appendChild(thead);

  // TBODY
  const tbody = document.createElement('tbody');
  for (let r=1;r<=ROWS;r++){
    const tr = document.createElement('tr');
    const rh = document.createElement('th');
    rh.className = 'row-hdr'; rh.textContent = r;
    tr.appendChild(rh);
    for (let c=1;c<=COLS;c++){
      const td = document.createElement('td');
      td.className = 'cell'; td.dataset.r=r; td.dataset.c=c;

      // merged coverage check
      const m = mergeCoverAt(r,c);
      if (m && !(m.r1===r && m.c1===c)) {
        // covered cell → hide
        td.hidden = true;
        tr.appendChild(td);
        continue;
      }
      // master: set spans
      if (m && m.r1===r && m.c1===c) {
        td.rowSpan = (m.r2 - m.r1 + 1);
        td.colSpan = (m.c2 - m.c1 + 1);
      }

      const input = document.createElement('input');
      input.value = cells[r-1][c-1].value ?? '';
      applyStyleToInput(input, cells[r-1][c-1].style);
      input.addEventListener('focus', ()=> setActive(r,c,true));
      input.addEventListener('input', (e)=> {
        const prev = cells[r-1][c-1].value ?? '';
        cells[r-1][c-1].value = e.target.value;
        pushUndo({ type:'set', r, c, prev, next:e.target.value });
        redoStack = [];
        refreshStatsUI();
      });
      input.addEventListener('mousedown', (e)=> startSelection(e, r, c));
      td.appendChild(input);

      // fill handle (only for visible cells)
      const fh = document.createElement('div');
      fh.className = 'fill-handle';
      fh.addEventListener('mousedown', (e)=> startFillDrag(e, r, c));
      td.appendChild(fh);

      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
  table.appendChild(tbody);

  container.innerHTML = '';
  container.appendChild(table);
}

function rebuild(){
  // Re-render keeping data/styles/merges/widths
  const prevActive = { ...active };
  render();
  setActive(prevActive.r, prevActive.c);
  drawSelection();
  refreshStatsUI();
}


/////////////////////////// Selection ///////////////////////////

function startSelection(e, r, c){
  sel = { r1:r, c1:c, r2:r, c2:c };
  drawSelection();

  const move = (ev)=>{
    const td = ev.target.closest('.cell'); if (!td) return;
    const rr = +td.dataset.r, cc = +td.dataset.c;
    sel.r2 = rr; sel.c2 = cc;
    drawSelection();
  };
  const up = ()=>{
    window.removeEventListener('mousemove', move);
    window.removeEventListener('mouseup', up);
  };
  window.addEventListener('mousemove', move);
  window.addEventListener('mouseup', up);
}

function drawSelection(){
  container.querySelectorAll('.sel-rect').forEach(el=>el.classList.remove('sel-rect'));
  const R = normSel(sel);
  for (let r=R.r1;r<=R.r2;r++){
    for (let c=R.c1;c<=R.c2;c++){
      cellTd(r,c)?.classList.add('sel-rect');
    }
  }
}

/////////////////////////// Fill handle & Series ///////////////////////////

function startFillDrag(e, r, c){
  e.stopPropagation(); e.preventDefault();

  // Seed from current selection if it's 1xN or Nx1; else single cell
  const sr = normSel(sel);
  const isRowSeed = (sr.r1===sr.r2 && sr.c1!==sr.c2);
  const isColSeed = (sr.c1===sr.c2 && sr.r1!==sr.r2);
  const seeds = collectSeeds(sr, isRowSeed, isColSeed, r, c);

  let lastSig = '';

  const move = (ev)=>{
    const td = ev.target.closest('.cell'); if (!td) return;
    const rr = +td.dataset.r, cc = +td.dataset.c;

    const dRow = Math.abs(rr - r), dCol = Math.abs(cc - c);
    const horizontal = dCol > dRow;

    const target = computeFillTarget(sr, horizontal, rr, cc);
    const sig = `${target.r1},${target.c1},${target.r2},${target.c2},${horizontal}`;
    if (sig === lastSig) return;
    lastSig = sig;

    applyFill(target, seeds, horizontal);
    sel = target; drawSelection();
  };

  const up = ()=>{
    window.removeEventListener('mousemove', move);
    window.removeEventListener('mouseup', up);
  };

  window.addEventListener('mousemove', move);
  window.addEventListener('mouseup', up);
}

function collectSeeds(rect, isRowSeed, isColSeed, fr, fc){
  if (isRowSeed) {
    const vals = [];
    for (let c=rect.c1; c<=rect.c2; c++) vals.push(cells[rect.r1-1][c-1].value ?? '');
    return { type:'row', values: vals };
  }
  if (isColSeed) {
    const vals = [];
    for (let r=rect.r1; r<=rect.r2; r++) vals.push(cells[r-1][rect.c1-1].value ?? '');
    return { type:'col', values: vals };
  }
  return { type:'single', values: [ cells[fr-1][fc-1].value ?? '' ] };
}

function computeFillTarget(seedRect, horizontal, rr, cc){
  const sr = normSel(seedRect);
  if (horizontal) {
    const right = cc >= sr.c2;
    const c1 = right ? sr.c2+1 : cc;
    const c2 = right ? cc : sr.c1-1;
    return { r1: sr.r1, c1: Math.min(c1,c2), r2: sr.r2, c2: Math.max(c1,c2) };
  } else {
    const down = rr >= sr.r2;
    const r1 = down ? sr.r2+1 : rr;
    const r2 = down ? rr : sr.r1-1;
    return { r1: Math.min(r1,r2), c1: sr.c1, r2: Math.max(r1,r2), c2: sr.c2 };
  }
}

function applyFill(target, seeds, horizontal){
  if (target.r2 < target.r1 || target.c2 < target.c1) return;

  // Series detection when seeds length >=2
  let series = null;
  if (seeds.values.length >= 2) series = detectSeries(seeds.values);

  if (horizontal) {
    for (let r=target.r1; r<=target.r2; r++){
      const base = seeds.values.length ? seeds.values : [''];
      const count = (target.c2 - target.c1 + 1);
      const out = buildSeries(base, series, count);
      for (let i=0;i<count;i++) setCell(r, target.c1 + i, out[i]);
    }
  } else {
    for (let c=target.c1; c<=target.c2; c++){
      const base = seeds.values.length ? seeds.values : [''];
      const count = (target.r2 - target.r1 + 1);
      const out = buildSeries(base, series, count);
      for (let i=0;i<count;i++) setCell(target.r1 + i, c, out[i]);
    }
  }
}

function detectSeries(vals){
  // Try numbers
  const nums = vals.map(parseMaybeNumber);
  if (nums.every(v => typeof v === 'number')) {
    if (nums.length >= 2) {
      const step = nums[nums.length-1] - nums[nums.length-2];
      return { kind:'number', step };
    }
  }
  // Try dates
  const dates = vals.map(parseMaybeDate);
  if (dates.every(v => v instanceof Date)) {
    if (dates.length >= 2) {
      const stepMs = dates[dates.length-1] - dates[nums.length-2];
      const stepDays = Math.round(stepMs / 86400000);
      return { kind:'date', stepDays: (isFinite(stepDays)? stepDays : 1) };
    }
  }
  return null;
}

function buildSeries(seedVals, series, count){
  if (!series) {
    const last = seedVals[seedVals.length-1] ?? '';
    return Array.from({length:count}, ()=> String(last));
  }
  if (series.kind === 'number') {
    const start = parseMaybeNumber(seedVals[seedVals.length-1]) ?? 0;
    return Array.from({length:count}, (_,i)=> String(start + series.step*(i+1)));
  }
  if (series.kind === 'date') {
    const start = parseMaybeDate(seedVals[seedVals.length-1]) ?? new Date();
    return Array.from({length:count}, (_,i)=>{
      const d = new Date(+start + (series.stepDays * (i+1) * 86400000));
      return isoDate(d);
    });
  }
  const last = seedVals[seedVals.length-1] ?? '';
  return Array.from({length:count}, ()=> String(last));
}


/////////////////////////// Column Resize ///////////////////////////

function startColResize(e, col, th){
  e.preventDefault(); e.stopPropagation();
  const startX = e.clientX;
  const startW = th.offsetWidth;
  const move = (ev)=>{
    const w = Math.max(40, startW + (ev.clientX - startX));
    th.style.width = w + 'px';
  };
  const up = (ev)=>{
    const w = Math.max(40, startW + (ev.clientX - startX));
    colWidths[col] = w;
    localStorage.setItem('colWidths', JSON.stringify(colWidths));
    window.removeEventListener('mousemove', move);
    window.removeEventListener('mouseup', up);
  };
  window.addEventListener('mousemove', move);
  window.addEventListener('mouseup', up);
}


/////////////////////////// Merge helpers ///////////////////////////

function mergeCoverAt(r,c){
  return merges.find(m => r>=m.r1 && r<=m.r2 && c>=m.c1 && c<=m.c2) || null;
}


/////////////////////////// Clipboard ///////////////////////////

async function copySelectionTSV(){
  const R = normSel(sel);
  const lines = [];
  for (let r=R.r1; r<=R.r2; r++){
    const row=[];
    for (let c=R.c1; c<=R.c2; c++) row.push(cells[r-1][c-1].value ?? '');
    lines.push(row.join('\t'));
  }
  const tsv = lines.join('\n');
  try { await navigator.clipboard.writeText(tsv); }
  catch {
    // Fallback
    const ta=document.createElement('textarea');
    ta.value=tsv; ta.style.position='fixed'; ta.style.opacity='0';
    document.body.appendChild(ta); ta.select(); document.execCommand('copy');
    document.body.removeChild(ta);
  }
}

function pasteTSVAtTopLeft(tsv){
  const R = normSel(sel);
  const lines = tsv.replace(/\r/g,'').split('\n');
  const matrix = lines.map(line => line.split('\t'));
  for (let i=0;i<matrix.length;i++){
    for (let j=0;j<matrix[i].length;j++){
      const r = R.r1 + i, c = R.c1 + j;
      if (r<=ROWS && c<=COLS) setCell(r,c,matrix[i][j]);
    }
  }
  drawSelection();
}


/////////////////////////// Undo/Redo ///////////////////////////

function pushUndo(entry){ undoStack.push(entry); }
function undo(){
  const last = undoStack.pop(); if (!last) return;
  if (last.type === 'set'){
    const cur = cells[last.r-1][last.c-1].value ?? '';
    cells[last.r-1][last.c-1].value = last.prev ?? '';
    const inp = cellInput(last.r,last.c); if (inp) inp.value = last.prev ?? '';
    redoStack.push({ ...last, prev: cur, next: last.prev ?? '' });
  }
  drawSelection(); refreshStatsUI();
}
function redo(){
  const last = redoStack.pop(); if (!last) return;
  if (last.type === 'set'){
    const cur = cells[last.r-1][last.c-1].value ?? '';
    cells[last.r-1][last.c-1].value = last.next ?? '';
    const inp = cellInput(last.r,last.c); if (inp) inp.value = last.next ?? '';
    undoStack.push({ ...last, prev: cur });
  }
  drawSelection(); refreshStatsUI();
}


/////////////////////////// Utilities ///////////////////////////

function colLabel(n){ let s=''; while (n>0){ const m=(n-1)%26; s=String.fromCharCode(65+m)+s; n=Math.floor((n-1)/26);} return s; }
function clamp(v,min,max){ return Math.max(min, Math.min(max, v)); }
function cellTd(r,c){ return container.querySelector(`.cell[data-r="${r}"][data-c="${c}"]`); }
function cellInput(r,c){ return cellTd(r,c)?.querySelector('input') || null; }
function normSel(a){ return { r1:Math.min(a.r1,a.r2), c1:Math.min(a.c1,a.c2), r2:Math.max(a.r1,a.r2), c2:Math.max(a.c1,a.c2) }; }
function parseMaybeNumber(s){ const t=String(s??'').trim().replace(',','.'); const n=Number(t); return Number.isFinite(n)? n : NaN; }
function parseMaybeDate(s){ const t=String(s??'').trim(); const d=new Date(t); return isNaN(+d)? null : d; }
function isoDate(d){ return d.toISOString().slice(0,10); }

function applyStyleToInput(inp, st){
  // extend later if you wire the formatting toolbar
  inp.style.fontWeight = st?.bold ? '700' : '';
  inp.style.fontSize = st?.size || '';
  inp.style.color = st?.color || '';
  inp.style.backgroundColor = st?.bg || '';
  inp.style.whiteSpace = st?.wrap ? 'normal' : 'nowrap';
  if (st?.align==='center') inp.style.textAlign='center';
  else if (st?.align==='right') inp.style.textAlign='right';
  else inp.style.textAlign='left';
}

function attachGlobalShortcuts(){
  // Right-click selection is handled in main.js; here we handle shortcuts
  document.addEventListener('keydown', async (e)=>{
    const meta = e.ctrlKey || e.metaKey;
    if (meta && e.key.toLowerCase()==='c'){ e.preventDefault(); await copySelectionTSV(); }
    if (meta && e.key.toLowerCase()==='x'){ e.preventDefault(); await copySelectionTSV(); deleteSelection(); }
    if (meta && e.key.toLowerCase()==='v'){ e.preventDefault(); let t=''; try{ t=await navigator.clipboard.readText(); }catch{} if(t){ pasteTSVAtTopLeft(t); e.preventDefault(); } }
    if (e.key==='Delete' || e.key==='Backspace'){ if (!meta) { e.preventDefault(); deleteSelection(); } }
    // arrows move active if not typing inside a cell
    const inCell = document.activeElement?.closest?.('.cell');
    if (!inCell){
      if (e.key==='ArrowRight'){ e.preventDefault(); setActive(active.r, clamp(active.c+1,1,COLS)); }
      if (e.key==='ArrowLeft'){ e.preventDefault(); setActive(active.r, clamp(active.c-1,1,COLS)); }
      if (e.key==='ArrowDown'){ e.preventDefault(); setActive(clamp(active.r+1,1,ROWS), active.c); }
      if (e.key==='ArrowUp'){ e.preventDefault(); setActive(clamp(active.r-1,1,ROWS), active.c); }
    }
  });
}

/////////////////////////// Exports used by index/main ///////////////////////////

export {
  // accessors
  setActive,
  // selection info
  sel as _selectionInternal
};
