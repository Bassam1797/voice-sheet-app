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
  cells = Array.from({ length: ROWS }, () =>
    Array.from({ length: COLS }, () => ({ value: '', style: {} }))
  );
  merges = [];
  undoStack = [];
  redoStack = [];

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
  for (let r = 1; r <= ROWS; r++) {
    for (let c = 1; c <= COLS; c++) {
      cells[r-1][c-1].value = '';
      const inp = cellInput(r,c); 
      if (inp) inp.value = '';
    }
  }
  undoStack = [];
  redoStack = [];
  merges = [];
  drawSelection();
  refreshStatsUI();
}

export function getData() {
  return cells.map(row => row.map(o => o.value ?? ''));
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

export function getActive() { 
  return `${colLabel(active.c)}${active.r}`; 
}

export function setActive(r, c, viaFocus=false) {
  active.r = clamp(r,1,ROWS); 
  active.c = clamp(c,1,COLS);
  container.querySelectorAll('.cell.active').forEach(el=>el.classList.remove('active'));
  const td = cellTd(active.r, active.c);
  td?.classList.add('active');
  if (!viaFocus) {
    sel = { r1: active.r, c1: active.c, r2: active.r, c2: active.c };
  }
  const inp = td?.querySelector('input'); 
  if (inp) inp.focus({ preventScroll:true });
  window.__onActiveChange && window.__onActiveChange();
  drawSelection();
}

export function getSelection(){ 
  return { ...sel }; 
}

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
  for (let r=R.r1;r<=R.r2;r++) {
    for (let c=R.c1;c<=R.c2;c++) setCell(r,c,'');
  }
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
  merges = merges.filter(m => !(m.r1>=active.r && m.r2<=active.r)); 
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
  let filled=0;
  for (let r=R.r1;r<=R.r2;r++) 
    for (let c=R.c1;c<=R.c2;c++) 
      if (cells[r-1][c-1].value) filled++;
  if (filled>1 && !confirm('Merging will keep only the top-left value and clear the rest. Continue?')) return;
  for (let r=R.r1;r<=R.r2;r++) 
    for (let c=R.c1;c<=R.c2;c++) 
      if (!(r===R.r1 && c===R.c1)) setCell(r,c,'');
  merges.push(R);
  rebuild();
  setActive(R.r1, R.c1);
}

export function unmergeSelection(){
  const R = normSel(sel);
  merges = merges.filter(m => !(m.r1===R.r1 && m.c1===R.c1 && m.r2===R.r2 && m.c2===R.c2));
  rebuild();
}

export function refreshStatsUI(){ /* optional hook */ }

export { ROWS, COLS, sel as _selectionInternal };
