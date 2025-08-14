import { colToLabel, labelToCol } from './utils.js';

let ROWS = 100, COLS = 26;
let active = { r: 1, c: 1 };
let undoStack = [], redoStack = [];

export function createGrid(container, rows, cols) {
  ROWS = rows; COLS = cols;
  container.innerHTML = '';
  const table = document.createElement('table');
  table.className = 'grid';

  const thead = document.createElement('thead');
  const hr = document.createElement('tr');
  const corner = document.createElement('th');
  hr.appendChild(corner);
  for (let c = 1; c <= COLS; c++) {
    const th = document.createElement('th');
    th.textContent = colToLabel(c);
    hr.appendChild(th);
  }
  thead.appendChild(hr);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  for (let r = 1; r <= ROWS; r++) {
    const tr = document.createElement('tr');
    const rowHdr = document.createElement('th');
    rowHdr.className = 'row-hdr';
    rowHdr.textContent = r;
    tr.appendChild(rowHdr);

    for (let c = 1; c <= COLS; c++) {
      const td = document.createElement('td');
      td.className = 'cell';
      td.dataset.r = r; td.dataset.c = c;
      const input = document.createElement('input');
      input.type = 'text';
      input.addEventListener('focus', () => setActive(r, c));
      input.addEventListener('input', (e) => recordEdit(r, c, e.target.value));
      input.addEventListener('keydown', (e) => {
        if (e.key === 'ArrowRight') setActive(r, Math.min(COLS, c+1));
        if (e.key === 'ArrowLeft') setActive(r, Math.max(1, c-1));
        if (e.key === 'ArrowUp') setActive(Math.max(1, r-1), c);
        if (e.key === 'ArrowDown') setActive(Math.min(ROWS, r+1), c);
      });
      td.appendChild(input);
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
  table.appendChild(tbody);
  container.appendChild(table);
  setActive(1,1);
}

export function newGrid(container, rows, cols) { createGrid(container, rows, cols); }
export function clearGrid() {
  document.querySelectorAll('.cell input').forEach(inp => inp.value = '');
  undoStack = []; redoStack = [];
}

export function getActive() { return `${colToLabel(active.c)}${active.r}`; }
export function setActive(r, c) {
  active = { r, c };
  document.querySelectorAll('.active').forEach(el=>el.classList.remove('active'));
  const td = document.querySelector(`.cell[data-r="${r}"][data-c="${c}"]`);
  if (td) td.classList.add('active');
  const input = td?.querySelector('input');
  input?.focus();
  if (window.__onActiveChange) window.__onActiveChange();
}
export function setCell(r, c, val) {
  const input = document.querySelector(`.cell[data-r="${r}"][data-c="${c}"] input`);
  if (!input) return;
  const prev = input.value;
  if (prev === String(val)) return;
  input.value = val ?? '';
  pushUndo({ r, c, prev, next: input.value });
  recalcAll();
}
export function onInputEdit(action) {
  if (action === 'undo') return undo();
  if (action === 'redo') return redo();
}
function recordEdit(r, c, v) { pushUndo({ r, c, prev: null, next: v, replace: true }); recalcAll(); }
function pushUndo(entry) { undoStack.push(entry); redoStack = []; }
function undo() {
  const last = undoStack.pop(); if (!last) return;
  const input = document.querySelector(`.cell[data-r="${last.r}"][data-c="${last.c}"] input`);
  if (!input) return;
  const cur = input.value;
  input.value = last.prev ?? '';
  redoStack.push({ ...last, prev: cur, next: input.value });
  recalcAll();
}
function redo() {
  const last = redoStack.pop(); if (!last) return;
  const input = document.querySelector(`.cell[data-r="${last.r}"][data-c="${last.c}"] input`);
  if (!input) return;
  const cur = input.value;
  input.value = last.prev ?? '';
  undoStack.push({ ...last, prev: cur, next: input.value });
  recalcAll();
}

export function getData() {
  const matrix = [];
  for (let r = 1; r <= ROWS; r++) {
    const row = [];
    for (let c = 1; c <= COLS; c++) {
      const v = document.querySelector(`.cell[data-r="${r}"][data-c="${c}"] input`).value || '';
      row.push(v);
    }
    matrix.push(row);
  }
  return matrix;
}

// Simple formula engine
function getCellRaw(r, c) {
  const input = document.querySelector(`.cell[data-r="${r}"][data-c="${c}"] input`);
  return input ? input.value : '';
}
function setCellDisplay(r, c, display) {
  const input = document.querySelector(`.cell[data-r="${r}"][data-c="${c}"] input`);
  if (input && input.value !== display) input.value = display;
}
function parseAddr(addr){
  const m = addr.toUpperCase().match(/^([A-Z]+)(\d{1,4})$/);
  if (!m) return null;
  return { r: parseInt(m[2],10), c: labelToCol(m[1]) };
}
function cellsInRange(r1,c1,r2,c2){
  const cells = [];
  for (let r=Math.min(r1,r2); r<=Math.max(r1,r2); r++) {
    for (let c=Math.min(c1,c2); c<=Math.max(c1,c2); c++) cells.push({r,c});
  }
  return cells;
}
function evalExpr(expr){
  let e = expr.trim();
  if (!e.startsWith('=')) return null;
  e = e.slice(1).trim();
  const ref = parseAddr(e);
  if (ref) {
    const v = parseFloat(getCellRaw(ref.r, ref.c));
    return isNaN(v) ? getCellRaw(ref.r, ref.c) : v;
  }
  const m = e.match(/^(SUM|AVG)\((.+)\)$/i);
  if (!m) return null;
  const fn = m[1].toUpperCase();
  const args = m[2].split(',').map(s=>s.trim());
  let values = [];
  for (const a of args) {
    const rm = a.match(/^(\w+):(\w+)$/);
    if (rm) {
      const a1 = parseAddr(rm[1]); const a2 = parseAddr(rm[2]);
      if (a1 && a2) {
        for (const cell of cellsInRange(a1.r,a1.c,a2.r,a2.c)) {
          const v = parseFloat(getCellRaw(cell.r, cell.c));
          if (!isNaN(v)) values.push(v);
        }
      }
    } else {
      const a1 = parseAddr(a);
      if (a1) {
        const v = parseFloat(getCellRaw(a1.r, a1.c));
        if (!isNaN(v)) values.push(v);
      } else {
        const v = parseFloat(a);
        if (!isNaN(v)) values.push(v);
      }
    }
  }
  if (fn === 'SUM') return values.reduce((a,b)=>a+b,0);
  if (fn === 'AVG') return values.length? (values.reduce((a,b)=>a+b,0)/values.length):0;
  return null;
}
function recalcAll(){
  for (let r=1;r<=ROWS;r++){
    for (let c=1;c<=COLS;c++){
      const raw = getCellRaw(r,c);
      if (raw && raw.trim().startsWith('=')) {
        const val = evalExpr(raw);
        if (val !== null) setCellDisplay(r,c,String(val));
      }
    }
  }
}

export function setRowValues(r, values, startCol=1) {
  for (let i = 0; i < values.length; i++) setCell(r, startCol + i, values[i]);
  setActive(r, startCol + Math.max(0, values.length-1));
}
export function setColumnValues(c, values, startRow=1) {
  for (let i = 0; i < values.length; i++) setCell(startRow + i, c, values[i]);
  setActive(startRow + Math.max(0, values.length-1), c);
}

export function addrToRC(addr) {
  const m = addr.toUpperCase().match(/^([A-Z]+)(\d{1,4})$/);
  if (!m) return null;
  const col = labelToCol(m[1]);
  const row = parseInt(m[2],10);
  return { r: row, c: col };
}

// Movement with wrap and block width
export function moveActive(dir, steps=1, opts={wrapRows:true, blockWidth:26}){
  dir = String(dir || '').toLowerCase();
  let r = active.r, c = active.c;
  const BW = Math.max(1, Math.min(COLS, opts.blockWidth || COLS));
  const startBlockCol = Math.floor((c-1)/BW)*BW + 1;
  const endBlockCol = Math.min(startBlockCol + BW - 1, COLS);

  if (dir === 'right') {
    c += steps;
    if (opts.wrapRows) {
      while (c > endBlockCol) {
        c = startBlockCol + (c - endBlockCol - 1);
        r = Math.min(ROWS, r + 1);
      }
    } else { c = Math.min(COLS, c); }
  } else if (dir === 'left') {
    c -= steps;
    if (opts.wrapRows) {
      while (c < startBlockCol) {
        c = endBlockCol - (startBlockCol - c - 1);
        r = Math.max(1, r - 1);
      }
    } else { c = Math.max(1, c); }
  } else if (dir === 'down') {
    r = Math.min(ROWS, r + steps);
  } else if (dir === 'up') {
    r = Math.max(1, r - steps);
  }
  setActive(r, c);
}
