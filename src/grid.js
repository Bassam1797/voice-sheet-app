import { colToLabel, labelToCol } from './utils.js';

let ROWS = 100, COLS = 26;
let active = { r: 1, c: 1 };
let undoStack = [], redoStack = [];

export function createGrid(container, rows, cols) {
  ROWS = rows; COLS = cols;
  container.innerHTML = '';
  const table = document.createElement('table');
  table.className = 'grid';

  // header row
  const thead = document.createElement('thead');
  const hr = document.createElement('tr');
  const corner = document.createElement('th'); // top-left corner
  hr.appendChild(corner);
  for (let c = 1; c <= COLS; c++) {
    const th = document.createElement('th');
    th.textContent = colToLabel(c);
    hr.appendChild(th);
  }
  thead.appendChild(hr);
  table.appendChild(thead);

  // body
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

export function getActive() {
  return `${colToLabel(active.c)}${active.r}`;
}
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
}
export function onInputEdit(action) {
  if (action === 'undo') return undo();
  if (action === 'redo') return redo();
}
function recordEdit(r, c, v) {
  // coalesce simple edits by pushing previous on first change
  // For simplicity, each input event is an undo step
  pushUndo({ r, c, prev: null, next: v, replace: true });
}
function pushUndo(entry) { undoStack.push(entry); redoStack = []; }
function undo() {
  const last = undoStack.pop(); if (!last) return;
  const input = document.querySelector(`.cell[data-r="${last.r}"][data-c="${last.c}"] input`);
  if (!input) return;
  const cur = input.value;
  input.value = last.prev ?? '';
  redoStack.push({ ...last, prev: cur, next: input.value });
}
function redo() {
  const last = redoStack.pop(); if (!last) return;
  const input = document.querySelector(`.cell[data-r="${last.r}"][data-c="${last.c}"] input`);
  if (!input) return;
  const cur = input.value;
  input.value = last.prev ?? '';
  undoStack.push({ ...last, prev: cur, next: input.value });
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

export function setRowValues(r, values, startCol=1) {
  for (let i = 0; i < values.length; i++) setCell(r, startCol + i, values[i]);
  setActive(r, startCol + Math.max(0, values.length-1));
}
export function setColumnValues(c, values, startRow=1) {
  for (let i = 0; i < values.length; i++) setCell(startRow + i, c, values[i]);
  setActive(startRow + Math.max(0, values.length-1), c);
}

// Helpers to parse addresses like "A5"
export function addrToRC(addr) {
  const m = addr.toUpperCase().match(/^([A-Z]+)(\d{1,4})$/);
  if (!m) return null;
  const col = labelToCol(m[1]);
  const row = parseInt(m[2],10);
  return { r: row, c: col };
}
