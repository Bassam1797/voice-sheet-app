import { letters, clamp, copyToClipboard, readFromClipboard, parseTSV, toTSV } from './utils.js';

interface GridOptions {
  rows?: number;
  cols?: number;
}

interface CellChange {
  r: number;
  c: number;
  val: string;
  old?: string;
}

interface Selection {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

interface ActiveCell {
  r: number;
  c: number;
}

interface GridState {
  rows: number;
  cols: number;
  data: string[][];
  merges: Array<{ r1: number; c1: number; r2: number; c2: number }>;
  colWidths: number[];
  selection: Selection;
  active: ActiveCell;
  undo: CellChange[][];
  redo: CellChange[][];
  isMouseDown: boolean;
  selStart: { r: number; c: number } | null;
}

export function createGrid(container: HTMLElement, opts: GridOptions = {}) {
  const rows = opts.rows ?? 50;
  const cols = opts.cols ?? 26;

  const state: GridState = {
    rows,
    cols,
    data: Array.from({ length: rows }, () => Array(cols).fill('')),
    merges: [],
    colWidths: Array(cols).fill(100),
    selection: { r1: 0, c1: 0, r2: 0, c2: 0 },
    active: { r: 0, c: 0 },
    undo: [],
    redo: [],
    isMouseDown: false,
    selStart: null,
  };

  const gridEl = container;
  gridEl.innerHTML = '';
  gridEl.classList.add('grid');

  const table = document.createElement('table');
  table.className = 'table';
  table.setAttribute('role', 'grid');
  table.setAttribute('aria-label', 'Voice input spreadsheet grid');

  const thead = document.createElement('thead');
  const tbody = document.createElement('tbody');
  table.appendChild(thead);
  table.appendChild(tbody);
  gridEl.appendChild(table);

  const headRow = document.createElement('tr');
  const corner = document.createElement('th');
  corner.className = 'corner';
  corner.textContent = '';
  corner.setAttribute('role', 'columnheader');
  headRow.appendChild(corner);

  for (let c = 0; c < state.cols; c++) {
    const th = document.createElement('th');
    th.className = 'th-col';
    th.dataset.c = c.toString();
    th.textContent = letters(c);
    th.style.width = state.colWidths[c] + 'px';
    th.setAttribute('role', 'columnheader');
    th.setAttribute('aria-colindex', (c + 1).toString());

    const resizer = document.createElement('div');
    resizer.className = 'col-resizer';
    let startX = 0, startW = 0;

    resizer.addEventListener('mousedown', (e: MouseEvent) => {
      startX = e.clientX;
      startW = th.offsetWidth;
      const move = (ev: MouseEvent) => {
        const dx = ev.clientX - startX;
        const w = Math.max(24, startW + dx);
        th.style.width = w + 'px';
        state.colWidths[c] = w;
        api.emit('colWidthChange', api.getColumnWidths());
      };
      const up = () => {
        document.removeEventListener('mousemove', move);
        document.removeEventListener('mouseup', up);
      };
      document.addEventListener('mousemove', move);
      document.addEventListener('mouseup', up);
      e.stopPropagation();
      e.preventDefault();
    });

    th.appendChild(resizer);
    headRow.appendChild(th);
  }

  thead.appendChild(headRow);

  for (let r = 0; r < state.rows; r++) {
    const tr = document.createElement('tr');

    const th = document.createElement('th');
    th.className = 'th-row';
    th.textContent = (r + 1).toString();
    th.setAttribute('role', 'rowheader');
    th.setAttribute('aria-rowindex', (r + 1).toString());
    tr.appendChild(th);

    for (let c = 0; c < state.cols; c++) {
      const td = document.createElement('td');
      td.className = 'cell';
      td.dataset.r = r.toString();
      td.dataset.c = c.toString();
      td.contentEditable = 'true';
      td.spellcheck = false;

      td.setAttribute('role', 'gridcell');
      td.setAttribute('tabindex', '-1');
      td.setAttribute('aria-selected', 'false');
      td.setAttribute('aria-rowindex', (r + 1).toString());
      td.setAttribute('aria-colindex', (c + 1).toString());

      const dot = document.createElement('div');
      dot.className = 'active-dot';
      td.appendChild(dot);

      tr.appendChild(td);
    }

    tbody.appendChild(tr);
  }

  function getCell(r: number, c: number): HTMLTableCellElement {
    return tbody.children[r].children[c + 1] as HTMLTableCellElement;
  }

  function setActive(r: number, c: number): void {
    state.active = { r, c };
    const allCells = gridEl.querySelectorAll('[role="gridcell"]');
    allCells.forEach(el => {
      el.setAttribute('tabindex', '-1');
      el.setAttribute('aria-selected', 'false');
      el.classList.remove('active');
    });

    const activeCell = getCell(r, c);
    activeCell.classList.add('active');
    activeCell.setAttribute('tabindex', '0');
    activeCell.setAttribute('aria-selected', 'true');
    activeCell.focus();
  }

  function setSelection(r1: number, c1: number, r2: number, c2: number): void {
    state.selection = {
      r1: Math.min(r1, r2),
      c1: Math.min(c1, c2),
      r2: Math.max(r1, r2),
      c2: Math.max(c1, c2),
    };
  }

  // Initial focus
  setActive(0, 0);
  setSelection(0, 0, 0, 0);

  const api = {
    get data() {
      return state.data;
    },
    set data(d: string[][]) {
      state.data = d;
    },
    setColumnWidths(w: number[]) {
      state.colWidths = w.slice(0, state.cols);
    },
    getColumnWidths() {
      return state.colWidths.slice();
    },
    moveSelection(dir: 'left' | 'right' | 'up' | 'down') {
      const { r, c } = state.active;
      let nr = r, nc = c;
      if (dir === 'right') nc = clamp(c + 1, 0, state.cols - 1);
      if (dir === 'left') nc = clamp(c - 1, 0, state.cols - 1);
      if (dir === 'down') nr = clamp(r + 1, 0, state.rows - 1);
      if (dir === 'up') nr = clamp(r - 1, 0, state.rows - 1);
      setActive(nr, nc);
      setSelection(nr, nc, nr, nc);
    },
    editCurrentCell(val: string, { pushUndo = true }: { pushUndo?: boolean } = {}) {
      const { r, c } = state.active;
      const ch: CellChange = { r, c, old: state.data[r][c], val };
      state.data[r][c] = val;
      const td = getCell(r, c);
      td.textContent = val;
    },
    on(event: string, fn: Function) {
      (this._ev ||= {})[event] = (this._ev[event] || []).concat(fn);
    },
    emit(event: string, payload: any) {
      (this._ev?.[event] || []).forEach(fn => fn(payload));
    },
  };

  return api;
}
