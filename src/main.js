// src/main.js
import {
  createGrid,
  setActive,
  getActive,
  getSelection,
  copySelection,
  cutSelection,
  pasteFromClipboard,
  deleteSelection,
  insertRow,
  insertCol,
  deleteRow,
  deleteCol,
  moveRow,
  moveCol,
  mergeSelection,
  unmergeSelection,
  autoWidthFitSelected,
  newGrid,
  clearGrid
} from './grid.js';

let gridHost = document.getElementById('grid-container');

// Init grid
createGrid(gridHost, 100, 26);

// ===== Toolbar buttons ===== //
document.getElementById('btn-insert-row-above').onclick = () => insertRow('above');
document.getElementById('btn-insert-row-below').onclick = () => insertRow('below');
document.getElementById('btn-insert-col-left').onclick = () => insertCol('left');
document.getElementById('btn-insert-col-right').onclick = () => insertCol('right');
document.getElementById('btn-delete-row').onclick = () => deleteRow();
document.getElementById('btn-delete-col').onclick = () => deleteCol();
document.getElementById('btn-move-row-up').onclick = () => moveRow('up');
document.getElementById('btn-move-row-down').onclick = () => moveRow('down');
document.getElementById('btn-move-col-left').onclick = () => moveCol('left');
document.getElementById('btn-move-col-right').onclick = () => moveCol('right');

document.getElementById('btn-merge').onclick = () => mergeSelection();
document.getElementById('btn-unmerge').onclick = () => unmergeSelection();
document.getElementById('btn-autofit').onclick = () => autoWidthFitSelected();

// ===== File import/export ===== //
document.getElementById('btn-export').onclick = () => {
  const sel = getSelection();
  const r1 = sel.r1, c1 = sel.c1, r2 = sel.r2, c2 = sel.c2;
  let tsv = '';
  for (let r = r1; r <= r2; r++) {
    const row = [];
    for (let c = c1; c <= c2; c++) {
      const td = document.querySelector(`.cell[data-r="${r}"][data-c="${c}"] input`);
      row.push(td ? td.value : '');
    }
    tsv += row.join('\t') + '\n';
  }
  const blob = new Blob([tsv], { type: 'text/tab-separated-values' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'selection.tsv';
  a.click();
};

document.getElementById('btn-import').onclick = async () => {
  const file = document.createElement('input');
  file.type = 'file';
  file.accept = '.tsv,.txt';
  file.onchange = async () => {
    const text = await file.files[0].text();
    await pasteFromClipboard(text);
  };
  file.click();
};

// ===== Context menu ===== //
const menu = document.getElementById('context-menu');
document.addEventListener('contextmenu', (e) => {
  const cell = e.target.closest('.cell');
  if (!cell) return;
  e.preventDefault();
  menu.style.display = 'block';
  menu.style.left = e.pageX + 'px';
  menu.style.top = e.pageY + 'px';
});

document.addEventListener('click', () => {
  menu.style.display = 'none';
});

document.getElementById('ctx-copy').onclick = async () => { await copySelection(); };
document.getElementById('ctx-cut').onclick = async () => { await cutSelection(); };
document.getElementById('ctx-paste').onclick = async () => { await pasteFromClipboard(); };
document.getElementById('ctx-delete').onclick = () => { deleteSelection(); };

// ===== Keyboard shortcuts are in grid.js ===== //

// ===== Voice command hooks (optional) ===== //
window.__onVoiceCommand = function (cmd) {
  // Example: "move right 3"
  const m = cmd.match(/move\s+(right|left|up|down)\s+(\d+)/i);
  if (m) {
    const dir = m[1];
    const steps = parseInt(m[2], 10);
    moveActive(dir, steps);
  }
};

// ===== Status UI ===== //
window.__onActiveChange = function () {
  document.getElementById('status-cell').textContent = getActive();
};

// ===== New/Clear Grid Buttons ===== //
document.getElementById('btn-new-grid').onclick = () => {
  const r = parseInt(prompt('Rows?', '100'), 10);
  const c = parseInt(prompt('Cols?', '26'), 10);
  if (r && c) newGrid(gridHost, r, c);
};
document.getElementById('btn-clear-grid').onclick = () => clearGrid();
