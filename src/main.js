// =======================
// Voice Sheet App - main.js
// =======================

// Track active cell coordinates
let activeCell = null;

// ===== GRID INITIALIZATION =====
function initGrid(rows = 20, cols = 10) {
  const container = document.getElementById("grid-container");
  container.innerHTML = "";
  for (let r = 1; r <= rows; r++) {
    const row = document.createElement("div");
    row.className = "row";
    for (let c = 1; c <= cols; c++) {
      const cell = document.createElement("div");
      cell.className = "cell";
      cell.dataset.r = r;
      cell.dataset.c = c;

      const input = document.createElement("input");
      input.type = "text";
      input.addEventListener("focus", () => setActiveCell(cell));
      cell.appendChild(input);

      row.appendChild(cell);
    }
    container.appendChild(row);
  }
  activeCell = null;
}

// ===== ACTIVE CELL & BLUE DOT =====
function setActiveCell(cell) {
  // Remove all blue dots
  document.querySelectorAll(".cell").forEach(c => c.classList.remove("active-cell"));
  // Mark this one
  cell.classList.add("active-cell");
  activeCell = cell;
  const colLetter = String.fromCharCode(64 + parseInt(cell.dataset.c));
  document.getElementById("status-cell").textContent = `${colLetter}${cell.dataset.r}`;
}

// ===== ROW/COL OPS =====
function insertRowAbove() {
  if (!activeCell) return;
  const r = parseInt(activeCell.dataset.r);
  const container = document.getElementById("grid-container");
  const row = container.children[r - 1];
  const newRow = row.cloneNode(true);
  newRow.querySelectorAll("input").forEach(i => i.value = "");
  container.insertBefore(newRow, row);
  reindexGrid();
}

function insertRowBelow() {
  if (!activeCell) return;
  const r = parseInt(activeCell.dataset.r);
  const container = document.getElementById("grid-container");
  const row = container.children[r - 1];
  const newRow = row.cloneNode(true);
  newRow.querySelectorAll("input").forEach(i => i.value = "");
  container.insertBefore(newRow, row.nextSibling);
  reindexGrid();
}

function deleteRow() {
  if (!activeCell) return;
  const r = parseInt(activeCell.dataset.r);
  const container = document.getElementById("grid-container");
  if (container.children.length > 1) {
    container.removeChild(container.children[r - 1]);
    reindexGrid();
  }
}

function insertColLeft() {
  if (!activeCell) return;
  const c = parseInt(activeCell.dataset.c);
  document.querySelectorAll(".row").forEach(row => {
    const cell = row.children[c - 1];
    const newCell = cell.cloneNode(true);
    newCell.querySelector("input").value = "";
    row.insertBefore(newCell, cell);
  });
  reindexGrid();
}

function insertColRight() {
  if (!activeCell) return;
  const c = parseInt(activeCell.dataset.c);
  document.querySelectorAll(".row").forEach(row => {
    const cell = row.children[c - 1];
    const newCell = cell.cloneNode(true);
    newCell.querySelector("input").value = "";
    row.insertBefore(newCell, cell.nextSibling);
  });
  reindexGrid();
}

function deleteCol() {
  if (!activeCell) return;
  const c = parseInt(activeCell.dataset.c);
  document.querySelectorAll(".row").forEach(row => {
    if (row.children.length > 1) {
      row.removeChild(row.children[c - 1]);
    }
  });
  reindexGrid();
}

// ===== REINDEX GRID =====
function reindexGrid() {
  const rows = document.querySelectorAll(".row");
  rows.forEach((row, ri) => {
    const cells = row.querySelectorAll(".cell");
    cells.forEach((cell, ci) => {
      cell.dataset.r = ri + 1;
      cell.dataset.c = ci + 1;
      cell.querySelector("input").onfocus = () => setActiveCell(cell);
    });
  });
}

// ===== IMPORT/EXPORT HELPERS =====
function getGridData() {
  const rows = document.querySelectorAll(".row");
  return Array.from(rows).map(row =>
    Array.from(row.querySelectorAll(".cell input")).map(cell => cell.value)
  );
}

function fillGridFromMatrix(matrix, startRow = 1, startCol = 1) {
  for (let r = 0; r < matrix.length; r++) {
    for (let c = 0; c < matrix[r].length; c++) {
      const cell = document.querySelector(`.cell[data-r="${startRow + r}"][data-c="${startCol + c}"] input`);
      if (cell) cell.value = matrix[r][c] ?? '';
    }
  }
}

function exportToExcel() {
  const data = getGridData();
  const ws = XLSX.utils.aoa_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  XLSX.writeFile(wb, "grid.xlsx");
}

function exportToCSV() {
  const data = getGridData();
  const ws = XLSX.utils.aoa_to_sheet(data);
  const csv = XLSX.utils.sheet_to_csv(ws);
  downloadFile("grid.csv", csv);
}

function exportToTSV() {
  const data = getGridData();
  const ws = XLSX.utils.aoa_to_sheet(data);
  const tsv = XLSX.utils.sheet_to_csv(ws, { FS: "\t" });
  downloadFile("grid.tsv", tsv);
}

function downloadFile(filename, content) {
  const blob = new Blob([content], { type: "text/plain;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

async function importFromFile(accept, parserFn) {
  const file = document.createElement('input');
  file.type = 'file';
  file.accept = accept;
  return new Promise(resolve => {
    file.onchange = async () => {
      const f = file.files[0];
      if (!f) return resolve();
      const arrayBuffer = await f.arrayBuffer();
      const startR = activeCell ? parseInt(activeCell.dataset.r) : 1;
      const startC = activeCell ? parseInt(activeCell.dataset.c) : 1;
      parserFn(arrayBuffer, startR, startC);
      resolve();
    };
    file.click();
  });
}

function importExcelParser(arrayBuffer, startR, startC) {
  const wb = XLSX.read(arrayBuffer, { type: 'array' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const matrix = XLSX.utils.sheet_to_json(ws, { header: 1 });
  fillGridFromMatrix(matrix, startR, startC);
}

function importCSVParser(arrayBuffer, startR, startC) {
  const text = new TextDecoder().decode(arrayBuffer);
  const matrix = text.split(/\r?\n/).map(line => line.split(','));
  fillGridFromMatrix(matrix, startR, startC);
}

function importTSVParser(arrayBuffer, startR, startC) {
  const text = new TextDecoder().decode(arrayBuffer);
  const matrix = text.split(/\r?\n/).map(line => line.split('\t'));
  fillGridFromMatrix(matrix, startR, startC);
}

// ===== EVENT WIRING =====
// Export
document.getElementById('export-xlsx').onclick = (e) => { e.preventDefault(); exportToExcel(); };
document.getElementById('export-csv').onclick  = (e) => { e.preventDefault(); exportToCSV(); };
document.getElementById('export-tsv').onclick  = (e) => { e.preventDefault(); exportToTSV(); };

// Import
document.getElementById('import-xlsx').onclick = (e) => { e.preventDefault(); importFromFile('.xlsx', importExcelParser); };
document.getElementById('import-csv').onclick  = (e) => { e.preventDefault(); importFromFile('.csv', importCSVParser); };
document.getElementById('import-tsv').onclick  = (e) => { e.preventDefault(); importFromFile('.tsv,.txt', importTSVParser); };

// ===== INIT =====
initGrid();
