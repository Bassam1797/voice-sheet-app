// ===== Export helpers =====
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

function getGridData() {
  const rows = document.querySelectorAll(".row");
  return Array.from(rows).map(row =>
    Array.from(row.querySelectorAll(".cell input")).map(cell => cell.value)
  );
}

// ===== Import helpers =====
async function importFromFile(accept, parserFn) {
  const file = document.createElement('input');
  file.type = 'file';
  file.accept = accept;
  return new Promise(resolve => {
    file.onchange = async () => {
      const f = file.files[0];
      if (!f) return resolve();
      const arrayBuffer = await f.arrayBuffer();
      parserFn(f.name, arrayBuffer);
      resolve();
    };
    file.click();
  });
}

function fillGridFromMatrix(matrix) {
  for (let r = 0; r < matrix.length; r++) {
    for (let c = 0; c < matrix[r].length; c++) {
      const cell = document.querySelector(`.cell[data-r="${r+1}"][data-c="${c+1}"] input`);
      if (cell) cell.value = matrix[r][c] ?? '';
    }
  }
}

function importExcelParser(filename, arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: 'array' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const matrix = XLSX.utils.sheet_to_json(ws, { header: 1 });
  fillGridFromMatrix(matrix);
}

function importCSVParser(filename, arrayBuffer) {
  const text = new TextDecoder().decode(arrayBuffer);
  const matrix = text.split(/\r?\n/).map(line => line.split(','));
  fillGridFromMatrix(matrix);
}

function importTSVParser(filename, arrayBuffer) {
  const text = new TextDecoder().decode(arrayBuffer);
  const matrix = text.split(/\r?\n/).map(line => line.split('\t'));
  fillGridFromMatrix(matrix);
}

// ===== Wire dropdown events =====
// Export
document.getElementById('export-xlsx').onclick = (e) => { e.preventDefault(); exportToExcel(); };
document.getElementById('export-csv').onclick  = (e) => { e.preventDefault(); exportToCSV(); };
document.getElementById('export-tsv').onclick  = (e) => { e.preventDefault(); exportToTSV(); };

// Import
document.getElementById('import-xlsx').onclick = (e) => { e.preventDefault(); importFromFile('.xlsx', importExcelParser); };
document.getElementById('import-csv').onclick  = (e) => { e.preventDefault(); importFromFile('.csv', importCSVParser); };
document.getElementById('import-tsv').onclick  = (e) => { e.preventDefault(); importFromFile('.tsv,.txt', importTSVParser); };
