export function exportToExcel(matrix) {
  // Convert matrix to worksheet
  const ws_data = matrix;
  const ws = XLSX.utils.aoa_to_sheet(ws_data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  XLSX.writeFile(wb, 'voice-sheet.xlsx');
}

export function exportToCsv(matrix) {
  const ws = XLSX.utils.aoa_to_sheet(matrix);
  const csv = XLSX.utils.sheet_to_csv(ws);
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'voice-sheet.csv';
  a.click();
  URL.revokeObjectURL(url);
}
