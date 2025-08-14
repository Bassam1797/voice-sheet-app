function stamp() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, '0');
  return (
    d.getFullYear() +
    pad(d.getMonth() + 1) +
    pad(d.getDate()) + '-' +
    pad(d.getHours()) +
    pad(d.getMinutes()) +
    pad(d.getSeconds())
  );
}
function safeBase(name) {
  const base = (name || 'voice-sheet').trim() || 'voice-sheet';
  return base.replace(/[^a-z0-9._-]+/gi, '_').slice(0, 60);
}

export function exportToExcel(matrix, sheetName) {
  const ws = XLSX.utils.aoa_to_sheet(matrix);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  const fname = `${safeBase(sheetName)}_${stamp()}.xlsx`;
  XLSX.writeFile(wb, fname);
}

export function exportToCsv(matrix, sheetName) {
  const ws = XLSX.utils.aoa_to_sheet(matrix);
  const csv = XLSX.utils.sheet_to_csv(ws);
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `${safeBase(sheetName)}_${stamp()}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

export async function importFromFile(file){
  const buf = await file.arrayBuffer();
  let data = [];
  if (file.name.toLowerCase().endsWith('.csv')) {
    const text = new TextDecoder('utf-8').decode(new Uint8Array(buf));
    data = text.split(/\r?\n/).map(line => line.split(','));
  } else {
    const wb = XLSX.read(buf, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    data = XLSX.utils.sheet_to_json(ws, { header: 1 });
  }
  const maxLen = Math.max(1, ...data.map(r => r.length));
  data = data.map(r => {
    const row = Array.from(r);
    while (row.length < maxLen) row.push('');
    return row;
  });
  return data;
}
