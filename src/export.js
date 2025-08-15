import { toTSV } from './utils.js';

export function setupExportImport(grid){
  // Export
  document.getElementById('export-tsv').addEventListener('click', ()=>{
    const tsv = toTSV(grid.data);
    downloadBlob(new Blob([tsv], {type:'text/tab-separated-values'}), 'voice-sheet.tsv');
  });
  document.getElementById('export-csv').addEventListener('click', ()=>{
    const ws = XLSX.utils.aoa_to_sheet(grid.data);
    const csv = XLSX.utils.sheet_to_csv(ws);
    downloadBlob(new Blob([csv], {type:'text/csv'}), 'voice-sheet.csv');
  });
  document.getElementById('export-xlsx').addEventListener('click', ()=>{
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(grid.data);
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    const out = XLSX.write(wb, {bookType:'xlsx', type:'array'});
    downloadBlob(new Blob([out], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}), 'voice-sheet.xlsx');
  });

  // Import
  const importers = [
    ['file-import-xlsx', handleXLSX],
    ['file-import-csv',  handleCSV],
    ['file-import-tsv',  handleTSV],
  ];
  importers.forEach(([id,handler])=>{
    const el = document.getElementById(id);
    el.addEventListener('change', async (e)=>{
      const file = e.target.files?.[0]; if (!file) return;
      const arr = await file.arrayBuffer();
      await handler(arr, grid);
      e.target.value='';
    });
  });
}

function handleXLSX(buf, grid){
  const wb = XLSX.read(buf, {type:'array'});
  const ws = wb.Sheets[wb.SheetNames[0]];
  const aoa = XLSX.utils.sheet_to_json(ws, {header:1, raw:true});
  grid.data = normalizeRect(aoa);
}

function handleCSV(buf, grid){
  const text = new TextDecoder().decode(new Uint8Array(buf));
  const wb = XLSX.read(text, {type:'string'});
  const ws = wb.Sheets[wb.SheetNames[0]];
  const aoa = XLSX.utils.sheet_to_json(ws, {header:1, raw:true});
  grid.data = normalizeRect(aoa);
}

function handleTSV(buf, grid){
  const text = new TextDecoder().decode(new Uint8Array(buf));
  const rows = text.replace(/\r/g,'').split('\n').map(r=>r.split('\t'));
  grid.data = normalizeRect(rows);
}

function normalizeRect(aoa){
  const cols = Math.max(...aoa.map(r => r.length));
  return aoa.map(r => Array.from({length: cols}, (_,i)=> r[i] ?? ''));
}

function downloadBlob(blob, filename){
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);
}
