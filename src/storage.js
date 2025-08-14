export function saveSheet(name, matrix) {
  if (!name) { alert('Please enter a sheet name first.'); return; }
  localStorage.setItem('sheet::' + name, JSON.stringify(matrix));
  alert('Saved "' + name + '".');
}
export async function loadSheet(name) {
  if (!name) { alert('Please enter a sheet name first.'); return null; }
  const raw = localStorage.getItem('sheet::' + name);
  return raw ? JSON.parse(raw) : null;
}

// History snapshots (keep last 10 per sheet)
function nowLabel(){
  const d = new Date();
  const pad = (n)=>String(n).padStart(2,'0');
  return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}
export function snapshotNow(sheetName, matrix){
  const name = sheetName || 'untitled';
  const key = `hist::${name}::${Date.now()}`;
  try{
    localStorage.setItem(key, JSON.stringify(matrix));
    const idxKey = `hist-index::${name}`;
    const idx = JSON.parse(localStorage.getItem(idxKey) || '[]');
    idx.unshift({ key, label: nowLabel() });
    const trimmed = idx.slice(0,10);
    localStorage.setItem(idxKey, JSON.stringify(trimmed));
    return true;
  }catch(e){
    alert('Failed to save snapshot (localStorage full?).');
    return false;
  }
}
export function listSnapshots(sheetName){
  const name = sheetName || 'untitled';
  const idxKey = `hist-index::${name}`;
  return JSON.parse(localStorage.getItem(idxKey) || '[]');
}
export async function restoreSnapshot(key){
  const raw = localStorage.getItem(key);
  return raw ? JSON.parse(raw) : null;
}
