// Very lightweight persistence: localStorage for names, IndexedDB optional later.
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
