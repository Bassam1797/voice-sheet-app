import { letters, clamp, copyToClipboard, readFromClipboard, parseTSV, toTSV } from './utils.js';

export function createGrid(container, opts = {}){
  const rows = opts.rows ?? 50;
  const cols = opts.cols ?? 26;
  const state = {
    rows, cols,
    data: Array.from({length: rows}, ()=> Array(cols).fill('')),
    merges: [], // {r1,c1,r2,c2}
    colWidths: Array(cols).fill(100),
    selection: { r1:0, c1:0, r2:0, c2:0 },
    active: { r:0, c:0 },
    undo: [], redo: [],
    isMouseDown: false, selStart: null,
  };

  const gridEl = container;
  gridEl.innerHTML = '';
  gridEl.classList.add('grid');

  const table = document.createElement('table');
  table.className = 'table';
  const thead = document.createElement('thead');
  const tbody = document.createElement('tbody');
  table.appendChild(thead); table.appendChild(tbody);
  gridEl.appendChild(table);

  // Corner + column headers
  const headRow = document.createElement('tr');
  const corner = document.createElement('th');
  corner.className = 'corner';
  corner.textContent = '';
  headRow.appendChild(corner);
  for (let c=0;c<state.cols;c++){
    const th = document.createElement('th');
    th.className = 'th-col';
    th.dataset.c = c;
    th.style.width = state.colWidths[c]+'px';
    th.textContent = letters(c);
    // resizer
    const resizer = document.createElement('div');
    resizer.className = 'col-resizer';
    let startX=0, startW=0;
    resizer.addEventListener('mousedown', (e)=>{
      startX = e.clientX; startW = th.offsetWidth;
      const move = (ev)=>{
        const dx = ev.clientX - startX;
        const w = Math.max(24, startW + dx); // allow smaller widths
        th.style.width = w+'px';
        state.colWidths[c] = w;
        api.emit('colWidthChange', api.getColumnWidths());
      };
      const up = ()=>{ document.removeEventListener('mousemove', move); document.removeEventListener('mouseup', up); };
      document.addEventListener('mousemove', move);
      document.addEventListener('mouseup', up);
      e.stopPropagation();
      e.preventDefault();
    });
    th.appendChild(resizer);
    headRow.appendChild(th);
  }
  thead.appendChild(headRow);

  // Build body
  for (let r=0;r<state.rows;r++){
    const tr = document.createElement('tr');
    // row header
    const th = document.createElement('th');
    th.className = 'th-row';
    th.textContent = (r+1);
    tr.appendChild(th);
    for (let c=0;c<state.cols;c++){
      const td = document.createElement('td');
      td.className = 'cell';
      td.dataset.r = r; td.dataset.c = c;
      td.contentEditable = 'true';
      td.spellcheck = false;
      td.addEventListener('input', handleInput);
      td.addEventListener('focus', handleFocus);
      td.addEventListener('pointerdown', handlePointerDown);
      td.addEventListener('keydown', handleKeyDown);
      td.appendChild(document.createElement('div')).className='active-dot';
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }

  // Selection rectangle + fill handle
  const selRect = document.createElement('div');
  selRect.className = 'selection-rect';
  selRect.style.display = 'none';
  gridEl.appendChild(selRect);
  const fillHandle = document.createElement('div');
  fillHandle.className = 'fill-handle';
  selRect.appendChild(fillHandle);

  // Context menu
  const menu = buildContextMenu();
  document.body.appendChild(menu);

  function buildContextMenu(){
    const menu = document.createElement('div');
    menu.style.position='fixed'; menu.style.display='none'; menu.style.zIndex='1000';
    menu.style.background='white'; menu.style.border='1px solid #e5e7eb'; menu.style.borderRadius='8px';
    menu.style.boxShadow='0 8px 24px rgba(0,0,0,.08)';
    const addBtn=(label,fn)=>{
      const b=document.createElement('button'); b.textContent=label; b.style.display='block'; b.style.padding='8px 12px'; b.style.width='160px'; b.style.border='none'; b.style.background='white'; b.style.textAlign='left';
      b.addEventListener('click', ()=>{ fn(); hide(); });
      b.addEventListener('mouseover', ()=> b.style.background='#f3f4f6');
      b.addEventListener('mouseout', ()=> b.style.background='white');
      menu.appendChild(b);
    };
    const hide=()=> menu.style.display='none';
    addBtn('Copy', copySelection);
    addBtn('Cut', cutSelection);
    addBtn('Paste', async ()=> pasteTSV(await readFromClipboard()));
    addBtn('Delete', clearSelection);
    addBtn('Insert Row Above', ()=> insertRows(state.selection.r1,1));
    addBtn('Insert Row Below', ()=> insertRows(state.selection.r2+1,1));
    addBtn('Insert Column Left', ()=> insertCols(state.selection.c1,1));
    addBtn('Insert Column Right', ()=> insertCols(state.selection.c2+1,1));
    addBtn('Delete Rows', ()=> deleteRows(state.selection.r1, state.selection.r2-state.selection.r1+1));
    addBtn('Delete Columns', ()=> deleteCols(state.selection.c1, state.selection.c2-state.selection.c1+1));
    addBtn('Merge Cells', mergeSelection);
    addBtn('Unmerge Cells', unmergeSelection);
    window.addEventListener('click', hide);
    menu.hide = hide;
    return menu;
  }

  gridEl.addEventListener('contextmenu', (e)=>{
    e.preventDefault();
    const menu = document.body.lastChild;
    menu.style.left = e.clientX + 'px';
    menu.style.top = e.clientY + 'px';
    menu.style.display = 'block';
  });

  function getCell(r,c){ return tbody.children[r].children[c+1]; } // +1 for row header

  function handleInput(e){
    const td = e.currentTarget;
    const r = +td.dataset.r, c = +td.dataset.c;
    pushUndo([{r,c, old: state.data[r][c], val: td.textContent}]);
    state.data[r][c] = td.textContent;
  }

  function handleFocus(e){
    const td=e.currentTarget; const r=+td.dataset.r, c=+td.dataset.c;
    setActive(r,c);
  }

  function handlePointerDown(e){
    if (e.button!==0) return;
    const td=e.currentTarget; const r=+td.dataset.r, c=+td.dataset.c;
    state.isMouseDown=true; state.selStart={r,c};
    setSelection(r,c,r,c);
    document.addEventListener('pointermove', handlePointerMove);
    document.addEventListener('pointerup', handlePointerUp, {once:true});
  }

  function handlePointerMove(e){
    if (!state.isMouseDown) return;
    const el = document.elementFromPoint(e.clientX, e.clientY);
    if (el && el.classList.contains('cell')){
      const r=+el.dataset.r, c=+el.dataset.c;
      setSelection(state.selStart.r, state.selStart.c, r, c);
    }
  }
  function handlePointerUp(){ state.isMouseDown=false; document.removeEventListener('pointermove', handlePointerMove); }

  function setSelection(r1,c1,r2,c2){
    state.selection = { r1:Math.min(r1,r2), c1:Math.min(c1,c2), r2:Math.max(r1,r2), c2:Math.max(c1,c2) };
    updateSelectionVisuals();
  }
  function setActive(r,c){
    state.active = { r, c };
    // move active class
    gridEl.querySelectorAll('.cell.active').forEach(el=>el.classList.remove('active'));
    getCell(r,c).classList.add('active');
  }

  function updateSelectionVisuals(){
    const {r1,c1,r2,c2} = state.selection;
    // highlight range
    gridEl.querySelectorAll('.cell.selected').forEach(el=>el.classList.remove('selected'));
    for(let r=r1;r<=r2;r++) for(let c=c1;c<=c2;c++) getCell(r,c).classList.add('selected');
    // selection rect coords
    const a = getCell(r1,c1).getBoundingClientRect();
    const b = getCell(r2,c2).getBoundingClientRect();
    const g = gridEl.getBoundingClientRect();
    selRect.style.display = 'block';
    selRect.style.left = (a.left - g.left + gridEl.scrollLeft) + 'px';
    selRect.style.top = (a.top - g.top + gridEl.scrollTop) + 'px';
    selRect.style.width = (b.right - a.left) + 'px';
    selRect.style.height = (b.bottom - a.top) + 'px';
    // ensure last active cell shows dot only on last active
    gridEl.querySelectorAll('.active-dot').forEach(dot => dot.style.display='none');
    getCell(state.active.r, state.active.c).querySelector('.active-dot').style.display='block';
  }

  // Keyboard handling + navigation
  function handleKeyDown(e){
    const {r1,c1,r2,c2} = state.selection;
    // Enter/Tab navigation: respect boundaries of selection when multi-selected
    const moveWithin=(dir)=>{
      if (r1===r2 && c1===c2){
        moveSelection(dir);
      } else {
        // if selection is multi, move within bounds and wrap
        const {r,c} = state.active;
        let nr=r, nc=c;
        if (dir==='right'){ nc = c+1; if (nc>c2){ nc=c1; nr = Math.min(r+1, r2); } }
        if (dir==='left'){ nc = c-1; if (nc<c1){ nc=c2; nr = Math.max(r-1, r1); } }
        if (dir==='down'){ nr = r+1; if (nr>r2){ nr=r1; nc = Math.min(c+1, c2); } }
        if (dir==='up'){ nr = r-1; if (nr<r1){ nr=r2; nc = Math.max(c-1, c1); } }
        setActive(nr??r, nc??c); setSelection(nr??r, nc??c, nr??r, nc??c);
        getCell(nr,nc).focus();
      }
    };
    if (e.key==='Enter'){ e.preventDefault(); if (e.shiftKey) moveWithin('up'); else moveWithin('down'); }
    if (e.key==='Tab'){ e.preventDefault(); if (e.shiftKey) moveWithin('left'); else moveWithin('right'); }
    if (e.key==='ArrowRight'){ e.preventDefault(); moveWithin('right'); }
    if (e.key==='ArrowLeft'){ e.preventDefault(); moveWithin('left'); }
    if (e.key==='ArrowDown'){ e.preventDefault(); moveWithin('down'); }
    if (e.key==='ArrowUp'){ e.preventDefault(); moveWithin('up'); }
    if ((e.ctrlKey||e.metaKey) && e.key.toLowerCase()==='c'){ e.preventDefault(); copySelection(); }
    if ((e.ctrlKey||e.metaKey) && e.key.toLowerCase()==='x'){ e.preventDefault(); cutSelection(); }
    if ((e.ctrlKey||e.metaKey) && e.key.toLowerCase()==='v'){ e.preventDefault(); pasteFromClipboard(); }
    if ((e.ctrlKey||e.metaKey) && e.key.toLowerCase()==='z'){ e.preventDefault(); if (e.shiftKey) redo(); else undo(); }
    if ((e.ctrlKey||e.metaKey) && (e.key.toLowerCase()==='y')){ e.preventDefault(); redo(); }
    if (e.key==='Delete' || e.key==='Backspace'){ e.preventDefault(); clearSelection(); }
  }

  async function pasteFromClipboard(){ pasteTSV(await readFromClipboard()); }

  function pushUndo(changes){
    // changes: [{r,c, old, val}]
    state.undo.push(changes);
    state.redo.length = 0;
  }

  function applyChanges(changes, {recordUndo=false}={}){
    const inverse = [];
    for (const ch of changes){
      const td = getCell(ch.r, ch.c);
      const old = state.data[ch.r][ch.c];
      if (recordUndo) inverse.push({r:ch.r, c:ch.c, old: ch.val, val: old});
      state.data[ch.r][ch.c] = ch.val;
      td.textContent = ch.val;
    }
    if (recordUndo && inverse.length) state.undo.push(inverse);
  }

  function undo(){
    const changes = state.undo.pop();
    if (!changes) return;
    const inverse = changes.map(({r,c, old, val})=>({r,c, old:val, val:old}));
    applyChanges(changes);
    state.redo.push(inverse);
  }
  function redo(){
    const changes = state.redo.pop();
    if (!changes) return;
    const inverse = changes.map(({r,c, old, val})=>({r,c, old:val, val:old}));
    applyChanges(changes);
    state.undo.push(inverse);
  }

  function copySelection(){
    const {r1,c1,r2,c2} = state.selection;
    const rows=[];
    for(let r=r1;r<=r2;r++){
      const row=[];
      for(let c=c1;c<=c2;c++) row.push(state.data[r][c] ?? '');
      rows.push(row);
    }
    copyToClipboard(toTSV(rows));
  }
  function cutSelection(){
    const {r1,c1,r2,c2} = state.selection;
    const changes=[];
    for(let r=r1;r<=r2;r++) for(let c=c1;c<=c2;c++){
      changes.push({r,c, old: state.data[r][c], val:''});
    }
    pushUndo(changes);
    applyChanges(changes);
    copySelection();
  }
  function clearSelection(){
    const {r1,c1,r2,c2} = state.selection;
    const changes=[];
    for(let r=r1;r<=r2;r++) for(let c=c1;c<=c2;c++){
      if (state.data[r][c]!=='') changes.push({r,c, old: state.data[r][c], val:''});
    }
    if (changes.length){ pushUndo(changes); applyChanges(changes); }
  }
  function pasteTSV(tsv){
    if (!tsv) return;
    const data = parseTSV(tsv);
    const {r1,c1} = state.selection;
    const changes=[];
    for(let i=0;i<data.length;i++){
      for(let j=0;j<data[i].length;j++){
        const r=r1+i, c=c1+j; if (r>=state.rows || c>=state.cols) continue;
        changes.push({r,c, old: state.data[r][c], val: data[i][j]});
      }
    }
    pushUndo(changes); applyChanges(changes);
  }

  // Drag-fill (including numeric series)
  let isFilling=false, fillStart=null;
  fillHandle.addEventListener('mousedown', (e)=>{
    isFilling=true; fillStart={...state.selection};
    document.addEventListener('mousemove', onFillMove);
    document.addEventListener('mouseup', onFillEnd, {once:true});
    e.preventDefault(); e.stopPropagation();
  });
  function onFillMove(e){
    const el = document.elementFromPoint(e.clientX, e.clientY);
    if (el && el.classList.contains('cell')){
      const r=+el.dataset.r, c=+el.dataset.c;
      // Extend selection rectangle to here
      setSelection(fillStart.r1, fillStart.c1, r, c);
    }
  }
  function onFillEnd(){
    document.removeEventListener('mousemove', onFillMove);
    isFilling=false;
    const src = fillStart;
    const dst = state.selection;
    // determine fill direction and size
    const changes=[];
    const srcRows = src.r2 - src.r1 + 1;
    const srcCols = src.c2 - src.c1 + 1;

    // collect source values
    const srcVals = [];
    for(let r=src.r1;r<=src.r2;r++){
      const row=[]; for(let c=src.c1;c<=src.c2;c++) row.push(state.data[r][c] ?? ''); srcVals.push(row);
    }

    for(let r=dst.r1;r<=dst.r2;r++){
      for(let c=dst.c1;c<=dst.c2;c++){
        const rr = (r - dst.r1) % srcRows;
        const cc = (c - dst.c1) % srcCols;
        let base = srcVals[rr][cc];
        // numeric series if expanding beyond source
        if ((dst.r2-dst.r1+1) > srcRows || (dst.c2-dst.c1+1) > srcCols){
          // try parse number
          const num = parseFloat(base);
          if (!isNaN(num)){
            // compute offset based on total steps from original position
            const stepsR = Math.floor((r - src.r1) / srcRows);
            const stepsC = Math.floor((c - src.c1) / srcCols);
            const step = Math.max(stepsR, stepsC);
            base = String(num + step);
          }
        }
        changes.push({r,c, old: state.data[r][c], val: base});
      }
    }
    pushUndo(changes); applyChanges(changes);
    // restore selection to dst
    setSelection(dst.r1,dst.c1,dst.r2,dst.c2);
  }

  // Row/Column ops
  function insertRows(at, count){
    const newRows = Array.from({length: count}, ()=> Array(state.cols).fill(''));
    state.data.splice(at, 0, ...newRows);
    // rebuild body for simplicity
    rebuild();
  }
  function insertCols(at, count){
    for (let r=0;r<state.rows;r++) state.data[r].splice(at, 0, ...Array(count).fill(''));
    state.colWidths.splice(at, 0, ...Array(count).fill(100));
    rebuild();
  }
  function deleteRows(at, count){
    state.data.splice(at, count);
    rebuild();
  }
  function deleteCols(at, count){
    for (let r=0;r<state.rows;r++) state.data[r].splice(at, count);
    state.colWidths.splice(at, count);
    rebuild();
  }

  function mergeSelection(){
    const {r1,c1,r2,c2} = state.selection;
    if (r1===r2 && c1===c2) return;
    state.merges.push({r1,c1,r2,c2});
    rebuild();
  }
  function unmergeSelection(){
    const {r1,c1,r2,c2} = state.selection;
    state.merges = state.merges.filter(m => !(m.r1>=r1 && m.c1>=c1 && m.r2<=r2 && m.c2<=c2));
    rebuild();
  }

  function rebuild(){
    // Recreate table body and headers to reflect rows/cols/merges/widths
    while (thead.firstChild) thead.removeChild(thead.firstChild);
    while (tbody.firstChild) tbody.removeChild(tbody.firstChild);
    // headers
    const headRow = document.createElement('tr');
    const corner = document.createElement('th'); corner.className='corner'; headRow.appendChild(corner);
    for(let c=0;c<state.data[0].length;c++){
      const th = document.createElement('th'); th.className='th-col'; th.dataset.c=c; th.style.width=(state.colWidths[c]||100)+'px'; th.textContent=letters(c);
      const resizer = document.createElement('div'); resizer.className='col-resizer';
      let startX=0, startW=0;
      resizer.addEventListener('mousedown', (e)=>{
        startX = e.clientX; startW = th.offsetWidth;
        const move = (ev)=>{
          const dx = ev.clientX - startX; const w = Math.max(24, startW + dx);
          th.style.width = w+'px'; state.colWidths[c]=w; api.emit('colWidthChange', api.getColumnWidths());
        };
        const up = ()=>{ document.removeEventListener('mousemove', move); document.removeEventListener('mouseup', up); };
        document.addEventListener('mousemove', move); document.addEventListener('mouseup', up);
        e.stopPropagation(); e.preventDefault();
      });
      th.appendChild(resizer);
      headRow.appendChild(th);
    }
    thead.appendChild(headRow);
    state.rows = state.data.length; state.cols = state.data[0].length;
    // body
    for (let r=0;r<state.rows;r++){
      const tr = document.createElement('tr');
      const th = document.createElement('th'); th.className='th-row'; th.textContent=(r+1); tr.appendChild(th);
      for (let c=0;c<state.cols;c++){
        const td = document.createElement('td');
        td.className='cell'; td.dataset.r=r; td.dataset.c=c; td.contentEditable='true'; td.spellcheck=false;
        td.appendChild(document.createElement('div')).className='active-dot';
        td.textContent = state.data[r][c] ?? '';
        td.addEventListener('input', handleInput);
        td.addEventListener('focus', handleFocus);
        td.addEventListener('pointerdown', handlePointerDown);
        td.addEventListener('keydown', handleKeyDown);
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    }
    // apply merges
    applyMerges();
    // restore selection + active (clamp if needed)
    setActive(clamp(state.active.r,0,state.rows-1), clamp(state.active.c,0,state.cols-1));
    setSelection(0,0,0,0);
    updateSelectionVisuals();
  }

  function applyMerges(){
    // Clear all spans first
    for (let r=0;r<state.rows;r++){
      for(let c=0;c<state.cols;c++){
        const td=getCell(r,c);
        td.style.display='';
        td.removeAttribute('rowspan'); td.removeAttribute('colspan');
      }
    }
    // Apply merges
    for (const m of state.merges){
      const master = getCell(m.r1, m.c1);
      master.setAttribute('rowspan', (m.r2-m.r1+1));
      master.setAttribute('colspan', (m.c2-m.c1+1));
      for (let r=m.r1;r<=m.r2;r++){
        for (let c=m.c1;c<=m.c2;c++){
          if (r===m.r1 && c===m.c1) continue;
          getCell(r,c).style.display='none';
        }
      }
    }
  }

  function autoFitColumn(c){
    let max = 40; // min
    for (let r=0;r<state.rows;r++){
      const val = String(state.data[r][c] ?? '');
      max = Math.max(max, 10 + val.length * 8);
    }
    state.colWidths[c] = max;
    const th = thead.querySelector(`.th-col[data-c="${c}"]`);
    if (th) th.style.width = max + 'px';
    api.emit('colWidthChange', api.getColumnWidths());
  }

  const api = {
    on(event, fn){ (this._ev||(this._ev={}))[event] = (this._ev[event]||[]).concat(fn); },
    emit(event, payload){ (this._ev?.[event]||[]).forEach(fn=>fn(payload)); },
    get data(){ return state.data; },
    set data(d){ state.data = d; rebuild(); },
    clear(){ const ch=[]; for(let r=0;r<state.rows;r++)for(let c=0;c<state.cols;c++) if(state.data[r][c]!=='') ch.push({r,c,old:state.data[r][c], val:''}); pushUndo(ch); applyChanges(ch); },
    reset(r,c){ state.data = Array.from({length:r},()=>Array(c).fill('')); state.merges=[]; state.colWidths=Array(c).fill(100); rebuild(); },
    setColumnWidths(w){ state.colWidths = w.slice(0, state.cols); rebuild(); },
    getColumnWidths(){ return state.colWidths.slice(); },
    moveSelection(dir){ moveSelection(dir); },
    editCurrentCell(val, {pushUndo:true}={}){
      const {r,c}=state.active;
      const ch={r,c, old: state.data[r][c], val};
      if (pushUndo) pushUndo([ch]);
      applyChanges([ch]);
      getCell(r,c).focus();
    },
    exportRange(){
      const {r1,c1,r2,c2}=state.selection;
      const out=[];
      for(let r=r1;r<=r2;r++){ const row=[]; for(let c=c1;c<=c2;c++) row.push(state.data[r][c]??''); out.push(row); }
      return out;
    },
    importArray2D(arr, startR=0, startC=0){
      const changes=[];
      for(let i=0;i<arr.length;i++){
        for(let j=0;j<arr[i].length;j++){
          const r=startR+i, c=startC+j;
          if (r>=state.rows || c>=state.cols) continue;
          changes.push({r,c, old: state.data[r][c], val: arr[i][j]});
        }
      }
      pushUndo(changes); applyChanges(changes);
    },
    autoFitColumn,
  };

  function moveSelection(dir){
    const {r,c} = state.active;
    let nr=r, nc=c;
    if (dir==='right') nc = clamp(c+1, 0, state.cols-1);
    if (dir==='left')  nc = clamp(c-1, 0, state.cols-1);
    if (dir==='down')  nr = clamp(r+1, 0, state.rows-1);
    if (dir==='up')    nr = clamp(r-1, 0, state.rows-1);
    setActive(nr,nc); setSelection(nr,nc,nr,nc);
    getCell(nr,nc).focus();
  }

  // Double-click column header to auto-fit
  thead.addEventListener('dblclick', (e)=>{
    const th = e.target.closest('.th-col'); if (!th) return;
    autoFitColumn(+th.dataset.c);
  });

  // Initial focus
  setActive(0,0); setSelection(0,0,0,0); updateSelectionVisuals();

  return api;
}
