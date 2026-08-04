/* ============================================================
   EDITOR DE OFICINA — documento (Word) + hoja de cálculo (Excel)
   ============================================================ */

const toast = document.getElementById('toast');
function showToast(msg){
  toast.textContent = msg;
  toast.classList.add('show');
  clearTimeout(showToast._t);
  showToast._t = setTimeout(()=> toast.classList.remove('show'), 2400);
}

/* ============================================================
   TABS
   ============================================================ */
document.querySelectorAll('.tab').forEach(tab=>{
  tab.addEventListener('click', ()=>{
    document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
    document.querySelectorAll('.view').forEach(v=>v.classList.remove('active'));
    tab.classList.add('active');
    document.getElementById('view-' + tab.dataset.view).classList.add('active');
  });
});

/* ============================================================
   DOCUMENTO — editor tipo Word
   ============================================================ */
const page = document.getElementById('page');
document.execCommand('defaultParagraphSeparator', false, 'p');

document.querySelectorAll('button.tool[data-cmd]').forEach(btn=>{
  btn.addEventListener('click', ()=>{
    page.focus();
    document.execCommand(btn.dataset.cmd, false, null);
    updateToolbarState();
  });
});
document.getElementById('fontName').addEventListener('change', e=>{
  page.focus(); document.execCommand('fontName', false, e.target.value);
});
document.getElementById('fontSize').addEventListener('change', e=>{
  page.focus(); document.execCommand('fontSize', false, e.target.value);
});
document.getElementById('textColor').addEventListener('input', e=>{
  page.focus(); document.execCommand('foreColor', false, e.target.value);
});
document.getElementById('hiliteColor').addEventListener('input', e=>{
  page.focus(); document.execCommand('hiliteColor', false, e.target.value);
});
document.getElementById('clearFormat').addEventListener('click', ()=>{
  page.focus(); document.execCommand('removeFormat', false, null);
});
document.getElementById('newPageBtn').addEventListener('click', ()=>{
  if(page.innerText.trim().length && !confirm('¿Borrar todo el contenido y empezar una página nueva?')) return;
  page.innerHTML = '';
  page.focus();
  updateWordCount();
});

// Pegado limpio (funciona igual con contenido copiado de Word, Copilot, webs, etc.)
page.addEventListener('paste', (e)=>{
  e.preventDefault();
  const html = (e.clipboardData || window.clipboardData).getData('text/html');
  const text = (e.clipboardData || window.clipboardData).getData('text/plain');
  if(html){
    const tmp = document.createElement('div');
    tmp.innerHTML = html;
    tmp.querySelectorAll('script,style,meta,link').forEach(n=>n.remove());
    tmp.querySelectorAll('*').forEach(n=>{
      n.removeAttribute('class');
      n.removeAttribute('id');
      ['margin','position','max-width','min-height'].forEach(p=>n.style.removeProperty(p));
    });
    document.execCommand('insertHTML', false, tmp.innerHTML);
  } else {
    document.execCommand('insertText', false, text);
  }
  updateWordCount();
});

function updateToolbarState(){
  document.querySelectorAll('button.tool[data-cmd]').forEach(btn=>{
    try{ btn.classList.toggle('active', document.queryCommandState(btn.dataset.cmd)); }catch(_){}
  });
}
page.addEventListener('keyup', ()=>{ updateToolbarState(); updateWordCount(); updateTableToolbar(); });
page.addEventListener('mouseup', ()=>{ updateToolbarState(); updateTableToolbar(); });
page.addEventListener('input', updateWordCount);

function updateWordCount(){
  const text = page.innerText.trim();
  const count = text.length ? text.split(/\s+/).length : 0;
  document.getElementById('wordCount').textContent = count + (count===1 ? ' palabra' : ' palabras');
}

page.addEventListener('keydown', (e)=>{
  if(e.ctrlKey || e.metaKey){
    const k = e.key.toLowerCase();
    if(k==='b'){ e.preventDefault(); document.execCommand('bold'); }
    if(k==='i'){ e.preventDefault(); document.execCommand('italic'); }
    if(k==='u'){ e.preventDefault(); document.execCommand('underline'); }
  }
});

/* ----- Insertar tabla (modal) ----- */
const tableModal = document.getElementById('tableModal');
document.getElementById('insertTableBtn').addEventListener('click', ()=>{
  tableModal.classList.remove('hidden');
});
document.getElementById('tableModalCancel').addEventListener('click', ()=>{
  tableModal.classList.add('hidden');
});
document.getElementById('tableModalOk').addEventListener('click', ()=>{
  const rows = Math.max(1, Math.min(30, parseInt(document.getElementById('tRows').value) || 3));
  const cols = Math.max(1, Math.min(12, parseInt(document.getElementById('tCols').value) || 3));
  insertTable(rows, cols);
  tableModal.classList.add('hidden');
});

function insertTable(rows, cols){
  page.focus();
  let html = '<table><tbody>';
  for(let r=0;r<rows;r++){
    html += '<tr>';
    for(let c=0;c<cols;c++){ html += '<td>&nbsp;</td>'; }
    html += '</tr>';
  }
  html += '</tbody></table><p><br></p>';
  document.execCommand('insertHTML', false, html);
}

/* ----- Barra flotante contextual para tablas ----- */
const tableToolbar = document.getElementById('tableToolbar');
let activeTable = null, activeCell = null;

function updateTableToolbar(){
  const sel = window.getSelection();
  if(!sel.rangeCount){ hideTableToolbar(); return; }
  let node = sel.anchorNode;
  if(node && node.nodeType === 3) node = node.parentElement;
  const cell = node ? node.closest('td, th') : null;
  const table = node ? node.closest('table') : null;
  if(cell && table && page.contains(table)){
    activeCell = cell; activeTable = table;
    const rect = table.getBoundingClientRect();
    tableToolbar.style.top = Math.max(8, rect.top - 44) + 'px';
    tableToolbar.style.left = Math.min(window.innerWidth - 320, Math.max(8, rect.left)) + 'px';
    tableToolbar.classList.remove('hidden');
  } else {
    hideTableToolbar();
  }
}
function hideTableToolbar(){
  tableToolbar.classList.add('hidden');
  activeTable = null; activeCell = null;
}
document.addEventListener('click', (e)=>{
  if(!page.contains(e.target) && !tableToolbar.contains(e.target)) hideTableToolbar();
});

tableToolbar.addEventListener('click', (e)=>{
  const action = e.target.dataset.action;
  if(!action || !activeCell || !activeTable) return;
  const row = activeCell.parentElement;
  const cellIndex = Array.from(row.children).indexOf(activeCell);

  if(action === 'rowAbove' || action === 'rowBelow'){
    const newRow = document.createElement('tr');
    for(let i=0;i<row.children.length;i++){
      const td = document.createElement('td');
      td.innerHTML = '&nbsp;';
      newRow.appendChild(td);
    }
    if(action === 'rowAbove') row.parentElement.insertBefore(newRow, row);
    else row.parentElement.insertBefore(newRow, row.nextSibling);
  }
  if(action === 'colLeft' || action === 'colRight'){
    Array.from(activeTable.querySelectorAll('tr')).forEach(tr=>{
      const td = document.createElement('td');
      td.innerHTML = '&nbsp;';
      const refCell = tr.children[cellIndex];
      if(action === 'colLeft') tr.insertBefore(td, refCell);
      else tr.insertBefore(td, refCell.nextSibling);
    });
  }
  if(action === 'delRow'){
    const body = row.parentElement;
    row.remove();
    if(!body.children.length) activeTable.remove();
  }
  if(action === 'delCol'){
    Array.from(activeTable.querySelectorAll('tr')).forEach(tr=>{
      if(tr.children[cellIndex]) tr.children[cellIndex].remove();
    });
    if(!activeTable.querySelector('tr')?.children.length) activeTable.remove();
  }
  if(action === 'delTable'){
    activeTable.remove();
  }
  hideTableToolbar();
  page.focus();
  updateWordCount();
});

/* ----- Exportar documento como .doc (compatible con Word) ----- */
document.getElementById('downloadDocBtn').addEventListener('click', ()=>{
  const content = page.innerHTML;
  const htmlDoc = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
  <head>
    <meta charset="utf-8">
    <title>Documento</title>
    <!--[if gte mso 9]>
    <xml><w:WordDocument><w:View>Print</w:View><w:Zoom>100</w:Zoom><w:DoNotOptimizeForBrowser/></w:WordDocument></xml>
    <![endif]-->
    <style>
      body{ font-family: Calibri, Arial, sans-serif; font-size:12pt; }
      p{ margin:0 0 10pt 0; }
      table{ border-collapse:collapse; }
      td, th{ border:1px solid #999; padding:4px 8px; }
    </style>
  </head>
  <body>${content}</body>
  </html>`;
  downloadBlob(htmlDoc, 'application/msword', `documento-${today()}.doc`);
  showToast('Documento descargado ✓ Ábrelo con Word');
});

/* ============================================================
   HOJA DE CÁLCULO — tipo Excel con fórmulas básicas
   ============================================================ */
const sheetEl = document.getElementById('sheet');
const formulaInput = document.getElementById('formulaInput');
const cellRefEl = document.getElementById('cellRef');

let numCols = 8;
let numRows = 20;
let selectedCell = null;

function colLetter(index){ // 0 -> A, 25 -> Z, 26 -> AA ...
  let s = '';
  index++;
  while(index > 0){
    const rem = (index - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    index = Math.floor((index - 1) / 26);
  }
  return s;
}

function buildSheet(){
  sheetEl.innerHTML = '';
  const thead = document.createElement('thead');
  const headRow = document.createElement('tr');
  headRow.innerHTML = '<th class="corner"></th>';
  for(let c=0;c<numCols;c++){
    const th = document.createElement('th');
    th.textContent = colLetter(c);
    headRow.appendChild(th);
  }
  thead.appendChild(headRow);
  sheetEl.appendChild(thead);

  const tbody = document.createElement('tbody');
  for(let r=0;r<numRows;r++){
    const tr = document.createElement('tr');
    const th = document.createElement('th');
    th.textContent = r + 1;
    tr.appendChild(th);
    for(let c=0;c<numCols;c++){
      const td = document.createElement('td');
      td.dataset.cell = colLetter(c) + (r + 1);
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
  sheetEl.appendChild(tbody);
}
buildSheet();

/* ----- Selección de celdas y barra de fórmulas ----- */
sheetEl.addEventListener('click', (e)=>{
  const td = e.target.closest('td');
  if(!td) return;
  selectCell(td);
});
sheetEl.addEventListener('focusin', (e)=>{
  const td = e.target.closest('td');
  if(!td) return;
  selectCell(td, true);
});

function selectCell(td, editing){
  if(selectedCell) selectedCell.classList.remove('selected');
  selectedCell = td;
  td.classList.add('selected');
  cellRefEl.textContent = td.dataset.cell;
  formulaInput.value = td.dataset.formula || td.textContent;
  if(editing) revealFormula(td);
}
function revealFormula(td){
  if(td.dataset.formula) td.textContent = td.dataset.formula;
}

formulaInput.addEventListener('input', ()=>{
  if(!selectedCell) return;
  setCellRaw(selectedCell, formulaInput.value);
});
formulaInput.addEventListener('keydown', (e)=>{
  if(e.key === 'Enter'){ e.preventDefault(); recalcAll(); selectedCell && selectedCell.blur(); }
});

sheetEl.addEventListener('input', (e)=>{
  const td = e.target.closest('td');
  if(!td) return;
  td.dataset.formula = td.textContent.startsWith('=') ? td.textContent : '';
  if(td === selectedCell) formulaInput.value = td.textContent;
});
sheetEl.addEventListener('focusout', (e)=>{
  const td = e.target.closest('td');
  if(!td) return;
  const raw = td.dataset.formula || td.textContent.trim();
  if(raw.startsWith('=')){
    td.dataset.formula = raw;
  } else {
    delete td.dataset.formula;
  }
  recalcAll();
});

/* ----- Formato de texto en celdas ----- */
document.querySelectorAll('button[data-scmd]').forEach(btn=>{
  btn.addEventListener('click', ()=>{
    if(selectedCell) selectedCell.focus();
    document.execCommand(btn.dataset.scmd, false, null);
  });
});

/* ----- Filas / columnas ----- */
document.getElementById('addRowBtn').addEventListener('click', ()=>{ numRows++; rebuildKeepingData(); });
document.getElementById('addColBtn').addEventListener('click', ()=>{ numCols++; rebuildKeepingData(); });
document.getElementById('delRowBtn').addEventListener('click', ()=>{
  if(numRows > 1){ numRows--; rebuildKeepingData(); }
});
document.getElementById('delColBtn').addEventListener('click', ()=>{
  if(numCols > 1){ numCols--; rebuildKeepingData(); }
});
document.getElementById('clearSheetBtn').addEventListener('click', ()=>{
  if(!confirm('¿Borrar todo el contenido de la hoja?')) return;
  numCols = 8; numRows = 20;
  buildSheet();
  formulaInput.value = '';
  cellRefEl.textContent = '—';
  selectedCell = null;
});

function rebuildKeepingData(){
  // Guarda los valores/fórmulas actuales antes de reconstruir la tabla
  const data = {};
  sheetEl.querySelectorAll('td[data-cell]').forEach(td=>{
    const raw = td.dataset.formula || td.textContent;
    if(raw) data[td.dataset.cell] = raw;
  });
  buildSheet();
  Object.keys(data).forEach(ref=>{
    const td = sheetEl.querySelector(`[data-cell="${ref}"]`);
    if(td) setCellRaw(td, data[ref], false);
  });
  recalcAll();
}

function setCellRaw(td, raw, updateInput){
  if(raw.startsWith('=')){
    td.dataset.formula = raw;
    td.textContent = raw;
  } else {
    delete td.dataset.formula;
    td.textContent = raw;
  }
  if(updateInput !== false) formulaInput.value = raw;
}

/* ----- Motor de fórmulas: + - * / ( ), SUMA/SUM, PROMEDIO/AVERAGE, MAX, MIN, CONTAR/COUNT ----- */
function cellRefToRC(ref){
  const m = ref.match(/^([A-Z]+)(\d+)$/i);
  if(!m) return null;
  let col = 0;
  const letters = m[1].toUpperCase();
  for(let i=0;i<letters.length;i++){ col = col*26 + (letters.charCodeAt(i) - 64); }
  return { col: col - 1, row: parseInt(m[2]) - 1 };
}

function expandRange(a, b){
  const ra = cellRefToRC(a), rb = cellRefToRC(b);
  if(!ra || !rb) return [];
  const refs = [];
  for(let r = Math.min(ra.row, rb.row); r <= Math.max(ra.row, rb.row); r++){
    for(let c = Math.min(ra.col, rb.col); c <= Math.max(ra.col, rb.col); c++){
      refs.push(colLetter(c) + (r + 1));
    }
  }
  return refs;
}

const evalStack = new Set();

function getRawCellValue(ref){
  const td = sheetEl.querySelector(`[data-cell="${ref}"]`);
  if(!td) return 0;
  const raw = td.dataset.formula || td.textContent.trim();
  if(!raw) return 0;
  if(raw.startsWith('=')){
    if(evalStack.has(ref)) return '#CIRC!';
    evalStack.add(ref);
    const result = evalFormula(raw.slice(1));
    evalStack.delete(ref);
    return result;
  }
  const n = parseFloat(raw.replace(',', '.'));
  return isNaN(n) ? raw : n;
}

function evalFormula(expr){
  try{
    let e = expr.toUpperCase();

    // Funciones con rango: SUMA/SUM(A1:B3), PROMEDIO/AVERAGE, MAX, MIN, CONTAR/COUNT
    const funcMap = { SUMA:'SUM', SUM:'SUM', PROMEDIO:'AVERAGE', AVERAGE:'AVERAGE', MAX:'MAX', MIN:'MIN', CONTAR:'COUNT', COUNT:'COUNT' };
    const funcRegex = /(SUMA|SUM|PROMEDIO|AVERAGE|MAX|MIN|CONTAR|COUNT)\(([A-Z]+\d+):([A-Z]+\d+)\)/;
    let m;
    while((m = e.match(funcRegex))){
      const fn = funcMap[m[1]];
      const refs = expandRange(m[2], m[3]);
      const values = refs.map(getRawCellValue).map(v => typeof v === 'number' ? v : 0);
      let result = 0;
      if(fn === 'SUM') result = values.reduce((a,b)=>a+b, 0);
      if(fn === 'AVERAGE') result = values.length ? values.reduce((a,b)=>a+b,0) / values.length : 0;
      if(fn === 'MAX') result = values.length ? Math.max(...values) : 0;
      if(fn === 'MIN') result = values.length ? Math.min(...values) : 0;
      if(fn === 'COUNT') result = refs.filter(r => typeof getRawCellValue(r) === 'number').length;
      e = e.replace(m[0], String(result));
    }

    // Sustituir referencias de celda sueltas (A1, B12, etc.)
    e = e.replace(/[A-Z]+\d+/g, (ref)=>{
      const v = getRawCellValue(ref);
      return typeof v === 'number' ? v : 0;
    });

    // Validar que sólo queden números y operadores antes de evaluar
    if(!/^[\d\s+\-*/().]+$/.test(e)) return '#ERROR';
    // eslint-disable-next-line no-new-func
    const result = Function('"use strict"; return (' + e + ')')();
    return typeof result === 'number' && isFinite(result) ? Math.round(result * 10000) / 10000 : '#ERROR';
  }catch(_){
    return '#ERROR';
  }
}

function recalcAll(){
  sheetEl.querySelectorAll('td[data-cell]').forEach(td=>{
    if(td === document.activeElement) return; // no pisar lo que se está editando
    const raw = td.dataset.formula;
    if(raw && raw.startsWith('=')){
      evalStack.clear();
      const val = evalFormula(raw.slice(1));
      td.textContent = val;
    }
  });
}

/* ----- Exportar hoja como .xls (compatible con Excel) ----- */
document.getElementById('downloadXlsBtn').addEventListener('click', ()=>{
  recalcAll();
  let table = '<table border="1">';
  const rows = sheetEl.querySelectorAll('tbody tr');
  rows.forEach(tr=>{
    table += '<tr>';
    tr.querySelectorAll('td').forEach(td=>{
      const val = td.textContent.trim();
      table += `<td>${val}</td>`;
    });
    table += '</tr>';
  });
  table += '</table>';

  const htmlDoc = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
  <head>
    <meta charset="utf-8">
    <!--[if gte mso 9]>
    <xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet>
      <x:Name>Hoja1</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions>
    </x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml>
    <![endif]-->
  </head>
  <body>${table}</body>
  </html>`;
  downloadBlob(htmlDoc, 'application/vnd.ms-excel', `hoja-${today()}.xls`);
  showToast('Hoja de cálculo descargada ✓ Ábrela con Excel');
});

/* ============================================================
   Utilidades comunes
   ============================================================ */
function downloadBlob(content, mime, filename){
  const blob = new Blob(['\ufeff', content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
function today(){ return new Date().toISOString().slice(0,10); }

updateWordCount();
