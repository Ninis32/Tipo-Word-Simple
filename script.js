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

/* ----- Abrir un .docx existente para leerlo y editarlo ----- */
document.getElementById('openDocBtn').addEventListener('click', ()=>{
  document.getElementById('openDocInput').click();
});
document.getElementById('openDocInput').addEventListener('change', (e)=>{
  const file = e.target.files[0];
  if(!file) return;
  if(!file.name.toLowerCase().endsWith('.docx')){
    showToast('Solo se pueden abrir archivos .docx (Word moderno)');
    e.target.value = '';
    return;
  }
  const reader = new FileReader();
  reader.onload = function(ev){
    mammoth.convertToHtml({ arrayBuffer: ev.target.result })
      .then(result=>{
        page.innerHTML = result.value || '<p><br></p>';
        updateWordCount();
        showToast('Documento cargado ✓ ya puedes editarlo');
      })
      .catch(err=>{
        console.error(err);
        showToast('No se pudo leer el archivo. Asegúrate de que sea un .docx válido.');
      });
  };
  reader.readAsArrayBuffer(file);
  e.target.value = '';
});

/* ----- Abrir un PDF existente para leerlo y editarlo ----- */
if(window.pdfjsLib){
  pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
}

document.getElementById('openPdfBtn').addEventListener('click', ()=>{
  document.getElementById('openPdfInput').click();
});
document.getElementById('openPdfInput').addEventListener('change', (e)=>{
  const file = e.target.files[0];
  if(!file) return;
  const reader = new FileReader();
  reader.onload = async function(ev){
    try{
      const pdf = await pdfjsLib.getDocument({ data: ev.target.result }).promise;
      let html = '';
      for(let i = 1; i <= pdf.numPages; i++){
        const pdfPage = await pdf.getPage(i);
        const content = await pdfPage.getTextContent();
        let lines = [];
        let currentLine = '';
        let lastY = null;
        content.items.forEach(item=>{
          if(lastY !== null && Math.abs(item.transform[5] - lastY) > 2){
            if(currentLine.trim()) lines.push(currentLine.trim());
            currentLine = '';
          }
          currentLine += item.str + ' ';
          lastY = item.transform[5];
        });
        if(currentLine.trim()) lines.push(currentLine.trim());
        html += lines.map(l => `<p>${escapeHtml(l)}</p>`).join('');
        if(i < pdf.numPages) html += '<p style="page-break-after:always;"><br></p>';
      }
      page.innerHTML = html || '<p><br></p>';
      updateWordCount();
      showToast('PDF cargado ✓ el texto ya es editable (el diseño original no se conserva)');
    }catch(err){
      console.error(err);
      showToast('No se pudo leer el PDF. Puede estar escaneado o protegido.');
    }
  };
  reader.readAsArrayBuffer(file);
  e.target.value = '';
});

function escapeHtml(str){
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

/* ----- Descargar el documento como PDF ----- */
document.getElementById('downloadPdfBtn').addEventListener('click', ()=>{
  showToast('Generando PDF…');
  const opt = {
    margin: 15,
    filename: `documento-${today()}.pdf`,
    image: { type: 'jpeg', quality: 0.98 },
    html2canvas: { scale: 2, useCORS: true },
    jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' },
    pagebreak: { mode: ['css', 'legacy'] }
  };
  html2pdf().set(opt).from(page).save();
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

/* ----- Abrir un Excel/CSV existente para leerlo y editarlo ----- */
document.getElementById('openXlsBtn').addEventListener('click', ()=>{
  document.getElementById('openXlsInput').click();
});
document.getElementById('openXlsInput').addEventListener('change', (e)=>{
  const file = e.target.files[0];
  if(!file) return;
  const reader = new FileReader();
  reader.onload = function(ev){
    try{
      const data = new Uint8Array(ev.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const ws = workbook.Sheets[firstSheetName];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      loadDataIntoSheet(rows);
      showToast(`Hoja "${firstSheetName}" cargada ✓ ya puedes editarla`);
    }catch(err){
      console.error(err);
      showToast('No se pudo leer el archivo. Verifica que sea .xlsx, .xls o .csv válido.');
    }
  };
  reader.readAsArrayBuffer(file);
  e.target.value = '';
});

function loadDataIntoSheet(rows){
  const maxCols = rows.reduce((max, r) => Math.max(max, r.length), 1);
  numRows = Math.max(rows.length, 20);
  numCols = Math.max(maxCols, 8);
  buildSheet();
  rows.forEach((row, r)=>{
    row.forEach((val, c)=>{
      if(val === '' || val === undefined || val === null) return;
      const ref = colLetter(c) + (r + 1);
      const td = sheetEl.querySelector(`[data-cell="${ref}"]`);
      if(td) td.textContent = String(val);
    });
  });
  formulaInput.value = '';
  cellRefEl.textContent = '—';
  selectedCell = null;
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

/* ============================================================
   VER / ANOTAR PDF — visor fiel al original con anotaciones
   ============================================================ */
(function(){
  const openBtn = document.getElementById('openPdfViewBtn');
  const input = document.getElementById('openPdfViewInput');
  const emptyState = document.getElementById('pdfEmptyState');
  const stage = document.getElementById('pdfStage');
  const baseCanvas = document.getElementById('pdfBaseCanvas');
  const annotCanvas = document.getElementById('pdfAnnotCanvas');
  const baseCtx = baseCanvas.getContext('2d');
  const annotCtx = annotCanvas.getContext('2d');
  const pageInfo = document.getElementById('pdfPageInfo');
  const zoomInfo = document.getElementById('zoomInfo');

  let pdfDoc = null;
  let currentPage = 1;
  let scale = 1.3;
  let currentTool = 'pen';
  let drawing = false;
  let currentStroke = null;
  // annotations[pageNum] = array de { tool, color, width, points:[{x,y}...] } o { tool:'text', x,y,text,color }
  const annotations = {};
  const history = []; // número de página de cada acción, para deshacer

  if(!window.pdfjsLib){ return; }
  pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

  openBtn.addEventListener('click', ()=> input.click());
  input.addEventListener('change', async (e)=>{
    const file = e.target.files[0];
    if(!file) return;
    try{
      const buf = await file.arrayBuffer();
      pdfDoc = await pdfjsLib.getDocument({ data: buf }).promise;
      currentPage = 1;
      Object.keys(annotations).forEach(k => delete annotations[k]);
      emptyState.style.display = 'none';
      stage.style.display = 'flex';
      await renderPage();
      showToast(`PDF cargado ✓ ${pdfDoc.numPages} página(s)`);
    }catch(err){
      console.error(err);
      showToast('No se pudo abrir el PDF.');
    }
    e.target.value = '';
  });

  async function renderPage(){
    if(!pdfDoc) return;
    const p = await pdfDoc.getPage(currentPage);
    const viewport = p.getViewport({ scale });
    baseCanvas.width = annotCanvas.width = viewport.width;
    baseCanvas.height = annotCanvas.height = viewport.height;
    await p.render({ canvasContext: baseCtx, viewport }).promise;
    redrawAnnotations();
    pageInfo.textContent = `${currentPage} / ${pdfDoc.numPages}`;
    zoomInfo.textContent = Math.round(scale / 1.3 * 100) + '%';
  }

  function redrawAnnotations(){
    annotCtx.clearRect(0, 0, annotCanvas.width, annotCanvas.height);
    const list = annotations[currentPage] || [];
    list.forEach(a => drawAnnotation(a));
  }

  function drawAnnotation(a){
    if(a.tool === 'text'){
      annotCtx.fillStyle = a.color;
      annotCtx.font = `${18 * (scale/1.3)}px 'Segoe UI', Arial, sans-serif`;
      annotCtx.fillText(a.text, a.x * scale, a.y * scale);
      return;
    }
    if(!a.points || a.points.length < 2) return;
    annotCtx.beginPath();
    annotCtx.moveTo(a.points[0].x * scale, a.points[0].y * scale);
    for(let i=1;i<a.points.length;i++){ annotCtx.lineTo(a.points[i].x * scale, a.points[i].y * scale); }
    annotCtx.strokeStyle = a.color;
    annotCtx.lineWidth = a.width * (scale/1.3);
    annotCtx.lineCap = 'round';
    annotCtx.lineJoin = 'round';
    annotCtx.globalAlpha = a.tool === 'highlighter' ? 0.35 : 1;
    annotCtx.stroke();
    annotCtx.globalAlpha = 1;
  }

  /* ----- Navegación ----- */
  document.getElementById('prevPageBtn').addEventListener('click', ()=>{
    if(!pdfDoc || currentPage <= 1) return;
    currentPage--; renderPage();
  });
  document.getElementById('nextPageBtn').addEventListener('click', ()=>{
    if(!pdfDoc || currentPage >= pdfDoc.numPages) return;
    currentPage++; renderPage();
  });
  document.getElementById('zoomInBtn').addEventListener('click', ()=>{
    if(!pdfDoc) return;
    scale = Math.min(scale + 0.2, 3.5); renderPage();
  });
  document.getElementById('zoomOutBtn').addEventListener('click', ()=>{
    if(!pdfDoc) return;
    scale = Math.max(scale - 0.2, 0.6); renderPage();
  });

  /* ----- Herramientas ----- */
  document.querySelectorAll('.annot-tool').forEach(btn=>{
    btn.addEventListener('click', ()=>{
      document.querySelectorAll('.annot-tool').forEach(b=>b.classList.remove('active'));
      btn.classList.add('active');
      currentTool = btn.dataset.tool;
    });
  });

  function getPos(e){
    const rect = annotCanvas.getBoundingClientRect();
    const clientX = e.touches ? e.touches[0].clientX : e.clientX;
    const clientY = e.touches ? e.touches[0].clientY : e.clientY;
    return { x: (clientX - rect.left) / scale, y: (clientY - rect.top) / scale };
  }

  annotCanvas.addEventListener('pointerdown', (e)=>{
    if(!pdfDoc) return;
    const pos = getPos(e);

    if(currentTool === 'text'){
      const text = prompt('Texto a insertar:');
      if(text && text.trim()){
        pushAnnotation({ tool:'text', x: pos.x, y: pos.y, text: text.trim(), color: document.getElementById('annotColor').value });
      }
      return;
    }
    if(currentTool === 'eraser'){
      eraseNear(pos);
      return;
    }
    drawing = true;
    currentStroke = {
      tool: currentTool,
      color: document.getElementById('annotColor').value,
      width: parseInt(document.getElementById('annotWidth').value),
      points: [pos]
    };
  });
  annotCanvas.addEventListener('pointermove', (e)=>{
    if(!drawing || !currentStroke) return;
    currentStroke.points.push(getPos(e));
    redrawAnnotations();
    drawAnnotation(currentStroke);
  });
  annotCanvas.addEventListener('pointerup', ()=>{
    if(drawing && currentStroke){ pushAnnotation(currentStroke); }
    drawing = false; currentStroke = null;
  });
  annotCanvas.addEventListener('pointerleave', ()=>{
    if(drawing && currentStroke){ pushAnnotation(currentStroke); }
    drawing = false; currentStroke = null;
  });

  function pushAnnotation(a){
    if(!annotations[currentPage]) annotations[currentPage] = [];
    annotations[currentPage].push(a);
    history.push(currentPage);
    redrawAnnotations();
  }

  function eraseNear(pos){
    const list = annotations[currentPage] || [];
    const threshold = 14;
    for(let i = list.length - 1; i >= 0; i--){
      const a = list[i];
      let hit = false;
      if(a.tool === 'text'){
        hit = Math.hypot(a.x - pos.x, a.y - pos.y) < threshold;
      } else {
        hit = a.points.some(p => Math.hypot(p.x - pos.x, p.y - pos.y) < threshold);
      }
      if(hit){ list.splice(i, 1); redrawAnnotations(); return; }
    }
  }

  document.getElementById('undoAnnotBtn').addEventListener('click', ()=>{
    for(let i = history.length - 1; i >= 0; i--){
      if(history[i] === currentPage){
        history.splice(i, 1);
        if(annotations[currentPage] && annotations[currentPage].length){
          annotations[currentPage].pop();
        }
        redrawAnnotations();
        return;
      }
    }
  });
  document.getElementById('clearAnnotBtn').addEventListener('click', ()=>{
    if(!confirm('¿Borrar todas las anotaciones de esta página?')) return;
    annotations[currentPage] = [];
    redrawAnnotations();
  });

  /* ----- Descargar PDF anotado ----- */
  document.getElementById('downloadAnnotPdfBtn').addEventListener('click', async ()=>{
    if(!pdfDoc){ showToast('Primero abre un PDF.'); return; }
    if(!window.jspdf){ showToast('No se pudo cargar el generador de PDF.'); return; }
    showToast('Generando PDF anotado…');
    const { jsPDF } = window.jspdf;
    const exportScale = 2; // buena calidad de imagen
    let outPdf = null;

    for(let n = 1; n <= pdfDoc.numPages; n++){
      const p = await pdfDoc.getPage(n);
      const viewport = p.getViewport({ scale: exportScale });
      const tmpCanvas = document.createElement('canvas');
      tmpCanvas.width = viewport.width;
      tmpCanvas.height = viewport.height;
      const tmpCtx = tmpCanvas.getContext('2d');
      await p.render({ canvasContext: tmpCtx, viewport }).promise;

      // dibuja las anotaciones de esta página encima, escaladas a exportScale
      const list = annotations[n] || [];
      list.forEach(a=>{
        if(a.tool === 'text'){
          tmpCtx.fillStyle = a.color;
          tmpCtx.font = `${18 * exportScale}px 'Segoe UI', Arial, sans-serif`;
          tmpCtx.fillText(a.text, a.x * exportScale, a.y * exportScale);
          return;
        }
        if(!a.points || a.points.length < 2) return;
        tmpCtx.beginPath();
        tmpCtx.moveTo(a.points[0].x * exportScale, a.points[0].y * exportScale);
        for(let i=1;i<a.points.length;i++){ tmpCtx.lineTo(a.points[i].x * exportScale, a.points[i].y * exportScale); }
        tmpCtx.strokeStyle = a.color;
        tmpCtx.lineWidth = a.width * exportScale;
        tmpCtx.lineCap = 'round';
        tmpCtx.lineJoin = 'round';
        tmpCtx.globalAlpha = a.tool === 'highlighter' ? 0.35 : 1;
        tmpCtx.stroke();
        tmpCtx.globalAlpha = 1;
      });

      const imgData = tmpCanvas.toDataURL('image/jpeg', 0.95);
      const pageWidthMm = (viewport.width / exportScale) * 0.264583;
      const pageHeightMm = (viewport.height / exportScale) * 0.264583;

      if(!outPdf){
        outPdf = new jsPDF({ unit: 'mm', format: [pageWidthMm, pageHeightMm], orientation: pageWidthMm > pageHeightMm ? 'landscape' : 'portrait' });
      } else {
        outPdf.addPage([pageWidthMm, pageHeightMm], pageWidthMm > pageHeightMm ? 'landscape' : 'portrait');
      }
      outPdf.addImage(imgData, 'JPEG', 0, 0, pageWidthMm, pageHeightMm);
    }

    outPdf.save(`pdf-anotado-${today()}.pdf`);
    showToast('PDF anotado descargado ✓');
  });
})();

