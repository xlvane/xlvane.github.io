// Demo: inspect an XLSX file in the browser with xlvane. Text from the file is only ever inserted
// with textContent, never as HTML.
import { createWorkbook, inspect, readWorkbook, toCSV, writeWorkbook } from '../vendor/xlvane/index.js';

const $ = (id) => document.getElementById(id);
const PREVIEW_ROWS = 50;
const PREVIEW_COLUMNS = 26;
let current = null;

function el(tag, text, className) {
  const node = document.createElement(tag);
  if (text !== undefined) node.textContent = text;
  if (className) node.className = className;
  return node;
}

function show(value) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return value.toISOString().replace('T', ' ').replace(/(\s00:00:00)?\.000Z$/, '');
  return String(value);
}

function setStatus(text) {
  $('status').textContent = text;
}

function showError(error) {
  const box = $('error');
  box.replaceChildren();
  const code = error && typeof error === 'object' && 'code' in error ? String(error.code) : 'Error';
  box.append(el('p', `${code}: ${error instanceof Error ? error.message : String(error)}`));
  if (error && typeof error === 'object' && 'hint' in error && error.hint) box.append(el('p', `Hint: ${String(error.hint)}`));
  box.hidden = false;
  $('result').hidden = true;
}

function open(bytes, name) {
  $('error').hidden = true;
  const started = performance.now();
  try {
    const report = inspect(bytes);
    const workbook = readWorkbook(bytes);
    current = { bytes, name, report, workbook };
    render();
    setStatus(`${name}: ${(bytes.byteLength / 1024).toFixed(1)} KiB, read in ${Math.round(performance.now() - started)} ms.`);
  } catch (error) {
    current = null;
    setStatus(`${name}: could not be read.`);
    showError(error);
  }
}

function render() {
  const { report } = current;
  const body = $('sheets').tBodies[0];
  body.replaceChildren();
  for (const s of report.sheets) {
    const tr = el('tr');
    tr.append(el('td', s.name), el('td', s.kind === 'worksheet' ? s.state : `${s.state} (${s.kind})`), el('td', s.usedRange ?? '-'), el('td', String(s.rowCount), 'num'), el('td', String(s.columnCount), 'num'), el('td', s.headerRow ? `${s.headerRow.row} (${s.headerRow.confidence})` : '-'));
    body.append(tr);
  }
  const warnings = $('warnings');
  warnings.replaceChildren();
  for (const w of report.warnings.slice(0, 20)) warnings.append(el('p', `${w.code}: ${w.message}`));
  if (report.warnings.length > 20) warnings.append(el('p', `… and ${report.warnings.length - 20} more warnings (see the report below).`));
  warnings.hidden = report.warnings.length === 0;

  const select = $('sheet');
  select.replaceChildren();
  for (const s of report.sheets.filter((x) => x.kind === 'worksheet')) {
    const option = el('option', s.name);
    option.value = s.name;
    select.append(option);
  }
  const json = JSON.stringify(report, null, 2);
  $('json').textContent = json.length > 200_000 ? `${json.slice(0, 200_000)}\n… (cut at 200 000 characters)` : json;
  $('result').hidden = false;
  renderSheet();
}

function renderSheet() {
  const name = $('sheet').value;
  const info = current.report.sheets.find((s) => s.name === name);
  const columns = $('columns').tBodies[0];
  columns.replaceChildren();
  const thead = $('preview').tHead;
  const tbody = $('preview').tBodies[0];
  thead.replaceChildren();
  tbody.replaceChildren();
  $('csv').disabled = !info;
  if (!info) {
    $('preview-note').textContent = 'This workbook has no worksheet with cells.';
    return;
  }
  for (const c of info.columns) {
    const tr = el('tr');
    tr.append(el('td', c.letter), el('td', c.header ?? ''), el('td', c.dominantType), el('td', c.samples.map(show).join(', ')));
    columns.append(tr);
  }
  const sheet = current.workbook.getSheet(name);
  const grid = new Map();
  let minRow = Infinity;
  let minCol = Infinity;
  for (const cell of sheet.cells()) {
    minRow = Math.min(minRow, cell.row);
    minCol = Math.min(minCol, cell.col);
  }
  if (minRow === Infinity) {
    $('preview-note').textContent = 'The sheet is empty.';
    return;
  }
  for (const cell of sheet.cells()) {
    if (cell.row >= minRow + PREVIEW_ROWS || cell.col >= minCol + PREVIEW_COLUMNS) continue;
    grid.set(`${cell.row}:${cell.col}`, cell);
  }
  const lastCol = Math.min(minCol + PREVIEW_COLUMNS - 1, minCol + Math.max(0, info.columnCount - 1));
  const lastRow = Math.min(minRow + PREVIEW_ROWS - 1, minRow + Math.max(0, info.rowCount - 1));
  const head = el('tr');
  head.append(el('th', ''));
  for (let c = minCol; c <= lastCol; c++) head.append(el('th', letter(c)));
  thead.append(head);
  for (let r = minRow; r <= lastRow; r++) {
    const tr = el('tr');
    tr.append(el('th', String(r)));
    for (let c = minCol; c <= lastCol; c++) {
      const cell = grid.get(`${r}:${c}`);
      const td = el('td', cell ? show(cell.value) : '', cell && typeof cell.value === 'number' ? 'num' : undefined);
      if (cell?.formula?.text) td.title = `=${cell.formula.text}`;
      tr.append(td);
    }
    tbody.append(tr);
  }
  $('preview-note').textContent = `Rows ${minRow} to ${lastRow} and columns ${letter(minCol)} to ${letter(lastCol)} of ${info.usedRange ?? ''}. Hover a cell to see its formula.`;
}

function letter(col) {
  let s = '';
  for (let n = col; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

function downloadCsv() {
  if (!current) return;
  const name = $('sheet').value;
  try {
    const csv = toCSV(current.bytes, { sheet: name });
    const url = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
    const a = document.createElement('a');
    a.href = url;
    a.download = `${current.name.replace(/\.[^.]+$/, '')}-${name}.csv`.replace(/[\\/:*?"<>|]/g, '_');
    document.body.append(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  } catch (error) {
    showError(error);
  }
}

function sample() {
  const wb = createWorkbook({ properties: { title: 'Sample', creator: 'xlvane demo' } });
  const sales = wb.addSheet('Sales');
  const head = { font: { bold: true }, fill: { color: 'D9E1F2' } };
  sales.set('A1', 'Date', head).set('B1', 'Region', head).set('C1', 'Product', head).set('D1', 'Units', head).set('E1', 'Revenue', head);
  const rows = [
    [Date.UTC(2026, 0, 5), 'North', 'Tea', 12, 30],
    [Date.UTC(2026, 0, 6), 'South', 'Coffee', 8, 36],
    [Date.UTC(2026, 0, 7), 'North', 'Coffee', 5, 22.5],
    [Date.UTC(2026, 0, 8), '東京', '緑茶', 20, 48],
  ];
  rows.forEach(([d, region, product, units, revenue], i) => {
    const r = i + 2;
    sales.set(`A${r}`, new Date(d)).set(`B${r}`, region).set(`C${r}`, product).set(`D${r}`, units).set(`E${r}`, revenue, { numFmt: '#,##0.00' });
  });
  sales.set('E6', { formula: 'SUM(E2:E5)', result: 136.5 }, { font: { bold: true }, numFmt: '#,##0.00' });
  sales.freeze({ rows: 1 }).autoFilter('A1:E5').column('A', { width: 12 });
  wb.addSheet('Notes', { state: 'hidden' }).set('A1', 'A hidden sheet');
  open(writeWorkbook(wb), 'sample.xlsx');
}

async function readFile(file) {
  if (!file) return;
  setStatus(`Reading ${file.name}…`);
  open(new Uint8Array(await file.arrayBuffer()), file.name);
}

$('file').addEventListener('change', (e) => readFile(e.target.files?.[0]));
$('sample').addEventListener('click', sample);
$('sheet').addEventListener('change', renderSheet);
$('csv').addEventListener('click', downloadCsv);
const drop = $('drop');
drop.addEventListener('dragover', (e) => {
  e.preventDefault();
  drop.classList.add('over');
});
drop.addEventListener('dragleave', () => drop.classList.remove('over'));
drop.addEventListener('drop', (e) => {
  e.preventDefault();
  drop.classList.remove('over');
  readFile(e.dataTransfer?.files?.[0]);
});
