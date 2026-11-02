// Demo: inspect an XLSX file in the browser with xlvane. Text from the file is only ever inserted
// with textContent (or set as a property), never as HTML, and it is cut to a few hundred
// characters, so a file with huge strings cannot freeze the page while it is laid out.
import { createWorkbook, DEFAULT_READ_LIMITS, inspect, readWorkbook, toCSV, writeWorkbook } from '../vendor/xlvane/index.js';

const $ = (id) => document.getElementById(id);
const PREVIEW_ROWS = 50;
const PREVIEW_COLUMNS = 26;
/** Longest text shown in one table cell or tooltip. */
const MAX_TEXT = 200;
/** Sheets and columns per sheet kept in the JSON shown on the page. */
const JSON_SHEETS = 20;
const JSON_COLUMNS = 50;
/** Largest rectangle (rows × columns from A1 to the end of the used range) exported as CSV. */
const MAX_CSV_CELLS = 2_000_000;
let current = null;
/** Cells of the preview by "row:col", to read a formula only when the pointer is on its cell. */
let previewCells = new Map();

function el(tag, text, className) {
  const node = document.createElement(tag);
  if (text !== undefined) node.textContent = cut(text);
  if (className) node.className = className;
  return node;
}

function cut(text, max = MAX_TEXT) {
  return text.length > max ? `${text.slice(0, max)}…` : text;
}

function show(value) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return value.toISOString().replace('T', ' ').replace(/(\s00:00:00)?\.000Z$/, '');
  return cut(String(value));
}

function setStatus(text) {
  $('status').textContent = cut(text, 400);
}

function showError(error, keepResult = false) {
  const box = $('error');
  box.replaceChildren();
  const code = error && typeof error === 'object' && 'code' in error ? String(error.code) : 'Error';
  box.append(el('p', `${code}: ${error instanceof Error ? error.message : String(error)}`));
  // Hints about options (such as raising a limit) are for programs; the demo uses the defaults.
  if (error && typeof error === 'object' && 'hint' in error && error.hint && code !== 'XLV_E_LIMIT_EXCEEDED') box.append(el('p', `Hint: ${String(error.hint)}`));
  box.hidden = false;
  if (!keepResult) $('result').hidden = true;
}

/** Lets the browser paint the status line before the file is read on the main thread. */
const paint = () => new Promise((resolve) => requestAnimationFrame(() => setTimeout(resolve, 0)));

async function open(bytes, name) {
  $('error').hidden = true;
  setStatus(`Reading ${name}…`);
  await paint();
  const started = performance.now();
  try {
    // The file is read once; inspect() and toCSV() take the workbook.
    const workbook = readWorkbook(bytes);
    const report = inspect(workbook);
    current = { name, report, workbook };
    render();
    setStatus(`${name}: ${(bytes.byteLength / 1024).toFixed(1)} KiB, read in ${Math.round(performance.now() - started)} ms.`);
  } catch (error) {
    current = null;
    setStatus(`${name}: could not be read.`);
    showError(error);
  }
}

/** The report with at most JSON_SHEETS sheets and JSON_COLUMNS columns per sheet, for display. */
function shortReport(report) {
  const sheets = report.sheets.slice(0, JSON_SHEETS).map((s) => (s.columns.length > JSON_COLUMNS ? { ...s, columns: s.columns.slice(0, JSON_COLUMNS), columnsShownOnThisPage: JSON_COLUMNS } : s));
  return { ...report, sheets, warnings: report.warnings.slice(0, 100), ...(report.sheets.length > JSON_SHEETS ? { sheetsShownOnThisPage: JSON_SHEETS } : {}) };
}

function render() {
  const { report } = current;
  const body = $('sheets').tBodies[0];
  body.replaceChildren();
  for (const s of report.sheets.slice(0, 500)) {
    const tr = el('tr');
    tr.append(el('td', s.name), el('td', s.kind === 'worksheet' ? s.state : `${s.state} (${s.kind})`), el('td', s.usedRange ?? '-'), el('td', String(s.rowCount), 'num'), el('td', String(s.columnCount), 'num'), el('td', s.headerRow ? `${s.headerRow.row} (${s.headerRow.confidence})` : '-'));
    body.append(tr);
  }
  if (report.sheets.length > 500) body.append(el('tr', `… ${report.sheets.length - 500} more sheets`));
  const warnings = $('warnings');
  warnings.replaceChildren();
  for (const w of report.warnings.slice(0, 20)) warnings.append(el('p', `${w.code}: ${w.message}`));
  if (report.warnings.length > 20) warnings.append(el('p', `… and ${report.warnings.length - 20} more warnings.`));
  warnings.hidden = report.warnings.length === 0;

  // Options carry the sheet's position: two sheets of a damaged file may have the same name.
  const select = $('sheet');
  select.replaceChildren();
  for (const s of report.sheets.slice(0, 500)) {
    if (s.kind !== 'worksheet') continue;
    const option = el('option', s.name);
    option.value = String(s.index);
    select.append(option);
  }
  // Long strings (headers, formulas, names) are cut while the JSON is built, not afterwards.
  const json = JSON.stringify(shortReport(report), (_key, value) => (typeof value === 'string' ? cut(value, 500) : value), 2);
  $('json').textContent = json.length > 200_000 ? `${json.slice(0, 200_000)}\n… (cut at 200 000 characters)` : json;
  $('result').hidden = false;
  renderSheet();
}

function renderSheet() {
  const index = Number($('sheet').value);
  const info = current.report.sheets.find((s) => s.index === index);
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
  for (const c of info.columns.slice(0, 200)) {
    const tr = el('tr');
    tr.append(el('td', c.letter), el('td', c.header ?? ''), el('td', c.dominantType), el('td', c.samples.map(show).join(', ')));
    columns.append(tr);
  }
  const sheet = current.workbook.sheets[index];
  const grid = new Map();
  previewCells = grid;
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
      // The formula text is read on hover: a shared formula is expanded for each cell that reads it.
      if (cell?.formula) td.dataset.cell = `${r}:${c}`;
      tr.append(td);
    }
    tbody.append(tr);
  }
  $('preview-note').textContent = cut(`Rows ${minRow} to ${lastRow} and columns ${letter(minCol)} to ${letter(lastCol)} of ${info.usedRange ?? ''}. Hover a cell to see its formula; long text is cut at ${MAX_TEXT} characters.`, 400);
}

function letter(col) {
  let s = '';
  for (let n = col; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

function downloadCsv() {
  if (!current) return;
  const index = Number($('sheet').value);
  const sheet = current.workbook.sheets[index];
  const info = current.report.sheets.find((s) => s.index === index);
  if (!sheet || !info) return;
  // toCSV writes every position from A1 to the end of the used range.
  const end = /([A-Z]+)(\d+)$/.exec(info.usedRange ?? '');
  const area = end ? Number(end[2]) * end[1].split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) : 0;
  if (area > MAX_CSV_CELLS) {
    showError(new Error(`The sheet spans ${area.toLocaleString('en')} cells from A1 to the end of its used range; the demo exports at most ${MAX_CSV_CELLS.toLocaleString('en')}.`), true);
    return;
  }
  if (current.workbook.sheets.filter((s) => s.name === sheet.name).length > 1) {
    showError(new Error(`The workbook has more than one sheet named "${cut(sheet.name, 100)}", so the demo cannot export this one by name.`), true);
    return;
  }
  $('error').hidden = true;
  try {
    const csv = toCSV(current.workbook, { sheet: sheet.name });
    const url = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
    const a = document.createElement('a');
    a.href = url;
    a.download = cut(`${current.name.replace(/\.[^.]+$/, '')}-${sheet.name}`, 120).replace(/[\\/:*?"<>|\u0000-\u001f\u007f-\u009f]/g, '_') + '.csv';
    document.body.append(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  } catch (error) {
    showError(error, true);
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
  void open(writeWorkbook(wb), 'sample.xlsx');
}

async function readFile(file) {
  if (!file) return;
  $('error').hidden = true;
  // The library refuses input above maxInputBytes; checking first avoids loading such a file into memory.
  if (file.size > DEFAULT_READ_LIMITS.maxInputBytes) {
    setStatus(`${file.name}: could not be read.`);
    showError(new Error(`The file is ${(file.size / 1048576).toFixed(0)} MiB; the demo reads files up to ${(DEFAULT_READ_LIMITS.maxInputBytes / 1048576).toFixed(0)} MiB.`));
    return;
  }
  setStatus(`Reading ${file.name}…`);
  let bytes;
  try {
    bytes = new Uint8Array(await file.arrayBuffer());
  } catch (error) {
    setStatus(`${file.name}: could not be read.`);
    showError(error);
    return;
  }
  await open(bytes, file.name);
}

$('file').addEventListener('change', (e) => void readFile(e.target.files?.[0]));
$('sample').addEventListener('click', sample);
$('sheet').addEventListener('change', renderSheet);
$('preview').addEventListener('mouseover', (e) => {
  const td = e.target instanceof Element ? e.target.closest('td[data-cell]') : null;
  if (!td || td.title) return;
  const cell = previewCells.get(td.dataset.cell);
  const text = cell?.formula?.text;
  if (text) td.title = cut(`=${text}`, 500);
});
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
  void readFile(e.dataTransfer?.files?.[0]);
});
