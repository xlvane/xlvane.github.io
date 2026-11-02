import { $ as fillXml, At as EMPTY_STYLES, B as tokenize, Bt as suggest, Cn as escapeText, Ct as relationshipsPartName, En as XlvError, Ft as normalizeStyle, G as DEFAULT_FONT, H as X14_DV_URI, I as WarningCollector, It as stableKey, J as borderXml, Lt as isDateFormat, M as openZip, Mt as READING_ORDERS, Nt as checkStyleCount, Pt as mergeStyles, Rt as DEFAULT_READ_LIMITS, Sn as escapeAttr, St as openPackage, Tt as zipNameToPartName, V as X14_CF_URI, _ as isoDate, an as columnLetters, at as reusableBuiltinNumFmt, b as readWorkbook, bn as parseXml, bt as parseIsoDateTime, cn as parseCellRef, en as SML_NS, et as fontXml, f as createWorkbook, ft as stringItemXml, gn as parseXmlTree, h as detectHeader, hn as childrenNamed, i as readOptions, in as MAX_ROWS, jt as parseStyles, ln as parseRange, m as writeWorkbook, mn as childNamed, mt as textElement, n as oneOfOption, nn as relationshipTypes, on as formatCellRef, q as attr, r as openInput, rn as MAX_COLUMNS, sn as formatRange, t as booleanOption, tn as classifyRelationship, tt as formatNumber, vn as decodeXml, vt as writeZip, wt as partKey, x as toBytes, xn as encodeXstring, xt as parseSharedStrings, yn as normalizeLineEnds, z as shiftFormula, zt as resolveLimits } from "./options-CwSp2pA4.js";
// src/table.ts
const READ_OPTIONS = [
	"sheet",
	"table",
	"header",
	"range",
	"schema",
	"dates",
	"skipEmptyRows",
	"limits",
	"onWarning"
];
const HEADER_SCAN_ROWS = 20;
/** The sheet to read: by name, or the first visible worksheet (or the first sheet). */
function pickSheet(workbook, name) {
	if (name !== void 0) {
		if (typeof name !== "string") throw new XlvError("XLV_E_INVALID_OPTION", "The option \"sheet\" must be a sheet name.", {
			hint: `Available sheets: ${workbook.sheets.map((s) => s.name).join(", ")}.`,
			details: { option: "sheet" }
		});
		return workbook.getSheet(name);
	}
	const sheet = workbook.sheets.find((s) => s.kind === "worksheet" && s.state === "visible") ?? workbook.sheets[0];
	if (sheet === void 0) throw new XlvError("XLV_E_SHEET_NOT_FOUND", "The workbook has no sheets.", { hint: "Pass a workbook that contains at least one worksheet." });
	return sheet;
}
function checkSchema(schema) {
	if (schema === void 0) return void 0;
	const props = typeof schema === "object" || typeof schema === "function" ? schema?.["~standard"] : void 0;
	if (typeof props !== "object" || props === null || props.version !== 1 || typeof props.validate !== "function") throw new XlvError("XLV_E_SCHEMA_INVALID", "The option \"schema\" is not a Standard Schema (version 1).", {
		hint: "Pass a schema from a library that implements Standard Schema, such as Zod 3.24+, Valibot 1.0+ or ArkType 2.0+ (for example z.object({ ... })).",
		details: { option: "schema" }
	});
	return schema;
}
/** Sets an own enumerable data property, also for keys such as `__proto__`. */
function setOwn(target, key, value) {
	Object.defineProperty(target, key, {
		value,
		enumerable: true,
		writable: true,
		configurable: true
	});
}
function cellValue(cell, dates) {
	if (cell === void 0) return null;
	const v = cell.value;
	if (v instanceof Date) return dates === "iso" ? isoDate(v) : v;
	return v;
}
function headerText(cell) {
	const v = cell?.value;
	if (v === void 0 || v === null) return null;
	if (v instanceof Date) return isoDate(v);
	return String(v);
}
/**
* The sheet, range (header and data rows) and header row of the table named by the option
* `table`, or `null` without it (TD-12).
*/
function pickTable(workbook, o) {
	const name = o.get("table");
	if (name === void 0) return null;
	if (typeof name !== "string") throw new XlvError("XLV_E_INVALID_OPTION", "The option \"table\" must be a table name.", {
		hint: "For example { table: 'Sales' }.",
		details: { option: "table" }
	});
	for (const key of [
		"sheet",
		"header",
		"range"
	]) if (o.get(key) !== void 0) throw new XlvError("XLV_E_INVALID_OPTION", `The options "table" and "${key}" cannot be used together.`, {
		hint: "A table already fixes its sheet, range and header row; leave the other option out.",
		details: { option: key }
	});
	const all = workbook.sheets.flatMap((sheet) => sheet.tables.map((t) => ({
		sheet,
		t
	})));
	const lower = name.toLowerCase();
	const found = all.find(({ t }) => t.displayName.toLowerCase() === lower || t.name.toLowerCase() === lower);
	if (found === void 0) {
		const names = all.map(({ t }) => t.displayName);
		const close = suggest(name, names)[0];
		throw new XlvError("XLV_E_TABLE_NOT_FOUND", `The workbook has no table named "${name.slice(0, 60)}".`, {
			hint: names.length === 0 ? "The workbook has no tables; read a sheet with { sheet } instead." : `${close === void 0 ? "" : `Did you mean "${close}"? `}Tables: ${names.slice(0, 20).map((n) => `"${n}"`).join(", ")}.`,
			details: {
				table: name.slice(0, 60),
				available: names.slice(0, 50)
			}
		});
	}
	const r = parseRange(found.t.ref);
	if (r === null) throw new XlvError("XLV_E_TABLE_NOT_FOUND", `Table "${name.slice(0, 60)}" has no valid range.`, {
		hint: "Open and re-save the file in a spreadsheet application.",
		details: { table: name.slice(0, 60) }
	});
	const end = r.end.row - found.t.totalsRowCount;
	const range = `${formatCellRef(r.start.row, r.start.col)}:${formatCellRef(Math.max(end, r.start.row), r.end.col)}`;
	return {
		sheet: found.sheet,
		range,
		headerRow: found.t.headerRowCount > 0 ? r.start.row : null
	};
}
function prepare(input, options) {
	const o = readOptions(options, READ_OPTIONS, "readTable options");
	const workbook = openInput(input, o.get("limits"), o.get("onWarning"));
	const table = pickTable(workbook, o);
	const sheet = table?.sheet ?? pickSheet(workbook, o.get("sheet"));
	const schema = checkSchema(o.get("schema"));
	const dates = oneOfOption("dates", o.get("dates"), ["date", "iso"], "date");
	const skipEmpty = booleanOption("skipEmptyRows", o.get("skipEmptyRows"), true);
	const headerOption = table !== null ? table.headerRow ?? false : o.get("header") ?? "auto";
	if (headerOption !== "auto" && headerOption !== false && !(typeof headerOption === "number" && Number.isInteger(headerOption) && headerOption >= 1 && headerOption <= 1048576)) throw new XlvError("XLV_E_INVALID_OPTION", "The option \"header\" must be \"auto\", a row number or false.", {
		hint: "Use header: \"auto\" to detect the header row, header: 3 for row 3, or header: false for no header.",
		details: { option: "header" }
	});
	let range = null;
	const rawRange = table !== null ? table.range : o.get("range");
	if (rawRange !== void 0) {
		range = typeof rawRange === "string" ? parseRange(rawRange) : null;
		if (range === null) throw new XlvError("XLV_E_INVALID_CELL_REF", `The range ${JSON.stringify(String(rawRange).slice(0, 40))} is not valid.`, {
			location: { sheet: sheet.name },
			hint: "Pass an A1-style range such as \"A1:D100\".",
			details: { ref: String(rawRange).slice(0, 40) }
		});
	}
	const byRow = /* @__PURE__ */ new Map();
	let minRow = Infinity;
	let maxRow = 0;
	let minCol = Infinity;
	let maxCol = 0;
	for (const cell of sheet.cells()) {
		if (range && (cell.row < range.start.row || cell.row > range.end.row || cell.col < range.start.col || cell.col > range.end.col)) continue;
		let row = byRow.get(cell.row);
		if (!row) {
			row = /* @__PURE__ */ new Map();
			byRow.set(cell.row, row);
		}
		row.set(cell.col, cell);
		if (cell.row < minRow) minRow = cell.row;
		if (cell.row > maxRow) maxRow = cell.row;
		if (cell.col < minCol) minCol = cell.col;
		if (cell.col > maxCol) maxCol = cell.col;
	}
	if (range) {
		minRow = range.start.row;
		maxRow = range.end.row;
		minCol = range.start.col;
		maxCol = range.end.col;
	}
	if (maxRow === 0 || maxCol === 0) {
		if (typeof headerOption === "number") throw headerNotFound(sheet.name, headerOption);
		return {
			sheet: sheet.name,
			rows: [],
			rowNumbers: [],
			columns: [],
			headerRow: null,
			range: null,
			schema
		};
	}
	let headerRow = null;
	if (headerOption === "auto") {
		const ordered = [...byRow.keys()].sort((a, b) => a - b).map((r) => [...byRow.get(r)?.values() ?? []].sort((a, b) => a.col - b.col));
		headerRow = detectHeader(ordered, maxCol - minCol + 1, HEADER_SCAN_ROWS)?.row ?? null;
	} else if (typeof headerOption === "number") {
		if (headerOption < minRow || headerOption > maxRow || (byRow.get(headerOption)?.size ?? 0) === 0) throw headerNotFound(sheet.name, headerOption);
		headerRow = headerOption;
	}
	const columns = [];
	const used = /* @__PURE__ */ new Set();
	const headerCells = headerRow === null ? void 0 : byRow.get(headerRow);
	for (let col = minCol; col <= maxCol; col++) {
		const letter = columnLetters(col);
		const header = headerRow === null ? null : headerText(headerCells?.get(col));
		const base = header === null || header.trim() === "" ? letter : header.trim();
		let key = base;
		for (let n = 2; used.has(key); n++) key = `${base}_${String(n)}`;
		used.add(key);
		columns.push({
			key,
			header: header === null || header.trim() === "" ? null : header,
			col,
			letter
		});
	}
	const rows = [];
	const rowNumbers = [];
	const first = headerRow === null ? minRow : headerRow + 1;
	const numbers = skipEmpty ? [...byRow.keys()].filter((r) => r >= first && r <= maxRow).sort((a, b) => a - b) : rangeOf(first, maxRow);
	for (const r of numbers) {
		const cells = byRow.get(r);
		if (skipEmpty && (cells === void 0 || [...cells.values()].every((c) => c.value === null))) continue;
		const row = {};
		for (const column of columns) setOwn(row, column.key, cellValue(cells?.get(column.col), dates));
		rows.push(row);
		rowNumbers.push(r);
	}
	const lastRow = rowNumbers.length > 0 ? Math.max(rowNumbers[rowNumbers.length - 1] ?? 0, headerRow ?? 0) : headerRow ?? maxRow;
	const startRow = headerRow ?? rowNumbers[0] ?? minRow;
	return {
		sheet: sheet.name,
		rows,
		rowNumbers,
		columns,
		headerRow,
		range: `${formatCellRef(startRow, minCol)}:${formatCellRef(Math.max(lastRow, startRow), maxCol)}`,
		schema
	};
}
function rangeOf(from, to) {
	const out = [];
	for (let r = from; r <= to; r++) out.push(r);
	return out;
}
function headerNotFound(sheet, row) {
	return new XlvError("XLV_E_HEADER_NOT_FOUND", `Row ${String(row)} of sheet "${sheet}" has no cells to use as a header.`, {
		location: { sheet },
		hint: "Pass the row that holds the column names, header: \"auto\" to detect it, or header: false to use column letters as keys.",
		details: { row }
	});
}
function isThenable(value) {
	return typeof value === "object" && value !== null && typeof value.then === "function";
}
/** Converts a schema issue to a {@link TableIssue}, mapping the first path key to a column. */
function toIssue(issue, row, columns) {
	const path = (issue.path ?? []).map((segment) => {
		const key = typeof segment === "object" && segment !== null ? segment.key : segment;
		return typeof key === "number" ? key : String(key);
	});
	const first = path[0];
	const column = typeof first === "string" ? columns.find((c) => c.key === first) : void 0;
	return {
		row,
		cell: column === void 0 ? null : formatCellRef(row, column.col),
		column: column === void 0 ? null : column.key,
		message: typeof issue.message === "string" ? issue.message : String(issue.message),
		path
	};
}
function collect(prepared, results) {
	const rows = [];
	const rowNumbers = [];
	const issues = [];
	results.forEach((result, i) => {
		const row = prepared.rowNumbers[i] ?? 0;
		if (result.issues) for (const issue of result.issues) issues.push(toIssue(issue, row, prepared.columns));
		else {
			rows.push(result.value);
			rowNumbers.push(row);
		}
	});
	return {
		rows,
		rowNumbers,
		issues
	};
}
function readTable(input, options) {
	const prepared = prepare(input, options);
	const base = {
		sheet: prepared.sheet,
		columns: prepared.columns,
		headerRow: prepared.headerRow,
		range: prepared.range
	};
	const schema = prepared.schema;
	if (schema === void 0) return {
		...base,
		rows: prepared.rows,
		rowNumbers: prepared.rowNumbers,
		issues: []
	};
	const results = prepared.rows.map((row) => {
		const result = schema["~standard"].validate(row);
		if (isThenable(result)) {
			Promise.resolve(result).catch(() => void 0);
			throw new XlvError("XLV_E_ASYNC_SCHEMA", "The schema validates asynchronously, so readTable cannot use it.", {
				location: { sheet: prepared.sheet },
				hint: "Use await readTableAsync(input, options) with this schema, or a schema without asynchronous checks."
			});
		}
		return result;
	});
	return {
		...base,
		...collect(prepared, results)
	};
}
async function readTableAsync(input, options) {
	const prepared = prepare(input, options);
	const base = {
		sheet: prepared.sheet,
		columns: prepared.columns,
		headerRow: prepared.headerRow,
		range: prepared.range
	};
	const schema = prepared.schema;
	if (schema === void 0) return {
		...base,
		rows: prepared.rows,
		rowNumbers: prepared.rowNumbers,
		issues: []
	};
	const results = [];
	for (const row of prepared.rows) results.push(await schema["~standard"].validate(row));
	return {
		...base,
		...collect(prepared, results)
	};
}
/**
* Reads a sheet as JSON-safe data: like {@link readTable} without a schema, with dates as ISO
* 8601 strings. `JSON.stringify` of the result never loses information.
*
* @example
* ```ts
* import { toJSON } from 'xlvane';
*
* declare const bytes: Uint8Array;
* console.log(JSON.stringify(toJSON(bytes, { sheet: 'Data' }).rows, null, 2));
* ```
*/
function toJSON(input, options) {
	const o = readOptions(options, [
		"sheet",
		"table",
		"header",
		"range",
		"skipEmptyRows",
		"limits",
		"onWarning"
	], "toJSON options");
	const table = readTable(input, {
		...Object.fromEntries(o),
		dates: "iso"
	});
	return {
		sheet: table.sheet,
		headerRow: table.headerRow,
		range: table.range,
		columns: table.columns.map((c) => c.key),
		rows: table.rows
	};
}
const WRITE_OPTIONS = [
	"columns",
	"sheet",
	"preset",
	"dateSystem"
];
/** Header style of the `report` preset (D-304). */
const REPORT_HEADER = {
	font: { bold: true },
	fill: { color: "D9E1F2" },
	border: { bottom: { style: "thin" } }
};
const MIN_WIDTH = 8;
const MAX_FITTED_WIDTH = 60;
function tableError(message, hint, details = {}) {
	return new XlvError("XLV_E_INVALID_ROWS", message, {
		hint,
		details
	});
}
function resolveColumns(spec, rows, arrays) {
	if (spec === void 0) {
		if (arrays) {
			const width = rows.reduce((n, r) => Math.max(n, r.length), 0);
			return Array.from({ length: width }, (_, i) => ({
				key: String(i),
				header: "",
				numFmt: void 0,
				width: void 0
			}));
		}
		const keys = /* @__PURE__ */ new Set();
		for (const row of rows) for (const key of Object.keys(row)) keys.add(key);
		return [...keys].map((key) => ({
			key,
			header: key,
			numFmt: void 0,
			width: void 0
		}));
	}
	if (!Array.isArray(spec)) throw new XlvError("XLV_E_INVALID_OPTION", "The option \"columns\" must be an array.", {
		hint: "Pass column keys or objects, for example columns: [\"name\", { key: \"total\", numFmt: \"#,##0.00\" }].",
		details: { option: "columns" }
	});
	return spec.map((entry, i) => {
		if (typeof entry === "string") return {
			key: arrays ? String(i) : entry,
			header: entry,
			numFmt: void 0,
			width: void 0
		};
		const o = readOptions(entry, [
			"key",
			"header",
			"numFmt",
			"width"
		], `columns[${String(i)}]`);
		const key = o.get("key");
		const header = o.get("header");
		const numFmt = o.get("numFmt");
		const width = o.get("width");
		if (!arrays && typeof key !== "string") throw new XlvError("XLV_E_INVALID_OPTION", `columns[${String(i)}] has no "key".`, {
			hint: "Give each column the property name it shows, for example { key: \"total\", header: \"Total\" }.",
			details: { option: "columns" }
		});
		if (header !== void 0 && typeof header !== "string") throw new XlvError("XLV_E_INVALID_OPTION", `columns[${String(i)}].header must be text.`, {
			hint: "Use a string such as \"Total\".",
			details: { option: "columns" }
		});
		if (numFmt !== void 0 && typeof numFmt !== "string") throw new XlvError("XLV_E_INVALID_OPTION", `columns[${String(i)}].numFmt must be text.`, {
			hint: "Use a format code such as \"#,##0.00\".",
			details: { option: "columns" }
		});
		if (width !== void 0 && typeof width !== "number") throw new XlvError("XLV_E_INVALID_OPTION", `columns[${String(i)}].width must be a number.`, {
			hint: "Use a width in characters, for example 20.",
			details: { option: "columns" }
		});
		return {
			key: arrays ? String(i) : key,
			header: header ?? (typeof key === "string" ? key : ""),
			numFmt,
			width
		};
	});
}
function displayLength(value) {
	if (typeof value === "string") return Math.min(value.length, 200);
	if (typeof value === "number") return String(value).length;
	if (typeof value === "boolean") return 5;
	if (value instanceof Date) return value.getTime() % 864e5 === 0 ? 10 : 19;
	return 0;
}
/** Adds a table as a new sheet of `workbook`; shared by writeTable, fromJSON and fromCSV. */
function addTableSheet(workbook, rows, options) {
	if (!Array.isArray(rows)) throw tableError(`Expected an array of rows, got ${rows === null ? "null" : typeof rows}.`, "Pass an array of objects, such as [{ name: \"Alice\", total: 10 }], or an array of arrays.");
	if (rows.length > 1048575) throw new XlvError("XLV_E_LIMIT_EXCEEDED", `${String(rows.length)} rows do not fit in one sheet.`, {
		hint: `A sheet has ${String(MAX_ROWS)} rows, one of them for the header. Split the data over several sheets.`,
		details: {
			limit: "rows",
			value: MAX_ROWS - 1
		}
	});
	const arrays = rows.length > 0 && Array.isArray(rows[0]);
	rows.forEach((row, i) => {
		if (!(arrays ? Array.isArray(row) : typeof row === "object" && row !== null && !Array.isArray(row) && !(row instanceof Date))) throw tableError(`Row ${String(i + 1)} is ${row === null ? "null" : Array.isArray(row) ? "an array" : typeof row}, but the first row is ${arrays ? "an array" : "an object"}.`, "Give every row the same shape: all objects (keys become columns) or all arrays (columns by position).", { row: i + 2 });
	});
	const columns = resolveColumns(options.get("columns"), rows, arrays);
	if (columns.length > 16384) throw new XlvError("XLV_E_LIMIT_EXCEEDED", `${String(columns.length)} columns do not fit in one sheet.`, {
		hint: `A sheet has ${String(MAX_COLUMNS)} columns. Pass fewer columns with the "columns" option.`,
		details: {
			limit: "columns",
			value: MAX_COLUMNS
		}
	});
	const preset = oneOfOption("preset", options.get("preset"), ["report", "plain"], "report");
	const sheetName = options.get("sheet") ?? "Sheet1";
	const sheet = workbook.addSheet(sheetName);
	const hasHeader = !arrays || columns.some((c) => c.header !== "");
	const offset = hasHeader ? 1 : 0;
	const widths = columns.map((c) => displayLength(c.header));
	columns.forEach((column, c) => {
		if (hasHeader && column.header !== "") sheet.set(formatCellRef(1, c + 1), column.header, preset === "report" ? REPORT_HEADER : void 0);
	});
	rows.forEach((row, i) => {
		const r = i + 1 + offset;
		columns.forEach((column, c) => {
			const record = row;
			const raw = arrays ? row[c] : Object.prototype.hasOwnProperty.call(record, column.key) ? record[column.key] : void 0;
			const value = raw === void 0 ? null : raw;
			if (!(value === null || typeof value === "number" || typeof value === "string" || typeof value === "boolean" || value instanceof Date)) throw tableError(`Row ${String(i + 1)}, column "${column.key}" holds ${Array.isArray(value) ? "an array" : typeof value}, which is not a cell value.`, "Cell values are numbers, strings, booleans, Dates or null. Convert nested data first, for example with JSON.stringify.", {
				row: r,
				key: column.key
			});
			if (value === null && column.numFmt === void 0) return;
			sheet.set(formatCellRef(r, c + 1), value, column.numFmt === void 0 ? void 0 : { numFmt: column.numFmt });
			widths[c] = Math.max(widths[c] ?? 0, displayLength(value));
		});
	});
	columns.forEach((column, c) => {
		if (column.width !== void 0) sheet.column(c + 1, { width: column.width });
		else if (preset === "report") sheet.column(c + 1, { width: Math.min(MAX_FITTED_WIDTH, Math.max(MIN_WIDTH, (widths[c] ?? 0) + 2)) });
	});
	if (preset === "report" && hasHeader && columns.length > 0) {
		sheet.freeze({ rows: 1 });
		sheet.autoFilter(`A1:${formatCellRef(Math.max(1, rows.length + offset), columns.length)}`);
	}
}
/**
* Writes rows as a formatted XLSX table and returns the file's bytes. Rows are objects (keys
* become columns) or arrays (columns by position, headers from `columns`). The default
* `"report"` preset makes a readable report: bold header with a fill and a bottom border,
* frozen header row, autofilter, and column widths fitted to the content. Dates get a date
* format.
*
* @example
* ```ts
* import { writeTable } from 'xlvane';
*
* const bytes = writeTable(
*   [
*     { customer: 'Alice', total: 120.5, date: new Date(Date.UTC(2024, 0, 31)) },
*     { customer: 'Bob', total: 80, date: new Date(Date.UTC(2024, 1, 2)) },
*   ],
*   { sheet: 'Orders', columns: ['date', 'customer', { key: 'total', header: 'Total', numFmt: '#,##0.00' }] },
* );
* ```
*/
function writeTable(rows, options) {
	const o = readOptions(options, WRITE_OPTIONS, "writeTable options");
	const dateSystem = o.get("dateSystem");
	const workbook = createWorkbook(dateSystem === void 0 ? void 0 : { dateSystem });
	addTableSheet(workbook, rows, o);
	return writeWorkbook(workbook);
}
/**
* Writes JSON data as a formatted XLSX table: the same as {@link writeTable}, but also accepts
* JSON text. Throws `XLV_E_JSON_INVALID` for text that is not valid JSON and
* `XLV_E_INVALID_ROWS` when the data is not an array of objects or arrays. JSON keys such as
* `__proto__` become ordinary columns.
*
* @example
* ```ts
* import { fromJSON } from 'xlvane';
*
* const bytes = fromJSON('[{"sku":"A-1","stock":12},{"sku":"B-7","stock":0}]', { sheet: 'Stock' });
* ```
*/
function fromJSON(input, options) {
	let rows = input;
	if (typeof input === "string") try {
		rows = JSON.parse(input);
	} catch (error) {
		throw new XlvError("XLV_E_JSON_INVALID", `The text is not valid JSON: ${error instanceof Error ? error.message.slice(0, 200) : "parse error"}.`, {
			hint: "Check the JSON text (the message gives the position of the problem), or pass the parsed array instead.",
			cause: error
		});
	}
	if (!Array.isArray(rows)) throw tableError("The JSON data is not an array.", "Pass an array of row objects, for example [{\"name\":\"Alice\"}]. For an object that holds the rows, pass that property, e.g. data.rows.");
	return writeTable(rows, options);
}
// src/csv.ts
/** First characters that make spreadsheet applications treat a CSV field as a formula (CSV-10). */
const INJECTION_START = /* @__PURE__ */ new Set([
	"=",
	"+",
	"-",
	"@",
	"	",
	"\r",
	"\n",
	"＝",
	"＋",
	"－",
	"＠"
]);
/**
* Converts a sheet to CSV text. Fields that contain the delimiter, a quote, CR or LF are
* enclosed in quotes, and quotes are doubled (RFC 4180). Numbers are written in JavaScript's
* shortest form, booleans as `TRUE`/`FALSE`, dates as ISO 8601, error cells as their code and
* formula cells as their cached value. By default, text that a spreadsheet would run as a
* formula is escaped (`formulaInjection: "escape"`).
*
* @example
* ```ts
* import { toCSV } from 'xlvane';
*
* declare const bytes: Uint8Array;
* const csv = toCSV(bytes, { sheet: 'Orders' });
* console.log(csv.split('\r\n')[0]); // the header line
* ```
*/
function toCSV(input, options) {
	const o = readOptions(options, [
		"sheet",
		"range",
		"delimiter",
		"newline",
		"bom",
		"formulaInjection",
		"limits",
		"onWarning"
	], "toCSV options");
	const delimiter = o.get("delimiter") ?? ",";
	if (typeof delimiter !== "string" || delimiter.length !== 1 || delimiter === "\"" || delimiter === "\r" || delimiter === "\n") throw new XlvError("XLV_E_INVALID_OPTION", "The option \"delimiter\" must be one character other than a quote, CR or LF.", {
		hint: "Use \",\" (default), \";\" or \"\\t\".",
		details: { option: "delimiter" }
	});
	const newline = oneOfOption("newline", o.get("newline"), ["\r\n", "\n"], "\r\n");
	const bom = booleanOption("bom", o.get("bom"), false);
	const injection = oneOfOption("formulaInjection", o.get("formulaInjection"), ["escape", "off"], "escape");
	const sheet = pickSheet(openInput(input, o.get("limits"), o.get("onWarning")), o.get("sheet"));
	let startRow = 1;
	let startCol = 1;
	let endRow = 0;
	let endCol = 0;
	const rawRange = o.get("range");
	if (rawRange !== void 0) {
		const range = typeof rawRange === "string" ? parseRange(rawRange) : null;
		if (range === null) throw new XlvError("XLV_E_INVALID_CELL_REF", `The range ${JSON.stringify(String(rawRange).slice(0, 40))} is not valid.`, {
			location: { sheet: sheet.name },
			hint: "Pass an A1-style range such as \"A1:D100\".",
			details: { ref: String(rawRange).slice(0, 40) }
		});
		({row: startRow, col: startCol} = range.start);
		({row: endRow, col: endCol} = range.end);
	} else for (const cell of sheet.cells()) {
		if (cell.row > endRow) endRow = cell.row;
		if (cell.col > endCol) endCol = cell.col;
	}
	const needsQuotes = (s) => s.includes(delimiter) || s.includes("\"") || s.includes("\r") || s.includes("\n");
	const field = (value) => {
		if (value === null || value === void 0) return "";
		if (typeof value === "number") return String(value);
		if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
		if (value instanceof Date) return isoDate(value);
		let s = String(value);
		let quote = needsQuotes(s);
		if (injection === "escape" && s.length > 0 && INJECTION_START.has(s[0] ?? "")) {
			s = `'${s}`;
			quote = true;
		}
		return quote ? `"${s.replace(/"/g, "\"\"")}"` : s;
	};
	const lines = [];
	for (let r = startRow; r <= endRow; r++) {
		const fields = [];
		for (let c = startCol; c <= endCol; c++) {
			const cell = sheet.cell(formatCellRef(r, c));
			fields.push(field(cell?.value));
		}
		lines.push(fields.join(delimiter));
	}
	const body = lines.length === 0 ? "" : lines.join(newline) + newline;
	return (bom ? "﻿" : "") + body;
}
const DEFAULT_MAX_CHARS = 268435456;
function malformed(message, line, column, hint) {
	return new XlvError("XLV_E_CSV_MALFORMED", `${message} (line ${String(line)}, column ${String(column)}).`, {
		hint,
		details: {
			line,
			column
		}
	});
}
/**
* Splits CSV text into rows of fields (RFC 4180 with CRLF, LF or CR line ends). Throws
* `XLV_E_CSV_MALFORMED` for an unterminated quoted field or text after a closing quote.
*/
function parseCsvRows(text, delimiter, maxCells = Number.POSITIVE_INFINITY) {
	const rows = [];
	if (text.length === 0) return rows;
	let cells = 0;
	let row = [];
	let i = 0;
	let line = 1;
	let lineStart = 0;
	const n = text.length;
	const limitCheck = () => {
		if (++cells > maxCells) throw new XlvError("XLV_E_LIMIT_EXCEEDED", `The CSV text has more than ${String(maxCells)} fields (line ${String(line)}).`, {
			hint: "Raise the maxCells option if you trust the input, or split the file.",
			details: {
				limit: "maxCells",
				value: maxCells,
				line
			}
		});
		if (row.length > 16384) throw new XlvError("XLV_E_LIMIT_EXCEEDED", `Line ${String(line)} has more than ${String(MAX_COLUMNS)} fields.`, {
			hint: `A sheet has ${String(MAX_COLUMNS)} columns. Check the delimiter option; the file may use another separator.`,
			details: {
				limit: "columns",
				value: MAX_COLUMNS,
				line
			}
		});
	};
	while (true) {
		if (i < n && text[i] === "\"") {
			const quoteLine = line;
			const quoteColumn = i - lineStart + 1;
			i++;
			let value = "";
			let segment = i;
			while (true) {
				const q = text.indexOf("\"", i);
				if (q === -1) throw malformed("A quoted field is not closed", quoteLine, quoteColumn, "Add the closing quote, or double quotes that are part of the text (\"\" stands for one quote).");
				for (let k = i; k < q; k++) {
					const ch = text.charCodeAt(k);
					if (ch === 10 || ch === 13 && text.charCodeAt(k + 1) !== 10) {
						line++;
						lineStart = k + 1;
					}
				}
				if (text[q + 1] === "\"") {
					value += text.slice(segment, q + 1);
					i = q + 2;
					segment = i;
					continue;
				}
				value += text.slice(segment, q);
				i = q + 1;
				break;
			}
			row.push(value);
			if (i < n && text[i] !== delimiter && text[i] !== "\r" && text[i] !== "\n") throw malformed("Text follows a closing quote", line, i - lineStart + 1, "Enclose the whole field in quotes, or remove the text after the closing quote.");
		} else {
			let j = i;
			while (j < n) {
				const ch = text[j];
				if (ch === delimiter || ch === "\r" || ch === "\n") break;
				j++;
			}
			row.push(text.slice(i, j));
			i = j;
		}
		limitCheck();
		if (i >= n) {
			rows.push(row);
			break;
		}
		const ch = text[i];
		if (ch === delimiter) {
			i++;
			continue;
		}
		i += ch === "\r" && text[i + 1] === "\n" ? 2 : 1;
		rows.push(row);
		row = [];
		line++;
		lineStart = i;
		if (rows.length > 1048576) throw new XlvError("XLV_E_LIMIT_EXCEEDED", `The CSV text has more than ${String(MAX_ROWS)} lines.`, {
			hint: `A sheet has ${String(MAX_ROWS)} rows. Split the file.`,
			details: {
				limit: "rows",
				value: MAX_ROWS
			}
		});
		if (i >= n) break;
	}
	return rows;
}
/** Plain decimal number: optional minus, no leading zeros, optional fraction and exponent (D-307). */
function inferNumber(s) {
	let i = 0;
	const n = s.length;
	if (s[i] === "-") i++;
	const intStart = i;
	while (i < n && s.charCodeAt(i) >= 48 && s.charCodeAt(i) <= 57) i++;
	const intDigits = i - intStart;
	if (intDigits === 0 || intDigits > 1 && s[intStart] === "0") return null;
	let fracDigits = 0;
	if (s[i] === ".") {
		i++;
		const fracStart = i;
		while (i < n && s.charCodeAt(i) >= 48 && s.charCodeAt(i) <= 57) i++;
		fracDigits = i - fracStart;
		if (fracDigits === 0) return null;
	}
	if (s[i] === "e" || s[i] === "E") {
		i++;
		if (s[i] === "+" || s[i] === "-") i++;
		const expStart = i;
		while (i < n && s.charCodeAt(i) >= 48 && s.charCodeAt(i) <= 57) i++;
		if (i === expStart) return null;
	}
	if (i !== n) return null;
	const value = Number(s);
	if (!Number.isFinite(value)) return null;
	const digits = s.slice(intStart, intStart + intDigits + (fracDigits > 0 ? fracDigits + 1 : 0)).replace(".", "").replace(/^0+/, "");
	if (value === 0 ? /[1-9]/.test(digits) : Math.abs(value) < MIN_NORMAL) return null;
	if (digits.length > 15) {
		if (fracDigits === 0 && !/[eE]/.test(s) || String(value) !== s) return null;
	}
	return value;
}
/** Smallest positive normal double (2^-1022); smaller values have fewer significant digits. */
const MIN_NORMAL = 22250738585072014e-324;
/** The value of a CSV field with `types: "infer"` (also used by `xlvane edit --set`). */
function inferValue(s, minDate) {
	if (s === "") return null;
	const number = inferNumber(s);
	if (number !== null) return number;
	const lower = s.toLowerCase();
	if (lower === "true") return true;
	if (lower === "false") return false;
	if (s.length >= 10 && s[4] === "-" && s[7] === "-") {
		const date = parseIsoDateTime(s);
		if (date !== null && date.getTime() >= minDate) return date;
	}
	return s;
}
/**
* Builds an XLSX workbook from CSV text (or UTF-8 bytes) and returns the file's bytes. Fields are
* never turned into formulas: a field such as `=HYPERLINK(...)` becomes text. With
* `types: "infer"` (default), plain numbers, booleans and ISO 8601 dates are converted; numbers
* with leading zeros, whole numbers with more than 15 digits and numbers a double cannot hold
* stay text, so codes and IDs keep their digits, and numbers written by `toCSV` read back as the
* same numbers. More than `maxCells` fields give `XLV_E_LIMIT_EXCEEDED`. Throws `XLV_E_CSV_MALFORMED` with the line and column of a broken quoted field.
*
* @example
* ```ts
* import { fromCSV } from 'xlvane';
*
* const bytes = fromCSV('sku,stock\r\nA-1,12\r\nB-7,0\r\n', { sheet: 'Stock', preset: 'report' });
* ```
*/
function fromCSV(input, options) {
	const o = readOptions(options, [
		"delimiter",
		"header",
		"types",
		"sheet",
		"preset",
		"dateSystem",
		"maxChars",
		"maxCells"
	], "fromCSV options");
	const delimiter = o.get("delimiter") ?? ",";
	if (typeof delimiter !== "string" || delimiter.length !== 1 || delimiter === "\"" || delimiter === "\r" || delimiter === "\n") throw new XlvError("XLV_E_INVALID_OPTION", "The option \"delimiter\" must be one character other than a quote, CR or LF.", {
		hint: "Use \",\" (default), \";\" or \"\\t\".",
		details: { option: "delimiter" }
	});
	const header = booleanOption("header", o.get("header"), true);
	const types = oneOfOption("types", o.get("types"), ["infer", "text"], "infer");
	const maxChars = o.get("maxChars") ?? DEFAULT_MAX_CHARS;
	if (typeof maxChars !== "number" || !(maxChars > 0)) throw new XlvError("XLV_E_INVALID_OPTION", "The option \"maxChars\" must be a positive number.", {
		hint: "For example { maxChars: 1e9 }, or Infinity.",
		details: { option: "maxChars" }
	});
	const maxCells = o.get("maxCells") ?? DEFAULT_READ_LIMITS.maxCells;
	if (typeof maxCells !== "number" || !(maxCells > 0)) throw new XlvError("XLV_E_INVALID_OPTION", "The option \"maxCells\" must be a positive number.", {
		hint: "For example { maxCells: 20_000_000 }, or Infinity.",
		details: { option: "maxCells" }
	});
	let text;
	if (typeof input === "string") text = input;
	else if (input instanceof Uint8Array) {
		if (input.length > maxChars * 4) throw new XlvError("XLV_E_LIMIT_EXCEEDED", "The CSV data is larger than the limit.", {
			hint: "Raise maxChars if you trust the input.",
			details: {
				limit: "maxChars",
				value: maxChars
			}
		});
		try {
			text = new TextDecoder("utf-8", { fatal: true }).decode(input);
		} catch (error) {
			throw new XlvError("XLV_E_CSV_MALFORMED", "The CSV bytes are not valid UTF-8.", {
				hint: "Convert the file to UTF-8 first, or decode it with the right encoding and pass the text.",
				details: {
					line: 0,
					column: 0
				},
				cause: error
			});
		}
	} else throw new XlvError("XLV_E_INVALID_INPUT", "fromCSV expects CSV text or UTF-8 bytes.", { hint: "Pass a string, or a Uint8Array with UTF-8 text." });
	if (text.length > maxChars) throw new XlvError("XLV_E_LIMIT_EXCEEDED", `The CSV text has ${String(text.length)} characters, more than the limit.`, {
		hint: "Raise maxChars if you trust the input.",
		details: {
			limit: "maxChars",
			value: maxChars
		}
	});
	if (text.charCodeAt(0) === 65279) text = text.slice(1);
	const parsed = parseCsvRows(text, delimiter, maxCells);
	const headerFields = header ? parsed.shift() ?? [] : [];
	const width = parsed.reduce((w, r) => Math.max(w, r.length), headerFields.length);
	const dateSystem = o.get("dateSystem");
	const minDate = dateSystem === "1904" ? Date.UTC(1904, 0, 1) : Date.UTC(1900, 0, 1);
	const rows = parsed.map((fields) => fields.map((f) => types === "infer" ? inferValue(f, minDate) : f === "" ? null : f));
	const columns = header ? Array.from({ length: width }, (_, i) => headerFields[i] ?? "") : void 0;
	if (!header && rows.length > 0 && (rows[0]?.length ?? 0) < width) {
		const first = rows[0] ?? [];
		while (first.length < width) first.push(null);
	}
	const workbook = createWorkbook(dateSystem === void 0 ? void 0 : { dateSystem });
	const tableOptions = /* @__PURE__ */ new Map([["sheet", o.get("sheet") ?? "Sheet1"], ["preset", oneOfOption("preset", o.get("preset"), ["plain", "report"], "plain")]]);
	if (columns) tableOptions.set("columns", columns);
	addTableSheet(workbook, rows, tableOptions);
	return writeWorkbook(workbook);
}
// src/diff.ts
const DEFAULT_MAX_CELLS = 1e4;
function toDiffCell(cell, styles) {
	const value = cell.value instanceof Date ? cell.value.toISOString() : cell.value;
	const out = {
		type: cell.type,
		value,
		formula: cell.formula?.text ?? null
	};
	if (styles) out.style = cell.style;
	return out;
}
function sameCell(a, b, styles) {
	if (a.type !== b.type) return false;
	const av = a.value instanceof Date ? a.value.getTime() : a.value;
	const bv = b.value instanceof Date ? b.value.getTime() : b.value;
	if (!Object.is(av, bv)) return false;
	if ((a.formula?.text ?? null) !== (b.formula?.text ?? null)) return false;
	return !styles || stableKey(a.style) === stableKey(b.style);
}
/**
* Compares two workbooks: added and removed sheets (by name), visibility changes, and cells
* whose value, type or formula differ (and style, with `styles: true`). The result is plain
* JSON, suitable for review by people or checks by programs. Formulas are compared as text and
* values as stored; xlvane does not recalculate.
*
* @example
* ```ts
* import { diff } from 'xlvane';
*
* declare const before: Uint8Array;
* declare const after: Uint8Array;
* const result = diff(before, after);
* for (const sheet of result.sheets) {
*   for (const change of sheet.cells) console.log(sheet.name, change.ref, change.kind, change.before?.value, '->', change.after?.value);
* }
* ```
*/
function diff(a, b, options) {
	const o = readOptions(options, [
		"styles",
		"maxCells",
		"limits"
	], "diff options");
	const styles = booleanOption("styles", o.get("styles"), false);
	const maxCells = o.get("maxCells") ?? DEFAULT_MAX_CELLS;
	if (typeof maxCells !== "number" || !(maxCells >= 0) || !Number.isInteger(maxCells) && maxCells !== Infinity) throw new XlvError("XLV_E_INVALID_OPTION", "The option \"maxCells\" must be a whole number of 0 or more.", {
		hint: "For example { maxCells: 1000 }, or Infinity to list every change.",
		details: { option: "maxCells" }
	});
	const first = openInput(a, o.get("limits"));
	const second = openInput(b, o.get("limits"));
	const firstNames = new Map(first.sheets.map((s) => [s.name, s]));
	const secondNames = new Set(second.sheets.map((s) => s.name));
	const summary = {
		sheetsAdded: 0,
		sheetsRemoved: 0,
		sheetsChanged: 0,
		cellsAdded: 0,
		cellsRemoved: 0,
		cellsChanged: 0
	};
	const sheetsAdded = second.sheets.filter((s) => !firstNames.has(s.name)).map((s) => s.name);
	const sheetsRemoved = first.sheets.filter((s) => !secondNames.has(s.name)).map((s) => s.name);
	summary.sheetsAdded = sheetsAdded.length;
	summary.sheetsRemoved = sheetsRemoved.length;
	const sheets = [];
	for (const after of second.sheets) {
		const before = firstNames.get(after.name);
		if (!before) continue;
		const fromFirst = [];
		const fromSecond = [];
		let count = 0;
		const record = (list, row, col, change) => {
			count++;
			if (change.kind === "added") summary.cellsAdded++;
			else if (change.kind === "removed") summary.cellsRemoved++;
			else summary.cellsChanged++;
			if (list.length < maxCells) list.push({
				row,
				col,
				change
			});
		};
		for (const cell of before.cells()) {
			const other = after.cell(cell.ref);
			if (other === void 0) record(fromFirst, cell.row, cell.col, {
				ref: cell.ref,
				kind: "removed",
				before: toDiffCell(cell, styles),
				after: null
			});
			else if (!sameCell(cell, other, styles)) record(fromFirst, cell.row, cell.col, {
				ref: cell.ref,
				kind: "changed",
				before: toDiffCell(cell, styles),
				after: toDiffCell(other, styles)
			});
		}
		for (const cell of after.cells()) if (before.cell(cell.ref) === void 0) record(fromSecond, cell.row, cell.col, {
			ref: cell.ref,
			kind: "added",
			before: null,
			after: toDiffCell(cell, styles)
		});
		const stateChanged = before.state !== after.state;
		if (count === 0 && !stateChanged) continue;
		summary.sheetsChanged++;
		const listed = [...fromFirst, ...fromSecond].sort((x, y) => x.row - y.row || x.col - y.col).slice(0, maxCells).map((e) => e.change);
		sheets.push({
			name: after.name,
			...stateChanged ? { state: {
				before: before.state,
				after: after.state
			} } : {},
			cells: listed,
			truncated: listed.length < count
		});
	}
	return {
		equal: sheetsAdded.length === 0 && sheetsRemoved.length === 0 && sheets.length === 0,
		sheetsAdded,
		sheetsRemoved,
		sheets,
		summary
	};
}
// src/edit.ts
/** `"x:"` for a qualified name `x:local`, `""` without a prefix. */
function prefixOf(qname) {
	const colon = qname.indexOf(":");
	return colon === -1 ? "" : qname.slice(0, colon + 1);
}
/** Adds a namespace prefix to every element of a snippet generated by xlvane (never to source text). */
function withPrefix(xml, prefix) {
	return prefix === "" ? xml : xml.replace(/<(\/?)([A-Za-z])/g, `<$1${prefix}$2`);
}
/** Sets (or adds) an attribute in a start tag given as source text. */
function setAttr(tag, name, value) {
	const pattern = new RegExp(`(\\s${name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\s*=\\s*)("[^"]*"|'[^']*')`);
	if (pattern.test(tag)) return tag.replace(pattern, (_m, head) => `${head}"${escapeAttr(value)}"`);
	const close = tag.endsWith("/>") ? tag.length - 2 : tag.length - 1;
	return `${tag.slice(0, close)} ${name}="${escapeAttr(value)}"${tag.slice(close)}`;
}
/** Turns an empty-element start tag `<x .../>` into an open tag `<x ...>`. */
function openTag(tag) {
	return tag.endsWith("/>") ? `${tag.slice(0, -2).trimEnd()}>` : tag;
}
/**
* Applies replacements given as [start, end, text] on `source`; ranges must not overlap.
* Several insertions at one position end up in reverse order of `edits`, before a replacement
* that starts there. The text is built in one pass, so the cost is linear in its size (D-R03).
*/
function splice(source, edits) {
	if (edits.length === 0) return source;
	const order = edits.map((edit, i) => ({
		edit,
		i
	})).sort((a, b) => a.edit[0] - b.edit[0] || a.edit[1] - b.edit[1] || b.i - a.i);
	const out = [];
	let pos = 0;
	for (const { edit } of order) {
		const [start, end, text] = edit;
		if (start < pos) return spliceSlow(source, edits);
		out.push(source.slice(pos, start), text);
		pos = end;
	}
	out.push(source.slice(pos));
	return out.join("");
}
/** The earlier end-to-start splicing, kept for overlapping edits (which callers do not produce). */
function spliceSlow(source, edits) {
	const sorted = [...edits].sort((a, b) => b[0] - a[0] || b[1] - a[1]);
	let out = source;
	for (const [start, end, text] of sorted) out = out.slice(0, start) + text + out.slice(end);
	return out;
}
/** An XML part as edited text, with a UTF-16 declaration corrected to UTF-8 (D-310). */
function asUtf8(text) {
	return text.replace(/^(<\?xml[^>]*?encoding\s*=\s*)(["'])utf-16(?:le|be)?\2/i, "$1\"UTF-8\"");
}
/** Scans a part and records the elements `keep` selects, with their spans. */
function scan(text, part, limits, keep) {
	const elements = [];
	const stack = [];
	const names = [];
	let root = null;
	parseXml(text, {
		start(ns, local, attrs, span) {
			const depth = names.length + 1;
			const parent = names[names.length - 1] ?? "";
			names.push(local);
			let element = null;
			if (depth === 1 || keep(local, depth, parent, ns)) {
				element = {
					local,
					ns,
					depth,
					attrs,
					start: span,
					end: span
				};
				if (depth === 1) root = element;
				else elements.push(element);
			}
			stack.push(element);
		},
		end(_ns, _local, span) {
			names.pop();
			const element = stack.pop();
			if (element) element.end = span;
		},
		text() {}
	}, {
		part,
		limits
	});
	if (root === null) throw new XlvError("XLV_E_PART_INVALID", "The part has no root element.", {
		location: { part },
		hint: "Re-save the file in a spreadsheet application."
	});
	return {
		root,
		elements
	};
}
/**
* The parts of a package being edited. Entries are inflated only when an edit reads them; parts
* that no edit touches are copied to the new archive in their compressed form, so their bytes do
* not change and unreferenced entries cost neither time nor memory (D-R03).
*/
var PackageEditor = class {
	limits;
	pkg;
	zip;
	entries;
	/** Entry names by part key (and `[content_types].xml`), so lookups do not scan all entries. */
	byKey = /* @__PURE__ */ new Map();
	data = /* @__PURE__ */ new Map();
	changed = /* @__PURE__ */ new Map();
	removed = /* @__PURE__ */ new Set();
	added = [];
	texts = /* @__PURE__ */ new Map();
	encoder = new TextEncoder();
	constructor(bytes, limits, pkg) {
		this.limits = limits;
		this.pkg = pkg;
		this.zip = openZip(bytes, limits);
		this.entries = this.zip.entries;
		for (const e of this.entries) {
			if (e.name === "" || e.name.endsWith("/")) continue;
			const key = e.name.toLowerCase() === "[content_types].xml" ? CONTENT_TYPES_KEY : partKey(zipNameToPartName(e.name));
			if (!this.byKey.has(key)) this.byKey.set(key, e);
		}
	}
	entryOf(partName) {
		return this.byKey.get(partName === "/[Content_Types].xml" ? CONTENT_TYPES_KEY : partKey(partName));
	}
	bytesOf(entry) {
		let bytes = this.data.get(entry.name);
		if (bytes === void 0) {
			bytes = this.zip.read(entry);
			this.data.set(entry.name, bytes);
		}
		return bytes;
	}
	/** The text of a part (or of `[Content_Types].xml`), normalised and as last written. */
	text(partName) {
		const cached = this.texts.get(partKey(partName));
		if (cached !== void 0) return cached;
		const entry = this.entryOf(partName);
		if (entry === void 0) return void 0;
		return normalizeLineEnds(decodeXml(this.bytesOf(entry), partName));
	}
	write(partName, text) {
		this.texts.set(partKey(partName), text);
		const data = this.encoder.encode(asUtf8(text));
		const entry = this.entryOf(partName);
		if (entry) {
			this.changed.set(entry.name, data);
			return;
		}
		const name = partName.slice(1);
		const existing = this.added.find((a) => a.name === name);
		if (existing) existing.data = data;
		else this.added.push({
			name,
			data
		});
	}
	remove(partName) {
		const entry = this.entryOf(partName);
		if (entry) this.removed.add(entry.name);
	}
	hasPart(partName) {
		return this.entryOf(partName) !== void 0 || this.added.some((a) => partKey(`/${a.name}`) === partKey(partName));
	}
	build() {
		const out = [];
		for (const e of this.entries) {
			if (this.removed.has(e.name)) continue;
			const changed = this.changed.get(e.name);
			if (changed !== void 0) {
				out.push({
					name: e.name,
					data: changed
				});
				continue;
			}
			const inflated = this.data.get(e.name);
			const raw = inflated === void 0 ? this.zip.raw?.(e) ?? null : null;
			out.push(raw !== null ? {
				name: e.name,
				data: EMPTY_BYTES,
				raw
			} : {
				name: e.name,
				data: inflated ?? this.bytesOf(e)
			});
		}
		out.push(...this.added);
		return writeZip(out);
	}
};
const CONTENT_TYPES_KEY = "\0[content_types]";
const EMPTY_BYTES = /* @__PURE__ */ new Uint8Array(0);
/** Adds an Override to `[Content_Types].xml`. */
function addContentType(edit, partName, contentType) {
	const text = edit.text("/[Content_Types].xml");
	if (text === void 0) return;
	const { root } = scan(text, "/[Content_Types].xml", edit.limits, () => false);
	const prefix = prefixOf(root.start.qname);
	const override = withPrefix(`<Override${attr("PartName", partName)}${attr("ContentType", contentType)}/>`, prefix);
	edit.write("/[Content_Types].xml", root.start.empty ? `${text.slice(0, root.start.start)}${openTag(text.slice(root.start.start, root.start.end))}${override}</${root.start.qname}>${text.slice(root.start.end)}` : splice(text, [[
		root.end.start,
		root.end.start,
		override
	]]));
}
/** Removes the Override of `partName` from `[Content_Types].xml`. */
function removeContentType(edit, partName) {
	const text = edit.text("/[Content_Types].xml");
	if (text === void 0) return;
	const { elements } = scan(text, "/[Content_Types].xml", edit.limits, (local, depth) => depth === 2 && local === "Override");
	const key = partKey(partName);
	const cut = elements.filter((e) => partKey(e.attrs.get("PartName") ?? "") === key).map((e) => [
		e.start.start,
		e.end.end,
		""
	]);
	if (cut.length > 0) edit.write("/[Content_Types].xml", splice(text, cut));
}
/** Adds a relationship from `source` and returns its Id. */
function addRelationship(edit, source, type, target) {
	const relsPart = relationshipsPartName(source);
	const text = edit.text(relsPart) ?? "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"></Relationships>";
	const { root, elements } = scan(text, relsPart, edit.limits, (local, depth) => depth === 2 && local === "Relationship");
	const ids = new Set(elements.map((e) => e.attrs.get("Id") ?? ""));
	let n = elements.length + 1;
	while (ids.has(`rId${String(n)}`)) n++;
	const id = `rId${String(n)}`;
	const rel = withPrefix(`<Relationship${attr("Id", id)}${attr("Type", type)}${attr("Target", target)}/>`, prefixOf(root.start.qname));
	const updated = root.start.empty ? `${text.slice(0, root.start.start)}${openTag(text.slice(root.start.start, root.start.end))}${rel}</${root.start.qname}>${text.slice(root.start.end)}` : splice(text, [[
		root.end.start,
		root.end.start,
		rel
	]]);
	edit.write(relsPart, updated);
	return id;
}
/** Removes the relationships of `source` whose Id is in `ids`. */
function removeRelationships(edit, source, ids) {
	const relsPart = relationshipsPartName(source);
	const text = edit.text(relsPart);
	if (text === void 0) return;
	const { elements } = scan(text, relsPart, edit.limits, (local, depth) => depth === 2 && local === "Relationship");
	const cut = elements.filter((e) => ids.has(e.attrs.get("Id") ?? "")).map((e) => [
		e.start.start,
		e.end.end,
		""
	]);
	if (cut.length > 0) edit.write(relsPart, splice(text, cut));
}
var SharedStringsEditor = class {
	edit;
	part;
	workbookPart;
	conformance;
	index = /* @__PURE__ */ new Map();
	added = [];
	references = 0;
	existing;
	constructor(edit, part, workbookPart, conformance, warnings) {
		this.edit = edit;
		this.part = part;
		this.workbookPart = workbookPart;
		this.conformance = conformance;
		const text = part === null ? void 0 : edit.text(part);
		const rich = /* @__PURE__ */ new Map();
		const strings = text === void 0 || part === null ? [] : parseSharedStrings(text, {
			part,
			limits: edit.limits,
			warnings
		}, rich);
		strings.forEach((s, i) => {
			const key = textElement(s);
			if (!rich.has(i) && !this.index.has(key)) this.index.set(key, i);
		});
		this.existing = strings.length;
	}
	/** Adds a string item given as its content (stringItemXml) and returns its index. */
	add(item) {
		this.references++;
		let i = this.index.get(item);
		if (i === void 0) {
			i = this.existing + this.added.length;
			this.added.push(item);
			this.index.set(item, i);
		}
		return i;
	}
	finish() {
		if (this.references === 0) return;
		if (this.part === null) {
			let name = "/xl/sharedStrings.xml";
			for (let n = 2; this.edit.hasPart(name); n++) name = `/xl/sharedStrings${String(n)}.xml`;
			this.part = name;
			const base = this.workbookPart.slice(0, this.workbookPart.lastIndexOf("/") + 1);
			const target = name.startsWith(base) ? name.slice(base.length) : name;
			addRelationship(this.edit, this.workbookPart, relationshipTypes("sharedStrings")[this.conformance === "strict" ? 0 : 1], target);
			addContentType(this.edit, name, "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
			this.edit.write(name, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="${SML_NS[this.conformance]}" count="0" uniqueCount="0"></sst>`);
		}
		const part = this.part;
		const text = this.edit.text(part) ?? "";
		const { root } = scan(text, part, this.edit.limits, () => false);
		const prefix = prefixOf(root.start.qname);
		const items = this.added.map((item) => withPrefix(`<si>${item}</si>`, prefix)).join("");
		const count = Number(root.attrs.get("count") ?? this.existing) + this.references;
		let tag = setAttr(text.slice(root.start.start, root.start.end), "count", String(Number.isFinite(count) ? count : this.existing + this.references));
		tag = setAttr(tag, "uniqueCount", String(this.existing + this.added.length));
		const updated = root.start.empty ? `${text.slice(0, root.start.start)}${openTag(tag)}${items}</${root.start.qname}>${text.slice(root.start.end)}` : splice(text, [[
			root.start.start,
			root.start.end,
			tag
		], [
			root.end.start,
			root.end.start,
			items
		]]);
		this.edit.write(part, updated);
	}
};
/** Containers of the styles part in schema order (W2-10). */
const STYLE_CONTAINERS = [
	"numFmts",
	"fonts",
	"fills",
	"borders",
	"cellStyleXfs",
	"cellXfs",
	"cellStyles",
	"dxfs",
	"tableStyles",
	"colors",
	"extLst"
];
/** Records prepended when a container has to be created, so that index 0 keeps its meaning (D-311). */
const CONTAINER_DEFAULTS = {
	fonts: [fontXml({})],
	fills: ["<fill><patternFill patternType=\"none\"/></fill>", "<fill><patternFill patternType=\"gray125\"/></fill>"],
	borders: ["<border><left/><right/><top/><bottom/><diagonal/></border>"],
	cellXfs: ["<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"]
};
const COUNT_LIMITS = {
	numFmts: "customNumFmts",
	fonts: "fonts",
	fills: "fills",
	borders: "borders",
	cellXfs: "cellXfs"
};
const ITEM_NAMES = {
	numFmts: "numFmt",
	fonts: "font",
	fills: "fill",
	borders: "border",
	cellXfs: "xf"
};
var StylesEditor = class {
	edit;
	part;
	workbookPart;
	conformance;
	styles;
	root;
	ns;
	existingCount = /* @__PURE__ */ new Map();
	additions = /* @__PURE__ */ new Map();
	addedIds = /* @__PURE__ */ new Map();
	numFmtIds = /* @__PURE__ */ new Map();
	nextNumFmtId = 164;
	cache = /* @__PURE__ */ new Map();
	constructor(edit, part, workbookPart, conformance, warnings) {
		this.edit = edit;
		this.part = part;
		this.workbookPart = workbookPart;
		this.conformance = conformance;
		const text = part === null ? void 0 : edit.text(part);
		if (text === void 0 || part === null) {
			this.styles = EMPTY_STYLES;
			this.root = null;
			this.ns = SML_NS[conformance];
			return;
		}
		this.styles = parseStyles(text, {
			part,
			limits: edit.limits,
			warnings
		});
		this.root = parseXmlTree(text, {
			part,
			limits: edit.limits
		});
		this.ns = this.root.ns;
		for (const container of [
			"numFmts",
			"fonts",
			"fills",
			"borders",
			"cellXfs"
		]) {
			const el = childNamed(this.root, this.ns, container);
			this.existingCount.set(container, el === void 0 ? -1 : childrenNamed(el, this.ns, ITEM_NAMES[container] ?? "").length);
		}
		const numFmts = childNamed(this.root, this.ns, "numFmts");
		for (const nf of numFmts === void 0 ? [] : childrenNamed(numFmts, this.ns, "numFmt")) {
			const id = Number(nf.attrs.get("numFmtId"));
			const code = nf.attrs.get("formatCode");
			if (!Number.isInteger(id) || code === void 0) continue;
			if (!this.numFmtIds.has(code)) this.numFmtIds.set(code, id);
			this.nextNumFmtId = Math.max(this.nextNumFmtId, id + 1);
		}
	}
	/** `true` if the styles part has a cell format at this index. */
	hasCellFormat(index) {
		return index < (this.existingCount.get("cellXfs") ?? -1);
	}
	/** The cell style of a cell format index, as the reader resolves it. */
	styleOf(index) {
		return this.styles.cellStyle(index, null);
	}
	add(container, xml) {
		let ids = this.addedIds.get(container);
		if (!ids) {
			ids = /* @__PURE__ */ new Map();
			this.addedIds.set(container, ids);
		}
		const known = ids.get(xml);
		if (known !== void 0) return known;
		let list = this.additions.get(container);
		if (!list) {
			list = [];
			this.additions.set(container, list);
		}
		const existing = this.existingCount.get(container) ?? -1;
		const id = (existing >= 0 ? existing : CONTAINER_DEFAULTS[container]?.length ?? 0) + list.length;
		const limit = COUNT_LIMITS[container];
		if (limit !== void 0) checkStyleCount(limit, id + 1);
		list.push(xml);
		ids.set(xml, id);
		return id;
	}
	numFmtId(code) {
		if (typeof code === "number") return code;
		const builtin = reusableBuiltinNumFmt(code);
		if (builtin !== void 0) return builtin;
		let id = this.numFmtIds.get(code);
		if (id === void 0) {
			id = this.nextNumFmtId++;
			this.numFmtIds.set(code, id);
			this.add("numFmts", `<numFmt${attr("numFmtId", id)}${attr("formatCode", encodeXstring(code))}/>`);
		}
		return id;
	}
	/** A cell format: the format at `base` with the parts of `patch` replaced or merged (D-311). */
	derive(base, patch) {
		const key = `${String(base)}|${stableKey(patch)}`;
		const cached = this.cache.get(key);
		if (cached !== void 0) return cached;
		const cellXfs = this.root === null ? void 0 : childNamed(this.root, this.ns, "cellXfs");
		const baseXf = cellXfs === void 0 ? void 0 : childrenNamed(cellXfs, this.ns, "xf")[base];
		const attrs = /* @__PURE__ */ new Map();
		for (const [k, v] of baseXf?.attrs ?? []) if (!k.startsWith("{")) attrs.set(k, v);
		for (const [k, v] of [
			["numFmtId", "0"],
			["fontId", "0"],
			["fillId", "0"],
			["borderId", "0"]
		]) if (!attrs.has(k)) attrs.set(k, v);
		const current = normalizeStyle(this.styleOf(base), "existing");
		const merged = mergeStyles(current, patch);
		if (patch.font) {
			const defaults = this.styles.defaultFont();
			const fontDefaults = {
				name: defaults?.name ?? DEFAULT_FONT.name,
				size: defaults?.size ?? DEFAULT_FONT.size
			};
			attrs.set("fontId", String(this.add("fonts", fontXml(merged.font ?? {}, fontDefaults))));
			attrs.set("applyFont", "1");
		}
		if (patch.fill && merged.fill) {
			attrs.set("fillId", String(this.add("fills", fillXml(merged.fill))));
			attrs.set("applyFill", "1");
		}
		if (patch.border && merged.border) {
			attrs.set("borderId", String(this.add("borders", borderXml(merged.border))));
			attrs.set("applyBorder", "1");
		}
		if (patch.numFmt !== void 0) {
			attrs.set("numFmtId", String(this.numFmtId(patch.numFmt)));
			attrs.set("applyNumberFormat", "1");
		}
		let children = "";
		const alignmentAttrs = /* @__PURE__ */ new Map();
		const baseAlignment = baseXf === void 0 ? void 0 : childNamed(baseXf, this.ns, "alignment");
		for (const [k, v] of baseAlignment?.attrs ?? []) if (!k.startsWith("{")) alignmentAttrs.set(k, v);
		if (patch.alignment) {
			const a = patch.alignment;
			if (a.horizontal !== void 0) alignmentAttrs.set("horizontal", a.horizontal);
			if (a.vertical !== void 0) alignmentAttrs.set("vertical", a.vertical);
			if (a.wrapText) alignmentAttrs.set("wrapText", "1");
			if (a.indent !== void 0) alignmentAttrs.set("indent", String(a.indent));
			if (a.textRotation !== void 0) alignmentAttrs.set("textRotation", String(a.textRotation));
			if (a.shrinkToFit) alignmentAttrs.set("shrinkToFit", "1");
			if (a.justifyLastLine) alignmentAttrs.set("justifyLastLine", "1");
			if (a.readingOrder !== void 0) alignmentAttrs.set("readingOrder", String(READING_ORDERS.indexOf(a.readingOrder)));
			attrs.set("applyAlignment", "1");
		}
		if (alignmentAttrs.size > 0) children += `<alignment${[...alignmentAttrs].map(([k, v]) => attr(k, v)).join("")}/>`;
		const protection = baseXf === void 0 ? void 0 : childNamed(baseXf, this.ns, "protection");
		const protectionAttrs = /* @__PURE__ */ new Map();
		for (const [k, v] of protection?.attrs ?? []) if (!k.startsWith("{")) protectionAttrs.set(k, v);
		if (patch.protection) {
			for (const key of ["locked", "hidden"]) {
				const v = patch.protection[key];
				if (v !== void 0) protectionAttrs.set(key, v ? "1" : "0");
			}
			attrs.set("applyProtection", "1");
		}
		if (protectionAttrs.size > 0) children += `<protection${[...protectionAttrs].map(([k, v]) => attr(k, v)).join("")}/>`;
		const xf = `<xf${[...attrs].map(([k, v]) => attr(k, v)).join("")}${children === "" ? "/>" : `>${children}</xf>`}`;
		const id = this.add("cellXfs", xf);
		this.cache.set(key, id);
		return id;
	}
	finish() {
		if (this.additions.size === 0) return;
		if (this.part === null) {
			let name = "/xl/styles.xml";
			for (let n = 2; this.edit.hasPart(name); n++) name = `/xl/styles${String(n)}.xml`;
			this.part = name;
			const base = this.workbookPart.slice(0, this.workbookPart.lastIndexOf("/") + 1);
			addRelationship(this.edit, this.workbookPart, relationshipTypes("styles")[this.conformance === "strict" ? 0 : 1], name.startsWith(base) ? name.slice(base.length) : name);
			addContentType(this.edit, name, "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
			this.edit.write(name, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="${this.ns}"></styleSheet>`);
		}
		const part = this.part;
		const text = this.edit.text(part) ?? "";
		const { root, elements } = scan(text, part, this.edit.limits, (_local, depth) => depth === 2);
		const prefix = prefixOf(root.start.qname);
		const changes = [];
		let rootTag = text.slice(root.start.start, root.start.end);
		let appendAtEnd = "";
		for (const [container, items] of this.additions) {
			const element = elements.find((e) => e.local === container);
			if (element) {
				const existing = this.existingCount.get(container) ?? 0;
				const tag = setAttr(text.slice(element.start.start, element.start.end), "count", String(existing + items.length));
				const body = withPrefix(items.join(""), prefix);
				if (element.start.empty) changes.push([
					element.start.start,
					element.start.end,
					`${openTag(tag)}${body}</${element.start.qname}>`
				]);
				else changes.push([
					element.start.start,
					element.start.end,
					tag
				], [
					element.end.start,
					element.end.start,
					body
				]);
				continue;
			}
			const all = [...CONTAINER_DEFAULTS[container] ?? [], ...items];
			const created = withPrefix(`<${container}${attr("count", all.length)}>${all.join("")}</${container}>`, prefix);
			const order = STYLE_CONTAINERS.indexOf(container);
			const next = elements.find((e) => STYLE_CONTAINERS.indexOf(e.local) > order);
			if (next) changes.push([
				next.start.start,
				next.start.start,
				created
			]);
			else if (root.start.empty) appendAtEnd += created;
			else changes.push([
				root.end.start,
				root.end.start,
				created
			]);
		}
		if (root.start.empty) {
			rootTag = openTag(rootTag);
			changes.push([
				root.start.start,
				root.start.end,
				`${rootTag}${appendAtEnd}</${root.start.qname}>`
			]);
		}
		this.edit.write(part, splice(text, changes));
	}
};
function valueParts(v, sst, inFormula) {
	switch (v.kind) {
		case "number": return {
			t: "",
			v: formatNumber(v.n)
		};
		case "string": return inFormula ? {
			t: "str",
			v: escapeText(encodeXstring(v.s))
		} : {
			t: "s",
			v: String(sst.add(stringItemXml(v)))
		};
		case "boolean": return {
			t: "b",
			v: v.b ? "1" : "0"
		};
		case "error": return {
			t: "e",
			v: escapeText(v.e)
		};
	}
}
function newCellXml(ref, value, xf, sst) {
	const s = xf === 0 ? "" : attr("s", xf);
	if (value === null) return `<c r="${ref}"${s}/>`;
	if (value.kind === "formula") {
		const f = `<f>${escapeText(encodeXstring(value.text))}</f>`;
		if (value.result === null) return `<c r="${ref}"${s}>${f}</c>`;
		const r = valueParts(value.result, sst, true);
		return `<c r="${ref}"${s}${r.t === "" ? "" : attr("t", r.t)}>${f}<v>${r.v}</v></c>`;
	}
	const r = valueParts(value, sst, false);
	return `<c r="${ref}"${s}${r.t === "" ? "" : attr("t", r.t)}><v>${r.v}</v></c>`;
}
/** Inserts `r="…"` after the qualified name of a start tag given as source text. */
function insertRef(tag, qname, ref) {
	const at = 1 + qname.length;
	return `${tag.slice(0, at)} r="${ref}"${tag.slice(at)}`;
}
/**
* The cell format a new cell of a sheet starts from: the `style` of the `col` element that covers
* its column, which ECMA-376 Part 1 §18.3.1.13 says affects cells not yet allocated (W2-03). A row
* with its own format (`row@s` with `customFormat`) keeps the earlier behaviour, format 0, since
* which of the two wins is not stated in the sources read (D-A04). Indexes beyond the cell
* formats of the file are ignored.
*/
function newCellStyles(text, part, limits, styles) {
	const { elements } = scan(text, part, limits, (local, depth, parent) => depth === 3 && local === "col" && parent === "cols" || depth === 3 && local === "row" && parent === "sheetData");
	const cols = [];
	const rowsWithFormat = /* @__PURE__ */ new Set();
	let rowNumber = 0;
	for (const e of elements) if (e.local === "col") {
		const [min, max, style] = [
			"min",
			"max",
			"style"
		].map((k) => {
			const v = e.attrs.get(k)?.trim() ?? "";
			return /^[0-9]{1,10}$/.test(v) ? Number(v) : NaN;
		});
		if (Number.isInteger(min) && Number.isInteger(max) && Number.isInteger(style) && style !== 0 && styles.hasCellFormat(style ?? 0)) cols.push({
			min: min ?? 0,
			max: max ?? 0,
			style: style ?? 0
		});
	} else {
		const r = e.attrs.get("r")?.trim();
		rowNumber = r !== void 0 && /^[0-9]{1,7}$/.test(r) ? Number(r) : rowNumber + 1;
		const custom = e.attrs.get("customFormat")?.trim();
		if (e.attrs.has("s") && (custom === "1" || custom === "true")) rowsWithFormat.add(rowNumber);
	}
	return (row, col) => rowsWithFormat.has(row) ? 0 : cols.find((c) => col >= c.min && col <= c.max)?.style ?? 0;
}
function editSheetXml(text, part, sheet, planned, limits) {
	const { root, elements } = scan(text, part, limits, (local, depth, parent) => depth === 2 && (local === "sheetData" || local === "dimension") || depth === 3 && local === "row" && parent === "sheetData" || depth === 4 && local === "c" && parent === "row");
	const prefix = prefixOf(root.start.qname);
	const sheetData = elements.find((e) => e.depth === 2 && e.local === "sheetData");
	if (!sheetData) throw new XlvError("XLV_E_PART_INVALID", `The worksheet part of sheet "${sheet}" has no sheetData element.`, {
		location: {
			part,
			sheet
		},
		hint: "The sheet part is damaged. Re-save the file in a spreadsheet application."
	});
	const rows = [];
	let currentRow = 0;
	let current = null;
	let lastCol = 0;
	const invalid = (ref) => {
		throw new XlvError("XLV_E_INVALID_CELL_REF", `Sheet "${sheet}" contains the invalid reference "${ref.slice(0, 20)}", so it cannot be edited safely.`, {
			location: {
				part,
				sheet
			},
			hint: "Open and re-save the file in a spreadsheet application, which drops the damaged cells, then edit the re-saved file.",
			details: { ref: ref.slice(0, 40) }
		});
	};
	for (const e of elements) if (e.depth === 3) {
		const r = e.attrs.get("r");
		const n = r === void 0 ? currentRow + 1 : /^[0-9]{1,7}$/.test(r.trim()) ? Number(r.trim()) : NaN;
		if (!Number.isInteger(n) || n < 1 || n > 1048576) invalid(r ?? String(n));
		currentRow = n;
		current = {
			number: currentRow,
			hasR: r !== void 0,
			element: e,
			cells: []
		};
		rows.push(current);
		lastCol = 0;
	} else if (e.depth === 4 && current) {
		const r = e.attrs.get("r");
		const address = r === void 0 || r === "" ? null : parseCellRef(r);
		if (r !== void 0 && r !== "" && address === null) invalid(r);
		const col = address?.col ?? lastCol + 1;
		lastCol = col;
		current.cells.push({
			col,
			hasR: r !== void 0,
			element: e
		});
	}
	const byRow = /* @__PURE__ */ new Map();
	for (const p of planned) {
		const list = byRow.get(p.row) ?? [];
		byRow.set(p.row, list);
		list.push(p);
	}
	for (const list of byRow.values()) list.sort((a, b) => a.col - b.col);
	const forceRefs = rows.some((r) => !r.hasR);
	const emitted = /* @__PURE__ */ new Set();
	const pending = [...byRow.keys()].filter((n) => !rows.some((r) => r.number === n)).sort((a, b) => a - b);
	const p = (xml) => withPrefix(xml, prefix);
	const newRow = (n) => {
		const cells = (byRow.get(n) ?? []).filter((c) => c.xml !== null).map((c) => p(c.xml ?? "")).join("");
		return `${p(`<row r="${String(n)}">`)}${cells}</${prefix}row>`;
	};
	let inner = "";
	for (const row of rows) {
		while (pending.length > 0 && (pending[0] ?? 0) < row.number) inner += newRow(pending.shift() ?? 0);
		const edits = byRow.get(row.number);
		const needsRefs = forceRefs || row.cells.some((c) => !c.hasR);
		if (!edits && row.hasR && !needsRefs) {
			inner += text.slice(row.element.start.start, row.element.end.end);
			continue;
		}
		let tag = text.slice(row.element.start.start, row.element.start.end);
		if (!row.hasR) tag = insertRef(tag, row.element.start.qname, String(row.number));
		let body = "";
		const edited = new Set((edits ?? []).map((c) => c.col));
		const toInsert = emitted.has(row.number) ? [] : (edits ?? []).filter((c) => c.xml !== null);
		emitted.add(row.number);
		let k = 0;
		for (const cell of row.cells) {
			while (k < toInsert.length && (toInsert[k]?.col ?? 0) < cell.col) body += p(toInsert[k++]?.xml ?? "");
			if (edited.has(cell.col)) continue;
			let cellText = text.slice(cell.element.start.start, cell.element.end.end);
			if (!cell.hasR) cellText = insertRef(cellText, cell.element.start.qname, formatCellRef(row.number, cell.col));
			body += cellText;
		}
		while (k < toInsert.length) body += p(toInsert[k++]?.xml ?? "");
		inner += row.element.start.empty ? `${openTag(tag)}${body}</${row.element.start.qname}>` : `${tag}${body}${text.slice(row.element.end.start, row.element.end.end)}`;
	}
	while (pending.length > 0) inner += newRow(pending.shift() ?? 0);
	const changes = [];
	if (sheetData.start.empty) changes.push([
		sheetData.start.start,
		sheetData.start.end,
		`${openTag(text.slice(sheetData.start.start, sheetData.start.end))}${inner}</${sheetData.start.qname}>`
	]);
	else changes.push([
		sheetData.start.end,
		sheetData.end.start,
		inner
	]);
	const dimension = elements.find((e) => e.depth === 2 && e.local === "dimension");
	if (dimension) {
		const old = parseRange(dimension.attrs.get("ref") ?? "");
		const values = planned.filter((c) => c.xml !== null);
		if (values.length > 0) {
			let minRow = old?.start.row ?? Infinity;
			let minCol = old?.start.col ?? Infinity;
			let maxRow = old?.end.row ?? 0;
			let maxCol = old?.end.col ?? 0;
			for (const c of values) {
				minRow = Math.min(minRow, c.row);
				minCol = Math.min(minCol, c.col);
				maxRow = Math.max(maxRow, c.row);
				maxCol = Math.max(maxCol, c.col);
			}
			const ref = minRow === maxRow && minCol === maxCol ? formatCellRef(minRow, minCol) : `${formatCellRef(minRow, minCol)}:${formatCellRef(maxRow, maxCol)}`;
			changes.push([
				dimension.start.start,
				dimension.start.end,
				setAttr(text.slice(dimension.start.start, dimension.start.end), "ref", ref)
			]);
		}
	}
	return splice(text, changes);
}
const CALC_PR_LATER = [
	"oleSize",
	"customWorkbookViews",
	"pivotCaches",
	"smartTagPr",
	"smartTagTypes",
	"webPublishing",
	"fileRecoveryPr",
	"webPublishObjects",
	"extLst"
];
/** Sets `calcPr@fullCalcOnLoad="1"`, adding `calcPr` in schema order if needed (W2-22, PROVISIONAL). */
function requestFullCalculation(edit, workbookPart) {
	const text = edit.text(workbookPart);
	if (text === void 0) return;
	const { root, elements } = scan(text, workbookPart, edit.limits, (_local, depth) => depth === 2);
	const calcPr = elements.find((e) => e.local === "calcPr");
	if (calcPr) {
		if (calcPr.attrs.get("fullCalcOnLoad") === "1") return;
		edit.write(workbookPart, splice(text, [[
			calcPr.start.start,
			calcPr.start.end,
			setAttr(text.slice(calcPr.start.start, calcPr.start.end), "fullCalcOnLoad", "1")
		]]));
		return;
	}
	const created = withPrefix("<calcPr fullCalcOnLoad=\"1\"/>", prefixOf(root.start.qname));
	const next = elements.find((e) => CALC_PR_LATER.includes(e.local));
	edit.write(workbookPart, splice(text, [[
		next ? next.start.start : root.end.start,
		next ? next.start.start : root.end.start,
		created
	]]));
}
/**
* editCells writes plain formulas only, and does not change cells whose formula covers other
* cells: an array formula's range (P1 §18.17.6.3) or a shared formula master with dependents
* (E-28). Replacing a cell that shares a master's formula is allowed: its own value overrides
* the group (E-28).
*/
function checkFormulaGroups(sheet, ref, at, value, cache) {
	const refuse = (message, hint) => new XlvError("XLV_E_NOT_SUPPORTED", message, {
		location: {
			part: sheet.part ?? "",
			sheet: sheet.name,
			cell: ref
		},
		hint,
		details: { cell: ref }
	});
	if (value !== null && value.kind === "formula" && (value.group !== void 0 || value.master !== void 0)) throw refuse(`editCells writes plain formulas only; the value for ${ref} is an array or shared formula.`, "Write { formula: \"…\" } for each cell, or build the sheet with createWorkbook to write array and shared formulas.");
	let arrays = cache.get(sheet.name);
	if (arrays === void 0) {
		arrays = [];
		for (const c of sheet.cells()) {
			const range = c.formula?.kind === "array" && c.formula.ref !== null ? parseRange(c.formula.ref) : null;
			if (range !== null) arrays.push({
				ref: c.formula.ref,
				range
			});
		}
		cache.set(sheet.name, arrays);
	}
	const array = arrays.find((a) => at.row >= a.range.start.row && at.row <= a.range.end.row && at.col >= a.range.start.col && at.col <= a.range.end.col);
	if (array !== void 0) throw refuse(`${ref} belongs to the array formula ${array.ref}, which editCells cannot change cell by cell.`, `Change the array formula in a spreadsheet application, or rebuild the sheet with createWorkbook; cells outside ${array.ref} can be edited.`);
	const f = sheet.cell(ref)?.formula;
	const range = f?.kind === "shared" && f.ref !== null ? parseRange(f.ref) : null;
	if (range !== null && (range.start.row !== range.end.row || range.start.col !== range.end.col)) throw refuse(`${ref} is the master of a shared formula used by cells in ${f.ref}; replacing it would leave them without a formula.`, `Edit the other cells of ${f.ref} first, or rebuild the sheet with createWorkbook.`);
}
/** Column names of a table: 1 to 255 characters on one line, unique ignoring case (TD-04). */
const MAX_COLUMN_NAME = 255;
/**
* Keeps the Table Definition parts of an edited sheet in step with its cells:
* - a header cell of a table holds the column name (OI §2.1.681 i, TD-05), so editing it renames
*   `tableColumn@name`; a value that is not text, an empty or multi-line name and a name another
*   column of the table already has are refused (`XLV_E_INVALID_VALUE`). Structured references
*   to the old name in formulas are not rewritten;
* - rows appended directly below a table's last row (appendRows) join the table: `ref` and an
*   `autoFilter` that covers the header and data rows grow with them, as when rows are inserted
*   inside a table (TD-11). A table whose totals row is the last used row is refused instead of
*   moving its totals row, with a hint to insert the rows above the totals row with insertRows.
*/
function syncTables(edit, plan, appended) {
	const sheet = plan.sheet;
	const tables = sheet.tables;
	for (const table of tables) {
		const range = parseRange(table.ref.trim().toUpperCase());
		if (range === null) continue;
		const header = table.headerRowCount > 0 ? range.start.row : null;
		const totals = table.totalsRowCount > 0 ? 1 : 0;
		const where = {
			part: table.part,
			sheet: sheet.name
		};
		const renames = /* @__PURE__ */ new Map();
		if (header !== null) {
			for (const [ref, c] of plan.cells) {
				if (c.row !== header || c.col < range.start.col || c.col > range.end.col) continue;
				const index = c.col - range.start.col;
				const value = c.value;
				const name = value !== null && value.kind === "string" ? value.s : null;
				const location = {
					part: sheet.part ?? "",
					sheet: sheet.name,
					cell: ref
				};
				if (name === null || name.length === 0 || name.length > MAX_COLUMN_NAME || /[\r\n]/.test(name)) throw new XlvError("XLV_E_INVALID_VALUE", `${ref} is a header cell of table "${table.displayName}", so it must hold the column name as text of 1 to ${String(MAX_COLUMN_NAME)} characters on one line.`, {
					location,
					hint: "Write the new column name as a string, for example \"Revenue\". To change the data, edit the rows below the header.",
					details: {
						table: table.displayName,
						cell: ref
					}
				});
				renames.set(index, name);
			}
			if (renames.size > 0) {
				const names = table.columns.map((column, i) => renames.get(i) ?? column.name);
				const seen = /* @__PURE__ */ new Map();
				names.forEach((name, i) => {
					const key = name.toLowerCase();
					const other = seen.get(key);
					if (other !== void 0 && (renames.has(i) || renames.has(other))) {
						const cell = formatCellRef(header, range.start.col + (renames.has(i) ? i : other));
						throw new XlvError("XLV_E_INVALID_VALUE", `Table "${table.displayName}" would have two columns named "${name.slice(0, 40)}".`, {
							location: {
								part: sheet.part ?? "",
								sheet: sheet.name,
								cell
							},
							hint: "Column names must be unique in a table (case is ignored). Choose another name, for example with a suffix such as \"_2\".",
							details: {
								table: table.displayName,
								cell
							}
						});
					}
					seen.set(key, i);
				});
			}
		}
		let newEnd = null;
		if (appended !== null && appended.count > 0 && range.end.row >= appended.lastRow && range.start.row <= appended.lastRow) {
			if (totals === 1) throw new XlvError("XLV_E_NOT_SUPPORTED", `Table "${table.displayName}" of sheet "${sheet.name}" ends with a totals row on row ${String(range.end.row)}, so rows appended below it would fall outside the table.`, {
				location: {
					...where,
					cell: formatCellRef(range.end.row, range.start.col)
				},
				hint: `Insert the rows above the totals row with insertRows(bytes, { sheet: "${sheet.name}", at: ${String(range.end.row)}, count: ${String(appended.count)} }), which grows the table, then fill them with editCells.`,
				details: { table: table.displayName }
			});
			const end = appended.lastRow + appended.count;
			if (end > range.end.row) {
				newEnd = end;
				const grown = {
					start: range.start,
					end: {
						row: end,
						col: range.end.col
					}
				};
				for (const other of tables) {
					if (other === table) continue;
					const r = parseRange(other.ref.trim().toUpperCase());
					if (r !== null && r.start.row <= grown.end.row && r.end.row >= grown.start.row && r.start.col <= grown.end.col && r.end.col >= grown.start.col) throw new XlvError("XLV_E_NOT_SUPPORTED", `Growing table "${table.displayName}" to ${formatRange(grown)} for the appended rows would overlap table "${other.displayName}" (${other.ref}).`, {
						location: where,
						hint: "Tables must not overlap. Append the rows to another sheet, or move one of the tables first.",
						details: { table: table.displayName }
					});
				}
			}
		}
		if (renames.size === 0 && newEnd === null) continue;
		const text = edit.text(table.part);
		if (text === void 0) continue;
		const { root, elements } = scan(text, table.part, edit.limits, (local, depth, parent) => depth === 2 && local === "autoFilter" || depth === 3 && local === "tableColumn" && parent === "tableColumns");
		const edits = [];
		if (newEnd !== null) {
			const end = newEnd;
			const grow = (r) => formatRange({
				start: r.start,
				end: {
					row: end,
					col: r.end.col
				}
			});
			edits.push([
				root.start.start,
				root.start.end,
				setAttr(text.slice(root.start.start, root.start.end), "ref", grow(range))
			]);
			const filter = elements.find((e) => e.local === "autoFilter");
			const filterRef = filter?.attrs.get("ref");
			const filterRange = filterRef === void 0 ? null : parseRange(filterRef.trim().toUpperCase());
			if (filter !== void 0 && filterRange !== null && filterRange.end.row === range.end.row) edits.push([
				filter.start.start,
				filter.start.end,
				setAttr(text.slice(filter.start.start, filter.start.end), "ref", grow(filterRange))
			]);
		}
		elements.filter((e) => e.local === "tableColumn").forEach((e, i) => {
			const name = renames.get(i);
			if (name !== void 0) edits.push([
				e.start.start,
				e.start.end,
				setAttr(text.slice(e.start.start, e.start.end), "name", name)
			]);
		});
		edit.write(table.part, splice(text, edits));
	}
}
function inputBytes(input) {
	if (input instanceof Uint8Array || input instanceof ArrayBuffer || ArrayBuffer.isView(input)) return toBytes(input);
	throw new XlvError("XLV_E_INVALID_INPUT", "Editing needs the bytes of the original file.", { hint: "Pass the file contents as a Uint8Array or ArrayBuffer (not a workbook from readWorkbook): unchanged parts are copied from the original bytes." });
}
function isFormula(value) {
	return value !== null && value.kind === "formula";
}
/** Validates and converts edits with the writer's rules, collecting them per sheet. */
function planEdits(workbook, edits) {
	const scratch = createWorkbook({ dateSystem: workbook.dateSystem });
	const scratchSheet = scratch.addSheet("Edit");
	const plans = /* @__PURE__ */ new Map();
	const arrayRanges = /* @__PURE__ */ new Map();
	edits.forEach((raw, i) => {
		const o = readOptions(raw, [
			"sheet",
			"cell",
			"value",
			"style"
		], `edits[${String(i)}]`);
		const sheetName = o.get("sheet");
		if (typeof sheetName !== "string") throw new XlvError("XLV_E_INVALID_OPTION", `edits[${String(i)}] has no sheet name.`, {
			hint: "Give each edit a sheet, for example { sheet: \"Data\", cell: \"B2\", value: 5 }.",
			details: { option: "sheet" }
		});
		const sheet = workbook.getSheet(sheetName);
		if (sheet.kind !== "worksheet" || sheet.part === null) throw new XlvError("XLV_E_INVALID_INPUT", `Sheet "${sheet.name}" is a ${sheet.kind} without editable cells.`, {
			location: { sheet: sheet.name },
			hint: "Edit cells of worksheets only."
		});
		const cellRef = o.get("cell");
		const address = typeof cellRef === "string" ? parseCellRef(cellRef) : null;
		if (address === null) throw new XlvError("XLV_E_INVALID_CELL_REF", `"${String(cellRef).slice(0, 40)}" is not a valid cell reference.`, {
			location: { sheet: sheet.name },
			hint: "Pass an A1-style reference such as \"B7\" (column letters A to XFD, rows 1 to 1048576).",
			details: { ref: String(cellRef).slice(0, 40) }
		});
		const ref = formatCellRef(address.row, address.col);
		const value = o.get("value") ?? null;
		let style = null;
		try {
			style = o.has("style") ? normalizeStyle(o.get("style")) : null;
			scratchSheet.set(ref, value, void 0);
		} catch (error) {
			if (error instanceof XlvError) throw new XlvError(error.code, error.message, {
				location: {
					part: sheet.part,
					sheet: sheet.name,
					cell: ref
				},
				hint: error.hint,
				details: error.details,
				cause: error
			});
			throw error;
		}
		const stored = scratchSheet.cellMap.get(address.row)?.get(address.col);
		checkFormulaGroups(sheet, ref, address, stored?.value ?? null, arrayRanges);
		const scratchStyle = scratch.styleOf(stored?.style ?? 0);
		scratchSheet.set(ref, null, {});
		const dateValue = value instanceof Date ? value : typeof value === "object" && value !== null && "result" in value && value.result instanceof Date ? value.result : null;
		if (dateValue !== null && workbook.conformance === "strict" && dateValue.getTime() < Date.UTC(1900, 2, 1)) throw new XlvError("XLV_E_INVALID_VALUE", `The date ${dateValue.toISOString()} for ${ref} cannot be written to a Strict workbook.`, {
			location: {
				part: sheet.part,
				sheet: sheet.name,
				cell: ref
			},
			hint: "Dates before 1900-03-01 are not supported when editing Strict workbooks. Write the date as text."
		});
		let plan = plans.get(sheet.name);
		if (!plan) {
			plan = {
				sheet,
				cells: /* @__PURE__ */ new Map()
			};
			plans.set(sheet.name, plan);
		}
		plan.cells.set(ref, {
			row: address.row,
			col: address.col,
			value: stored?.value ?? null,
			style,
			date: scratchStyle.numFmt ?? null
		});
	});
	return plans;
}
function applyEdits$1(bytes, workbook, plans, limits, appended = null) {
	const pkg = openPackage(openZip(bytes, limits), limits);
	const edit = new PackageEditor(bytes, limits, pkg);
	const warnings = new WarningCollector();
	const workbookPart = pkg.relationships(null).find((r) => classifyRelationship(r.type)?.name === "officeDocument")?.target ?? "/xl/workbook.xml";
	const rels = pkg.relationships(workbookPart);
	const relOf = (name) => rels.find((r) => !r.external && classifyRelationship(r.type)?.name === name);
	const sstRel = relOf("sharedStrings");
	const stylesRel = relOf("styles");
	const sst = new SharedStringsEditor(edit, sstRel !== void 0 && pkg.hasPart(sstRel.target) ? sstRel.target : null, workbookPart, workbook.conformance, warnings);
	const styles = new StylesEditor(edit, stylesRel !== void 0 && pkg.hasPart(stylesRel.target) ? stylesRel.target : null, workbookPart, workbook.conformance, warnings);
	let formulasChanged = false;
	for (const plan of plans.values()) {
		const part = plan.sheet.part ?? "";
		const planned = [];
		const text = edit.text(part);
		const columnStyle = text === void 0 ? () => 0 : newCellStyles(text, part, limits, styles);
		let styled = null;
		const styledCell = (ref) => (styled ??= new Map([...plan.sheet.styledCells()].map((cell) => [cell.ref, cell]))).get(ref);
		const merges = plan.sheet.merges.map((m) => ({
			ref: m,
			range: parseRange(m.trim().toUpperCase())
		}));
		for (const [ref, c] of plan.cells) {
			if (c.value !== null) {
				const merge = merges.find(({ range: r }) => r !== null && c.row >= r.start.row && c.row <= r.end.row && c.col >= r.start.col && c.col <= r.end.col && (c.row !== r.start.row || c.col !== r.start.col));
				if (merge !== void 0 && merge.range !== null) {
					const topLeft = formatCellRef(merge.range.start.row, merge.range.start.col);
					warnings.add("XLV_W_MERGED_CELL_WRITE", `A value was written into a merged range outside its top-left cell (first: ${ref} in ${merge.ref}). Spreadsheet applications show only the top-left cell (${topLeft}), so the value will not be visible. Write it to the top-left cell, or unmerge the range.`, {
						part,
						sheet: plan.sheet.name,
						cell: ref
					});
				}
			}
			const existing = plan.sheet.cell(ref) ?? styledCell(ref);
			if (isFormula(c.value) || existing?.formula) formulasChanged = true;
			if (c.value === null && c.style === null) {
				planned.push({
					row: c.row,
					col: c.col,
					ref,
					xml: null
				});
				continue;
			}
			const base = existing?.styleIndex ?? columnStyle(c.row, c.col);
			let patch = c.style ?? {};
			const existingIsDate = existing?.numFmt != null && isDateFormat(existing.numFmt);
			if (c.date !== null && patch.numFmt === void 0 && !existingIsDate) patch = {
				...patch,
				numFmt: c.date
			};
			const xf = Object.keys(patch).length === 0 ? base : styles.derive(base, patch);
			planned.push({
				row: c.row,
				col: c.col,
				ref,
				xml: newCellXml(ref, c.value, xf, sst)
			});
		}
		if (text === void 0) continue;
		syncTables(edit, plan, appended !== null && appended.sheet === plan.sheet.name ? appended : null);
		edit.write(part, editSheetXml(text, part, plan.sheet.name, planned, limits));
	}
	sst.finish();
	styles.finish();
	if (formulasChanged || workbook.sheets.some((s) => [...s.cells()].some((cell) => cell.formula !== null))) {
		requestFullCalculation(edit, workbookPart);
		warnings.add("XLV_W_FORMULAS_NOT_RECALCULATED", "The workbook contains formulas. xlvane does not recalculate them: cached results may be out of date until the file is opened in a spreadsheet application, which is asked to recalculate on load.", { part: workbookPart });
	}
	if (formulasChanged) {
		const chains = rels.filter((r) => !r.external && classifyRelationship(r.type)?.name === "calcChain");
		for (const chain of chains) {
			edit.remove(chain.target);
			removeContentType(edit, chain.target);
		}
		if (chains.length > 0) removeRelationships(edit, workbookPart, new Set(chains.map((c) => c.id)));
	}
	return {
		bytes: edit.build(),
		warnings: warnings.list()
	};
}
/**
* Changes cells of an existing XLSX file and returns the new file. Only the edited sheets'
* cell data changes, plus new entries in the shared strings and styles: charts, comments,
* conditional formatting, data validation, defined names and every other part stay exactly as
* they were. Each edit sets a value (as `NewSheet.set` does, with the same validation) and can
* merge a style into the cell's current style. When the workbook contains formulas, the file
* asks spreadsheet applications to recalculate on load (`XLV_W_FORMULAS_NOT_RECALCULATED`).
* Editing a table's header cell renames the table column, so the value must be a unique,
* non-empty name (`XLV_E_INVALID_VALUE` otherwise). A value written into a merged range outside
* its top-left cell gives `XLV_W_MERGED_CELL_WRITE`.
*
* @example
* ```ts
* import { editCells } from 'xlvane';
*
* declare const bytes: Uint8Array;
* const { bytes: updated, warnings } = editCells(bytes, [
*   { sheet: 'Sales', cell: 'B7', value: 1250.5 },
*   { sheet: 'Sales', cell: 'C7', value: 'checked', style: { font: { color: '008000' } } },
* ]);
* console.log(updated.length, warnings.length);
* ```
*/
function editCells(input, edits, options) {
	const bytes = inputBytes(input);
	if (!Array.isArray(edits)) throw new XlvError("XLV_E_INVALID_INPUT", "editCells expects an array of edits.", { hint: "Pass an array such as [{ sheet: \"Data\", cell: \"A1\", value: 1 }]." });
	const o = readOptions(options, ["limits"], "editCells options");
	const limits = resolveLimits(o.get("limits"));
	const workbook = readWorkbook(bytes, { limits });
	return applyEdits$1(bytes, workbook, planEdits(workbook, edits), limits);
}
/**
* Appends rows below the last used row of a sheet (the rows of merged ranges count as used) in
* an existing XLSX file, editing the file surgically like {@link editCells}. Object rows are
* placed under the matching header columns (unknown keys throw `XLV_E_INVALID_ROWS` with
* suggestions); array rows start at the first used column. A table whose last row is the sheet's last used row grows to include the new
* rows (its range and filter); if that table has a totals row the call is refused with
* `XLV_E_NOT_SUPPORTED` and a hint to insert the rows above the totals row with `insertRows`.
* Calculated columns are not filled in the new rows.
*
* @example
* ```ts
* import { appendRows } from 'xlvane';
*
* declare const bytes: Uint8Array;
* const { bytes: updated } = appendRows(bytes, 'Orders', [
*   { Customer: 'Alice', Total: 120.5 },
*   { Customer: 'Bob', Total: 80 },
* ]);
* ```
*/
function appendRows(input, sheet, rows, options) {
	const bytes = inputBytes(input);
	const o = readOptions(options, ["limits", "header"], "appendRows options");
	const limits = resolveLimits(o.get("limits"));
	const workbook = readWorkbook(bytes, { limits });
	const target = pickSheet(workbook, sheet);
	if (!Array.isArray(rows)) throw new XlvError("XLV_E_INVALID_ROWS", "appendRows expects an array of rows.", { hint: "Pass an array of objects or arrays." });
	let lastRow = 0;
	let firstCol = Infinity;
	for (const cell of target.cells()) {
		lastRow = Math.max(lastRow, cell.row);
		firstCol = Math.min(firstCol, cell.col);
	}
	if (firstCol === Infinity) firstCol = 1;
	for (const ref of target.merges) {
		const merge = parseRange(ref);
		if (merge !== null && merge.end.row > lastRow) lastRow = merge.end.row;
	}
	if (lastRow + rows.length > 1048576) throw new XlvError("XLV_E_LIMIT_EXCEEDED", `Appending ${String(rows.length)} rows after row ${String(lastRow)} would pass the last row of the sheet.`, {
		location: { sheet: target.name },
		hint: `A sheet has ${String(MAX_ROWS)} rows. Append fewer rows or use another sheet.`,
		details: {
			limit: "rows",
			value: MAX_ROWS
		}
	});
	let columns = null;
	const edits = [];
	rows.forEach((row, i) => {
		const r = lastRow + 1 + i;
		if (Array.isArray(row)) {
			row.forEach((value, k) => {
				if (value !== null && value !== void 0) edits.push({
					sheet: target.name,
					cell: formatCellRef(r, firstCol + k),
					value
				});
			});
			return;
		}
		if (typeof row !== "object" || row === null) throw new XlvError("XLV_E_INVALID_ROWS", `Row ${String(i + 1)} is not an object or an array.`, {
			location: { sheet: target.name },
			hint: "Pass objects keyed by header text, or arrays of values.",
			details: { row: i + 1 }
		});
		if (columns === null) {
			const header = o.get("header") ?? "auto";
			const table = readTable(workbook, {
				sheet: target.name,
				header
			});
			if (table.headerRow === null) throw new XlvError("XLV_E_HEADER_NOT_FOUND", `Sheet "${target.name}" has no header row to place object rows under.`, {
				location: { sheet: target.name },
				hint: "Pass the header row with { header: 1 }, or pass rows as arrays of values.",
				details: { row: 0 }
			});
			columns = new Map(table.columns.map((c) => [c.key, c.col]));
		}
		const known = columns;
		for (const key of Object.keys(row)) {
			const col = known.get(key);
			if (col === void 0) {
				const close = suggest(key, [...known.keys()]);
				throw new XlvError("XLV_E_INVALID_ROWS", `Row ${String(i + 1)} has the key "${key.slice(0, 40)}", which is not a column of sheet "${target.name}".`, {
					location: { sheet: target.name },
					hint: `${close.length > 0 ? `Did you mean "${close[0] ?? ""}"? ` : ""}Columns: ${[...known.keys()].join(", ")}.`,
					details: {
						row: i + 1,
						key: key.slice(0, 40)
					}
				});
			}
			const value = row[key];
			if (value !== null && value !== void 0) edits.push({
				sheet: target.name,
				cell: formatCellRef(r, col),
				value
			});
		}
	});
	if (edits.length === 0) return {
		bytes: bytes.slice(),
		warnings: []
	};
	return applyEdits$1(bytes, workbook, planEdits(workbook, edits), limits, {
		sheet: target.name,
		lastRow,
		count: rows.length
	});
}
// src/sml/ref-adjust.ts
const REF = "#REF!";
function invalid(option, message, hint) {
	return new XlvError("XLV_E_INVALID_OPTION", message, {
		hint,
		details: { option }
	});
}
function check(change) {
	if (typeof change !== "object" || change === null) throw invalid("change", "The reference change must be an object.", "Pass { sheet: 'Sheet1', axis: 'row', at: 5, count: 2 }.");
	const { sheet, axis, at, count } = change;
	const formulaSheet = change.formulaSheet ?? sheet;
	if (typeof sheet !== "string" || sheet === "") throw invalid("sheet", "The option \"sheet\" must be a sheet name.", "For example { sheet: 'Sales' }.");
	if (typeof formulaSheet !== "string" || formulaSheet === "") throw invalid("formulaSheet", "The option \"formulaSheet\" must be a sheet name.", "For example { formulaSheet: 'Summary' }.");
	if (axis !== "row" && axis !== "col") throw invalid("axis", `The option "axis" must be "row" or "col", not ${JSON.stringify(axis)}.`, "Use { axis: 'row' } for rows or { axis: 'col' } for columns.");
	const max = axis === "row" ? MAX_ROWS : MAX_COLUMNS;
	if (!Number.isInteger(at) || at < 1 || at > max) throw invalid("at", `The option "at" must be a whole number from 1 to ${String(max)}.`, "For example { at: 5 } for row 5 or column E.");
	if (!Number.isInteger(count) || count === 0 || count < 0 && at - count - 1 > max || Math.abs(count) > max) throw invalid("count", `The option "count" must be a whole number other than 0, and a deletion must end at ${String(max)} or before.`, "Use a positive count to insert and a negative count to delete, for example { count: -2 }.");
	let sheets = null;
	if (change.sheets !== void 0) {
		if (!Array.isArray(change.sheets) || !change.sheets.every((s) => typeof s === "string")) throw invalid("sheets", "The option \"sheets\" must be an array of sheet names.", "For example { sheets: ['Jan', 'Feb', 'Mar'] }.");
		sheets = change.sheets.map((s) => s.toLowerCase());
	}
	return {
		sheet: sheet.toLowerCase(),
		formulaSheet: formulaSheet.toLowerCase(),
		row: axis === "row",
		at,
		count,
		max,
		sheets
	};
}
/** New position of one coordinate, or `null` when it is deleted or pushed beyond the grid. */
function moveOne(v, c) {
	if (c.count > 0) {
		const out = v >= c.at ? v + c.count : v;
		return out > c.max ? null : out;
	}
	const end = c.at - c.count - 1;
	if (v < c.at) return v;
	if (v > end) return v + c.count;
	return null;
}
/**
* New ends of a span v1..v2 (v1 ≤ v2), or `null` when it disappears (R-03 … R-12):
* - insertion moves every end at or after `at`; a start pushed beyond the grid gives `null`, an end
*   pushed beyond it stops at the last row or column;
* - deletion cuts the deleted band out; a span entirely inside it gives `null`; a span longer than
*   one row or column whose end is the last one keeps that end.
*/
function moveSpan(v1, v2, c) {
	if (c.count > 0) {
		const a = v1 >= c.at ? v1 + c.count : v1;
		if (a > c.max) return null;
		const b = v2 >= c.at ? v2 + c.count : v2;
		return [a, Math.min(b, c.max)];
	}
	const end = c.at - c.count - 1;
	if (v1 >= c.at && v2 <= end) return null;
	return [v1 < c.at ? v1 : v1 > end ? v1 + c.count : c.at, v2 === c.max && v1 < v2 ? v2 : v2 < c.at ? v2 : v2 > end ? v2 + c.count : c.at - 1];
}
/** Whether a reference with this prefix points into the changed sheet; `null` if unknown. */
function onChangedSheet(prefix, c) {
	if (prefix === null) return c.formulaSheet === c.sheet;
	if (prefix.external) return false;
	const names = prefix.sheets.map((s) => s.toLowerCase());
	if (names.length === 1) return names[0] === c.sheet;
	if (c.sheets === null) return null;
	const i = c.sheets.indexOf(names[0] ?? "");
	const j = c.sheets.indexOf(names[1] ?? "");
	if (i < 0 || j < 0) return null;
	const k = c.sheets.indexOf(c.sheet);
	return k >= Math.min(i, j) && k <= Math.max(i, j) ? null : false;
}
function cellText(p, col, row) {
	return (p.colAbs ? "$" : "") + columnLetters(col) + (p.rowAbs ? "$" : "") + String(row);
}
/** Adjusts a single cell reference; returns its new text or `#REF!`. */
function adjustCell(p, c) {
	const v = moveOne(c.row ? p.row : p.col, c);
	if (v === null) return REF;
	return c.row ? cellText(p, p.col, v) : cellText(p, v, p.row);
}
/** Adjusts a range of two cell references; returns its new text or `#REF!`. */
function adjustCellRange(first, second, c) {
	let a = first;
	let b = second;
	if (c.row ? a.row > b.row : a.col > b.col) {
		a = c.row ? {
			...first,
			row: second.row,
			rowAbs: second.rowAbs
		} : {
			...first,
			col: second.col,
			colAbs: second.colAbs
		};
		b = c.row ? {
			...second,
			row: first.row,
			rowAbs: first.rowAbs
		} : {
			...second,
			col: first.col,
			colAbs: first.colAbs
		};
	}
	const span = moveSpan(c.row ? a.row : a.col, c.row ? b.row : b.col, c);
	if (span === null) return REF;
	return c.row ? `${cellText(a, a.col, span[0])}:${cellText(b, b.col, span[1])}` : `${cellText(a, span[0], a.row)}:${cellText(b, span[1], b.row)}`;
}
function adjustLines(p, c) {
	const rows = p.kind === "rows";
	if (rows !== c.row) return render(p);
	let v1 = rows ? p.r1 : p.c1;
	let v2 = rows ? p.r2 : p.c2;
	let abs1 = rows ? p.r1Abs : p.c1Abs;
	let abs2 = rows ? p.r2Abs : p.c2Abs;
	if (v1 > v2) [v1, v2, abs1, abs2] = [
		v2,
		v1,
		abs2,
		abs1
	];
	const span = moveSpan(v1, v2, c);
	if (span === null) return REF;
	const show = (v, abs) => (abs ? "$" : "") + (rows ? String(v) : columnLetters(v));
	return `${show(span[0], abs1)}:${show(span[1], abs2)}`;
}
function render(p) {
	if (p.kind === "cell") return cellText(p, p.col, p.row);
	if (p.kind === "rows") return `${p.r1Abs ? "$" : ""}${String(p.r1)}:${p.r2Abs ? "$" : ""}${String(p.r2)}`;
	return `${p.c1Abs ? "$" : ""}${columnLetters(p.c1)}:${p.c2Abs ? "$" : ""}${columnLetters(p.c2)}`;
}
function adjustParts(parts, c) {
	let out = "";
	let i = 0;
	while (i < parts.length) {
		let p = parts[i];
		let prefix = null;
		if (typeof p === "string") {
			out += p;
			i++;
			continue;
		}
		if (p.kind === "prefix") {
			prefix = p;
			const next = parts[i + 1];
			if (typeof next === "string" && /^(?:#REF!|[A-Za-z_\\\u0080-\uFFFF])/.test(next)) {
				out += p.text;
				i++;
				continue;
			}
			if (next === void 0 || typeof next === "string" || next.kind === "prefix") return null;
			p = next;
			i++;
		}
		const colon = parts[i + 1];
		const second = parts[i + 2];
		const range = p.kind === "cell" && colon === ":" && second !== void 0 && typeof second !== "string" && second.kind === "cell" ? second : null;
		if (p.kind === "cell" && colon === ":" && second !== void 0 && typeof second !== "string" && second.kind === "prefix") return null;
		const on = onChangedSheet(prefix, c);
		if (on === null) return null;
		const head = prefix?.text ?? "";
		let text;
		if (!on) text = range === null ? render(p) : `${render(p)}:${render(range)}`;
		else if (p.kind === "cell") text = range === null ? adjustCell(p, c) : adjustCellRange(p, range, c);
		else text = adjustLines(p, c);
		if (text === null) return null;
		out += text === REF ? REF : head + text;
		i += range === null ? 1 : 3;
	}
	return out;
}
/**
* Rewrites the references of a formula after rows or columns of a sheet are inserted or deleted,
* as spreadsheet applications do: references at or after the change move, `$`-marked ones too;
* a range grows when rows or columns are inserted inside it and shrinks when some of its rows or
* columns are deleted; a reference whose cells are all deleted becomes `#REF!`. References to
* other sheets, other workbooks, defined names, structured references and text in strings are not
* changed. Returns `null` when the formula cannot be adjusted safely: text the tokenizer does not
* classify with confidence (the same rules as shared formulas, docs/facts/formulas.md), or a 3-D
* reference that may cover the changed sheet. The formula is given without the leading "=".
*
* @example
* ```ts
* import { adjustFormula } from 'xlvane';
*
* // Two rows inserted before row 5 of "Sales".
* adjustFormula('SUM(B2:B10)*$B$12', { sheet: 'Sales', axis: 'row', at: 5, count: 2 }); // 'SUM(B2:B12)*$B$14'
* // Rows 5 and 6 deleted; the formula is on another sheet.
* adjustFormula('Sales!B5+Sales!B9', { sheet: 'Sales', formulaSheet: 'Summary', axis: 'row', at: 5, count: -2 }); // '#REF!+Sales!B7'
* ```
*/
function adjustFormula(text, change) {
	const c = check(change);
	if (typeof text !== "string") throw invalid("text", "The formula must be a string.", "Pass the formula text without the leading '=', for example 'SUM(A1:A3)'.");
	const parts = tokenize(text);
	return parts === null ? null : adjustParts(parts, c);
}
/**
* Rewrites a list of cell ranges of the changed sheet, such as the `sqref` of a conditional format
* or the `ref` of a hyperlink (space-separated ranges like `A1:B5 D7`), after rows or columns are
* inserted or deleted, with the rules of {@link adjustFormula}. Ranges whose cells are all deleted
* are left out, so the result is `''` when nothing remains. Returns `null` if the text is not a
* list of A1 references.
*
* @example
* ```ts
* import { adjustRange } from 'xlvane';
*
* adjustRange('A1:B5 D7', { sheet: 'Sales', axis: 'row', at: 3, count: 1 }); // 'A1:B6 D8'
* adjustRange('A5:A6', { sheet: 'Sales', axis: 'row', at: 5, count: -2 }); // ''
* ```
*/
function adjustRange(sqref, change) {
	const c = check(change);
	if (typeof sqref !== "string") throw invalid("sqref", "The ranges must be a string.", "For example 'A1:B5 D7'.");
	const out = [];
	for (const item of sqref.split(" ")) {
		if (item === "") continue;
		const parts = tokenize(item);
		if (parts === null) return null;
		const isCell = (p) => p !== void 0 && typeof p !== "string" && p.kind === "cell";
		let text;
		if (parts.length === 1 && isCell(parts[0])) text = adjustCell(parts[0], {
			...c,
			formulaSheet: c.sheet
		});
		else if (parts.length === 3 && isCell(parts[0]) && parts[1] === ":" && isCell(parts[2])) {
			const t = adjustCellRange(parts[0], parts[2], c);
			if (t === null) return null;
			text = t;
		} else return null;
		if (text !== REF) out.push(text);
	}
	return out.join(" ");
}
// src/rows.ts
/** Relationship types of a worksheet whose parts do not depend on row positions or are handled here. */
const SAFE_SHEET_RELATIONSHIPS = /* @__PURE__ */ new Set([
	"hyperlink",
	"comments",
	"vmlDrawing",
	"printerSettings",
	"image",
	"customProperty",
	"drawing",
	"table"
]);
/** Elements of a worksheet whose content refers to cells and that xlvane does not rewrite. */
const REFUSED_ELEMENTS = {
	oleObjects: "embedded objects",
	controls: "form controls",
	scenarios: "scenarios",
	cellWatches: "cell watches",
	customSheetViews: "custom sheet views"
};
const WORKAROUND = "Insert or delete the rows in a spreadsheet application, or remove that content first.";
/** A character that cannot be part of a sheet name, used as the "sheet" of defined names. */
const NO_SHEET = "";
function refChange(c, formulaSheet) {
	return {
		sheet: c.sheet,
		formulaSheet,
		axis: "row",
		at: c.at,
		count: c.count,
		sheets: c.sheets
	};
}
/** New number of a row, or `null` when the row is deleted or pushed beyond the last row. */
function moveRow(n, c) {
	if (c.count > 0) {
		if (n < c.at) return n;
		return n + c.count > 1048576 ? null : n + c.count;
	}
	const end = c.at - c.count - 1;
	if (n < c.at) return n;
	return n > end ? n + c.count : null;
}
function decodeEntities(s) {
	return s.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|amp|lt|gt|quot|apos);/g, (_m, e) => {
		if (e === "amp") return "&";
		if (e === "lt") return "<";
		if (e === "gt") return ">";
		if (e === "quot") return "\"";
		if (e === "apos") return "'";
		return String.fromCodePoint(e.startsWith("#x") ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10));
	});
}
function removeAttr(tag, name) {
	return tag.replace(new RegExp(`\\s${name}\\s*=\\s*("[^"]*"|'[^']*')`), "");
}
function content(text, e) {
	return e.start.empty ? "" : text.slice(e.start.end, e.end.start);
}
function notSupported(message, location, details) {
	return new XlvError("XLV_E_NOT_SUPPORTED", message, {
		location,
		hint: WORKAROUND,
		...details ? { details } : {}
	});
}
/**
* Adjusts the formula text of an element (XML-escaped, as in the file). Returns the new escaped
* text, the same text when nothing changes, or `null` when the formula cannot be adjusted although
* it may refer to the changed sheet.
*/
function adjustText(raw, c, formulaSheet) {
	if (raw.includes("<")) return null;
	const text = decodeEntities(raw);
	const out = adjustDecoded(text, c, formulaSheet);
	if (out === null) return null;
	return out === text ? raw : escapeText(out);
}
/** {@link adjustText} for formula text that is already decoded, such as an attribute value. */
function adjustDecoded(text, c, formulaSheet) {
	const out = adjustFormula(text, refChange(c, formulaSheet));
	if (out === null) return formulaSheet.toLowerCase() !== c.sheet.toLowerCase() && !text.includes("!") ? text : null;
	return out;
}
/** `ext@uri` values of the worksheet extensions whose ranges and formulas are rewritten (SR-40, D-F04). */
const KNOWN_EXTENSIONS = new Set([X14_CF_URI, X14_DV_URI].map((u) => u.toLowerCase()));
const isKnownExtension = (e) => KNOWN_EXTENSIONS.has((e.attrs.get("uri") ?? "").trim().toLowerCase());
/** Start tags named `f` or `sqref`, with any prefix: the formulas and ranges an extension can hold. */
const FORMULA_OR_RANGE_TAG = /<(?:[\w.-]+:)?(?:f|sqref)[\s/>]/g;
const CFVO_PARENTS = /* @__PURE__ */ new Set([
	"colorScale",
	"dataBar",
	"iconSet"
]);
const within = (outer, inner) => inner.start.start >= outer.start.start && inner.end.end <= outer.end.end;
/**
* Applies non-overlapping replacements in one pass (linear in the text size, unlike edit.ts
* `splice`, which matters for sheets with many moved cells). An edit inside an earlier replaced
* span, such as a removed row, is dropped with it; an insertion at the start of a removed element
* is kept.
*/
function applyEdits(text, edits) {
	const sorted = [...edits].sort((a, b) => a[0] - b[0] || Number(b[0] === b[1]) - Number(a[0] === a[1]) || b[1] - a[1]);
	const out = [];
	let pos = 0;
	for (const [start, end, replacement] of sorted) {
		if (start < pos) continue;
		out.push(text.slice(pos, start), replacement);
		pos = end;
	}
	out.push(text.slice(pos));
	return out.join("");
}
const KEEP_DEPTH3 = /* @__PURE__ */ new Set([
	"row",
	"mergeCell",
	"hyperlink",
	"cfRule",
	"dataValidation",
	"brk",
	"ignoredError",
	"protectedRange"
]);
/** Rewrites one worksheet part: references everywhere, and row positions on the changed sheet. */
function rewriteSheet(text, part, sheet, c, target, style, limits) {
	const { elements } = scan(text, part, limits, (local, depth, parent, ns) => depth === 2 || depth === 3 && (KEEP_DEPTH3.has(local) || local === "ext" && parent === "extLst") || depth === 4 && (local === "c" && parent === "row" || local === "formula" && parent === "cfRule" || (local === "formula1" || local === "formula2") && parent === "dataValidation") || depth === 5 && (local === "f" && parent === "c" || local === "cfvo" && CFVO_PARENTS.has(parent)) || ns === "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" && (depth === 4 && (local === "conditionalFormattings" || local === "dataValidations") || depth === 5 && (local === "conditionalFormatting" || local === "dataValidation") || depth === 6 && (local === "cfRule" || local === "id" && parent === "ext")) || ns === "http://schemas.microsoft.com/office/excel/2006/main" && (depth === 6 && local === "sqref" || depth >= 7 && local === "f"));
	const edits = [];
	const tagOf = (e) => text.slice(e.start.start, e.start.end);
	const replaceTag = (e, tag) => {
		edits.push([
			e.start.start,
			e.start.end,
			tag
		]);
	};
	const removeElement = (e) => {
		edits.push([
			e.start.start,
			e.end.end,
			""
		]);
	};
	const where = (cell) => cell === void 0 ? {
		part,
		sheet: sheet.name
	} : {
		part,
		sheet: sheet.name,
		cell
	};
	if (target) for (const e of elements) {
		if (e.depth !== 2) continue;
		const what = REFUSED_ELEMENTS[e.local];
		if (what !== void 0) throw notSupported(`Sheet "${sheet.name}" has ${what}, which xlvane cannot move with its rows yet.`, where(), { element: e.local });
	}
	const knownExts = /* @__PURE__ */ new Set();
	{
		let current = null;
		let found = 0;
		const close = () => {
			if (current !== null && (text.slice(current.start.start, current.end.end).match(FORMULA_OR_RANGE_TAG) ?? []).length === found) knownExts.add(current);
		};
		for (const e of elements) if (e.depth === 3 && e.local === "ext") {
			close();
			current = isKnownExtension(e) ? e : null;
			found = 0;
		} else if (current !== null && e.ns === "http://schemas.microsoft.com/office/excel/2006/main" && (e.depth === 6 && e.local === "sqref" || e.depth >= 7 && e.local === "f") && within(current, e)) found++;
		close();
	}
	for (const ext of elements) {
		if (ext.depth !== 3 || ext.local !== "ext" || knownExts.has(ext)) continue;
		const body = text.slice(ext.start.start, ext.end.end);
		if (target && /sqref|<xm:f|sparkline|slicer/i.test(body)) throw notSupported(`Sheet "${sheet.name}" has extended content (such as sparklines or slicers) that refers to cells, which xlvane cannot move with its rows yet.`, where(), { element: "extLst" });
		if (!target && /<xm:f|sparkline/i.test(body) && body.toLowerCase().includes(c.sheet.toLowerCase())) throw notSupported(`Sheet "${sheet.name}" has extended content (such as sparklines) that may refer to sheet "${c.sheet}", which xlvane cannot rewrite yet.`, where(), { element: "extLst" });
	}
	const rows = [];
	let current = null;
	let currentRow = 0;
	let lastCol = 0;
	let lastCell = null;
	const formulas = [];
	const invalidRef = (ref) => {
		throw new XlvError("XLV_E_INVALID_CELL_REF", `Sheet "${sheet.name}" contains the invalid reference "${ref.slice(0, 20)}", so its rows cannot be moved safely.`, {
			location: where(),
			hint: "Open and re-save the file in a spreadsheet application, which drops the damaged cells, then try again.",
			details: { ref: ref.slice(0, 40) }
		});
	};
	let hasFormulas = false;
	let extension = null;
	for (const e of elements) {
		if (e.depth === 3 && e.local === "ext") extension = knownExts.has(e) ? e : null;
		if (e.depth === 3 && e.local === "row") {
			const r = e.attrs.get("r");
			const n = r === void 0 ? currentRow + 1 : /^[0-9]{1,7}$/.test(r.trim()) ? Number(r.trim()) : NaN;
			if (!Number.isInteger(n) || n < 1 || n > 1048576) invalidRef(r ?? String(n));
			currentRow = n;
			current = {
				number: n,
				hasR: r !== void 0,
				element: e,
				cells: []
			};
			rows.push(current);
			lastCol = 0;
		} else if (e.depth === 4 && e.local === "c" && current) {
			const r = e.attrs.get("r");
			const address = r === void 0 || r === "" ? null : parseCellRef(r);
			if (r !== void 0 && r !== "" && address === null) invalidRef(r);
			const col = address?.col ?? lastCol + 1;
			lastCol = col;
			lastCell = {
				col,
				hasR: r !== void 0 && r !== "",
				element: e,
				f: null
			};
			current.cells.push(lastCell);
		} else if (e.depth === 5 && e.local === "f" && lastCell && current) {
			lastCell.f = e;
			hasFormulas = true;
			formulas.push({
				cell: lastCell,
				row: current
			});
		} else if (e.depth === 4 && (e.local === "formula" || e.local === "formula1" || e.local === "formula2")) {
			const raw = content(text, e);
			const out = adjustText(raw, c, sheet.name);
			if (out === null) throw notSupported(`A ${e.local === "formula" ? "conditional format" : "data validation"} formula of sheet "${sheet.name}" cannot be adjusted safely: "${decodeEntities(raw).slice(0, 60)}".`, where(), { formula: decodeEntities(raw).slice(0, 80) });
			if (out !== raw) edits.push([
				e.start.end,
				e.end.start,
				out
			]);
		} else if (e.depth === 5 && e.local === "cfvo") {
			const val = e.attrs.get("val");
			if (val === void 0) continue;
			const out = adjustDecoded(val, c, sheet.name);
			if (out === null) throw notSupported(`A conditional format threshold of sheet "${sheet.name}" cannot be adjusted safely: "${val.slice(0, 60)}".`, where(), { formula: val.slice(0, 80) });
			if (out !== val) replaceTag(e, setAttr(tagOf(e), "val", out));
		} else if (e.ns === "http://schemas.microsoft.com/office/excel/2006/main" && e.local === "f" && extension !== null && within(extension, e)) {
			const raw = content(text, e);
			const out = adjustText(raw, c, sheet.name);
			if (out === null) throw notSupported(`A ${(extension.attrs.get("uri") ?? "").toLowerCase() === "{CCE6A557-97BC-4B89-ADB6-D9C93CAAB3DF}".toLowerCase() ? "data validation" : "conditional format"} formula of sheet "${sheet.name}" cannot be adjusted safely: "${decodeEntities(raw).slice(0, 60)}".`, where(), { formula: decodeEntities(raw).slice(0, 80) });
			if (out !== raw) edits.push([
				e.start.end,
				e.end.start,
				out
			]);
		} else if (e.depth === 3 && e.local === "hyperlink") {
			const location = e.attrs.get("location");
			let tag = tagOf(e);
			if (location !== void 0) {
				const out = adjustFormula(location, refChange(c, sheet.name));
				if (out !== null && out !== location) tag = setAttr(tag, "location", out);
			}
			if (target) {
				const ref = adjustRange(e.attrs.get("ref") ?? "", refChange(c, sheet.name));
				if (ref === "") {
					removeElement(e);
					continue;
				}
				if (ref !== null && ref !== e.attrs.get("ref")) tag = setAttr(tag, "ref", ref);
			}
			if (tag !== tagOf(e)) replaceTag(e, tag);
		}
	}
	const masters = /* @__PURE__ */ new Map();
	for (const { cell, row } of formulas) {
		const f = cell.f;
		if (f.attrs.get("t") === "shared" && f.attrs.has("ref") && f.attrs.has("si")) masters.set(f.attrs.get("si") ?? "", {
			row: row.number,
			col: cell.col,
			text: content(text, f)
		});
	}
	const expand = /* @__PURE__ */ new Set();
	for (const [si, m] of masters) {
		const out = adjustText(m.text, c, sheet.name);
		if (target || out === null || out !== m.text) expand.add(si);
	}
	const deletedRow = (n) => c.count < 0 && moveRow(n, c) === null;
	for (const { cell, row } of formulas) {
		const f = cell.f;
		if (target && deletedRow(row.number)) continue;
		const ref = formatCellRef(row.number, cell.col);
		const kind = f.attrs.get("t");
		const failed = (formula) => {
			throw notSupported(`The formula in ${ref} of sheet "${sheet.name}" cannot be adjusted safely: "${formula.slice(0, 60)}".`, where(ref), { formula: formula.slice(0, 80) });
		};
		if (kind === "shared") {
			const si = f.attrs.get("si") ?? "";
			const master = masters.get(si);
			if (master === void 0 || !expand.has(si)) continue;
			const source = decodeEntities(master.text);
			const own = shiftFormula(source, row.number - master.row, cell.col - master.col);
			if (own === null) failed(source);
			const adjusted = adjustFormula(own, refChange(c, sheet.name));
			if (adjusted === null) failed(own);
			let tag = removeAttr(removeAttr(removeAttr(tagOf(f), "t"), "ref"), "si");
			if (tag.endsWith("/>")) tag = `${tag.slice(0, -2).trimEnd()}>`;
			edits.push([
				f.start.start,
				f.end.end,
				`${tag}${escapeText(adjusted)}</${f.start.qname}>`
			]);
			continue;
		}
		if (kind === "dataTable") {
			if (target) throw notSupported(`Cell ${ref} of sheet "${sheet.name}" is part of a data table, which xlvane cannot move.`, where(ref));
			continue;
		}
		const raw = content(text, f);
		const out = adjustText(raw, c, sheet.name);
		if (out === null) failed(decodeEntities(raw));
		let tag = tagOf(f);
		if (kind === "array" && target) {
			const range = parseRange(f.attrs.get("ref") ?? "");
			if (range !== null) {
				const end = c.at - c.count - 1;
				if (c.count > 0 ? c.at > range.start.row && c.at <= range.end.row : !(range.start.row >= c.at && range.end.row <= end) && range.end.row >= c.at && range.start.row <= end) throw notSupported(`The array formula in ${f.attrs.get("ref") ?? ref} of sheet "${sheet.name}" would be split by the change.`, where(ref));
				const moved = adjustRange(f.attrs.get("ref") ?? "", refChange(c, sheet.name));
				if (moved !== null && moved !== "" && moved !== f.attrs.get("ref")) tag = setAttr(tag, "ref", moved);
			}
		}
		if (tag !== tagOf(f)) replaceTag(f, tag);
		if (out !== raw) edits.push([
			f.start.end,
			f.end.start,
			out
		]);
	}
	if (!target) return {
		text: edits.length === 0 ? text : applyEdits(text, edits),
		hasFormulas
	};
	const force = rows.some((r) => !r.hasR || r.cells.some((x) => !x.hasR));
	for (const row of rows) {
		const n = moveRow(row.number, c);
		if (n === null) {
			if (c.count > 0 && row.cells.some((x) => !x.element.start.empty)) throw new XlvError("XLV_E_LIMIT_EXCEEDED", `Inserting ${String(c.count)} rows would push row ${String(row.number)} of sheet "${sheet.name}", which holds data, past the last row of the sheet.`, {
				location: where(),
				hint: `A sheet has ${String(MAX_ROWS)} rows. Delete the rows at the end of the sheet first, or insert fewer rows.`,
				details: {
					limit: "rows",
					value: MAX_ROWS
				}
			});
			removeElement(row.element);
			continue;
		}
		if (n === row.number && !force) continue;
		replaceTag(row.element, setAttr(tagOf(row.element), "r", String(n)));
		for (const cell of row.cells) replaceTag(cell.element, setAttr(tagOf(cell.element), "r", formatCellRef(n, cell.col)));
	}
	if (c.count > 0 && style !== "none") {
		const source = rows.find((r) => r.number === (style === "above" ? c.at - 1 : c.at));
		if (source) {
			const prefix = source.element.start.qname.slice(0, source.element.start.qname.length - 3);
			const keep = [
				"s",
				"customFormat",
				"ht",
				"customHeight"
			].filter((a) => source.element.attrs.has(a));
			let created = "";
			for (let i = 0; i < c.count; i++) {
				const n = c.at + i;
				const attrs = keep.map((a) => ` ${a}="${escapeAttr(source.element.attrs.get(a) ?? "")}"`).join("");
				const cells = source.cells.filter((x) => (x.element.attrs.get("s") ?? "0") !== "0").map((x) => `<${prefix}c r="${formatCellRef(n, x.col)}" s="${escapeAttr(x.element.attrs.get("s") ?? "0")}"/>`).join("");
				created += cells === "" ? `<${prefix}row r="${String(n)}"${attrs}/>` : `<${prefix}row r="${String(n)}"${attrs}>${cells}</${prefix}row>`;
			}
			const next = rows.find((r) => r.number >= c.at);
			if (next) edits.push([
				next.element.start.start,
				next.element.start.start,
				created
			]);
			else {
				const sheetData = elements.find((e) => e.depth === 2 && e.local === "sheetData");
				if (sheetData && !sheetData.start.empty) edits.push([
					sheetData.end.start,
					sheetData.end.start,
					created
				]);
				else if (sheetData) edits.push([
					sheetData.start.start,
					sheetData.start.end,
					`${tagOf(sheetData).slice(0, -2).trimEnd()}>${created}</${sheetData.start.qname}>`
				]);
			}
		}
	}
	const removedIds = /* @__PURE__ */ new Set();
	const linkIds = /* @__PURE__ */ new Map();
	let format = null;
	for (const e of elements) if (e.depth === 2) format = e.local === "conditionalFormatting" ? e : null;
	else if (format !== null && e.depth === 6 && e.ns === "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" && e.local === "id") {
		const list = linkIds.get(format) ?? [];
		list.push(content(text, e).trim().toLowerCase());
		linkIds.set(format, list);
	}
	const containers = /* @__PURE__ */ new Map();
	let parent2 = null;
	for (const e of elements) {
		if (e.depth === 2) {
			parent2 = e;
			if (e.local === "dimension" || e.local === "autoFilter") {
				const ref = e.attrs.get("ref");
				if (ref === void 0) continue;
				const out = adjustRange(ref, refChange(c, sheet.name));
				if (out === null || out === ref) continue;
				if (e.local === "dimension") replaceTag(e, setAttr(tagOf(e), "ref", out === "" ? "A1" : out.split(" ")[0] ?? "A1"));
				else if (out === "") removeElement(e);
				else replaceTag(e, setAttr(tagOf(e), "ref", out));
			} else if (e.local === "conditionalFormatting") {
				const sqref = e.attrs.get("sqref") ?? "";
				const out = adjustRange(sqref, refChange(c, sheet.name));
				if (out === null) throw notSupported(`A conditional format of sheet "${sheet.name}" has a range xlvane cannot adjust: "${sqref.slice(0, 60)}".`, where());
				if (out === "") {
					removeElement(e);
					for (const id of linkIds.get(e) ?? []) removedIds.add(id);
				} else if (out !== sqref) replaceTag(e, setAttr(tagOf(e), "sqref", out));
			}
			continue;
		}
		if (e.depth !== 3 || parent2 === null) continue;
		const container = parent2;
		const count = containers.get(container) ?? {
			kept: 0,
			total: 0
		};
		containers.set(container, count);
		if (e.local === "mergeCell") {
			count.total++;
			const ref = e.attrs.get("ref") ?? "";
			const range = parseRange(ref);
			const out = range === null ? null : c.count < 0 && deletedRow(range.start.row) ? "" : adjustRange(ref, refChange(c, sheet.name));
			if (out === null) {
				count.kept++;
				continue;
			}
			const moved = out === "" ? null : parseRange(out);
			if (moved === null || moved.start.row === moved.end.row && moved.start.col === moved.end.col) {
				removeElement(e);
				continue;
			}
			count.kept++;
			if (out !== ref) replaceTag(e, setAttr(tagOf(e), "ref", out));
		} else if (e.local === "dataValidation" || e.local === "ignoredError" || e.local === "protectedRange") {
			count.total++;
			const sqref = e.attrs.get("sqref") ?? "";
			const out = adjustRange(sqref, refChange(c, sheet.name));
			if (out === null) throw notSupported(`Sheet "${sheet.name}" has a ${e.local} range xlvane cannot adjust: "${sqref.slice(0, 60)}".`, where());
			if (out === "") {
				removeElement(e);
				continue;
			}
			count.kept++;
			if (out !== sqref) replaceTag(e, setAttr(tagOf(e), "sqref", out));
		} else if (e.local === "hyperlink") {
			count.total++;
			if (adjustRange(e.attrs.get("ref") ?? "", refChange(c, sheet.name)) !== "") count.kept++;
		} else if (e.local === "brk" && container.local === "rowBreaks") {
			count.total++;
			const id = Number(e.attrs.get("id") ?? "NaN");
			if (!Number.isInteger(id)) {
				count.kept++;
				continue;
			}
			const n = moveRow(id, c);
			if (n === null) {
				removeElement(e);
				continue;
			}
			count.kept++;
			if (n !== id) replaceTag(e, setAttr(tagOf(e), "id", String(n)));
		}
	}
	for (const [container, n] of containers) {
		if (n.total === 0 || n.kept === n.total) continue;
		const counted = [
			"mergeCells",
			"dataValidations",
			"rowBreaks"
		].includes(container.local);
		if (n.kept === 0) removeElement(container);
		else if (counted) {
			let tag = tagOf(container);
			if (container.attrs.has("count")) tag = setAttr(tag, "count", String(n.kept));
			if (container.local === "rowBreaks" && container.attrs.has("manualBreakCount")) tag = setAttr(tag, "manualBreakCount", String(n.kept));
			if (tag !== tagOf(container)) replaceTag(container, tag);
		}
	}
	rewriteExtensionRanges(text, elements, knownExts, c, sheet, removedIds, edits, where);
	const result = applyEdits(text, edits);
	return {
		text: /sortState/.test(result) ? rewriteSortStates(result, part, sheet, c, limits) : result,
		hasFormulas
	};
}
/**
* Moves the ranges (`xm:sqref`) of the extension conditional formats and data validations of the
* changed sheet (SR-40, SR-44; R-20; D-F04). An item whose cells are all deleted is removed, as is
* an extension rule linked to a removed classic rule (SR-41); emptied containers, `ext` and
* `extLst` elements are removed, and `x14:dataValidations@count` is updated.
*/
function rewriteExtensionRanges(text, elements, knownExts, c, sheet, removedIds, edits, where) {
	const roots = [];
	const stack = [];
	for (const e of elements) {
		if (!(e.depth === 2 && e.local === "extLst" || e.depth === 3 && e.local === "ext" || e.ns === "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" && e.depth >= 4 && e.depth <= 6 || e.ns === "http://schemas.microsoft.com/office/excel/2006/main" && e.depth === 6 && e.local === "sqref")) continue;
		while (stack.length > 0 && !within(stack[stack.length - 1].e, e)) stack.pop();
		const node = {
			e,
			children: []
		};
		(stack.length > 0 ? stack[stack.length - 1].children : roots).push(node);
		stack.push(node);
	}
	const tagOf = (e) => text.slice(e.start.start, e.start.end);
	const remove = (e) => {
		edits.push([
			e.start.start,
			e.end.end,
			""
		]);
	};
	for (const list of roots) {
		if (list.e.local !== "extLst") continue;
		let extsKept = 0;
		for (const ext of list.children) {
			if (ext.e.local !== "ext" || !knownExts.has(ext.e)) {
				extsKept++;
				continue;
			}
			let containersKept = 0;
			let containersTotal = 0;
			for (const container of ext.children) {
				if (container.e.depth !== 4 || container.e.ns !== "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main") continue;
				containersTotal++;
				let kept = 0;
				let total = 0;
				for (const item of container.children) {
					if (item.e.depth !== 5 || item.e.ns !== "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main") continue;
					total++;
					let gone = false;
					const sqref = item.children.find((n) => n.e.ns === "http://schemas.microsoft.com/office/excel/2006/main" && n.e.local === "sqref")?.e;
					if (sqref) {
						const normal = decodeEntities(content(text, sqref)).trim().split(/\s+/).join(" ");
						const out = normal === "" ? "" : adjustRange(normal, refChange(c, sheet.name));
						if (out === null) throw notSupported(`An extension ${item.e.local === "dataValidation" ? "data validation" : "conditional format"} of sheet "${sheet.name}" has a range xlvane cannot adjust: "${normal.slice(0, 60)}".`, where());
						if (out === "" && normal !== "") gone = true;
						else if (out !== normal) edits.push([
							sqref.start.end,
							sqref.end.start,
							escapeText(out)
						]);
					}
					if (!gone && removedIds.size > 0) {
						const rules = item.children.filter((n) => n.e.ns === "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" && n.e.local === "cfRule").map((n) => n.e);
						const linked = rules.filter((r) => removedIds.has((r.attrs.get("id") ?? "").trim().toLowerCase()));
						if (linked.length > 0 && linked.length === rules.length) gone = true;
						else for (const r of linked) remove(r);
					}
					if (gone) remove(item.e);
					else kept++;
				}
				if (kept === 0 && total > 0) {
					remove(container.e);
					continue;
				}
				containersKept++;
				if (kept < total && container.e.attrs.has("count")) edits.push([
					container.e.start.start,
					container.e.start.end,
					setAttr(tagOf(container.e), "count", String(kept))
				]);
			}
			if (containersTotal > 0 && containersKept === 0) remove(ext.e);
			else extsKept++;
		}
		if (list.children.length > 0 && extsKept === 0) remove(list.e);
	}
}
/**
* Moves the sort ranges of the changed sheet (`sortState` at sheet level or inside `autoFilter`,
* and their `sortCondition`s) as ranges move (D-B01). A sort range that is deleted entirely is
* removed; one that cannot be adjusted is refused. Runs on the text after the other edits, so
* a sort state inside a removed filter is already gone.
*/
function rewriteSortStates(text, part, sheet, c, limits) {
	const { elements } = scan(text, part, limits, (local) => local === "sortState" || local === "sortCondition");
	const edits = [];
	const removed = [];
	for (const e of elements) {
		if (removed.some(([a, b]) => e.start.start >= a && e.end.end <= b)) continue;
		const ref = e.attrs.get("ref");
		if (ref === void 0) continue;
		const out = adjustRange(ref, refChange(c, sheet.name));
		if (out === null) throw notSupported(`A sort range of sheet "${sheet.name}" cannot be adjusted: "${ref.slice(0, 60)}".`, {
			part,
			sheet: sheet.name
		});
		if (out === "") {
			edits.push([
				e.start.start,
				e.end.end,
				""
			]);
			removed.push([e.start.start, e.end.end]);
		} else if (out !== ref) edits.push([
			e.start.start,
			e.start.end,
			setAttr(text.slice(e.start.start, e.start.end), "ref", out)
		]);
	}
	return edits.length === 0 ? text : applyEdits(text, edits);
}
/**
* Reads the `row` and `rowOff` of the first `from` or `to` marker of an anchor with `indexOf`
* only, so that the work is linear in the anchor's size whatever its content (a regular
* expression here backtracked cubically on runs of white space, D-R03).
*/
function readMarker(anchor, pre, local) {
	const open = `<${pre}${local}>`;
	const start = anchor.indexOf(open);
	if (start < 0) return null;
	const close = anchor.indexOf(`</${pre}${local}>`, start + open.length);
	if (close < 0) return null;
	const child = (name) => {
		const tag = `<${pre}${name}>`;
		const a = anchor.indexOf(tag, start + open.length);
		if (a < 0 || a > close) return null;
		const b = anchor.indexOf(`</${pre}${name}>`, a + tag.length);
		if (b < 0 || b > close) return null;
		return {
			value: anchor.slice(a + tag.length, b),
			start: a + tag.length,
			end: b
		};
	};
	const row = child("row");
	const rowOff = child("rowOff");
	if (row === null || rowOff === null || !/^\s*\d+\s*$/.test(row.value)) return null;
	return {
		row: Number(row.value),
		rowOff: Number(rowOff.value.trim()) || 0,
		rowStart: row.start,
		rowEnd: row.end
	};
}
/**
* Calls `change` for each `<v:shape …>…</v:shape>` of a VML part and joins the results, finding
* the elements with `indexOf` so that the work is linear even in a part that is not well-formed
* (a lazy regular expression here was quadratic on unclosed shapes, D-R03).
*/
function eachVmlShape(text, change) {
	const out = [];
	let pos = 0;
	let from = 0;
	for (;;) {
		const at = text.indexOf("<v:shape", from);
		if (at < 0) break;
		const next = text.charCodeAt(at + 8);
		if (!(next === 32 || next === 62 || next === 47 || next === 9 || next === 10 || next === 13)) {
			from = at + 8;
			continue;
		}
		const end = text.indexOf("</v:shape>", at + 8);
		if (end < 0) break;
		const stop = end + 10;
		out.push(text.slice(pos, at), change(text.slice(at, stop)));
		pos = stop;
		from = stop;
	}
	out.push(text.slice(pos));
	return out.join("");
}
/** Rewrites the notes of the changed sheet: the Comments part and the note shapes of its VML part. */
function rewriteNotes(edit, pkg, sheetPart, sheetText, sheet, c) {
	const rels = pkg.relationships(sheetPart);
	const comments = rels.find((r) => !r.external && classifyRelationship(r.type)?.name === "comments");
	if (comments !== void 0) {
		const text = edit.text(comments.target);
		if (text !== void 0) {
			const { elements } = scan(text, comments.target, edit.limits, (local, depth) => depth === 3 && local === "comment");
			const edits = [];
			for (const e of elements) {
				const ref = e.attrs.get("ref") ?? "";
				const out = adjustRange(ref, refChange(c, sheet.name));
				if (out === null || out === ref) continue;
				if (out === "") edits.push([
					e.start.start,
					e.end.end,
					""
				]);
				else edits.push([
					e.start.start,
					e.start.end,
					setAttr(text.slice(e.start.start, e.start.end), "ref", out)
				]);
			}
			if (edits.length > 0) edit.write(comments.target, applyEdits(text, edits));
		}
	}
	const id = /<(?:[A-Za-z_][\w.-]*:)?legacyDrawing\b[^>]*?\br:id\s*=\s*["']([^"']*)["']/.exec(sheetText)?.[1];
	const vml = id === void 0 ? void 0 : rels.find((r) => r.id === id && !r.external);
	if (vml === void 0) return;
	const text = edit.text(vml.target);
	if (text === void 0) return;
	let changed = false;
	const out = eachVmlShape(text, (shape) => {
		if (!/ObjectType\s*=\s*["']Note["']/.test(shape)) return shape;
		const row = /<x:Row>\s*(\d+)\s*<\/x:Row>/.exec(shape);
		if (row === null) return shape;
		const n = moveRow(Number(row[1]) + 1, c);
		changed = true;
		if (n === null) return "";
		const moveAnchor = (v) => {
			const one = v + 1;
			if (c.count > 0) return one >= c.at ? v + c.count : v;
			const end = c.at - c.count - 1;
			return one < c.at ? v : one > end ? v + c.count : c.at - 1;
		};
		return shape.replace(/<x:Row>\s*\d+\s*<\/x:Row>/, `<x:Row>${String(n - 1)}</x:Row>`).replace(/<x:Anchor>([^<]*)<\/x:Anchor>/, (all, list) => {
			const v = list.split(",").map((s) => s.trim());
			if (v.length !== 8 || !v.every((s) => /^\d+$/.test(s))) return all;
			const nums = v.map(Number);
			nums[2] = moveAnchor(nums[2] ?? 0);
			nums[6] = Math.max(nums[2] ?? 0, moveAnchor(nums[6] ?? 0));
			return `<x:Anchor>${nums.join(", ")}</x:Anchor>`;
		});
	});
	if (changed) edit.write(vml.target, out);
}
/**
* Moves the pictures of the changed sheet's drawing with their rows (TD-38): the `row` of each
* `from` and `to` marker of two-cell and one-cell anchors; absolute anchors do not depend on
* rows. Drawings with charts, shapes, groups or other objects are refused, and so is a deletion
* that would cut a picture (a deleted row inside the rows it covers or under its anchor cell).
*/
function rewriteDrawing(edit, pkg, sheetPart, sheetText, sheet, c) {
	const id = /<(?:[A-Za-z_][\w.-]*:)?drawing\b[^>]*?\br:id\s*=\s*["']([^"']*)["']/.exec(sheetText)?.[1];
	const rels = pkg.relationships(sheetPart);
	const rel = id === void 0 ? void 0 : rels.find((r) => r.id === id && !r.external && classifyRelationship(r.type)?.name === "drawing");
	if (rel === void 0) return;
	const text = edit.text(rel.target);
	if (text === void 0) return;
	const where = {
		part: rel.target,
		sheet: sheet.name
	};
	for (const r of pkg.relationships(rel.target)) {
		const name = classifyRelationship(r.type)?.name;
		if (!r.external && name !== "image" && name !== "hyperlink") throw notSupported(`The drawing of sheet "${sheet.name}" has a related part of type "${name ?? r.type}" (${r.target}), which xlvane cannot move with its rows yet.`, where, { type: r.type });
	}
	const { elements } = scan(text, rel.target, edit.limits, (_local, depth) => depth === 2);
	const anchorKinds = /* @__PURE__ */ new Set([
		"twoCellAnchor",
		"oneCellAnchor",
		"absoluteAnchor"
	]);
	const others = () => notSupported(`The drawing of sheet "${sheet.name}" has charts, shapes or other objects besides pictures, which xlvane cannot move with its rows yet.`, where);
	for (const e of elements) {
		if (!anchorKinds.has(e.local) || e.start.empty) throw others();
		const pre = e.start.qname.slice(0, e.start.qname.length - e.local.length);
		if (!new RegExp(`<${pre.replace(/[.]/g, "\\.")}pic[\\s>]`).test(text.slice(e.start.start, e.end.end))) throw others();
	}
	const end = c.at - c.count - 1;
	const moveMarker = (m, boundary = false) => {
		const one = boundary ? m : m + 1;
		if (c.count > 0) return one >= c.at ? m + c.count : m;
		return one < c.at ? m : one > end ? m + c.count : c.at - 1;
	};
	const edits = [];
	for (const e of elements) {
		const kind = e.local;
		if (kind === "absoluteAnchor") continue;
		const anchor = text.slice(e.start.start, e.end.end);
		const pre = e.start.qname.slice(0, e.start.qname.length - kind.length);
		const marker = (local) => readMarker(anchor, pre, local);
		const from = marker("from");
		const to = kind === "twoCellAnchor" ? marker("to") : null;
		if (from === null || kind === "twoCellAnchor" && to === null) throw notSupported(`A picture of sheet "${sheet.name}" has a position xlvane cannot read, so its rows cannot be moved.`, where);
		const first = from.row + 1;
		const last = to === null ? first : to.rowOff === 0 && to.row > from.row ? to.row : to.row + 1;
		if (c.count < 0 && first <= end && last >= c.at) throw notSupported(`Deleting rows ${String(c.at)} to ${String(end)} of sheet "${sheet.name}" would cut a picture (rows ${String(first)} to ${String(last)}), which xlvane does not do.`, where, { cell: formatCellRef(first, 1) });
		const rows = [moveMarker(from.row), to === null ? null : moveMarker(to.row, to.rowOff === 0 && to.row > from.row)];
		if (rows.some((r) => r !== null && r > 1048575)) throw notSupported(`Inserting rows would push a picture of sheet "${sheet.name}" beyond the last row.`, where);
		if (rows[0] === from.row && (to === null || rows[1] === to.row)) continue;
		const markerEdits = [[
			from.rowStart,
			from.rowEnd,
			String(rows[0])
		]];
		if (to !== null && rows[1] !== null && rows[1] !== void 0) markerEdits.push([
			to.rowStart,
			to.rowEnd,
			String(rows[1])
		]);
		edits.push([
			e.start.start,
			e.end.end,
			applyEdits(anchor, markerEdits)
		]);
	}
	if (edits.length > 0) edit.write(rel.target, applyEdits(text, edits));
}
/**
* Moves, grows or shrinks the tables of a sheet with its rows (TD-11): the table `ref` and its
* `autoFilter` follow as ranges do (adjustRange), formulas of calculated columns and totals are
* adjusted like cell formulas, and structured references are left as they are. A deletion that
* would remove a table's header row, totals row or all of its data rows is refused. For tables
* of other sheets (`target` false) only the formulas are adjusted.
*/
function rewriteTables(edit, pkg, sheetPart, sheet, c, target) {
	for (const rel of pkg.relationships(sheetPart)) {
		if (rel.external || classifyRelationship(rel.type)?.name !== "table") continue;
		const text = edit.text(rel.target);
		if (text === void 0) continue;
		const where = {
			part: rel.target,
			sheet: sheet.name
		};
		const { root, elements } = scan(text, rel.target, edit.limits, (local, depth, parent) => depth === 2 && local === "autoFilter" || depth === 4 && parent === "tableColumn" && (local === "calculatedColumnFormula" || local === "totalsRowFormula"));
		const name = root.attrs.get("displayName") ?? root.attrs.get("name") ?? "";
		const edits = [];
		let derivedFilter = null;
		if (target) {
			const ref = root.attrs.get("ref") ?? "";
			const range = parseRange(ref.trim().toUpperCase());
			if (range === null) throw notSupported(`Table "${name}" of sheet "${sheet.name}" has a range xlvane cannot read ("${ref.slice(0, 20)}").`, where);
			const header = (root.attrs.get("headerRowCount") ?? "1").trim() !== "0" ? 1 : 0;
			const totals = /^[1-9]/.test((root.attrs.get("totalsRowCount") ?? "0").trim()) ? 1 : 0;
			if (c.count < 0) {
				const end = c.at - c.count - 1;
				const cuts = (row) => row >= c.at && row <= end;
				const dataFirst = range.start.row + header;
				const dataLast = range.end.row - totals;
				if (header === 1 && cuts(range.start.row) || totals === 1 && cuts(range.end.row) || c.at <= dataFirst && end >= dataLast && range.start.row <= end && range.end.row >= c.at) throw notSupported(`Deleting rows ${String(c.at)} to ${String(end)} of sheet "${sheet.name}" would remove the header row, the totals row or every data row of table "${name}" (${formatRange(range)}).`, where, { table: name });
			}
			const newRef = adjustRange(ref, refChange(c, sheet.name));
			if (newRef === null || newRef === "") throw notSupported(`Table "${name}" of sheet "${sheet.name}" cannot be moved with its rows.`, where, { table: name });
			const moved = parseRange(newRef);
			if (moved !== null && moved.end.row - moved.start.row < range.end.row - range.start.row && c.count > 0) throw notSupported(`Inserting rows would push table "${name}" beyond the last row.`, where, { table: name });
			if (header === 1 && moved !== null) {
				const filterOf = (r) => formatRange({
					start: r.start,
					end: {
						col: r.end.col,
						row: r.end.row - totals
					}
				});
				derivedFilter = {
					from: filterOf(range),
					to: filterOf(moved)
				};
			}
			if (newRef !== ref) edits.push([
				root.start.start,
				root.start.end,
				setAttr(text.slice(root.start.start, root.start.end), "ref", newRef)
			]);
		}
		for (const e of elements) {
			if (e.local === "autoFilter") {
				if (!target) continue;
				const ref = e.attrs.get("ref");
				if (ref === void 0) continue;
				const out = derivedFilter !== null && ref.trim().toUpperCase() === derivedFilter.from ? derivedFilter.to : adjustRange(ref, refChange(c, sheet.name));
				if (out === null || out === "") throw notSupported(`The filter of table "${name}" of sheet "${sheet.name}" cannot be moved with its rows.`, where, { table: name });
				if (out !== ref) edits.push([
					e.start.start,
					e.start.end,
					setAttr(text.slice(e.start.start, e.start.end), "ref", out)
				]);
				continue;
			}
			const raw = content(text, e);
			const out = adjustText(raw, c, sheet.name);
			if (out === null) throw notSupported(`A formula of table "${name}" of sheet "${sheet.name}" cannot be adjusted safely: "${decodeEntities(raw).slice(0, 60)}".`, where, { table: name });
			if (out !== raw) edits.push([
				e.start.end,
				e.end.start,
				out
			]);
		}
		if (edits.length > 0) edit.write(rel.target, applyEdits(text, edits));
	}
}
/** Rewrites the defined names of the workbook part. */
function rewriteNames(edit, workbookPart, c) {
	const text = edit.text(workbookPart);
	if (text === void 0) return;
	const { elements } = scan(text, workbookPart, edit.limits, (local, depth, parent) => depth === 3 && local === "definedName" && parent === "definedNames");
	const edits = [];
	for (const e of elements) {
		const raw = content(text, e);
		const out = adjustText(raw, c, NO_SHEET);
		if (out === null) throw notSupported(`The defined name "${e.attrs.get("name") ?? ""}" cannot be adjusted safely: "${decodeEntities(raw).slice(0, 60)}".`, { part: workbookPart }, { name: e.attrs.get("name") ?? "" });
		if (out !== raw) edits.push([
			e.start.end,
			e.end.start,
			out
		]);
	}
	if (edits.length > 0) edit.write(workbookPart, splice(text, edits));
}
function checkOptions(options, insert) {
	const what = insert ? "insertRows options" : "deleteRows options";
	const o = readOptions(options, insert ? [
		"sheet",
		"at",
		"count",
		"style",
		"limits"
	] : [
		"sheet",
		"at",
		"count",
		"limits"
	], what);
	const sheet = o.get("sheet");
	if (typeof sheet !== "string") throw new XlvError("XLV_E_INVALID_OPTION", `${what} need the sheet name.`, {
		hint: "Pass { sheet: 'Sales', at: 5 }.",
		details: { option: "sheet" }
	});
	const at = o.get("at");
	if (typeof at !== "number" || !Number.isInteger(at) || at < 1 || at > 1048576) throw new XlvError("XLV_E_INVALID_OPTION", `The option "at" must be a row number from 1 to ${String(MAX_ROWS)}.`, {
		hint: "For example { at: 5 } for row 5.",
		details: { option: "at" }
	});
	const count = o.get("count") ?? 1;
	if (typeof count !== "number" || !Number.isInteger(count) || count < 1 || at + count - 1 > 1048576) throw new XlvError("XLV_E_INVALID_OPTION", `The option "count" must be a whole number of at least 1, and the rows must end at row ${String(MAX_ROWS)} or before.`, {
		hint: "For example { count: 2 } for two rows.",
		details: { option: "count" }
	});
	const style = o.get("style") ?? "none";
	if (style !== "above" && style !== "below" && style !== "none") throw new XlvError("XLV_E_INVALID_OPTION", `The option "style" must be "above", "below" or "none", not ${JSON.stringify(style)}.`, {
		hint: "For example { style: 'above' } to copy the formatting of the row above.",
		details: { option: "style" }
	});
	return {
		sheet,
		at,
		count,
		style,
		limits: resolveLimits(o.get("limits"))
	};
}
function changeRows(input, sheetName, at, count, style, limits) {
	const bytes = inputBytes(input);
	const workbook = readWorkbook(bytes, { limits });
	const sheet = workbook.getSheet(sheetName);
	if (sheet.kind !== "worksheet" || sheet.part === null) throw new XlvError("XLV_E_INVALID_INPUT", `Sheet "${sheet.name}" is a ${sheet.kind} without rows.`, {
		location: { sheet: sheet.name },
		hint: "Insert or delete rows of worksheets only."
	});
	const pkg = openPackage(openZip(bytes, limits), limits);
	const edit = new PackageEditor(bytes, limits, pkg);
	const warnings = new WarningCollector();
	const workbookPart = pkg.relationships(null).find((r) => classifyRelationship(r.type)?.name === "officeDocument")?.target ?? "/xl/workbook.xml";
	const c = {
		sheet: sheet.name,
		at,
		count,
		sheets: workbook.sheets.map((s) => s.name)
	};
	for (const rel of pkg.relationships(sheet.part)) {
		if (rel.external) continue;
		const name = classifyRelationship(rel.type)?.name ?? null;
		if (name === null || !SAFE_SHEET_RELATIONSHIPS.has(name)) throw notSupported(`Sheet "${sheet.name}" has a related part of type "${name ?? rel.type}" (${rel.target}), which xlvane cannot move with its rows yet.`, {
			part: sheet.part,
			sheet: sheet.name
		}, { type: rel.type });
	}
	const lowerName = sheet.name.toLowerCase();
	for (const name of pkg.partNames) {
		if (!/\/charts\/[^/]+\.xml$|pivotCacheDefinition[^/]*\.xml$/i.test(name)) continue;
		const text = edit.text(name);
		if (text !== void 0 && text.toLowerCase().includes(lowerName)) throw notSupported(`The part ${name} (a chart or pivot table) may refer to sheet "${sheet.name}", and xlvane cannot rewrite it yet.`, {
			part: name,
			sheet: sheet.name
		});
	}
	let hasFormulas = false;
	for (const s of workbook.sheets) {
		if (s.kind !== "worksheet" || s.part === null) continue;
		const text = edit.text(s.part);
		if (text === void 0) continue;
		const target = s.part === sheet.part;
		const out = rewriteSheet(text, s.part, s, c, target, style, limits);
		hasFormulas ||= out.hasFormulas;
		if (out.text !== text) edit.write(s.part, out.text);
		rewriteTables(edit, pkg, s.part, s, c, target);
		if (target) {
			rewriteNotes(edit, pkg, s.part, text, s, c);
			rewriteDrawing(edit, pkg, s.part, text, s, c);
		}
	}
	rewriteNames(edit, workbookPart, c);
	if (hasFormulas) {
		requestFullCalculation(edit, workbookPart);
		warnings.add("XLV_W_FORMULAS_NOT_RECALCULATED", "The workbook contains formulas. xlvane does not recalculate them: cached results may be out of date until the file is opened in a spreadsheet application, which is asked to recalculate on load.", { part: workbookPart });
		const chains = pkg.relationships(workbookPart).filter((r) => !r.external && classifyRelationship(r.type)?.name === "calcChain");
		for (const chain of chains) {
			edit.remove(chain.target);
			removeContentType(edit, chain.target);
		}
		if (chains.length > 0) removeRelationships(edit, workbookPart, new Set(chains.map((x) => x.id)));
	}
	return {
		bytes: edit.build(),
		warnings: warnings.list()
	};
}
/**
* Inserts empty rows into a sheet of an existing XLSX file and returns the new file. Rows below
* move down, and every reference follows them as in a spreadsheet application: formulas on all
* sheets (ranges that the new rows fall into grow), merged cells, hyperlinks, notes, the
* autofilter, conditional formats (also those of the [MS-XLSX] extension, and thresholds that
* refer to cells), data validation (also of the extension), page breaks, defined names, tables
* (which grow when rows are inserted inside them) and pictures placed on the sheet. Every other
* part of the file stays exactly as it was.
*
* Refused with `XLV_E_NOT_SUPPORTED`, leaving the file unchanged: a drawing that holds charts,
* shapes or anything besides pictures; embedded objects, form controls, scenarios, cell watches,
* custom sheet views, other extended content that refers to cells (sparklines, slicers) and
* other related parts xlvane does not know; charts or pivot tables anywhere in the workbook that
* may refer to the sheet; data tables; an array formula the change would split; and formulas it
* cannot rewrite safely (docs/facts/ref-adjust.md).
*
* @example
* ```ts
* import { insertRows } from 'xlvane';
*
* declare const bytes: Uint8Array;
* // Two rows before row 5 of "Sales", formatted like row 4.
* const { bytes: updated, warnings } = insertRows(bytes, { sheet: 'Sales', at: 5, count: 2, style: 'above' });
* console.log(updated.length, warnings.map((w) => w.code));
* ```
*/
function insertRows(input, options) {
	const o = checkOptions(options, true);
	return changeRows(input, o.sheet, o.at, o.count, o.style, o.limits);
}
/**
* Deletes rows from a sheet of an existing XLSX file and returns the new file. Rows below move up,
* and references follow as in a spreadsheet application: ranges shrink, references to deleted
* cells become `#REF!`, merged cells whose top-left cell is deleted are removed, and so are notes,
* hyperlinks and other ranges that lie entirely in the deleted rows. Refusals are as for
* {@link insertRows}; in addition a deletion that would cut a picture, or remove a table's header
* row, totals row or all of its data rows, is refused.
*
* @example
* ```ts
* import { deleteRows } from 'xlvane';
*
* declare const bytes: Uint8Array;
* // Rows 5 and 6 of "Sales".
* const { bytes: updated } = deleteRows(bytes, { sheet: 'Sales', at: 5, count: 2 });
* ```
*/
function deleteRows(input, options) {
	const o = checkOptions(options, false);
	return changeRows(input, o.sheet, o.at, -o.count, "none", o.limits);
}
export { appendRows as a, fromCSV as c, fromJSON as d, pickSheet as f, writeTable as g, toJSON as h, adjustRange as i, inferValue as l, readTableAsync as m, insertRows as n, editCells as o, readTable as p, adjustFormula as r, diff as s, deleteRows as t, toCSV as u };
