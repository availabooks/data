// @ts-check

/**
 * The pure logic of Datasets: reading the list of datasets (index.json), and reading a dataset's .zip, as
 * SQL for Excel's Export data makes it (a CSV file for each table and a metadata.json), into tables of
 * Excel cells. No Office or network calls.
 *
 * A few of these functions are copied from SQL for Excel (sql/export.js, sql/dump.js and sql/data.js),
 * because installed apps can't import code from another app.
 *
 * @typedef {{ id: string, name: string, description: string, url: string, tables: number | null, rows: number | null }} Dataset
 *   url: the .zip's full address
 * @typedef {"text" | "integer" | "bigint" | "numeric" | "double precision" | "boolean" | "date" | "timestamp" | "time"} SqlType
 * @typedef {{ name: string, header: string[], rows: (string | number | boolean)[][], formats: string[] }} SheetTable
 *   a table ready for a sheet: each value as an Excel cell, and each column's number format
 */

export const INDEX_FORMAT = "jolt-datasets";
export const EXPORT_FORMAT = "sql-for-excel-export";

// ---- The list of datasets --------------------------------------------------------------------------------

/**
 * The datasets in index.json, with each file's address made absolute (a file is relative to the index).
 * Entries without a name or a file are skipped. Throws an Error with a message for the user if it isn't a
 * list of datasets this version can read.
 * @param {string} text index.json's text @param {string} indexUrl where it came from
 * @returns {Dataset[]}
 */
export function parseIndex(text, indexUrl) {
  /** @type {any} */
  let index;
  try {
    index = JSON.parse(text);
  } catch {
    throw new Error("The list of datasets isn't valid JSON.");
  }
  if (index?.format !== INDEX_FORMAT) throw new Error("The list of datasets isn't in the format this app reads.");
  if (index.version !== 1) throw new Error(`The list of datasets is version ${index.version}. This version of the app reads version 1. Update the app to see them.`);
  /** @param {unknown} n */
  const count = (n) => (typeof n === "number" && Number.isFinite(n) ? n : null);
  return (Array.isArray(index.datasets) ? index.datasets : [])
    .filter((d) => d?.name && d?.file)
    .map((d) => ({
      id: String(d.id ?? d.file),
      name: String(d.name),
      description: String(d.description ?? ""),
      url: new URL(String(d.file), indexUrl).href,
      tables: count(d.tables),
      rows: count(d.rows),
    }));
}

/** "3 tables · 1,250 rows", from what the index says, or "". @param {Dataset} d */
export function sizeLabel(d) {
  /** @param {number} n @param {string} w */
  const count = (n, w) => `${n.toLocaleString()} ${w}${n === 1 ? "" : "s"}`;
  return [d.tables === null ? "" : count(d.tables, "table"), d.rows === null ? "" : count(d.rows, "row")].filter(Boolean).join(" · ");
}

// ---- Reading a dataset -----------------------------------------------------------------------------------

/**
 * Read CSV text: fields separated by commas, quoted with double quotes ("" inside quotes is one quote), and
 * rows ending in \n or \r\n. An empty field with no quotes is null; a byte order mark at the start is skipped.
 * (From sql/export.js.)
 * @param {string} text
 * @returns {(string | null)[][]}
 */
export function parseCsv(text) {
  /** @type {(string | null)[][]} */
  const rows = [];
  /** @type {(string | null)[]} */
  let row = [];
  let i = text.charCodeAt(0) === 0xfeff ? 1 : 0;
  const n = text.length;
  while (i <= n) {
    let value = /** @type {string | null} */ (null);
    if (text[i] === '"') {
      let s = "";
      i++;
      while (i < n) {
        if (text[i] === '"') {
          if (text[i + 1] === '"') {
            s += '"';
            i += 2;
          } else {
            i++;
            break;
          }
        } else s += text[i++];
      }
      value = s;
      // Anything between the closing quote and the separator is kept, as spreadsheets do.
      while (i < n && text[i] !== "," && text[i] !== "\n" && text[i] !== "\r") value += text[i++];
    } else {
      let j = i;
      while (j < n && text[j] !== "," && text[j] !== "\n" && text[j] !== "\r") j++;
      value = j > i ? text.slice(i, j) : null;
      i = j;
    }
    row.push(value);
    if (text[i] === ",") {
      i++;
      continue;
    }
    // End of a row (or of the text).
    if (text[i] === "\r") i++;
    if (text[i] === "\n") i++;
    if (!(row.length === 1 && row[0] === null && i >= n)) rows.push(row);
    row = [];
    if (i >= n) break;
  }
  return rows;
}

/** The file name without folders, lowercase. @param {string} name */
const base = (name) => name.replace(/^.*[\\/]/, "").toLowerCase();

/**
 * A dataset's files as tables for sheets. With a metadata.json, its tables are loaded in its order, with its
 * column types; columns are matched by name, so a CSV whose columns were moved still loads. Without one,
 * each CSV is a table named after its file, with the types its values suggest. Throws an Error with a
 * message for the user when there's nothing to load.
 * @param {{ name: string, text: string }[]} files the CSVs and metadata.json from the .zip
 * @returns {{ tables: SheetTable[], problems: string[] }}
 */
export function readPackage(files) {
  const metaFile = files.find((f) => base(f.name) === "metadata.json");
  const csvs = new Map(files.filter((f) => /\.csv$/i.test(f.name)).map((f) => [base(f.name), f]));
  if (!csvs.size) throw new Error("This dataset has no CSV files in it.");
  /** @type {string[]} */
  const problems = [];
  /** @type {SheetTable[]} */
  const tables = [];
  if (!metaFile) {
    for (const f of csvs.values()) {
      const [header = [], ...rows] = parseCsv(f.text);
      const names = uniqueNames(header.map((h, c) => (h ?? "").trim() || `Column ${c + 1}`));
      const types = names.map((_, c) => guessType(rows.map((r) => r[c] ?? null)));
      tables.push(sheetTable(f.name.replace(/^.*[\\/]/, "").replace(/\.csv$/i, ""), names, types, rows));
    }
    return { tables, problems };
  }
  /** @type {any} */
  let meta;
  try {
    meta = JSON.parse(metaFile.text);
  } catch {
    throw new Error("This dataset's metadata.json isn't valid JSON.");
  }
  if (meta?.format !== EXPORT_FORMAT) throw new Error("This dataset's metadata.json isn't from an SQL for Excel export.");
  if (meta.version !== 1) throw new Error(`This dataset is version ${meta.version} of SQL for Excel's export. This app reads version 1.`);
  for (const t of Array.isArray(meta.tables) ? meta.tables : []) {
    if (!t?.name || !Array.isArray(t.columns)) continue;
    const f = csvs.get(base(String(t.file ?? `${t.name}.csv`)));
    if (!f) {
      problems.push(`${t.name}: its file ${t.file} is missing, so it wasn't loaded`);
      continue;
    }
    const columns = t.columns.filter((c) => c?.name).map((c) => ({ name: String(c.name), type: sqlFamily(String(c.type || "text")) }));
    const [header = [], ...rows] = parseCsv(f.text);
    const at = columns.map((c) => header.findIndex((h) => (h ?? "").trim().toLowerCase() === c.name.toLowerCase()));
    const lost = columns.filter((_, i) => at[i] < 0).map((c) => c.name);
    if (lost.length) problems.push(`${t.name}: ${lost.join(", ")} ${lost.length === 1 ? "isn't" : "aren't"} in ${f.name}, so ${lost.length === 1 ? "it's" : "they're"} empty`);
    const picked = rows.map((r) => at.map((i) => (i < 0 ? null : (r[i] ?? null))));
    tables.push(
      sheetTable(
        String(t.name),
        columns.map((c) => c.name),
        columns.map((c) => c.type),
        picked,
      ),
    );
  }
  return { tables, problems };
}

/** @param {string} name @param {string[]} header @param {SqlType[]} types @param {(string | null)[][]} rows @returns {SheetTable} */
function sheetTable(name, header, types, rows) {
  return {
    name,
    header,
    rows: rows.map((r) => types.map((t, c) => excelCell(r[c] ?? null, t))),
    formats: types.map((t) => EXCEL_FORMATS[t]),
  };
}

/**
 * The type a column's text values suggest, for CSVs without a metadata.json. Numbers with a leading zero
 * (like ZIP codes, "01234") stay text, so the zero isn't lost.
 * @param {(string | null)[]} cells
 * @returns {SqlType}
 */
export function guessType(cells) {
  const vals = /** @type {string[]} */ (cells.filter((v) => v !== null && v.trim() !== "")).map((v) => v.trim());
  if (!vals.length) return "text";
  const leadingZero = (/** @type {string} */ v) => /^-?0\d/.test(v);
  if (vals.every((v) => /^-?\d{1,15}$/.test(v) && !leadingZero(v))) return "integer";
  if (vals.every((v) => /^-?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i.test(v) && !leadingZero(v))) return "numeric";
  if (vals.every((v) => /^(true|false)$/i.test(v))) return "boolean";
  if (vals.every((v) => /^\d{4}-\d{2}-\d{2}$/.test(v))) return "date";
  if (vals.every((v) => /^\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}(:\d{2}(\.\d+)?)?$/.test(v))) return "timestamp";
  return "text";
}

// ---- Values as Excel cells (from sql/dump.js and sql/data.js) ---------------------------------------------

/**
 * Which of these types a PostgreSQL type is closest to: varchar and uuid are text, smallint is integer, real
 * is double precision, timestamptz is timestamp, and so on.
 * @param {string} declared
 * @returns {SqlType}
 */
export function sqlFamily(declared) {
  const t = declared.toLowerCase().trim();
  if (/^(smallint|integer|int|int2|int4|smallserial|serial)\b/.test(t)) return "integer";
  if (/^(bigint|int8|bigserial)\b/.test(t)) return "bigint";
  if (/^(numeric|decimal|money)\b/.test(t)) return "numeric";
  if (/^(real|double precision|float)/.test(t)) return "double precision";
  if (/^bool/.test(t)) return "boolean";
  if (/^date\b/.test(t)) return "date";
  if (/^timestamp/.test(t)) return "timestamp";
  if (/^time\b/.test(t)) return "time";
  return "text";
}

/** Excel number formats for each type's column. Text is "@" so "01234" stays text. */
export const EXCEL_FORMATS = /** @type {Record<SqlType, string>} */ ({
  text: "@",
  integer: "General",
  bigint: "General",
  numeric: "General",
  "double precision": "General",
  boolean: "General",
  date: "yyyy-mm-dd",
  timestamp: "yyyy-mm-dd hh:mm:ss",
  time: "hh:mm:ss",
});

/** "2024-01-02" or "2024-01-02 10:30:00" → an Excel serial number, or null. @param {string} iso */
export function isoToSerial(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?)?$/.exec(iso.trim());
  if (!m) return null;
  const ms = Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] ?? 0), +(m[5] ?? 0), +(m[6] ?? 0), Number(`0.${m[7] ?? 0}`) * 1000);
  return ms / 86400000 + 25569;
}

/** Drop a time zone after a time: "2024-01-02 10:00:00+00" → "2024-01-02 10:00:00". Excel has no time zones. @param {string} s */
const dropZone = (s) => s.replace(/(\d{2}:\d{2}(?::\d{2})?(?:\.\d+)?)\s*(?:Z|[+-]\d{2}(?::?\d{2})?)$/, "$1");

/**
 * A value (as text) as an Excel cell value: numbers as numbers, true/false as TRUE/FALSE, dates and times as
 * serial numbers, and anything that doesn't fit as text. Text that Excel would read as a formula gets a '.
 * @param {string | null} v @param {SqlType} type
 * @returns {string | number | boolean}
 */
export function excelCell(v, type) {
  if (v === null) return "";
  switch (type) {
    case "integer":
    case "bigint":
    case "numeric":
    case "double precision": {
      const n = Number(v);
      return v.trim() !== "" && Number.isFinite(n) ? n : v;
    }
    case "boolean":
      return /^(t|true)$/i.test(v.trim());
    case "date":
    case "timestamp": {
      const serial = isoToSerial(dropZone(v));
      return serial ?? v;
    }
    case "time": {
      const m = /^(\d{2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?/.exec(v);
      return m ? (+m[1] * 3600 + +m[2] * 60 + Number(m[3] ?? 0)) / 86400 : v;
    }
  }
  return /^[=+\-@]/.test(v) && !/^[-+]?\d/.test(v) ? "'" + v : v;
}

// ---- Names -----------------------------------------------------------------------------------------------

/** Make names unique, ignoring case, by adding _2, _3… @param {string[]} names */
export function uniqueNames(names) {
  const seen = new Set();
  return names.map((n) => {
    let name = n;
    for (let i = 2; seen.has(name.toLowerCase()); i++) name = `${n}_${i}`;
    seen.add(name.toLowerCase());
    return name;
  });
}

/** A sheet name Excel accepts: no []:*?/\ and at most 31 characters. @param {string} name */
export function sheetName(name) {
  const n = name.replace(/[[\]:*?/\\]/g, " ").replace(/^'+|'+$/g, "").trim().slice(0, 31);
  return n || "Data";
}

/** "Orders", then "Orders 2"… whichever isn't taken. @param {string} base @param {string[]} taken */
export function freeSheetName(base, taken) {
  const used = new Set(taken.map((t) => t.toLowerCase()));
  const b = sheetName(base);
  if (!used.has(b.toLowerCase())) return b;
  for (let i = 2; ; i++) {
    const suffix = ` ${i}`;
    const n = b.slice(0, 31 - suffix.length) + suffix;
    if (!used.has(n.toLowerCase())) return n;
  }
}

/**
 * A name Excel accepts for a Table, unique among `taken`: letters, digits, _ and ., starting with a letter
 * or _, and not something Excel would read as a cell address (like A1 or R1C1).
 * @param {string} name @param {string[]} taken
 */
export function excelTableName(name, taken) {
  let n = name.replace(/[^A-Za-z0-9_.]+/g, "_").replace(/^_+(?=.)/, "").slice(0, 200);
  if (!/^[A-Za-z_]/.test(n) || /^[A-Za-z]{1,3}\d+$/.test(n) || /^[RrCc]$/.test(n) || /^[Rr]\d*[Cc]\d*$/.test(n)) n = `T_${n}`;
  return uniqueNames([...taken, n]).at(-1) ?? n;
}
