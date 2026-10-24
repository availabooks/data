// @ts-check
import van from "van";
import { excelTableName, freeSheetName, parseIndex, readPackage, sizeLabel } from "./datasets.js";
import { readZip } from "./zip.js";

/**
 * Datasets: a list of datasets to download into the workbook. Click one to read what's in it, then Import
 * to load each of its tables onto a sheet of its own as an Excel Table, optionally deleting every other
 * sheet first so the workbook holds just the dataset.
 *
 * The list is datasets/index.json in the availabooks/data repo on GitHub, read through jsDelivr. Each
 * dataset is a .zip that SQL for Excel's Export data makes: a CSV file for each table and a metadata.json.
 *
 * @typedef {import("./datasets.js").Dataset} Dataset
 * @typedef {import("./datasets.js").SheetTable} SheetTable
 */

/**
 * Where the list of datasets is. jsDelivr caches main for up to 12 hours, so a GitHub Action
 * (.github/workflows/purge-datasets.yml) clears its cache after each push that changes datasets/.
 */
export const INDEX_URL = "https://cdn.jsdelivr.net/gh/availabooks/data@main/datasets/index.json";

const { main, header, h1, p, button, input, label, ul, li, div, small } = van.tags;

const datasets = van.state(/** @type {Dataset[]} */ ([]));
const loading = van.state(false);
const loadError = van.state("");
/** The id of the dataset whose description is open. */
const open = van.state("");
const deleteOthers = van.state(false);
/** What's happening during an import, or "". */
const busy = van.state("");

/** @param {unknown} e */
const message = (e) => (e instanceof Error ? e.message : String(e));

/** Read the list of datasets. */
async function refresh() {
  loading.val = true;
  loadError.val = "";
  try {
    /** @type {Response} */
    let res;
    try {
      res = await fetch(INDEX_URL, { cache: "no-cache" });
    } catch {
      throw new Error("Couldn't download the list of datasets. Check your internet connection, then click Refresh.");
    }
    if (!res.ok) throw new Error(`Couldn't download the list of datasets: the site answered ${res.status}${res.statusText ? ` ${res.statusText}` : ""}.`);
    datasets.val = parseIndex(await res.text(), INDEX_URL);
  } catch (e) {
    loadError.val = message(e);
  } finally {
    loading.val = false;
  }
}

/** Download a dataset and read its tables. @param {Dataset} d */
async function download(d) {
  /** @type {Response} */
  let res;
  try {
    res = await fetch(d.url);
  } catch {
    throw new Error(`Couldn't download ${d.name}. Check your internet connection and try again.`);
  }
  if (!res.ok) throw new Error(`Couldn't download ${d.name}: the site answered ${res.status}${res.statusText ? ` ${res.statusText}` : ""}.`);
  busy.val = `Unpacking ${d.name}…`;
  const dec = new TextDecoder();
  const files = (await readZip(await res.arrayBuffer())).filter((f) => /\.(csv|json)$/i.test(f.name)).map((f) => ({ name: f.name, text: dec.decode(f.data) }));
  return readPackage(files);
}

// ---- Excel -----------------------------------------------------------------------------------------------

/**
 * Delete every sheet in the workbook. A workbook needs at least one sheet, so an empty one is added first;
 * it's returned, to be deleted once the dataset's sheets are in.
 */
async function clearWorkbook() {
  return jolt.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();
    const old = [...sheets.items];
    const holder = sheets.add(freeSheetName("Sheet1", old.map((s) => s.name)));
    holder.activate();
    await context.sync();
    for (const s of old) {
      // Hidden sheets go too; making them visible first avoids Excel refusing to delete a very hidden one.
      s.visibility = Excel.SheetVisibility.visible;
      s.delete();
    }
    await context.sync();
    holder.load("name");
    await context.sync();
    return holder.name;
  });
}

/** Delete a sheet, unless it's the only one left. @param {string} name */
async function deleteSheet(name) {
  await jolt.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();
    if (sheets.items.length > 1) sheets.getItemOrNullObject(name).delete();
    await context.sync();
  });
}

/**
 * Write a table to a new sheet as an Excel Table: the header, then the rows in batches so a big table
 * doesn't go over what one request to Excel can carry. Each batch's number formats are set before its
 * values, so text columns ("@") keep values like 01234 as text. Returns the sheet's name.
 * @param {SheetTable} t
 */
async function writeTableSheet(t) {
  return jolt.run(async (context) => {
    const wb = context.workbook;
    wb.worksheets.load("items/name");
    wb.tables.load("items/name");
    await context.sync();
    const name = freeSheetName(t.name, wb.worksheets.items.map((s) => s.name));
    const ws = wb.worksheets.add(name);
    const width = t.header.length;
    const n = t.rows.length;
    ws.getRangeByIndexes(0, 0, 1, width).values = [t.header];
    await context.sync();
    const batch = Math.max(1, Math.floor(20_000 / Math.max(1, width)));
    for (let r = 0; r < n; r += batch) {
      const part = t.rows.slice(r, r + batch);
      const range = ws.getRangeByIndexes(1 + r, 0, part.length, width);
      range.numberFormat = part.map(() => t.formats);
      range.values = part;
      await context.sync();
    }
    // An Excel Table needs a row under its header, even an empty one.
    const all = ws.getRangeByIndexes(0, 0, Math.max(n, 1) + 1, width);
    const table = ws.tables.add(all, true);
    table.name = excelTableName(t.name, wb.tables.items.map((x) => x.name));
    all.format.autofitColumns();
    ws.freezePanes.freezeRows(1);
    await context.sync();
    return name;
  });
}

/** @param {string} name */
async function activate(name) {
  await jolt.run(async (context) => {
    const ws = context.workbook.worksheets.getItem(name);
    ws.activate();
    ws.getRange("A1").select();
    await context.sync();
  });
}

/** Download a dataset and write its tables to new sheets, deleting the other sheets first if asked. @param {Dataset} d */
async function importDataset(d) {
  const clear = deleteOthers.val;
  if (
    clear &&
    !(await jolt.confirm(`Delete every sheet in this workbook, then load ${d.name}? The sheets that are there now can't be brought back.`, {
      title: "Delete other sheets",
      ok: "Delete and import",
      cancel: "Cancel",
    }))
  )
    return;
  busy.val = `Downloading ${d.name}…`;
  /** @type {string | null} */
  let holder = null;
  try {
    // Everything that can fail without touching the workbook happens before any sheet is deleted.
    const { tables, problems } = await download(d);
    if (!tables.length) throw new Error(`${d.name} has no tables to load${problems.length ? `: ${problems[0]}` : "."}`);
    if (clear) {
      busy.val = "Deleting the other sheets…";
      holder = await clearWorkbook();
    }
    /** @type {string[]} */
    const made = [];
    for (const t of tables) {
      busy.val = `Writing ${t.name} (${t.rows.length.toLocaleString()} rows)…`;
      made.push(await writeTableSheet(t));
    }
    if (holder) await deleteSheet(holder);
    holder = null;
    await activate(made[0]);
    const rows = tables.reduce((n, t) => n + t.rows.length, 0);
    jolt.notify(`Loaded ${d.name}: ${tables.length} ${tables.length === 1 ? "table" : "tables"}, ${rows.toLocaleString()} rows.`, "success");
    if (problems.length) jolt.notify(`Some of ${d.name} didn't load: ${problems.join("; ")}.`, "info");
  } catch (e) {
    jolt.notify(message(e), "error");
  } finally {
    // If writing failed part way, the empty sheet stays only when there's nothing else.
    if (holder) await deleteSheet(holder).catch(() => {});
    busy.val = "";
  }
}

// ---- The page --------------------------------------------------------------------------------------------

/** A description's paragraphs, split at blank lines. @param {string} text */
const Description = (text) =>
  div(
    { class: "description" },
    ...(text.trim() ? text.trim().split(/\n\s*\n/) : ["There's no description of this dataset."]).map((para) => p(para.trim())),
  );

/** @param {Dataset} d */
const Item = (d) =>
  li(
    { class: () => `dataset${open.val === d.id ? " open" : ""}` },
    button(
      { class: "name", "aria-expanded": () => String(open.val === d.id), onclick: () => (open.val = open.val === d.id ? "" : d.id) },
      d.name,
      sizeLabel(d) ? small(sizeLabel(d)) : null,
    ),
    () =>
      open.val === d.id
        ? div(
            { class: "details" },
            Description(d.description),
            label({ class: "check" }, input({ type: "checkbox", checked: deleteOthers, onchange: (/** @type {Event} */ e) => (deleteOthers.val = /** @type {HTMLInputElement} */ (e.target).checked) }), "Delete all other sheets in this workbook"),
            button({ class: "primary", onclick: () => importDataset(d), disabled: () => !!busy.val }, "Import"),
          )
        : div(),
  );

const App = () =>
  main(
    { class: "app" },
    header(h1("Datasets"), p("Data to practise with. Click a dataset to see what's in it, then import it. Each table goes on a sheet of its own as an Excel Table.")),
    () => (busy.val ? p({ class: "status" }, busy.val) : div()),
    () => {
      if (loading.val) return p({ class: "status" }, "Reading the list of datasets…");
      if (loadError.val) return div({ class: "error" }, p(loadError.val), button({ onclick: refresh }, "Refresh"));
      if (!datasets.val.length) return div(p({ class: "status" }, "There are no datasets yet."), button({ onclick: refresh }, "Refresh"));
      return ul({ class: "list" }, ...datasets.val.map(Item));
    },
  );

export async function autoRun() {
  van.add(/** @type {Element} */ (jolt.$("#app")), App());
  await refresh();
}
