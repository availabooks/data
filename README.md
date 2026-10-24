# Datasets

A [JOLT](https://jolt.availabooks.com/) app for Excel that lists datasets to practise with. Click one to read its description, then import it: each table goes on a sheet of its own as an Excel Table, with its numbers, dates and text kept as they should be. A checkbox deletes the workbook's other sheets first.

| File | What's in it |
| --- | --- |
| `jolt.json` | Name, version, description, permissions, and where to start |
| `index.html` | The page: one element that the app fills in |
| `styles.css` | The look, using JET's theme colours so dark mode works |
| `main.js` | The UI (built with [VanJS](https://vanjs.org/), which JOLT bundles), downloading, and every call to Excel, through `jolt.run` |
| `datasets.js` | The logic, as pure functions: reading the list of datasets and turning a dataset's CSV files into Excel cells |
| `zip.js` | Reads .zip files (copied from SQL for Excel) |
| `datasets/` | The datasets, and `index.json`, the list of them |

The app reads the datasets from this repo's `main` branch through jsDelivr: `https://cdn.jsdelivr.net/gh/availabooks/data@main/datasets/index.json`.

## Add a dataset

1. Load the data into SQL for Excel, then use **Data model > ☰ menu > Export data** to get a .zip of CSV files and a metadata.json. The metadata's column types decide how each value lands in Excel (numbers, dates, TRUE/FALSE, or text that keeps leading zeros). A .zip with only CSV files works too; then the types are guessed from the values.
2. Put the .zip in `datasets/`.
3. Add an entry to `datasets/index.json`: `name`, `description` (put a blank line, `\n\n`, between paragraphs), `file` (relative to index.json, or a full address on cdn.jsdelivr.net), and optionally `id`, `tables` and `rows`, which show next to the name.
4. Push to `main`. jsDelivr caches `main` for up to 12 hours, so a GitHub Action ([purge-datasets.yml](.github/workflows/purge-datasets.yml)) then clears its cache for every file in `datasets/`, and the app sees the change within a few minutes. To purge without a push, run the Action from the **Actions** tab.
