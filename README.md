# Splunk Excel Extraction

Download Splunk search results as a **native Excel (.xlsx) file** from a button in a Simple XML dashboard, instead of the CSV export Splunk offers out of the box.

Splunk cannot produce `.xlsx` files on its own (its export formats are CSV, JSON, XML and raw). This app builds the workbook **in the browser** with [SheetJS](https://sheetjs.com/): there is no server-side component, no external service and no change to `limits.conf`.

- App label: **Splunk Excel Extraction**
- App ID / folder name: `splunk_excel_extract` (used in URLs and in `script="splunk_excel_extract:export_excel.js"`)

---

## Table of contents

1. [Features](#features)
2. [Dependencies and requirements](#dependencies-and-requirements)
3. [App layout](#app-layout)
4. [Installation](#installation)
5. [Demo dashboards](#demo-dashboards)
6. [Adding an export button to a dashboard](#adding-an-export-button-to-a-dashboard)
7. [Button attributes reference](#button-attributes-reference)
8. [How values are converted](#how-values-are-converted)
9. [How it works](#how-it-works)
10. [Limitations](#limitations)
11. [Troubleshooting](#troubleshooting)
12. [Splunk Cloud](#splunk-cloud)
13. [Updating SheetJS](#updating-sheetjs)
14. [Packaging and uninstalling](#packaging-and-uninstalling)
15. [Licenses](#licenses)

---

## Features

- One click downloads a real `.xlsx` file. It opens in Excel without a format warning, and accents and Unicode are preserved.
- Exports any search or post-process search of the dashboard, referenced by its `id`.
- **Full result set**: the export contains every row of the search, not only the page displayed in the table.
- **Multi-sheet workbooks**: export several searches into one file, one sheet per search.
- **Typed cells**: numbers are written as numbers, `_time` as an Excel date, multivalue fields as multi-line cells.
- Auto-sized columns and a filter on the header row of every sheet.
- File name with an optional timestamp (`events_20261006_1705.xlsx`).
- Configured entirely through `data-*` attributes on the button: no JavaScript to write.
- Reusable from **any app**: a dashboard in another app loads the script with `script="splunk_excel_extract:export_excel.js"`, and SheetJS is always loaded from this app.
- No second search: the button reuses the job the dashboard already ran.
- Two demo dashboards that use data available on any instance (`makeresults` and `index=_internal`).

## Dependencies and requirements

| Item | Requirement |
|------|-------------|
| Splunk | Splunk Enterprise 8.2 or later, or Splunk Cloud (see [Splunk Cloud](#splunk-cloud)) |
| Dashboard type | **Classic dashboards (Simple XML)** only. Dashboard Studio does not allow custom JavaScript. |
| JavaScript library | **SheetJS Community Edition 0.18.5** (`xlsx.full.min.js`), bundled in `appserver/static/`. No CDN or Internet access needed. |
| Splunk JS framework | `jquery`, `splunkjs/mvc` and `splunkjs/mvc/simplexml/ready!`, provided by Splunk Web (RequireJS). |
| Browser | Any current browser (Chrome, Edge, Firefox, Safari). The file is generated and downloaded client-side. |
| Permissions | Admin rights to install the app. Users need **read** access to the app (granted to all roles by `metadata/default.meta`) so their browser can load the static files. |

## App layout

```
splunk_excel_extract/
├── README.md                              This file
├── appserver/
│   └── static/
│       ├── export_excel.js                Export logic (click handler, conversion)
│       └── xlsx.full.min.js               SheetJS Community Edition 0.18.5
├── default/
│   ├── app.conf                           App metadata (label "Splunk Excel Extraction")
│   └── data/
│       └── ui/
│           ├── nav/default.xml            App navigation
│           └── views/
│               ├── excel_export_demo.xml            Demo: single and multi-sheet export
│               └── kpi_drilldown_export_demo.xml    Demo: KPI drilldown + export
├── licenses/
│   └── SheetJS-LICENSE.txt                Apache-2.0 license of SheetJS
└── metadata/
    └── default.meta                       Permissions
```

## Installation

### Option A: Splunk Web

1. Package the folder (see [Packaging](#packaging-and-uninstalling)).
2. In Splunk Web go to **Apps → Manage Apps → Install app from file**.
3. Select the `.tgz` and click **Upload**.
4. Restart Splunk if prompted.

### Option B: Command line

```bash
tar -xzf splunk_excel_extract.tgz -C $SPLUNK_HOME/etc/apps/
$SPLUNK_HOME/bin/splunk restart
```

### Option C: Manual copy

Copy the `splunk_excel_extract` folder into `$SPLUNK_HOME/etc/apps/` and restart Splunk.

### Clear the static asset cache

Splunk and browsers cache the files in `appserver/static` aggressively. After installing, and **every time `export_excel.js` changes**:

1. Open `https://<your-splunk>/en-US/_bump` and click **Bump version**.
2. Hard-refresh the dashboard (Ctrl+F5 / Cmd+Shift+R).

## Demo dashboards

Both dashboards are in the app navigation menu.

### Excel Export Demo (`splunk_excel_extract`, default view)

Uses 250 generated rows (`makeresults`).

- **Download Excel (.xlsx)** exports `search_export` to `events_<timestamp>.xlsx`.
- **Download Excel (2 sheets)** exports `search_export` and `search_summary` into one workbook with the sheets `Events` and `Summary by host`.

### KPI Drilldown Export Demo (`kpi_drilldown_export_demo`)

Uses `index=_internal`, with a time range picker.

1. A Single Value shows the number of `ERROR` events.
2. Clicking it sets the token `show_detail` and reveals a table of errors by `component` and `sourcetype`.
3. Above the table, **Télécharger Excel (.xlsx)** exports the table's search (`search_detail`) to `detail_erreurs_<timestamp>.xlsx`, and **Masquer le tableau** hides the table again.

## Adding an export button to a dashboard

### 1. Load the script

On the root element of the dashboard (`<dashboard>` or `<form>`):

| Dashboard location | Attribute |
|--------------------|-----------|
| Inside this app | `script="export_excel.js"` or `script="splunk_excel_extract:export_excel.js"` |
| In any other app | `script="splunk_excel_extract:export_excel.js"` |

```xml
<form version="1.1" script="splunk_excel_extract:export_excel.js">
```

The `app:file` form tells Splunk to load the file from that app's `appserver/static/`, so nothing has to be copied into your own app. To combine with other scripts, separate them with commas: `script="my_app:other.js, splunk_excel_extract:export_excel.js"`.

### 2. Give the search an `id`

```xml
<search id="my_search">
  <query>index=main | stats count by host, sourcetype</query>
  <earliest>-24h</earliest>
  <latest>now</latest>
</search>
```

This works for global searches, searches inside a panel (`<table><search id="...">`) and post-process searches (`<search id="..." base="...">`). Exporting a post-process search exports exactly what the table shows, without pagination.

### 3. Add a button

Any element with the class `excel-export-btn` becomes an export button, usually inside an `<html>` panel:

```xml
<row>
  <panel>
    <html>
      <button class="btn btn-primary excel-export-btn"
              data-search="my_search"
              data-filename="host_report"
              data-sheet="Hosts">Download Excel</button>
    </html>
  </panel>
</row>
```

The `btn` / `btn-primary` classes are Splunk's Bootstrap styles and are optional.

### 4. Reload

Bump the static cache (see above) and reload the dashboard.

### Several searches in one workbook

```xml
<button class="btn excel-export-btn"
        data-search="search_a,search_b,search_c"
        data-sheet="Errors,Latency,Traffic"
        data-filename="weekly_report">Download full report</button>
```

Sheet names are matched to the searches in order. Missing names default to `Results`, `Results 2`, and so on.

## Button attributes reference

All attributes are optional, but `data-search` should always be set.

| Attribute | Default | Description |
|-----------|---------|-------------|
| `data-search` | `search_export` | `id` of a search or post-process search, or a comma-separated list of ids (one sheet per id, in order). Ids are case-sensitive. |
| `data-filename` | `splunk_export` | File name without extension. A trailing `.xlsx` is ignored. |
| `data-sheet` | `Results` | Sheet name, or comma-separated names matching `data-search`. The characters `\ / ? * [ ] :` are replaced by `_`, names are cut to 31 characters and made unique (`Results`, `Results 2`, ...). |
| `data-timestamp` | `true` | `false` removes the `_YYYYMMDD_HHmm` suffix from the file name. |
| `data-keep-text` | `false` | `true` writes every value as text (no number conversion). |
| `data-keep-internal` | `false` | `true` keeps all internal fields starting with `_`. By default only `_time` and `_raw` are kept. |
| `data-dates` | `true` | `false` keeps `_time` as the original ISO text instead of an Excel date. |

Constants that are not exposed as attributes can be changed at the top of `export_excel.js`:

| Constant | Value | Meaning |
|----------|-------|---------|
| `TIMEOUT_MS` | 5 minutes | Maximum wait for a running search before giving up |
| `MAX_COL_WIDTH` | 60 | Maximum column width (characters) |
| `DATE_FORMAT` | `yyyy-mm-dd hh:mm:ss` | Excel format applied to `_time` |

## How values are converted

| Splunk value | Excel cell |
|--------------|------------|
| Integer or decimal such as `42`, `-3.5` | Number |
| Value with a leading zero such as `007` or `000123` | Text (identifiers are never altered) |
| Number with more than 15 digits | Text (Excel keeps only 15 significant digits) |
| `_time` | Date/time formatted `yyyy-mm-dd hh:mm:ss`, in the browser's local time |
| Multivalue field | One cell, values separated by line breaks |
| Empty / null | Empty cell |
| Text longer than 32,767 characters | Truncated (Excel's limit per cell) |
| Internal fields other than `_time` and `_raw` (`_bkt`, `_cd`, `_indextime`, ...) | Dropped (unless `data-keep-internal="true"`) |

Values are always written as typed cells, never as formulas: a value such as `=SUM(A1:A9)` stays plain text.

**About dates.** `_time` is converted to an Excel serial number by the script itself, using the wall-clock time of the user's browser. SheetJS' own date handling is not used because it adds a few seconds of error in time zones that had an unusual UTC offset in 1899 (for example Europe/Paris). Use `data-dates="false"` to keep the original ISO text with its time zone offset.

## How it works

```
 Dashboard                                Browser
┌──────────────────────┐   click   ┌──────────────────────────────────────┐
│ <button class=       │ ────────▶ │ export_excel.js                      │
│  "excel-export-btn"> │           │  1. read data-* attributes           │
└──────────────────────┘           │  2. mvc.Components.get(searchId)     │
                                   │  3. manager.data('results',          │
┌──────────────────────┐           │       {count:0, output_mode:         │
│ <search id="...">    │ ◀──────── │        'json_rows'})                 │
│ (already running in  │  results  │  4. SheetJS builds the worksheet(s)  │
│  the dashboard)      │           │  5. XLSX.writeFile() → download      │
└──────────────────────┘           └──────────────────────────────────────┘
```

- The button reuses the job the dashboard already ran and fetches its **full** result set (`count: 0`) through the Splunk JS SDK. No second search is dispatched.
- If the search is still running, the button shows `Preparing...` and the download starts as soon as results are available.
- SheetJS is registered in RequireJS under the name `xlsx` (the AMD name SheetJS declares). The path is `../app/<app>/xlsx.full.min`, where `<app>` is the app that serves `export_excel.js` (read from the script's own URL, `splunk_excel_extract` by default). It does **not** depend on the app of the dashboard.

## Limitations

- **Classic Simple XML only.** Dashboard Studio cannot load custom JavaScript.
- **Client-side generation.** Memory and CPU of the user's machine are the limit. Very large exports (several hundred thousand cells) can make the browser slow or freeze; aggregate with SPL first.
- **Splunk result cap.** Rows are fetched through the REST API, which is capped by `maxresultrows` (`limits.conf`, `[restapi]`, default 50,000). Larger result sets are silently cut.
- **Excel limits.** 1,048,576 rows and 16,384 columns per sheet, 32,767 characters per cell, 31 characters per sheet name.
- **Basic formatting only.** Header row, auto-sized columns, filter and date format. SheetJS Community Edition does not write cell styles (colours, bold) or frozen panes.
- **Only searches with an `id`** can be exported. Searches without an id are invisible to the script.
- **Current state of the dashboard.** The export contains the results as currently run, with the current tokens and time range. A search that has not run yet (for example one depending on an unset token) makes the button wait until the 5-minute timeout.
- **Time zone.** `_time` is written in the browser's local time; the time zone itself is not stored in the cell.
- **Security.** A user can export only the data the dashboard already shows them. The export does not bypass any role or index restriction.

## Troubleshooting

| Symptom | Likely cause and fix |
|---------|----------------------|
| Button does nothing | Open the developer console (F12). Check that `export_excel.js` loaded without a 404 and that the dashboard root has the `script` attribute. Bump the static cache and hard-refresh. |
| Console: `Script error for "xlsx"` | `xlsx.full.min.js` could not be loaded. Open `https://<your-splunk>/en-US/static/app/splunk_excel_extract/xlsx.full.min.js`: if it fails, the app is not installed or the user's role has no read access to it. Then bump the static cache. |
| `404` on `export_excel.js` | Wrong `script` attribute. From another app, use `script="splunk_excel_extract:export_excel.js"`. |
| Console: `SheetJS ... could not be loaded` | Same causes as `Script error for "xlsx"`. |
| Button shows `Search "xxx" not found` | `data-search` does not match any search `id` in the dashboard (case-sensitive). |
| Button shows `No results to export` | The search finished with zero rows. |
| Button shows `The search failed.` | The search returned an error; fix the SPL first. |
| Button shows `Timed out waiting for results.` | The search took more than 5 minutes or never started (unset token). Increase `TIMEOUT_MS` or fix the search. |
| Changes to the JS have no effect | Static files are cached. Use `/_bump` and a hard refresh. |
| Fewer rows than expected | `maxresultrows` cap (see [Limitations](#limitations)). |

## Splunk Cloud

You cannot copy files to `appserver/static` on Splunk Cloud. Instead:

1. Package the app (including `appserver/static`) as a `.tgz`.
2. Install it as a **private app** through **Apps → Manage Apps → Install app from file**, or through the Admin Config Service (ACS).
3. The app goes through **AppInspect** vetting. Custom JavaScript is allowed in private apps; review the report for warnings about the bundled library.

## Updating SheetJS

The bundled `xlsx.full.min.js` is the Community Edition **0.18.5**, the last version published on npm. Newer releases are distributed from [cdn.sheetjs.com](https://cdn.sheetjs.com/). To upgrade:

1. Download the newest `xlsx.full.min.js` from the SheetJS CDN.
2. Replace the file in `appserver/static/`.
3. Bump the static cache and hard-refresh.

Security note: npm versions up to 0.19.2 have published advisories (prototype pollution and ReDoS) affecting the **reading/parsing** of untrusted spreadsheet files. This app only **writes** workbooks and never parses files, so these code paths are not used. Upgrading is still recommended if your security policy requires it.

## Packaging and uninstalling

### Build the package

From the folder that contains `splunk_excel_extract/`:

```bash
tar -czf splunk_excel_extract.tgz splunk_excel_extract
```

Splunk accepts `.tgz` and `.spl` (same format, different extension).

### Uninstall

In Splunk Web: **Apps → Manage Apps → Delete** next to the app. Or:

```bash
$SPLUNK_HOME/bin/splunk remove app splunk_excel_extract
```

Dashboards in other apps that reference `splunk_excel_extract:export_excel.js` stop exporting once the app is removed.

## Licenses

- This app's code is provided as an example; adapt it freely.
- **SheetJS Community Edition** is licensed under the Apache License 2.0. The license text is in `licenses/SheetJS-LICENSE.txt`.
