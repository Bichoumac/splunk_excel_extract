# Splunk Excel Extraction

Download Splunk search results as a **native Excel (.xlsx) file** from **Simple XML (Classic)** and **Dashboard Studio** dashboards, instead of the CSV export Splunk offers out of the box.

Splunk cannot produce `.xlsx` files on its own (its export formats are CSV, JSON, XML and raw). This app adds a REST endpoint, `/services/excel_export`, written in Python with [openpyxl](https://openpyxl.readthedocs.io/). It reads the results of a search job (or runs a search) **with the permissions of the calling user** and returns a ready-to-open workbook.

- App label: **Splunk Excel Extraction**
- App ID / folder name: `splunk_excel_extract`
- Endpoint: `https://<splunk-web>/<locale>/splunkd/__raw/services/excel_export` (browser), `https://<splunkd>:8089/services/excel_export` (REST clients)

---

## Table of contents

1. [Features](#features)
2. [Requirements](#requirements)
3. [Repository layout](#repository-layout)
4. [Installation](#installation)
5. [Demo dashboards](#demo-dashboards)
6. [REST endpoint reference](#rest-endpoint-reference)
7. [Simple XML: export buttons (export_excel.js)](#simple-xml-export-buttons-export_exceljs)
8. [Simple XML: plain link, no JavaScript](#simple-xml-plain-link-no-javascript)
9. [Dashboard Studio](#dashboard-studio)
10. [Calling the endpoint from scripts](#calling-the-endpoint-from-scripts)
11. [How values are converted](#how-values-are-converted)
12. [Configuration (excel_export.conf)](#configuration-excel_exportconf)
13. [Security](#security)
14. [Limitations](#limitations)
15. [Troubleshooting](#troubleshooting)
16. [Development](#development)
17. [Packaging, Splunk Cloud and uninstalling](#packaging-splunk-cloud-and-uninstalling)
18. [Upgrading from 1.x (SheetJS)](#upgrading-from-1x-sheetjs)
19. [Licenses](#licenses)

---

## Features

- Real `.xlsx` files: they open in Excel without a format warning; accents and Unicode are preserved.
- Works from **Simple XML** (button or plain link) and **Dashboard Studio** (markdown link or drilldown), and from any HTTP client (`curl`, Python, ...).
- Two modes:
  - **`sid`**: export a job that already ran (the dashboard's search), no second search;
  - **`search`**: run an SPL search and export its results.
- **Full result set**: results are read page by page, so the export is **not** cut at `maxresultrows` (50,000 rows by default) like a normal REST call.
- **Post-process searches** (Simple XML `base=`) are exported exactly as the panel shows them.
- **Multi-sheet workbooks**: one sheet per job or search.
- **Typed cells**: numbers as numbers, `_time` as an Excel date, multivalue fields as multi-line cells, identifiers like `007` kept as text.
- **Formatting**: bold coloured header, frozen header row, filter on the header, auto-sized columns.
- Permissions are those of the user: no system credentials, no role escalation.
- Guard rails: SPL commands that write or send data (`outputlookup`, `sendemail`, `delete`, ...) are refused; the `search` mode can be disabled.
- Logs in `$SPLUNK_HOME/var/log/splunk/excel_export.log` (searchable in `index=_internal`).

## Requirements

| Item | Requirement |
|------|-------------|
| Splunk | Splunk Enterprise 8.2 or later (Python 3), or Splunk Cloud (see [Splunk Cloud](#splunk-cloud)) |
| Python | The Python 3 shipped with Splunk (3.7 or later). Nothing to install: **openpyxl 3.1.2** and **et_xmlfile 1.1.0** are bundled in `bin/lib/`. |
| Dashboards | Simple XML (Classic) and Dashboard Studio. Studio job tokens (`$name:job.sid$`) need a recent Splunk 9.x; the `search` mode works on every version. |
| Browser | Any current browser (Chrome, Edge, Firefox, Safari). |
| Permissions | Admin rights to install the app. Users need **read** access to the app (granted to all roles by `metadata/default.meta`) to load `export_excel.js`, and the usual `search` capability to read or run jobs. |

## Repository layout

The Splunk app itself lives in `src/`. Everything else (README, license, tests, build script) is repository material; `build.sh` copies `README.md` and `LICENSE.txt` into the package.

```
splunk_excel_extract/                (git repository)
├── README.md                        This file
├── LICENSE.txt                      Apache-2.0 license of this app
├── build.sh                         Builds dist/splunk_excel_extract-<version>.tgz
├── requirements.txt                 Versions of the libraries vendored in src/bin/lib
├── tests/
│   └── test_excel_export.py         Unit tests (no Splunk needed)
└── src/                             ← the Splunk app (becomes splunk_excel_extract/)
    ├── appserver/static/
    │   └── export_excel.js          Simple XML buttons → calls the endpoint
    ├── bin/
    │   ├── excel_export_handler.py  REST handler: parameters, jobs, pagination, response
    │   ├── xlsx_builder.py          Workbook generation (openpyxl)
    │   └── lib/                     Bundled openpyxl + et_xmlfile
    ├── default/
    │   ├── app.conf
    │   ├── restmap.conf             Declares /services/excel_export
    │   ├── web.conf                 Exposes it to Splunk Web (/splunkd/__raw/...)
    │   ├── excel_export.conf        Endpoint settings
    │   └── data/ui/
    │       ├── nav/default.xml
    │       └── views/
    │           ├── excel_export_demo.xml          Simple XML: buttons, 2 sheets, plain link
    │           ├── excel_export_studio_demo.xml   Dashboard Studio: links and drilldown
    │           └── kpi_drilldown_export_demo.xml  Simple XML: KPI drilldown + export
    ├── README/
    │   └── excel_export.conf.spec
    ├── licenses/                    openpyxl and et_xmlfile licenses (MIT)
    ├── metadata/default.meta        Permissions
    └── static/appIcon*.png          App icons
```

## Installation

### 1. Build the package

From the repository root (Linux, macOS or WSL):

```bash
./build.sh
# -> dist/splunk_excel_extract-2.0.0.tgz
```

### 2. Install it

**Option A, Splunk Web:** **Apps → Manage Apps → Install app from file**, select the `.tgz`, tick **Upgrade app** if a previous version is installed, click **Upload**, then restart Splunk when prompted.

**Option B, command line:**

```bash
tar -xzf dist/splunk_excel_extract-2.0.0.tgz -C $SPLUNK_HOME/etc/apps/
$SPLUNK_HOME/bin/splunk restart
```

**Option C, development install:** link or copy `src/` as `$SPLUNK_HOME/etc/apps/splunk_excel_extract` and restart Splunk.

```bash
ln -s "$(pwd)/src" $SPLUNK_HOME/etc/apps/splunk_excel_extract
```

> A **restart is required** the first time (and whenever `restmap.conf` or `web.conf` changes) so that splunkd registers the endpoint and Splunk Web exposes it. On a search head cluster, deploy the app from the deployer (`splunk apply shcluster-bundle`).

### 3. Clear the static asset cache

Splunk and browsers cache `appserver/static` aggressively. After installing or upgrading, open `https://<your-splunk>/en-US/_bump`, click **Bump version**, then hard-refresh the dashboards (Ctrl+F5 / Cmd+Shift+R).

### 4. Check that it works

1. Open the app and the **Excel Export Demo** dashboard, click **Download Excel (.xlsx)**.
2. Or, from a shell (replace the credentials):

   ```bash
   curl -sk -u admin:changeme "https://localhost:8089/services/excel_export" \
        --data-urlencode "search=| makeresults count=5 | eval n=random()" \
        -d timestamp=false -o test.xlsx && file test.xlsx
   # test.xlsx: Microsoft Excel 2007+
   ```

## Demo dashboards

All three are in the app navigation menu.

| Dashboard | Type | What it shows |
|-----------|------|---------------|
| **Excel Export Demo** (`excel_export_demo`, default view) | Simple XML | Button exporting `search_export` (250 `makeresults` rows); button building a 2-sheet workbook (`Events` + post-process `Summary by host`); a plain link with no JavaScript. |
| **Excel Export Demo (Dashboard Studio)** (`excel_export_studio_demo`) | Dashboard Studio | Markdown link reusing the job of a data source (`sid`), markdown link running a search (`search`), and a table whose row click downloads its results. |
| **KPI Drilldown Export Demo** (`kpi_drilldown_export_demo`) | Simple XML | Clicking a Single Value (`ERROR` count in `index=_internal`) reveals a detail table with a **Télécharger Excel (.xlsx)** button. |

## REST endpoint reference

```
GET  /services/excel_export?<parameters>
POST /services/excel_export          (form-encoded or JSON body, same parameters)
```

From a browser, through Splunk Web: `/<locale>/splunkd/__raw/services/excel_export`, for example `/en-US/splunkd/__raw/services/excel_export`. A path without locale (`/splunkd/__raw/...`) is redirected to the user's locale by Splunk Web.

### Parameters

Each source is one sheet. Use **either** `sid` **or** `search` in one request, and repeat it for several sheets (`sid=a&sid=b`).

| Parameter | Repeatable | Default | Description |
|-----------|-----------|---------|-------------|
| `sid` | yes | | Search job id to export. The job must be visible to the user. If it is still running, the endpoint waits for it (up to `job_timeout`). |
| `postprocess` | yes, aligned with `sid` | none | Post-process SPL applied to the job results (e.g. `| stats count by host`). An empty value means "none" for that sid. |
| `search` | yes | | SPL to run. `search ` is prepended if the query does not start with `|` or `search`. The job is cancelled once exported. Refused when `allow_search = false`. |
| `earliest` / `latest` | yes, aligned with `search` | none (all time) | Time range of `search`, in any Splunk time format (`-24h@h`, `now`, epoch, ...). One value applies to every search. |
| `app` | no | `default_app` (`search`) | App namespace in which `search` runs (lookups, macros, event types). |
| `sheet` | yes | `Results`, `Results 2`, ... | Sheet names, in the order of the sources. `\ / ? * [ ] :` are replaced by `_`, names are cut to 31 characters and made unique. |
| `filename` | no | `splunk_export` | File name without extension. Path and reserved characters are replaced by `_`. |
| `timestamp` | no | `true` | `false` removes the `_YYYYMMDD_HHMM` suffix (server time) from the file name. |
| `keep_text` | no | `false` | `true` writes every value as text (no number conversion). |
| `keep_internal` | no | `false` | `true` keeps every internal field starting with `_`. By default only `_time` and `_raw` are kept. |
| `dates` | no | `true` | `false` keeps `_time` as its original ISO text instead of an Excel date. |

Booleans accept `true/false`, `1/0`, `yes/no`, `on/off`.

### Responses

| Status | Body | When |
|--------|------|------|
| `200` | The workbook, `Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`, `Content-Disposition: attachment; filename=...` | Success |
| `400` | `{"error": "...", "status": 400}` | Missing/invalid parameter, both `sid` and `search`, SPL error, failed search |
| `401` | JSON error | Not authenticated |
| `403` | JSON error | Risky command in the SPL, `search` mode disabled, job of another user |
| `404` | JSON error | Unknown or expired `sid`, or **no results to export** |
| `405` | JSON error | Method other than GET/POST |
| `504` | JSON error | The job did not finish within `job_timeout` |

## Simple XML: export buttons (export_excel.js)

The script turns any element with the class `excel-export-btn` into an export button. It finds the job of the referenced search in the dashboard and posts its `sid` (and post-process SPL, if any) to the endpoint, then saves the returned file. **No second search is run.**

### 1. Load the script

On the root element (`<dashboard>` or `<form>`), in this app or **any other app**:

```xml
<form version="1.1" script="splunk_excel_extract:export_excel.js">
```

To combine with other scripts: `script="my_app:other.js, splunk_excel_extract:export_excel.js"`.

### 2. Give the search an `id`

```xml
<search id="my_search">
  <query>index=main | stats count by host, sourcetype</query>
  <earliest>-24h</earliest>
  <latest>now</latest>
</search>
```

Global searches, searches inside a panel and post-process searches (`<search id="..." base="...">`, including chains of post-processes) are supported.

### 3. Add a button

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

Several searches in one workbook:

```xml
<button class="btn excel-export-btn"
        data-search="search_a,search_b,search_c"
        data-sheet="Errors,Latency,Traffic"
        data-filename="weekly_report">Download full report</button>
```

### Button attributes

| Attribute | Default | Endpoint parameter |
|-----------|---------|--------------------|
| `data-search` | `search_export` | Comma-separated search ids (case-sensitive), one sheet each → `sid` / `postprocess` |
| `data-filename` | `splunk_export` | `filename` |
| `data-sheet` | `Results` | Comma-separated → `sheet` |
| `data-timestamp` | `true` | `timestamp` |
| `data-keep-text` | `false` | `keep_text` |
| `data-keep-internal` | `false` | `keep_internal` |
| `data-dates` | `true` | `dates` |

While the file is generated the button shows `Preparing...`; errors returned by the endpoint (e.g. `No results to export.`) are shown on the button for 5 seconds and logged in the browser console.

## Simple XML: plain link, no JavaScript

Store the job id in a token and link to the endpoint:

```xml
<search id="search_export">
  <query>index=main | stats count by host</query>
  <done>
    <set token="export_sid">$job.sid$</set>
  </done>
</search>
...
<html>
  <a class="btn btn-primary"
     href="/splunkd/__raw/services/excel_export?sid=$export_sid$&amp;filename=hosts&amp;sheet=Hosts">Download Excel</a>
</html>
```

Remember to write `&` as `&amp;` in XML. The link also works in `<drilldown><link target="_blank">...</link></drilldown>`.

## Dashboard Studio

Dashboard Studio does not run custom JavaScript, so it calls the endpoint with a **link**: a markdown link or a **Link to custom URL** drilldown. The browser downloads the file and stays on the dashboard.

### Option 1: reuse the job of a data source (`sid`)

1. Give the data source a `name` and enable job tokens with `"enableSmartSources": true`:

   ```json
   "dataSources": {
     "ds_events": {
       "type": "ds.search",
       "name": "events",
       "options": {
         "query": "index=main | stats count by host",
         "enableSmartSources": true
       }
     }
   }
   ```

2. Use the `$events:job.sid$` token in a markdown visualization:

   ```json
   "viz_export": {
     "type": "splunk.markdown",
     "options": {
       "markdown": "[⬇ Download Excel](/splunkd/__raw/services/excel_export?sid=$events:job.sid$&filename=hosts&sheet=Hosts)"
     }
   }
   ```

   or as a drilldown on any visualization (here a click on a table row):

   ```json
   "eventHandlers": [
     {
       "type": "drilldown.customUrl",
       "options": {
         "url": "/splunkd/__raw/services/excel_export?sid=$events:job.sid$&filename=hosts",
         "newTab": false
       }
     }
   ]
   ```

For a chained data source (`ds.chain`), pass the sid of the **base** data source and the chain's SPL in `postprocess` (URL-encoded).

### Option 2: run the search from the link (`search`)

When job tokens are not available, put the URL-encoded SPL in the link and pass the time range tokens:

```
/splunkd/__raw/services/excel_export?search=index%3Dmain%20%7C%20stats%20count%20by%20host&earliest=$global_time.earliest$&latest=$global_time.latest$&app=my_app&filename=hosts
```

Encode the SPL once, for example with `python3 -c "import urllib.parse,sys; print(urllib.parse.quote(sys.argv[1], safe=''))" 'index=main | stats count by host'`. Pass `app=<your app>` if the search uses lookups or macros of that app.

### Several sheets from Studio

Repeat the parameters: `?sid=$errors:job.sid$&sid=$traffic:job.sid$&sheet=Errors&sheet=Traffic&filename=report`.

## Calling the endpoint from scripts

```bash
# Run a search, 2 sheets, through splunkd (port 8089) with a token
curl -sk -H "Authorization: Bearer $SPLUNK_TOKEN" "https://splunk:8089/services/excel_export" \
     --data-urlencode "search=index=_internal log_level=ERROR | stats count by component" \
     --data-urlencode "search=index=_internal | stats count by sourcetype" \
     -d earliest=-24h -d sheet=Errors -d sheet=Sourcetypes -d filename=internal \
     -OJ                                    # save under the name sent by the server

# Export an existing job
curl -sk -u admin:changeme "https://splunk:8089/services/excel_export?sid=1696761234.123" -o job.xlsx
```

```python
import requests
r = requests.post("https://splunk:8089/services/excel_export", verify=False,
                  headers={"Authorization": "Bearer " + token},
                  json={"search": ["index=main | stats count by host"], "earliest": "-7d", "filename": "hosts"})
r.raise_for_status()
open("hosts.xlsx", "wb").write(r.content)
```

## How values are converted

| Splunk value | Excel cell |
|--------------|------------|
| Integer or decimal such as `42`, `-3.5` | Number |
| Value with a leading zero such as `007` or `000123` | Text (identifiers are never altered) |
| Number with more than 15 integer digits | Text (Excel keeps only 15 significant digits) |
| `_time` | Date/time formatted `yyyy-mm-dd hh:mm:ss`, in the **time zone of the Splunk user** (the offset Splunk returns) |
| Multivalue field | One cell, values separated by line breaks (wrapped) |
| Empty / null | Empty cell |
| Text longer than 32,767 characters | Truncated (Excel's limit per cell) |
| Control characters not allowed in XML | Removed |
| Value starting with `=` | Text, never a formula |
| Internal fields other than `_time` and `_raw` (`_bkt`, `_cd`, `_indextime`, ...) | Dropped (unless `keep_internal=true`) |

Columns appear in the order of the search results (use `| table` to choose it).

## Configuration (excel_export.conf)

Settings are in `default/excel_export.conf`; override them in `local/excel_export.conf` (read on each request, no restart needed). Full description in `README/excel_export.conf.spec`.

```ini
[settings]
allow_search       = true      # false: only "sid" exports are allowed
default_app        = search    # namespace of "search" requests without "app"
job_timeout        = 300       # seconds to wait for a job to finish
page_size          = 50000     # rows per request to splunkd (<= maxresultrows)
max_rows_per_sheet = 1048575   # Excel limit
max_sheets         = 20
risky_commands     = collect, delete, dump, map, mcollect, meventcollect, outputcsv, outputlookup, outputtext, run, runshellscript, script, sendalert, sendemail, tscollect
```

## Security

- **User permissions only.** The handler uses the session token of the caller (`passSystemAuth = false`): it can read only the jobs the user can read, and runs searches with the user's roles, indexes and quotas.
- **Authentication required** (`requireAuthentication = true`). From Splunk Web, the session cookie authenticates `GET` requests; `POST` requests also need the `X-Splunk-Form-Key` header (CSRF protection), which `export_excel.js` sends.
- **Forged links.** Because a simple `GET` link can run a search, the `search` and `postprocess` parameters refuse the SPL commands listed in `risky_commands` (commands that write, delete or send data). Set `allow_search = false` if you only want to export existing jobs.
- **No formula injection.** Values are always written as typed cells; `=...` stays text.
- The `sid` must match `[A-Za-z0-9_.@:-]`; file names and sheet names are sanitised.

## Limitations

- **Memory on the search head.** The whole workbook is built in memory by the handler. Hundreds of thousands of rows work, but very large exports cost CPU and RAM on the search head; aggregate with SPL first and lower `max_rows_per_sheet` if needed.
- **Excel limits.** 1,048,576 rows and 16,384 columns per sheet, 32,767 characters per cell, 31 characters per sheet name. Rows beyond `max_rows_per_sheet` are dropped.
- **Job lifetime.** A `sid` export only works while the job exists (dashboard jobs expire after about 10 minutes of inactivity). Reload the dashboard if you get `Search job ... not found`.
- **Time zone.** `_time` is written as wall-clock time in the user's Splunk time zone; the time zone itself is not stored in the cell.
- **Dashboard Studio** has no button component: use markdown links or drilldowns.

## Troubleshooting

| Symptom | Likely cause and fix |
|---------|----------------------|
| `404` on `/splunkd/__raw/services/excel_export` | Splunk was not restarted after install, or `web.conf` is not loaded. Test splunkd directly: `curl -k -u admin https://localhost:8089/services/excel_export` must answer 400 `Missing parameter`. If splunkd answers but Splunk Web does not, check `$SPLUNK_HOME/bin/splunk btool web list expose:excel_export`. |
| `{"error": "Missing parameter: 'sid' ..."}` | The link has no `sid`/`search`, often because the token (`$export_sid$`, `$events:job.sid$`) is not set yet: wait for the search to finish, check `enableSmartSources` in Studio. |
| `Search job ... not found` | The job expired or belongs to another user. Re-run the dashboard. |
| `The search uses commands that are not allowed...` | The SPL contains a command from `risky_commands`. Change the SPL or the setting. |
| `Timed out waiting for search job` | The search is longer than `job_timeout`. Increase it in `local/excel_export.conf`. |
| Button does nothing | Open the browser console (F12): `export_excel.js` must load without 404 and the dashboard root must have the `script` attribute. Bump `/_bump` and hard-refresh. |
| Button shows `Search "xxx" not found` | `data-search` does not match any search `id` (case-sensitive). |
| `Export failed: ...` (HTTP 500) | Unexpected error: see `index=_internal source=*excel_export.log` and `source=*splunkd.log* excel_export`. |
| The file opens as text / JSON | The endpoint returned an error (JSON) instead of the workbook: read the message. |

Logs: `$SPLUNK_HOME/var/log/splunk/excel_export.log`, e.g. `index=_internal source=*excel_export.log "export ok"` gives who exported what, how many rows and how long it took.

## Development

```bash
python3 -m unittest discover -s tests -v     # unit tests, no Splunk needed (uses src/bin/lib)
./build.sh                                   # package
```

- `src/bin/xlsx_builder.py` contains the conversion and formatting logic (pure Python + openpyxl).
- `src/bin/excel_export_handler.py` contains the REST handler; splunkd access goes through `SplunkClient`, replaced by a fake in the tests.
- To refresh the vendored libraries, keep versions compatible with Python 3.7 (Splunk 9.0–9.2):

  ```bash
  rm -rf src/bin/lib && python3 -m pip install --no-deps --no-compile --target src/bin/lib -r requirements.txt
  ```

## Packaging, Splunk Cloud and uninstalling

### Validate before uploading

```bash
pip install splunk-appinspect
splunk-appinspect inspect dist/splunk_excel_extract-2.0.0.tgz --mode precert \
    --included-tags cloud --included-tags splunk_appinspect
```

### Splunk Cloud

Install the package as a **private app** (**Apps → Manage Apps → Install app from file**, or the Admin Config Service). It goes through AppInspect vetting; custom REST handlers in Python 3 are allowed in private apps. Users then call the endpoint through Splunk Web exactly as on Splunk Enterprise (`/splunkd/__raw/services/excel_export`).

### Uninstall

**Apps → Manage Apps → Delete**, or `$SPLUNK_HOME/bin/splunk remove app splunk_excel_extract`. Dashboards that reference `splunk_excel_extract:export_excel.js` or the endpoint stop exporting once the app is removed.

## Upgrading from 1.x (SheetJS)

Version 1.x built the workbook in the browser with SheetJS. Version 2.0 replaces it with the server-side endpoint:

- **Existing dashboards keep working unchanged**: same script (`splunk_excel_extract:export_excel.js`), same `excel-export-btn` class, same `data-*` attributes.
- `xlsx.full.min.js` (SheetJS, ~950 KB) is no longer shipped: the browser does not download or run it any more.
- Improvements: no 50,000-row cap, formatted header, frozen panes, Dashboard Studio support, `_time` in the user's Splunk time zone instead of the browser's.
- After upgrading: restart Splunk (new `restmap.conf`/`web.conf`), bump `/_bump`, hard-refresh.

## Licenses

- This app is licensed under the Apache License 2.0, Copyright 2026 Bichoumac. See `LICENSE.txt`.
- **openpyxl** (MIT) and **et_xmlfile** (MIT) are bundled in `bin/lib/`. Their licenses are in `licenses/`.
