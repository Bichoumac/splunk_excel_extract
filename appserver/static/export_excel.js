/*
 * export_excel.js
 *
 * Turns any <button class="excel-export-btn"> in a Splunk Simple XML dashboard
 * into a "Download as .xlsx" button. The workbook is built in the browser with
 * SheetJS (xlsx.full.min.js, bundled in this folder), so no server-side
 * component is needed.
 *
 * Usage in a dashboard:
 *
 *   <dashboard version="1.1" script="export_excel.js">
 *     <search id="my_search"> ... </search>
 *     ...
 *     <html>
 *       <button class="btn btn-primary excel-export-btn"
 *               data-search="my_search"
 *               data-filename="my_report"
 *               data-sheet="Results">Download Excel</button>
 *     </html>
 *
 * Supported data-* attributes (all optional except data-search is strongly advised):
 *
 *   data-search         Search/post-process id, or a comma-separated list of ids
 *                       (one sheet per id). Default: "search_export".
 *   data-filename       File name without extension. Default: "splunk_export".
 *   data-sheet          Sheet name, or comma-separated names matching data-search.
 *                       Default: "Results" (or "Results 2", "Results 3", ...).
 *   data-timestamp      "false" to NOT append _YYYYMMDD_HHmm to the file name.
 *   data-keep-text      "true" to keep every value as text (no number conversion).
 *   data-keep-internal  "true" to keep internal fields starting with "_"
 *                       (by default only _time and _raw are kept).
 *   data-dates          "false" to keep _time as text instead of an Excel date.
 *
 * The script works from any app: reference it as
 * script="splunk_excel_extract:export_excel.js" and it loads xlsx.full.min.js
 * from its own app (splunk_excel_extract/appserver/static/), not the dashboard's.
 */

(function () {
  'use strict';

  // ---------------------------------------------------------------------------
  // Locate the app that ships this script (and xlsx.full.min.js next to it),
  // then register SheetJS under the RequireJS name "xlsx". SheetJS registers
  // itself with the *named* AMD id "xlsx", so the path alias must use exactly
  // that name.
  //
  // The app is NOT taken from the page URL: a dashboard in another app that
  // loads script="splunk_excel_extract:export_excel.js" would otherwise look for
  // xlsx.full.min.js in its own app and fail with 'Script error for "xlsx"'.
  // ---------------------------------------------------------------------------
  var APP = 'splunk_excel_extract';
  var scripts = document.getElementsByTagName('script');
  for (var s = 0; s < scripts.length; s++) {
    var src = scripts[s].getAttribute('src') || '';
    var found = src.match(/\/app\/([^\/]+)\/export_excel\.js/);
    if (found) { APP = found[1]; break; }
  }

  require.config({
    paths: {
      xlsx: '../app/' + APP + '/xlsx.full.min'
    }
  });

  require([
    'jquery',
    'splunkjs/mvc',
    'xlsx',
    'splunkjs/mvc/simplexml/ready!'
  ], function ($, mvc, XLSXModule) {

    var XLSX = (XLSXModule && XLSXModule.utils) ? XLSXModule : window.XLSX;
    if (!XLSX || !XLSX.utils) {
      console.error('[export_excel] SheetJS (xlsx.full.min.js) could not be loaded from app "' + APP + '".');
      return;
    }

    var DEFAULT_SEARCH_ID = 'search_export';
    var DEFAULT_FILENAME = 'splunk_export';
    var DEFAULT_SHEET = 'Results';
    var TIMEOUT_MS = 5 * 60 * 1000;        // give up waiting for a search after 5 minutes
    var MAX_CELL_CHARS = 32767;            // Excel hard limit per cell
    var MAX_COL_WIDTH = 60;
    var DATE_FORMAT = 'yyyy-mm-dd hh:mm:ss';

    // One results model per search id, created lazily on first click and then reused.
    var modelCache = {};

    // -------------------------------------------------------------------------
    // Small helpers
    // -------------------------------------------------------------------------
    function pad(n) { return n < 10 ? '0' + n : String(n); }

    function timestamp() {
      var d = new Date();
      return d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate()) +
        '_' + pad(d.getHours()) + pad(d.getMinutes());
    }

    function splitList(value) {
      return String(value || '').split(',').map(function (s) {
        return s.replace(/^\s+|\s+$/g, '');
      }).filter(function (s) { return s.length > 0; });
    }

    function isTrue(value) { return String(value).toLowerCase() === 'true'; }
    function isFalse(value) { return String(value).toLowerCase() === 'false'; }

    // Excel sheet names: max 31 chars, none of  \ / ? * [ ] :  and must be unique.
    function makeSheetNames(requested, count) {
      var used = {};
      var names = [];
      for (var i = 0; i < count; i++) {
        var raw = requested[i] || (i === 0 ? DEFAULT_SHEET : DEFAULT_SHEET + ' ' + (i + 1));
        var name = String(raw).replace(/[\\\/\?\*\[\]:]/g, '_').substring(0, 31) || DEFAULT_SHEET;
        var base = name;
        var suffix = 2;
        while (used[name.toLowerCase()]) {
          var tail = ' ' + suffix++;
          name = base.substring(0, 31 - tail.length) + tail;
        }
        used[name.toLowerCase()] = true;
        names.push(name);
      }
      return names;
    }

    function fieldName(field) {
      return (typeof field === 'string') ? field : (field && field.name) || '';
    }

    // Excel stores a date/time as the number of days since 1899-12-30, using the
    // local wall-clock time. SheetJS' own Date handling is deliberately avoided:
    // it adds a few seconds of error in time zones that had an odd offset in 1899
    // (e.g. Europe/Paris, Europe/Amsterdam), so the serial number is computed here.
    function excelSerial(d) {
      var wallClockAsUtc = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(),
        d.getHours(), d.getMinutes(), d.getSeconds(), d.getMilliseconds());
      return (wallClockAsUtc - Date.UTC(1899, 11, 30)) / 86400000;
    }

    // Convert a raw Splunk value into a value suitable for an Excel cell.
    // Dates are returned as {serial: <number>} so the caller can apply a date format.
    function toCell(value, field, opts) {
      if (value === null || value === undefined) { return ''; }
      if (Object.prototype.toString.call(value) === '[object Array]') {
        value = value.join('\n');                 // multivalue field
      }
      var text = String(value);

      if (field === '_time' && opts.dates) {
        var parsed = new Date(text);
        if (!isNaN(parsed.getTime())) { return { serial: excelSerial(parsed) }; }
      }

      if (!opts.keepText && /^-?(0|[1-9]\d{0,14})(\.\d+)?$/.test(text)) {
        return Number(text);                      // "42" -> 42, but "007" stays text
      }

      if (text.length > MAX_CELL_CHARS) { text = text.substring(0, MAX_CELL_CHARS); }
      return text;
    }

    // Build one worksheet from a Splunk json_rows payload ({fields: [...], rows: [[...]]}).
    function buildSheet(data, opts) {
      var allFields = (data.fields || []).map(fieldName);
      var keep = [];
      allFields.forEach(function (name, idx) {
        if (opts.keepInternal || name.charAt(0) !== '_' || name === '_time' || name === '_raw') {
          keep.push(idx);
        }
      });

      var header = keep.map(function (idx) { return allFields[idx]; });
      var dateCells = [];                         // cells that need the date number format
      var rows = (data.rows || []).map(function (row, r) {
        return keep.map(function (idx, c) {
          var cell = toCell(row[idx], allFields[idx], opts);
          if (cell && typeof cell === 'object') {
            dateCells.push({ r: r + 1, c: c });   // +1: row 0 is the header
            return cell.serial;
          }
          return cell;
        });
      });

      var ws = XLSX.utils.aoa_to_sheet([header].concat(rows));

      // Column widths (based on content).
      var widths = header.map(function (h) { return String(h).length; });
      rows.forEach(function (row) {
        row.forEach(function (cell, c) {
          var len = String(cell).split('\n')[0].length;
          if (len > widths[c]) { widths[c] = len; }
        });
      });
      dateCells.forEach(function (pos) {
        widths[pos.c] = Math.max(widths[pos.c], DATE_FORMAT.length);
      });
      ws['!cols'] = widths.map(function (w) {
        return { wch: Math.min(MAX_COL_WIDTH, Math.max(8, w + 2)) };
      });

      // Excel date format for _time.
      dateCells.forEach(function (pos) {
        var cell = ws[XLSX.utils.encode_cell({ r: pos.r, c: pos.c })];
        if (cell) { cell.z = DATE_FORMAT; }
      });

      if (ws['!ref']) { ws['!autofilter'] = { ref: ws['!ref'] }; }
      return ws;
    }

    // -------------------------------------------------------------------------
    // Fetching results from a search manager
    // -------------------------------------------------------------------------
    function getResultsModel(manager, searchId) {
      var entry = modelCache[searchId];
      // Re-create the model if the search manager behind this id has been replaced.
      if (!entry || entry.manager !== manager) {
        entry = modelCache[searchId] = {
          manager: manager,
          model: manager.data('results', { count: 0, output_mode: 'json_rows' })
        };
      }
      return entry.model;
    }

    function readyData(model) {
      var d = model.data();
      return (d && d.rows) ? d : null;
    }

    // Calls onReady(data) as soon as the model holds results, or onError(message).
    function whenReady(model, manager, onReady, onError) {
      var data = readyData(model);
      if (data) { onReady(data); return; }

      var timer = null;

      function finish() {
        if (timer) { clearTimeout(timer); }
        model.off('data', onData);
        manager.off('search:failed', onFail);
        manager.off('search:error', onFail);
      }
      function onData() {
        var d = readyData(model);
        if (d) { finish(); onReady(d); }
      }
      function onFail() {
        finish();
        onError('The search failed.');
      }

      model.on('data', onData);
      manager.on('search:failed', onFail);
      manager.on('search:error', onFail);
      timer = setTimeout(function () {
        finish();
        onError('Timed out waiting for results.');
      }, TIMEOUT_MS);
    }

    // -------------------------------------------------------------------------
    // Button state
    // -------------------------------------------------------------------------
    function setBusy($btn, busy) {
      if (busy) {
        $btn.data('excelLabel', $btn.text());
        $btn.data('excelBusy', true);
        $btn.prop('disabled', true);
        $btn.text('Preparing...');
      } else {
        $btn.data('excelBusy', false);
        $btn.prop('disabled', false);
        $btn.text($btn.data('excelLabel') || 'Download Excel');
      }
    }

    function flash($btn, message) {
      var label = $btn.data('excelLabel') || $btn.text();
      $btn.data('excelLabel', label);
      $btn.text(message);
      setTimeout(function () { $btn.text(label); }, 4000);
      console.warn('[export_excel] ' + message);
    }

    // -------------------------------------------------------------------------
    // Click handler
    // -------------------------------------------------------------------------
    function readConfig($btn) {
      var searchIds = splitList($btn.attr('data-search'));
      if (!searchIds.length) { searchIds = [DEFAULT_SEARCH_ID]; }

      var base = String($btn.attr('data-filename') || DEFAULT_FILENAME).replace(/\.xlsx$/i, '');
      var stamp = isFalse($btn.attr('data-timestamp')) ? '' : '_' + timestamp();

      return {
        searchIds: searchIds,
        sheetNames: makeSheetNames(splitList($btn.attr('data-sheet')), searchIds.length),
        fileName: base + stamp + '.xlsx',
        sheetOpts: {
          keepText: isTrue($btn.attr('data-keep-text')),
          keepInternal: isTrue($btn.attr('data-keep-internal')),
          dates: !isFalse($btn.attr('data-dates'))
        }
      };
    }

    $(document).on('click', '.excel-export-btn', function (event) {
      event.preventDefault();

      var $btn = $(this);
      if ($btn.data('excelBusy')) { return; }

      var cfg = readConfig($btn);
      var managers = [];
      for (var i = 0; i < cfg.searchIds.length; i++) {
        var manager = mvc.Components.get(cfg.searchIds[i]);
        if (!manager || typeof manager.data !== 'function') {
          flash($btn, 'Search "' + cfg.searchIds[i] + '" not found');
          return;
        }
        managers.push(manager);
      }

      setBusy($btn, true);

      var collected = new Array(managers.length);
      var remaining = managers.length;
      var failed = false;

      function finishExport() {
        try {
          var total = 0;
          var wb = XLSX.utils.book_new();
          collected.forEach(function (data, idx) {
            total += (data.rows || []).length;
            XLSX.utils.book_append_sheet(wb, buildSheet(data, cfg.sheetOpts), cfg.sheetNames[idx]);
          });
          if (total === 0) {
            setBusy($btn, false);
            flash($btn, 'No results to export');
            return;
          }
          XLSX.writeFile(wb, cfg.fileName);
          setBusy($btn, false);
        } catch (err) {
          console.error('[export_excel] Export failed', err);
          setBusy($btn, false);
          flash($btn, 'Export failed (see console)');
        }
      }

      managers.forEach(function (manager, idx) {
        var model = getResultsModel(manager, cfg.searchIds[idx]);
        whenReady(model, manager, function (data) {
          if (failed) { return; }
          collected[idx] = data;
          remaining -= 1;
          if (remaining === 0) { finishExport(); }
        }, function (message) {
          if (failed) { return; }
          failed = true;
          setBusy($btn, false);
          flash($btn, message);
        });
      });
    });
  });
})();
