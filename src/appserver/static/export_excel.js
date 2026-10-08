/*
 * export_excel.js
 *
 * Turns any <button class="excel-export-btn"> in a Splunk Simple XML dashboard
 * into a "Download as .xlsx" button. The workbook is built server-side by the
 * /services/excel_export REST endpoint of this app (Python + openpyxl), from
 * the job the dashboard already ran: no second search is dispatched.
 *
 * Usage in a dashboard:
 *
 *   <dashboard version="1.1" script="splunk_excel_extract:export_excel.js">
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
 */

require([
  'jquery',
  'splunkjs/mvc',
  'splunk.util',
  'splunkjs/mvc/simplexml/ready!'
], function ($, mvc, splunkUtil) {
  'use strict';

  var ENDPOINT = '/splunkd/__raw/services/excel_export';
  var DEFAULT_SEARCH_ID = 'search_export';
  var DEFAULT_FILENAME = 'splunk_export';
  var SID_TIMEOUT_MS = 5 * 60 * 1000;   // give up waiting for a search to start after 5 minutes

  function splitList(value) {
    return String(value || '').split(',').map(function (s) {
      return s.replace(/^\s+|\s+$/g, '');
    }).filter(function (s) { return s.length > 0; });
  }

  function isTrue(value) { return String(value).toLowerCase() === 'true'; }
  function isFalse(value) { return String(value).toLowerCase() === 'false'; }

  // App of the current page, used as namespace (only matters for "search" requests).
  function currentApp() {
    var m = window.location.pathname.match(/\/app\/([^\/]+)/);
    return m ? m[1] : 'search';
  }

  // Resolved search string of a manager (tokens already substituted).
  function queryOf(manager) {
    if (manager.query && typeof manager.query.resolve === 'function') {
      return manager.query.resolve();
    }
    return manager.settings.get('search') || '';
  }

  // Walk up post-process managers. Returns {root: <SearchManager>, postprocess: "<SPL>"}.
  function resolveSource(manager) {
    var parts = [];
    var current = manager;
    var guard = 0;
    while (current && current.settings && current.settings.get('managerid') && guard++ < 20) {
      parts.unshift(queryOf(current));
      current = mvc.Components.get(current.settings.get('managerid'));
    }
    if (!current) { return null; }
    var postprocess = '';
    parts.forEach(function (part) {
      part = String(part || '').replace(/^\s+|\s+$/g, '');
      if (!part) { return; }
      if (!postprocess) { postprocess = part; return; }
      postprocess += (part.charAt(0) === '|' ? ' ' : ' | ') + part;
    });
    return { root: current, postprocess: postprocess };
  }

  // Calls done(sid) as soon as the job of a SearchManager has a sid.
  function whenSid(manager, done, fail) {
    function sid() { return manager.job && manager.job.sid; }
    if (sid()) { done(sid()); return; }

    var timer = setTimeout(function () { finish(); fail('Timed out waiting for the search to start.'); }, SID_TIMEOUT_MS);
    function finish() {
      clearTimeout(timer);
      manager.off('search:start search:progress search:done', onEvent);
      manager.off('search:failed search:error', onFail);
    }
    function onEvent() { if (sid()) { finish(); done(sid()); } }
    function onFail() { finish(); fail('The search failed.'); }
    manager.on('search:start search:progress search:done', onEvent);
    manager.on('search:failed search:error', onFail);
  }

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
    setTimeout(function () { $btn.text(label); }, 5000);
    console.warn('[export_excel] ' + message);
  }

  function fileNameFrom(xhr, fallback) {
    var header = xhr.getResponseHeader('Content-Disposition') || '';
    var utf8 = header.match(/filename\*=UTF-8''([^;]+)/i);
    if (utf8) { try { return decodeURIComponent(utf8[1]); } catch (e) { /* ignore */ } }
    var plain = header.match(/filename="?([^";]+)"?/i);
    return plain ? plain[1] : fallback;
  }

  function saveBlob(blob, name) {
    if (window.navigator && window.navigator.msSaveOrOpenBlob) {   // legacy Edge
      window.navigator.msSaveOrOpenBlob(blob, name);
      return;
    }
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(function () { document.body.removeChild(a); URL.revokeObjectURL(url); }, 1000);
  }

  function errorMessage(xhr, callback) {
    var blob = xhr.response;
    if (!blob || !blob.size) { callback('Export failed (HTTP ' + xhr.status + ')'); return; }
    var reader = new FileReader();
    reader.onload = function () {
      var msg = 'Export failed (HTTP ' + xhr.status + ')';
      try { msg = JSON.parse(reader.result).error || msg; } catch (e) { /* not JSON */ }
      callback(msg);
    };
    reader.readAsText(blob);
  }

  function download(params, fallbackName, $btn) {
    var xhr = new XMLHttpRequest();
    xhr.open('POST', splunkUtil.make_url(ENDPOINT), true);
    xhr.responseType = 'blob';
    xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded; charset=UTF-8');
    xhr.setRequestHeader('X-Requested-With', 'XMLHttpRequest');
    xhr.setRequestHeader('X-Splunk-Form-Key', splunkUtil.getFormKey());
    xhr.onload = function () {
      setBusy($btn, false);
      if (xhr.status === 200) {
        saveBlob(xhr.response, fileNameFrom(xhr, fallbackName));
      } else {
        errorMessage(xhr, function (msg) { flash($btn, msg); });
      }
    };
    xhr.onerror = function () {
      setBusy($btn, false);
      flash($btn, 'Export failed (network error)');
    };
    xhr.send($.param(params, true));
  }

  $(document).on('click', '.excel-export-btn', function (event) {
    event.preventDefault();
    var $btn = $(this);
    if ($btn.data('excelBusy')) { return; }

    var searchIds = splitList($btn.attr('data-search'));
    if (!searchIds.length) { searchIds = [DEFAULT_SEARCH_ID]; }

    var sources = [];
    for (var i = 0; i < searchIds.length; i++) {
      var manager = mvc.Components.get(searchIds[i]);
      var source = manager && manager.settings ? resolveSource(manager) : null;
      if (!source) {
        flash($btn, 'Search "' + searchIds[i] + '" not found');
        return;
      }
      sources.push(source);
    }

    var filename = String($btn.attr('data-filename') || DEFAULT_FILENAME).replace(/\.xlsx$/i, '');
    var params = {
      sid: [],
      postprocess: [],
      sheet: splitList($btn.attr('data-sheet')),
      filename: filename,
      timestamp: isFalse($btn.attr('data-timestamp')) ? 'false' : 'true',
      keep_text: isTrue($btn.attr('data-keep-text')) ? 'true' : 'false',
      keep_internal: isTrue($btn.attr('data-keep-internal')) ? 'true' : 'false',
      dates: isFalse($btn.attr('data-dates')) ? 'false' : 'true',
      app: currentApp()
    };

    setBusy($btn, true);
    var remaining = sources.length;
    var failed = false;
    sources.forEach(function (source, idx) {
      whenSid(source.root, function (sid) {
        if (failed) { return; }
        params.sid[idx] = sid;
        params.postprocess[idx] = source.postprocess;
        remaining -= 1;
        if (remaining === 0) { download(params, filename + '.xlsx', $btn); }
      }, function (message) {
        if (failed) { return; }
        failed = true;
        setBusy($btn, false);
        flash($btn, message);
      });
    });
  });
});
