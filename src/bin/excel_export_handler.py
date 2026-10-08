"""
excel_export_handler.py

Custom REST endpoint (persistent handler) that returns Splunk search results
as an .xlsx workbook built with openpyxl.

    GET|POST /services/excel_export
    (from the browser: /<locale>/splunkd/__raw/services/excel_export)

One sheet is produced per source. A source is either:
  - an existing search job:  sid=<sid>  [postprocess=<SPL>]
  - a new search:            search=<SPL>  [earliest=..] [latest=..]
Repeat the parameter to get several sheets (sid=a&sid=b, or search=..&search=..).

Other parameters: sheet (repeatable), filename, timestamp, keep_text,
keep_internal, dates, app. See README.md for the full reference.

All calls to splunkd use the session of the calling user, so the export never
returns more than what that user is allowed to search.
"""

import base64
import json
import logging
import os
import re
import sys
import time
from datetime import datetime

try:                                    # Python 3
    from configparser import ConfigParser
    from urllib.parse import quote
except ImportError:                     # pragma: no cover
    from ConfigParser import ConfigParser
    from urllib import quote

BIN_DIR = os.path.dirname(os.path.abspath(__file__))
APP_DIR = os.path.dirname(BIN_DIR)
APP_NAME = os.path.basename(APP_DIR)
LIB_DIR = os.path.join(BIN_DIR, "lib")
# Splunk's persistent server loads this file by path: neither bin/ (for
# xlsx_builder) nor bin/lib (for openpyxl) is on sys.path by default.
for _path in (LIB_DIR, BIN_DIR):
    if _path not in sys.path:
        sys.path.insert(0, _path)

try:
    import xlsx_builder  # noqa: E402
    IMPORT_ERROR = None
except Exception:                       # reported by handle() instead of breaking the protocol
    import traceback
    xlsx_builder = None
    IMPORT_ERROR = traceback.format_exc()

try:
    from splunk.persistconn.application import PersistentServerConnectionApplication
except ImportError:                     # outside Splunk (unit tests)
    class PersistentServerConnectionApplication(object):
        def __init__(self):
            pass

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
TRUE_VALUES = ("1", "true", "yes", "y", "t", "on")
FALSE_VALUES = ("0", "false", "no", "n", "f", "off")
SID_RE = re.compile(r"^[A-Za-z0-9_.@:\-]{1,256}$")
APP_RE = re.compile(r"^[A-Za-z0-9_.\-]{1,100}$")
DEFAULT_RISKY_COMMANDS = (
    "collect, delete, dump, map, mcollect, meventcollect, outputcsv, outputlookup, "
    "outputtext, run, runshellscript, script, sendalert, sendemail, tscollect"
)


def get_logger():
    logger = logging.getLogger("excel_export")
    if not logger.handlers:
        logger.setLevel(logging.INFO)
        home = os.environ.get("SPLUNK_HOME")
        if home:
            path = os.path.join(home, "var", "log", "splunk", "excel_export.log")
            try:
                handler = logging.FileHandler(path)
                handler.setFormatter(logging.Formatter(
                    "%(asctime)s %(levelname)s pid=%(process)d %(message)s"))
                logger.addHandler(handler)
            except (IOError, OSError):
                pass
        if not logger.handlers:
            logger.addHandler(logging.NullHandler())
    return logger


LOG = get_logger()


class ExportError(Exception):
    def __init__(self, message, status=400):
        Exception.__init__(self, message)
        self.status = status


# -----------------------------------------------------------------------------
# Configuration (default/excel_export.conf, overridden by local/excel_export.conf)
# -----------------------------------------------------------------------------
class Settings(object):
    def __init__(self, values=None):
        values = values or {}
        self.allow_search = parse_bool(values.get("allow_search"), True)
        self.default_app = values.get("default_app") or "search"
        self.job_timeout = to_int(values.get("job_timeout"), 300)
        self.page_size = max(1, to_int(values.get("page_size"), 50000))
        self.max_rows_per_sheet = max(1, min(to_int(values.get("max_rows_per_sheet"), xlsx_builder.MAX_ROWS_PER_SHEET),
                                             xlsx_builder.MAX_ROWS_PER_SHEET))
        self.max_sheets = max(1, to_int(values.get("max_sheets"), 20))
        commands = values.get("risky_commands")
        if commands is None:
            commands = DEFAULT_RISKY_COMMANDS
        self.risky_commands = set(c.strip().lower() for c in commands.split(",") if c.strip())

    @classmethod
    def load(cls, app_dir=APP_DIR):
        parser = ConfigParser(interpolation=None)
        parser.read([os.path.join(app_dir, d, "excel_export.conf") for d in ("default", "local")])
        values = dict(parser.items("settings")) if parser.has_section("settings") else {}
        return cls(values)


def to_int(value, default):
    try:
        return int(str(value).strip())
    except (TypeError, ValueError):
        return default


def parse_bool(value, default):
    if value is None:
        return default
    text = str(value).strip().lower()
    if text in TRUE_VALUES:
        return True
    if text in FALSE_VALUES:
        return False
    return default


# -----------------------------------------------------------------------------
# splunkd access
# -----------------------------------------------------------------------------
class SplunkClient(object):
    """Thin wrapper around splunk.rest.simpleRequest using the caller's session."""

    def __init__(self, session_key):
        self.session_key = session_key

    def request(self, method, path, getargs=None, postargs=None):
        import splunk
        import splunk.rest as rest
        try:
            response, content = rest.simpleRequest(
                path, sessionKey=self.session_key, method=method,
                getargs=getargs, postargs=postargs, raiseAllErrors=False, timeout=120)
        except splunk.AuthenticationFailed:
            raise ExportError("Authentication failed.", 401)
        except splunk.AuthorizationFailed:
            raise ExportError("You are not allowed to access %s." % path, 403)
        except splunk.ResourceNotFound:
            # Let the caller turn it into a meaningful message (e.g. expired job).
            return 404, json.dumps({"messages": [{"type": "ERROR", "text": "Not found: %s" % path}]})
        status = int(response.status)
        if isinstance(content, bytes):
            content = content.decode("utf-8", "replace")
        return status, content


def splunk_messages(content):
    try:
        messages = json.loads(content).get("messages") or []
        return "; ".join(m.get("text", "") for m in messages if m.get("text"))
    except (ValueError, AttributeError):
        return (content or "")[:300]


class SearchRunner(object):
    def __init__(self, client, settings, sleep=time.sleep, clock=time.time):
        self.client = client
        self.settings = settings
        self.sleep = sleep
        self.clock = clock

    def dispatch(self, user, app, search, earliest, latest):
        query = search.strip()
        if not (query.startswith("|") or query.lower().startswith("search ")):
            query = "search " + query
        postargs = {"search": query, "output_mode": "json", "exec_mode": "normal", "timeout": 120}
        if earliest:
            postargs["earliest_time"] = earliest
        if latest:
            postargs["latest_time"] = latest
        path = "/servicesNS/%s/%s/search/jobs" % (quote(user, safe=""), quote(app, safe=""))
        status, content = self.client.request("POST", path, postargs=postargs)
        if status >= 300:
            raise ExportError("Could not start the search: %s" % splunk_messages(content), 400 if status < 500 else 502)
        return json.loads(content)["sid"]

    def wait(self, sid):
        deadline = self.clock() + self.settings.job_timeout
        delay = 0.2
        while True:
            status, content = self.client.request(
                "GET", "/services/search/jobs/%s" % quote(sid, safe=""), getargs={"output_mode": "json"})
            if status == 404:
                raise ExportError("Search job %s not found: it has expired (dashboard jobs are kept about "
                                  "10 minutes), was cancelled, or is not visible to you. Reload the dashboard, "
                                  "or pass 'search' as well so the export can re-run it." % sid, 404)
            if status >= 300:
                raise ExportError("Could not read job %s: %s" % (sid, splunk_messages(content)), 502)
            job = json.loads(content)["entry"][0]["content"]
            if job.get("isFailed") or job.get("dispatchState") == "FAILED":
                msgs = job.get("messages") or {}
                text = "; ".join(m for v in msgs.values() for m in (v if isinstance(v, list) else [v]))
                raise ExportError("The search failed. %s" % text, 400)
            if job.get("isDone") or job.get("dispatchState") == "DONE":
                return job
            if self.clock() >= deadline:
                raise ExportError("Timed out waiting for search job %s." % sid, 504)
            self.sleep(delay)
            delay = min(delay * 2, 2.0)

    def read_page(self, sid, getargs, postprocess):
        """Post-processing is refused by the v1 results endpoint on recent Splunk
        versions ("Postprocessing search is blocked in API v1"): use a POST on
        search/v2 (v2 only accepts post-processing with POST), and fall back to
        v1 on versions that do not have it (before 9.0.1)."""
        quoted = quote(sid, safe="")
        if postprocess:
            status, content = self.client.request(
                "POST", "/services/search/v2/jobs/%s/results" % quoted, postargs=getargs)
            if status != 404:
                return status, content
        return self.client.request("GET", "/services/search/jobs/%s/results" % quoted, getargs=getargs)

    def results(self, sid, postprocess=None):
        """Read every result row of a finished job, page by page, so the
        [restapi] maxresultrows limit does not truncate the export."""
        fields = []
        index = {}
        rows = []
        offset = 0
        limit = self.settings.max_rows_per_sheet
        while len(rows) < limit:
            getargs = {"output_mode": "json_rows", "offset": offset,
                       "count": min(self.settings.page_size, limit - len(rows))}
            if postprocess:
                getargs["search"] = postprocess
            status, content = self.read_page(sid, getargs, bool(postprocess))
            if status == 204 or not content:
                break
            if status >= 300:
                raise ExportError("Could not read results of %s: %s" % (sid, splunk_messages(content)),
                                  400 if status < 500 else 502)
            page = json.loads(content)
            page_fields = [f if isinstance(f, str) else f.get("name", "") for f in page.get("fields") or []]
            page_rows = page.get("rows") or []
            if not page_rows:
                break
            # Fields can differ from one page to the next: map them onto one ordered list.
            mapping = []
            for name in page_fields:
                if name not in index:
                    index[name] = len(fields)
                    fields.append(name)
                mapping.append(index[name])
            for row in page_rows:
                out = [None] * len(fields)
                for i, value in enumerate(row):
                    if i < len(mapping):
                        out[mapping[i]] = value
                rows.append(out)
            offset += len(page_rows)
        width = len(fields)
        for row in rows:
            if len(row) < width:
                row.extend([None] * (width - len(row)))
        return fields, rows[:limit]

    def cancel(self, sid):
        try:
            self.client.request("POST", "/services/search/jobs/%s/control" % quote(sid, safe=""),
                                postargs={"action": "cancel"})
        except Exception:                           # best effort cleanup
            LOG.warning("could not cancel job sid=%s", sid)


# -----------------------------------------------------------------------------
# Request parsing
# -----------------------------------------------------------------------------
def risky_commands_in(spl, risky):
    """Return the risky commands used in an SPL string (after a pipe, at the
    start or at the start of a subsearch)."""
    found = []
    for match in re.finditer(r"(?:^|\||\[)\s*([A-Za-z_][A-Za-z0-9_]*)", spl or ""):
        name = match.group(1).lower()
        if name in risky and name not in found:
            found.append(name)
    return found


def safe_filename(name):
    name = re.sub(r"\.xlsx$", "", str(name or "").strip(), flags=re.I)
    name = re.sub(r'[\x00-\x1f\x7f\\/:*?"<>|]+', "_", name).strip(" .")
    return name[:150] or "splunk_export"


def content_disposition(filename):
    ascii_name = re.sub(r"[^A-Za-z0-9._\-]", "_", filename)
    return "attachment; filename=\"%s\"; filename*=UTF-8''%s" % (ascii_name, quote(filename, safe=""))


class ExportRequest(object):
    def __init__(self, params, settings, user):
        def many(key):
            return [v for v in params.get(key, []) if v is not None]

        def one(key, default=None):
            values = many(key)
            return values[-1] if values else default

        def aligned(key, count):
            values = many(key)
            if len(values) == 1:
                return values * count
            return [(values[i] if i < len(values) else "") for i in range(count)]

        self.searches = [s for s in many("search") if s.strip()]
        sids = [s.strip() for s in many("sid")]
        if self.searches:
            # sid + search: the sid is used while the job exists, the search re-runs it
            # otherwise (Dashboard Studio jobs expire ~10 minutes after they finish).
            # An empty or unresolved sid token ("$ds:job.sid$") simply means "run the search".
            if len([s for s in sids if s]) > len(self.searches):
                raise ExportError("When 'sid' and 'search' are combined, give one 'search' per 'sid' "
                                  "(the search is run when the job no longer exists).")
            count = len(self.searches)
            self.sids = [(sids[i] if i < len(sids) and SID_RE.match(sids[i]) else "") for i in range(count)]
        else:
            self.sids = [s for s in sids if s]
            count = len(self.sids)
            if not count:
                raise ExportError("Missing parameter: 'sid' (existing search job) or 'search' (SPL to run).")
            for sid in self.sids:
                if not SID_RE.match(sid):
                    raise ExportError("Invalid sid: %r" % sid)
        if count > settings.max_sheets:
            raise ExportError("Too many sheets requested (%d, max %d)." % (count, settings.max_sheets))

        self.postprocess = aligned("postprocess", count) if any(self.sids) else [""] * count
        self.earliest = aligned("earliest", count)
        self.latest = aligned("latest", count)

        self.app = (one("app") or settings.default_app).strip()
        if not APP_RE.match(self.app):
            raise ExportError("Invalid app: %r" % self.app)

        self.allow_search = settings.allow_search
        if self.searches and not settings.allow_search and not all(self.sids):
            raise ExportError("Running a new search through this endpoint is disabled "
                              "(allow_search = false). Pass the 'sid' of an existing job instead.", 403)
        for spl in self.searches + self.postprocess:
            risky = risky_commands_in(spl, settings.risky_commands)
            if risky:
                raise ExportError("The search uses commands that are not allowed in an export: %s."
                                  % ", ".join(risky), 403)

        self.sheet_names = xlsx_builder.make_sheet_names(many("sheet"), count)
        stamp = "" if not parse_bool(one("timestamp"), True) else datetime.now().strftime("_%Y%m%d_%H%M")
        self.filename = safe_filename(one("filename", "splunk_export")) + stamp + ".xlsx"
        self.options = xlsx_builder.ExportOptions(
            keep_text=parse_bool(one("keep_text"), False),
            keep_internal=parse_bool(one("keep_internal"), False),
            dates=parse_bool(one("dates"), True),
            max_rows=settings.max_rows_per_sheet)
        self.user = user


def collect_params(request):
    """Merge query string and form body parameters into {name: [values]}."""
    params = {}
    for source in ("query", "form"):
        for pair in request.get(source) or []:
            if isinstance(pair, (list, tuple)) and len(pair) == 2:
                params.setdefault(pair[0], []).append(pair[1])
    payload = request.get("payload")
    if payload and not request.get("form"):
        # JSON body: {"sid": ["..."], "sheet": "..."}
        try:
            body = json.loads(payload)
            if isinstance(body, dict):
                for key, value in body.items():
                    values = value if isinstance(value, list) else [value]
                    params.setdefault(key, []).extend("" if v is None else str(v) for v in values)
        except ValueError:
            pass
    return params


def export(req, runner):
    """Run the export described by an ExportRequest. Returns xlsx bytes."""
    sheets = []
    total = 0
    dispatched = []
    try:
        for i in range(len(req.sheet_names)):
            sid = req.sids[i] if i < len(req.sids) else ""
            postprocess = req.postprocess[i]
            if sid:
                try:
                    runner.wait(sid)
                except ExportError as err:
                    if err.status != 404 or not req.searches:
                        raise
                    LOG.info("job not found, running the search instead sid=%s user=%s", sid, req.user)
                    sid = ""
            if not sid:
                if not req.allow_search:
                    raise ExportError("The search job no longer exists and running a new search is "
                                      "disabled (allow_search = false). Reload the dashboard.", 404)
                sid = runner.dispatch(req.user, req.app, req.searches[i], req.earliest[i], req.latest[i])
                dispatched.append(sid)
                postprocess = ""          # the search already contains the full SPL
                runner.wait(sid)
            fields, rows = runner.results(sid, postprocess or None)
            total += len(rows)
            sheets.append(xlsx_builder.SheetData(req.sheet_names[i], fields, rows))
    finally:
        for sid in dispatched:
            runner.cancel(sid)
    if total == 0:
        raise ExportError("No results to export.", 404)
    return xlsx_builder.build_workbook(sheets, req.options), total


def json_response(status, message):
    return {
        "status": status,
        "headers": {"Content-Type": "application/json; charset=utf-8", "Cache-Control": "no-store"},
        "payload": json.dumps({"error": message, "status": status}),
    }


class ExcelExportHandler(PersistentServerConnectionApplication):
    def __init__(self, command_line=None, command_arg=None):
        PersistentServerConnectionApplication.__init__(self)

    def handle(self, in_string, client_factory=SplunkClient, settings=None):
        started = time.time()
        user = "-"
        if IMPORT_ERROR:
            LOG.error("cannot load xlsx_builder/openpyxl: %s", IMPORT_ERROR)
            return json_response(500, "The app could not load its Python libraries: %s"
                                 % IMPORT_ERROR.strip().splitlines()[-1])
        try:
            request = json.loads(in_string)
            method = (request.get("method") or "GET").upper()
            if method not in ("GET", "POST"):
                return json_response(405, "Method not allowed.")
            session = request.get("session") or {}
            user = session.get("user") or "-"
            token = session.get("authtoken")
            if not token:
                return json_response(401, "Not authenticated.")

            settings = settings or Settings.load()
            req = ExportRequest(collect_params(request), settings, user)
            runner = SearchRunner(client_factory(token), settings)
            data, total = export(req, runner)
            LOG.info("export ok user=%s app=%s sheets=%d rows=%d bytes=%d file=%s duration=%.2fs",
                     user, req.app, len(req.sheet_names), total, len(data), req.filename,
                     time.time() - started)
            return {
                "status": 200,
                "headers": {
                    "Content-Type": XLSX_MIME,
                    "Content-Disposition": content_disposition(req.filename),
                    "Cache-Control": "no-store",
                    "X-Content-Type-Options": "nosniff",
                },
                "payload_base64": base64.b64encode(data).decode("ascii"),
            }
        except ExportError as err:
            LOG.warning("export refused user=%s status=%d reason=%s", user, err.status, err)
            return json_response(err.status, str(err))
        except Exception as err:                    # pragma: no cover - unexpected
            LOG.exception("export failed user=%s", user)
            return json_response(500, "Export failed: %s" % err)

    def handleStream(self, handle, in_string):
        raise NotImplementedError("PersistentServerConnectionApplication.handleStream")

    def done(self):
        pass
