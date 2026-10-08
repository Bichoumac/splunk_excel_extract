#
# excel_export.conf.spec - Splunk Excel Extraction
#
# Settings of the /services/excel_export REST endpoint.
# Put overrides in $SPLUNK_HOME/etc/apps/splunk_excel_extract/local/excel_export.conf
# The file is read on every request: no restart is needed after a change.
#

[settings]

allow_search = <boolean>
* Whether the endpoint may run a new search passed in the "search" parameter.
* When false, only existing jobs ("sid" parameter) can be exported.
* Default: true

default_app = <string>
* App namespace used to run "search" requests when the caller does not pass
  the "app" parameter (lookups, macros and event types are resolved in it).
* Default: search

job_timeout = <integer>
* Maximum number of seconds to wait for a search job to finish.
* Default: 300

page_size = <integer>
* Number of rows read per request to splunkd. Keep it lower than or equal to
  maxresultrows in the [restapi] stanza of limits.conf.
* Default: 50000

max_rows_per_sheet = <integer>
* Maximum number of data rows written per sheet. Capped to 1048575, the Excel
  limit (1,048,576 rows including the header).
* Default: 1048575

max_sheets = <integer>
* Maximum number of sheets (sources) in one export.
* Default: 20

risky_commands = <comma-separated list>
* SPL commands refused in the "search" and "postprocess" parameters, so that a
  forged link cannot make a user run a search that writes or sends data.
* Set to an empty value to disable the check.
* Default: collect, delete, dump, map, mcollect, meventcollect, outputcsv,
  outputlookup, outputtext, run, runshellscript, script, sendalert, sendemail,
  tscollect
