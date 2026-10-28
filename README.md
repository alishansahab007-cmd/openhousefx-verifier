# OpenhouseFX Lead Verifier

A Streamlit workspace for reviewing CSV and Excel lead sheets, with the same navy-and-gold design, Inter typography, Plus Jakarta Sans headings and architectural house mark as OpenhouseFX Pocket Studio and Scraper Ultra.

Upload a sheet, confirm the email column, run verification, filter the results, and export the categories you need. Verification keeps the original rows and columns and appends clear outcomes.

## Run locally

Use Python 3.12. Run these commands from this folder, where `app.py` and `.streamlit/` are located.

Windows PowerShell:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m streamlit run app.py
```

macOS / Linux:

```bash
python3.12 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python -m streamlit run app.py
```

Open the local URL printed by Streamlit, normally `http://localhost:8501`. All core verification features run without Supabase or other API credentials. The pinned `streamlit==1.66.1` is required for deferred download generation; do not replace it with an older Streamlit release.

## Verification and interpretation

| Outcome | Meaning |
| --- | --- |
| Valid | Supported address syntax passed, explicit MX records exist, the row is unique, and no role/disposable flag was found. |
| Invalid | Missing or invalid syntax, a duplicate email row, an absent domain, explicit null MX, or no MX/A/AAAA route. |
| Risky | A role account, disposable provider, inconclusive DNS result, timeout, temporary server failure, or A/AAAA fallback without an explicit MX record. |

**Valid does not confirm mailbox existence or guarantee delivery.** This app does not perform SMTP recipient probes, detect catch-all mailboxes, or estimate bounce rates. Review Risky rows and retry temporary DNS failures later. Flags can overlap; a duplicate or invalid-domain row may also carry a role/disposable flag.

The pipeline:

1. Validate a practical RFC address subset with `email-validator`, without its DNS checks. Quoted local parts and internationalized domain names are supported. SMTPUTF8 local parts, domain literals, comments/display names and obsolete RFC forms are outside the supported input format.
2. Compare normalized emails after trimming surrounding whitespace and ignoring case. Plus tags and dots are retained. The first occurrence is classified normally; later occurrences stay in the full report as Invalid with `Is_Duplicate=True`.
3. Flag known generic business local parts, including `admin`, `info`, `support`, `sales` and `contact`, and match disposable domains, including subdomains, against a bundled snapshot.
4. Check each unique domain once per run, including across batch boundaries. A SQLite cache coordinates deduplicated work; each DNS worker also has a TTL-aware dnspython cache. No email local part or lead sheet is sent to DNS servers.
5. Use 20-50 configurable concurrent checks, with a shared two-second DNS lifetime budget per domain for MX followed by A/AAAA fallback. Null MX ends the check immediately. At most twice the worker count of futures are outstanding.

DNS normally uses the operating system resolver configuration. Set `VERIFIER_DNS_NAMESERVERS` to comma-separated resolver IP addresses to override it, for example `1.1.1.1,8.8.8.8`. Invalid entries are rejected before a run. If the operating system exposes no DNS servers, the app falls back to Cloudflare/Google public resolvers at those addresses. These DNS services receive queried domain names; email local parts and lead rows are not sent.

## Files, batches and original data

- Upload `.csv`, `.xlsx` or legacy `.xls`, up to 50 MB per file. Select an Excel worksheet when needed.
- Common headers such as `email`, `e-mail`, `mail`, `contact` and headers containing `email` are detected automatically. A dropdown always allows manual column mapping.
- CSV is read incrementally as text, preserving leading zeroes and literal text such as `NA`; UTF-8/BOM is preferred, with Latin-1 fallback. Comma, semicolon, tab and pipe delimiters are detected. Check unusual legacy encodings before upload.
- XLSX uses openpyxl read-only row iteration; values, date/time types and formula expressions are retained. XLS uses xlrd's on-demand sheet reader and is not fully streaming. Legacy XLS formulas are represented by their stored values.
- Verification uses batches of 1,000, 2,000 or 5,000 rows. Original values, results, duplicate keys and domain results are stored on disk in a per-session SQLite workspace; tables retrieve only the displayed page.
- Entirely empty rows are skipped. The first row is treated as the header. Short CSV rows are padded. Rows containing more data cells than available header columns are rejected with an error rather than truncated.
- Blank headers receive `Column_N`; duplicate headers get numeric suffixes. A source column matching a verification field is renamed with `_Original` (and a suffix if needed). The Run details panel lists these renames.
- Workbook formatting, macros, merged cells and formulas are not recreated as an editable source workbook. In the exported XLSX, formula-like strings are written as literal text. CSV exports retain original formula-like text; import untrusted CSVs as text in spreadsheet software.

## Dashboard and exports

KPIs show Total Processed, Valid Emails, Invalid Domains, Role Accounts and Disposable/Risky; invalid-domain count excludes syntax-only failures. The progress display includes stages, percentage, elapsed time and an approximate ETA. ETA becomes useful after rows have completed and is not a guarantee.

Results support status filtering, text search and 50/100/250-row pagination. These controls and downloads retain the completed result in Streamlit session state.

Each export contains all original columns followed by:

```text
Verification_Status, Reason, Is_Role, Is_Disposable, Is_Duplicate,
Normalized_Email, Email_Domain, DNS_Check
```

- **Full report - CSV:** every nonempty source row and every outcome.
- **Valid only - CSV:** unique rows passing syntax/MX checks without role/disposable flags. Its download filename remains `mx_verified_leads.csv`.
- **Invalid & risky - CSV:** the full removal/review subset.
- **Role & disposable - CSV:** every flagged row, including duplicate and invalid rows.
- **Full report - Excel:** generated on demand with openpyxl's write-only mode. Excel row/column/cell limits are enforced; use CSV if a report exceeds them.

Exports cover their complete category, regardless of the currently displayed filters or page. CSVs use UTF-8 with a BOM for Excel compatibility. Files are generated on disk. Downloads load only the chosen export when clicked, avoiding preloading all reports on every rerun.

## Capacity and sessions

The upload itself is held in memory by Streamlit, even though parsing and result processing use disk-backed batches. A requested download is also held in memory while served, and Excel shared strings or legacy XLS sheets can consume additional memory. There is no guarantee that an arbitrary 100,000-row sheet fits every host: file width, cell lengths, compressed XLSX expansion, simultaneous sessions, disk space and DNS latency all matter. Begin with 1,000-row batches and 20 workers on a small host; CSV is preferable for very large lists.

Ordinary widget reruns keep results. A browser reload/disconnection, app restart or redeploy can end the session and lose the temporary workspace. Completed runs can optionally be saved to a configured persistent directory, as described below. Incomplete jobs do not resume after interruption. Saved runs are not automatically expired; the server administrator manages retention and backups. This app has no built-in multi-user account/authentication layer, so put a privately deployed workspace behind your normal access controls.

## Optional saved-run recovery

Set `VERIFIER_DATA_DIR` to a directory the server can write, then enable **Save this run on the server** before starting a run. The app saves the completed SQLite database, reports and manifest. Copy the private UUID recovery code from Run details and use **Save & recovery** in the sidebar to restore the run after a reload.

The recovery code acts as access to the saved lead sheet. Keep it private. Optional saved runs are plain server files; protect and back up the directory. Setting the environment variable alone does not make a host's filesystem persistent.

PowerShell example:

```powershell
$env:VERIFIER_DATA_DIR = Join-Path (Get-Location) '.verifier-data'
.\.venv\Scripts\python.exe -m streamlit run app.py
```

POSIX example:

```bash
export VERIFIER_DATA_DIR="$PWD/.verifier-data"
.venv/bin/python -m streamlit run app.py
```

## Optional Supabase run totals

Run `supabase_schema.sql` in your Supabase project's SQL editor. Set these **server-side environment variables**, then enable **Log run totals to Supabase** in the sidebar:

```text
SUPABASE_URL=https://YOUR_PROJECT.supabase.co
SUPABASE_KEY=YOUR_SERVER_SIDE_SERVICE_ROLE_JWT_KEY
```

Use the project's legacy `service_role` JWT key for this REST client, never an anon/publishable key. Keep the key out of GitHub and browser code. The supplied schema enables RLS, revokes public/client-role table grants, and permits backend service-role writes without public policies, following the [Supabase RLS model](https://supabase.com/docs/guides/database/postgres/row-level-security).

Only batch UUID, timestamp, total validated and category counts are logged. Lead rows and filenames are not logged. Metadata logging cannot restore a lead sheet and does not replace saved-run recovery. If logging fails, the verification result and downloads still remain usable.

`.env.example` documents configuration; `.env` is not automatically loaded. Export environment variables in your shell or configure them in the hosting dashboard before starting the app.

## GitHub and Render

Back up your current repository, then copy the **contents of this complete project folder** into the repository root. Include hidden `.streamlit/`, `.python-version`, `assets/`, `static/` and `data/`; uploading only `app.py` is insufficient. The old `app.py.txt` duplicate is no longer needed. Install the updated dependencies. This deliverable does not push to GitHub or change a live deployment.

For an existing Render Web Service, use:

```text
Runtime: Python
Build command: python -m pip install -r requirements.txt
Start command: python -m streamlit run app.py --server.address 0.0.0.0 --server.port $PORT
Health check: /_stcore/health
```

The included `render.yaml` is an optional Blueprint for a new service; it starts on the Free plan for preview and does not provision a disk or database. Review the service plan for your workload. `.python-version` requests the latest Python 3.12 patch; remove an older conflicting `PYTHON_VERSION` setting or set it to a complete supported 3.12 patch version. See [Render Python versions](https://render.com/docs/python-version).

Render's normal filesystem is ephemeral. To preserve saved runs, attach a persistent disk on a paid web service, mount it at `/var/data`, and set `VERIFIER_DATA_DIR=/var/data/openhousefx-verifier`. Only data under the mount persists. Free instances cannot attach persistent disks. See [Render persistent disks](https://render.com/docs/disks) and [Free instance limitations](https://render.com/docs/free). Supabase totals alone do not preserve reports across restarts.

## Update the disposable-domain snapshot

The app does not download a list on startup. The bundled snapshot remains usable when the upstream source is unreachable. See `data/SOURCE.md` for provenance, hash, count and permission notice.

To deliberately refresh it from the fixed upstream HTTPS source:

```bash
python scripts/update_disposable_domains.py
```

The script bounds the response size, checks domain syntax and minimum list size, atomically replaces the list, and updates `data/SNAPSHOT.json`. Restart the app to invalidate its in-process list cache. Review the diff before committing a new snapshot; provider classifications can change.

## Tests

```bash
python -m pip install -r requirements-dev.txt
python -m pytest -q
```

Tests use controlled DNS responses; they do not require live DNS to classify fixtures. Live domain availability and cloud deployment depend on the host network and configuration.

## Project contents

```text
app.py                       Engine, disk-backed results, UI and optional integrations
requirements.txt             Runtime dependencies
requirements-dev.txt         Test dependencies
README.md                    Usage and deployment instructions
render.yaml                  Optional Render Blueprint
.python-version              Python 3.12 selection
.env.example                 Optional server environment settings
supabase_schema.sql          Optional metadata-only backend table
.streamlit/config.toml       Theme, font serving and upload size
assets/                      Architectural house SVG
static/                      Original local fonts and full OFL licenses
data/                        Disposable-domain list and provenance
scripts/                     Explicit snapshot updater
tests/                       Engine and state/integration regression checks
```

Bundled Inter and Plus Jakarta Sans font licenses are in `static/`. The logo geometry and exact design tokens originate from the user's Pocket Studio and Scraper Ultra projects. No external font requests are needed.


