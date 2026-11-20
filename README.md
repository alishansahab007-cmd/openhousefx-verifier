# OpenhouseFX Lead Verifier 3.0

Upload a lead sheet (CSV / Excel). Every lead gets a **Good / Review / Bad** verdict, a **0–100 score** and a plain-English reason. Behind that verdict:

| Check | What it tells you |
| --- | --- |
| **Email syntax & clean-up** | Fixes messy cells (`mailto:`, `Name <a@b.com>`, several addresses in one cell) and rejects malformed addresses. |
| **Duplicates** | Repeat emails (case/space-insensitive) and repeat phone numbers (compared in international format). |
| **Domain DNS / MX** | Domain exists, publishes mail servers, and those mail servers actually resolve. Null-MX and dead domains are Invalid. |
| **Mailbox check (SMTP)** | Asks the recipient's mail server whether the exact mailbox exists, **without sending an email**. Detects catch-all (accept-everything) domains. |
| **Typo catcher** | `gmial.com`, `gmail.co`, `hotmial.com` … with the corrected address suggested. |
| **List hygiene** | Role inboxes (info@, sales@), no-reply addresses, disposable providers (~10k domains), placeholder addresses (test@, asdf@), free providers flagged. |
| **Phone validation** | Valid / invalid for the number's country, E.164 format, type (mobile / landline / VoIP / toll-free), carrier, location. Fixes numbers Excel stripped the `+` from. |
| **Live phone line status** *(optional)* | Active / inactive / unreachable via Twilio Lookup (paid per lookup). |
| **Client age** | From a date-of-birth column (or an existing age column). |
| **Lead age** | How long ago the lead was created / signed up. |
| **Activity status** | **Active / Lapsing / Inactive** from the last-activity date (thresholds adjustable, default 90 / 365 days). |
| **Company domain age** | When the lead's company domain was registered (RDAP), from the website column or business email domain. |
| **Location** | Country and state/province standardized (`USA`, `u.s.a.`, `US` → United States; `ca` → California). Country falls back to the phone number. |

Then **sort** by name, country, state, score, age, activity…; **filter** by verdict, country, state; **group** by any column with a breakdown chart; and **export** good / review / bad CSVs, the exact current view, a full Excel report, or an **organized workbook with one sorted sheet per country / state / anything**.

All original columns and values (leading zeros included) are kept in every export.

## How accurate is it?

No verifier can promise 100% — but this one uses the same evidence the commercial services do, and tells you exactly which evidence each verdict rests on:

- **Mailbox: Exists** – the mail server confirmed the mailbox. This is the strongest proof available without sending an email.
- **Mailbox: Does not exist** – the server rejected it (e.g. Gmail `550 5.1.1`). Graded Bad.
- **Catch-all** – the server accepts *every* address, so nobody can confirm the mailbox. Graded Review.
- **Unverifiable** – the server refused the check or could not be reached; the email keeps its DNS result. Turn on **Strict mode** if only mailbox-confirmed emails should be Good.
- Phone validation checks the official numbering plan for every country; whether the line is switched on *right now* needs the optional Twilio lookup.

### Where the mailbox check works

The SMTP mailbox check needs **outbound TCP port 25**. Many hosts block it:

| Host | Port 25 | Mailbox check |
| --- | --- | --- |
| Your own computer (most home/office networks) | Usually open for some providers | ✅ (partial if your ISP filters some) |
| VPS with port 25 unblocked (Hetzner, OVH, AWS EC2 after request …) | Open | ✅ full |
| Render (all plans), Heroku, Google Cloud, Streamlit Community Cloud | Blocked | ⚠️ skipped automatically — or set `ZEROBOUNCE_API_KEY` to check mailboxes through ZeroBounce's API |

The app detects this itself: if port 25 is blocked it says so on the results page and continues with every other check. Mail providers that this network cannot reach are skipped quickly (no long timeouts).

## Run locally

Python 3.12 is recommended.

macOS / Linux:

```bash
python3.12 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python -m streamlit run app.py
```

Windows PowerShell:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m streamlit run app.py
```

Open `http://localhost:8501`. Try `examples/sample_leads.csv` first.

## Put it on GitHub and make it live

1. Create a new GitHub repository (or open your existing `openhousefx-verifier` repo).
2. Upload **everything in this folder**, including the hidden files `.streamlit/`, `.python-version`, `.gitignore`, `.env.example`. (On github.com: *Add file → Upload files*, drag the folder contents in, commit.) The old `app.py.txt` is no longer needed — delete it from the repo.
3. Deploy — pick one:

**Render (easiest, existing setup)** — Web Service, runtime Python:

```text
Build command:  python -m pip install -r requirements.txt
Start command:  python -m streamlit run app.py --server.address 0.0.0.0 --server.port $PORT
Health check:   /_stcore/health
Environment:    PYTHON_VERSION=3.12.8   (plus optional keys below)
```

Render blocks port 25, so add `ZEROBOUNCE_API_KEY` if you want mailbox-level checks there. Everything else works on the free plan. `render.yaml` is included as a Blueprint.

**Streamlit Community Cloud (free)** — *New app* → pick the repo → main file `app.py` → Advanced settings → Python 3.12. Add optional keys under *Secrets* as environment-style `KEY = "value"` lines.

**Any VPS / Docker (full mailbox checks)** — on a server where port 25 is open:

```bash
docker build -t lead-verifier .
docker run -d -p 80:8501 --restart unless-stopped \
  -e VERIFIER_SMTP_HELO=verify.yourdomain.com -e VERIFIER_SMTP_FROM=verify@yourdomain.com \
  lead-verifier
```

Put the app behind your normal login/access control: it has no built-in user accounts.

## Settings (environment variables)

All optional — see `.env.example`.

| Variable | Purpose |
| --- | --- |
| `VERIFIER_SMTP_HELO`, `VERIFIER_SMTP_FROM` | Identity used in the mailbox check. A real domain you own (with matching reverse DNS) gets the best answers from mail servers. |
| `ZEROBOUNCE_API_KEY` | Mailbox check through ZeroBounce's HTTPS API when port 25 is blocked. `VERIFIER_MAILBOX_PROVIDER=zerobounce` forces it always. |
| `TWILIO_ACCOUNT_SID`, `TWILIO_AUTH_TOKEN` | Enables the live phone-line toggle (Twilio Lookup v2, `line_type_intelligence` + `line_status`). |
| `VERIFIER_DEFAULT_REGION` | Default phone country (e.g. `GB`, `PK`, `AE`). Changeable in the sidebar. |
| `VERIFIER_SMTP_DEFAULT`, `VERIFIER_DOMAIN_AGE_DEFAULT`, `VERIFIER_STRICT_DEFAULT` | Default on/off for the sidebar toggles. |
| `VERIFIER_DNS_NAMESERVERS` | Comma-separated DNS resolver IPs (default: the OS resolver). |
| `VERIFIER_DATA_DIR` | Enables saving completed runs to disk with a private recovery code. |
| `SUPABASE_URL`, `SUPABASE_KEY` | Optional logging of run totals only (see `supabase_schema.sql`). |

## Grading rules

- **Good** – email Valid (mailbox confirmed, or domain working when the mailbox can't be checked) and the phone isn't invalid. Phone-only lists: a valid, non-duplicate number.
- **Review** – risky email (catch-all, role inbox, typo, full mailbox, DNS doubt), or a good email next to an invalid phone, or a bad email but a working phone ("reachable by phone only").
- **Bad** – duplicate, invalid syntax, dead domain, mailbox doesn't exist, no-reply, disposable or placeholder address — with no working phone.
- **Score** – email quality (70%) + phone quality (30%); −5 for lapsing and −15 for inactive clients. Duplicates score 0.

## Output columns

Every export = your original columns + these:

```text
Lead_Verdict, Lead_Score, Lead_Issues, Lead_Name, Lead_Company, Lead_Country, Lead_State,
Verification_Status, Reason, Mailbox_Check, Normalized_Email, Email_Domain, DNS_Check, MX_Host,
Is_Role, Is_Disposable, Is_Free_Provider, Is_Duplicate, Typo_Suggestion,
Phone_Status, Phone_E164, Phone_Type, Phone_Country, Phone_Carrier, Phone_Location,
Phone_Line_Status, Is_Duplicate_Phone, Age, Record_Age_Days, Record_Age,
Last_Activity_Days, Activity_Status, Domain_Created, Domain_Age_Years
```

## Files and capacity

- CSV (UTF-8, Windows-1252/Latin-1; comma, semicolon, tab or pipe), XLSX (worksheet picker; formula *values* are used) and legacy XLS, up to 50 MB.
- Columns are auto-detected (email, phone, name, first/last name, company, website, country, state, DOB, age, created date, last activity). Every mapping can be changed in **More columns**.
- Work is done in batches on disk (SQLite), so 100k-row sheets are fine: the offline 100,000-row test completes in about a minute. Each email domain is looked up once per run. Live mailbox checks add time per mail domain; Gmail-heavy lists are fastest.
- Excel exports write formula-like text as literal text. CSV exports keep it as-is; open untrusted CSVs as text.

## Tests

```bash
python -m pip install -r requirements-dev.txt
python -m pytest -q                    # ~100 tests, no network needed
RUN_STRESS=1 python -m pytest -q       # adds the 100,000-row test
```

## Project layout

```text
app.py                     Streamlit UI
verifier/engine.py         Batch pipeline, grading, storage, sorting/grouping, exports
verifier/email_checks.py   Syntax, typo/role/disposable flags, DNS + MX checks
verifier/smtp_check.py     Mailbox (RCPT TO) probe, catch-all detection, ZeroBounce fallback
verifier/phone.py          Phone validation (libphonenumber) + optional Twilio line status
verifier/enrich.py         Names, country/state, dates, ages, activity, company domain age (RDAP)
verifier/io.py             Streaming CSV / XLSX / XLS readers
verifier/columns.py        Output schema and column auto-detection
data/                      Disposable-domain list and provenance (refresh: scripts/update_disposable_domains.py)
tests/                     Unit, end-to-end and UI tests
Dockerfile, render.yaml    Deployment
```

Fonts: Inter and Plus Jakarta Sans (OFL, licenses in `static/`). Disposable list: CC0, see `data/SOURCE.md`.
