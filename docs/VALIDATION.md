# Release validation – OpenhouseFX Lead Verifier 3.0.0

Validated on 2026-11-20 with Python 3.12 and Streamlit 1.66.1, from a clean install of `requirements-dev.txt`.

## Automated tests

`python -m pytest -q` → **101 passed**, plus the 100,000-row stress test with `RUN_STRESS=1` (about a minute).

- **Checks (≈78 cases):** column auto-detection; messy-cell email extraction; syntax; typo suggestions incl. false-positive guards; role / no-reply / placeholder / disposable / free flags; DNS (MX, MX hosts that don't resolve, null MX, NXDOMAIN, A fallback, timeout, resolver override); SMTP reply classification (Gmail, Microsoft 365, greylisting, Spamhaus blocks, full / disabled mailboxes); a local fake SMTP server for exists / missing / catch-all / reconnects / unreachable; per-provider breaker; ZeroBounce mapping; phone parsing (national, international, `00` prefix, Excel-stripped `+`, Excel number cells, placeholders); Twilio parsing; country/state standardization; dates, ages, activity.
- **Engine (17 cases):** 15-row end-to-end run with every verdict path; each domain/mailbox checked once; DNS retry pass; people / location / activity / domain age; original values and column order; category exports partition the list; sort / filter / search / group; SQL-injection-safe sorting; Excel and organized workbooks; literal formula text; phone-only lists; port-25-blocked fallback; strict mode; XLSX / XLS / cp1252-semicolon inputs; failing lookups never sink a run.
- **UI (5 AppTest cases):** initial render, results tabs, sorting + country + verdict filters, group by state, saved-run recovery errors.

## Live checks (this machine's home network)

- Gmail, ProtonMail and Zoho mail servers confirmed real mailboxes (`postmaster@`) and rejected made-up ones (`550 5.1.1`).
- Outlook / Microsoft 365, Yahoo/AOL and iCloud could not be reached on port 25 from this network: they were reported as *Unverifiable* with the reason, and the provider breaker cut the run from 297 s to 32 s.
- RDAP returned registration dates (e.g. github.com → 2007-10-09); an unregistered test domain returned *Not registered*.
- The sample file was uploaded through the browser: verdicts, chips, filters, sorting, the group-by-country chart and all export buttons rendered correctly.

## Not performed

- GitHub push and Render deployment (no credentials used). Render blocks port 25 on all plans, so mailbox checks there need `ZEROBOUNCE_API_KEY`.
- Live Twilio and ZeroBounce calls (no keys supplied); both are covered by mocked tests.
