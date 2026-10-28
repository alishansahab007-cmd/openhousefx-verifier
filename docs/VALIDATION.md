# Release validation - OpenhouseFX Lead Verifier 2.0.0

Validated locally on 2026-10-28 using Python 3.12 and Streamlit 1.66.1.

## Automated verification

Final complete suite: **43 passed in 43.93 seconds**, with the 100,000-row test enabled.

- 27 engine/integration cases cover CSV (UTF-8/BOM/Latin-1), XLSX and legacy XLS; column detection and manual mapping; malformed syntax and internationalized domains; duplicates across chunks; role/disposable detection and subdomains; null MX, NXDOMAIN, timeout, temporary failure and A/AAAA fallback; original cell preservation; filtered and full CSV/XLSX exports; literal Excel formula-like strings; pagination; incomplete-run rejection; saved-disk recovery; and mocked metadata-only Supabase logging.
- 9 DNS configuration cases cover system resolvers, empty-system fallback, explicit overrides, invalid configuration, and thread-local configuration changes.
- 7 Streamlit AppTest cases cover initial rendering, retained report state across filters/settings, lazy exports, Excel preparation, manual column mapping, and graceful rejection of missing/corrupt saved runs.

The 100,000-row case uses deterministic DNS results, 1,000-row batches and 25 domains. It asserts exactly 25 domain lookups, complete exports and correct final-page results. This is a regression/stress fixture, not a guarantee of live-network throughput or arbitrary host capacity.

Reproduce from the project root:

```powershell
python -m pip install -r requirements-dev.txt
$env:RUN_STRESS = '1'
python -m pytest tests -q
```

On macOS/Linux: `RUN_STRESS=1 python -m pytest tests -q`.

## Live browser and DNS checks

- Local Streamlit health endpoint returned `ok`.
- Gmail and Outlook returned explicit MX records; a deliberately nonexistent test domain returned NXDOMAIN.
- The supplied synthetic seven-row sample was uploaded through the browser. The completed report showed 2 Valid, 2 Risky and 3 Invalid rows.
- Downloaded the Valid-only CSV through its deferred download button. Confirmed two Valid rows and retained leading zeroes in the original Phone column.
- Prepared the full Excel report in the browser; the result table and CSV downloads remained available.
- Visually inspected the original Inter/Plus Jakarta Sans fonts, house mark, navy/gold theme, input mapping, result table and export controls. Screenshots are in this folder.

## Deployment and integration boundaries

Supabase logging was tested with a mocked HTTPS service. No actual Supabase credentials were supplied and no external database was changed. Local SQLite save/restore was tested; persistence on Render requires a configured persistent disk. A browser reload or server restart still ends an unsaved session.

Render deployment and GitHub push were not performed. The included README, render.yaml, .streamlit configuration and environment example describe installation and deployment. Upload the complete project, including its hidden configuration, font assets and disposable-domain data.

Verification checks syntax, list hygiene and domain mail records. It does not confirm individual mailbox existence or guarantee email delivery.
