# Disposable-domain snapshot provenance

The verifier matches against `disposable_domains.txt` locally. No blocklist request occurs during app startup or verification.

Initial release snapshot:

- Retrieved: **2026-10-28**.
- Source project: [disposable-email-domains/disposable-email-domains](https://github.com/disposable-email-domains/disposable-email-domains).
- Source list: [disposable_email_blocklist.conf](https://raw.githubusercontent.com/disposable-email-domains/disposable-email-domains/main/disposable_email_blocklist.conf).
- Bundled domain count: **9,963**.
- Bundled file SHA-256: `c9c0b584f8d624a327c34a3d5b328e47948a55ee8a05a911eed29f423e09484c`.
- Source README preserved in full as `DISPOSABLE-SOURCE.md`, including its License section and CC0 link. The upstream repository did not provide a separate LICENSE file in the downloaded source used for this release.

The upstream License section permits copying, modification, distribution and use, including commercial use. Its README also notes that classifications are not guaranteed to remain current. The verifier therefore treats a matched provider as Risky for review, rather than claiming the mailbox is nonexistent.

`SNAPSHOT.json` records the current snapshot's date, source, count and hash. This document records the initial release snapshot even after updates.

## Maintenance

From the project root:

```bash
python scripts/update_disposable_domains.py --check-local
python scripts/update_disposable_domains.py
```

The second command intentionally accesses the fixed upstream HTTPS URL. The updater rejects malformed/undersized lists and responses above 2 MiB, normalizes and deduplicates domains, and atomically replaces each output file only after it has been written. DNS matching also covers subdomains of listed providers.

Review updates before committing them and restart Streamlit afterward to clear the in-process list cache. If the update fails before replacement, the existing bundled list remains available. This list is a screening aid, not an exhaustive registry of all disposable email providers.
