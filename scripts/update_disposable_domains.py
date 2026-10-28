"""Explicit maintenance utility; never invoked by the Streamlit app.

Run with --check-local to validate the current snapshot without network or writes.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import tempfile
import time
from datetime import datetime, timezone

import requests

SOURCE_URL = (
    "https://raw.githubusercontent.com/disposable-email-domains/"
    "disposable-email-domains/main/disposable_email_blocklist.conf"
)
MAX_BYTES = 2 * 1024 * 1024
MAX_ENTRIES = 200_000
DOMAIN_PATTERN = re.compile(
    r"^(?=.{1,253}$)(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+"
    r"[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?$"
)
DATA_DIR = Path(__file__).resolve().parents[1] / "data"


def validate_snapshot(raw: bytes) -> list[str]:
    """Reject unexpected HTML, corruption and undersized replacement lists."""
    if len(raw) > MAX_BYTES:
        raise ValueError("The list exceeds the 2 MiB size bound.")
    text = raw.decode("utf-8-sig")
    domains = set()
    for number, line in enumerate(text.splitlines(), 1):
        value = line.strip().lower()
        if not value or value.startswith("#"):
            continue
        if not DOMAIN_PATTERN.fullmatch(value):
            raise ValueError(f"Invalid domain syntax on upstream line {number}.")
        domains.add(value)
        if len(domains) > MAX_ENTRIES:
            raise ValueError("The list exceeds the entry bound.")
    if len(domains) < 100:
        raise ValueError("Refusing a replacement list with fewer than 100 domains.")
    return sorted(domains)


def atomic_write(path: Path, content: bytes) -> None:
    """Replace a file only after its complete new content is written to disk."""
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(
            dir=path.parent, prefix=".snapshot-", delete=False
        ) as handle:
            temporary = Path(handle.name)
            handle.write(content)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary, path)
    finally:
        if temporary is not None and temporary.exists():
            temporary.unlink()


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--check-local",
        action="store_true",
        help="Validate the local list; no network or writes",
    )
    arguments = parser.parse_args()
    target = DATA_DIR / "disposable_domains.txt"
    if arguments.check_local:
        raw = target.read_bytes()
        domains = validate_snapshot(raw)
        print(
            f"Local snapshot valid: {len(domains):,} domains; SHA-256 {hashlib.sha256(raw).hexdigest()}"
        )
        return 0

    started = time.monotonic()
    chunks = bytearray()
    # One fixed HTTPS source; no redirects or arbitrary user-supplied URLs.
    with requests.get(
        SOURCE_URL, stream=True, timeout=(5, 15), allow_redirects=False
    ) as response:
        response.raise_for_status()
        if response.status_code != 200:
            raise ValueError("The upstream source did not return HTTP 200.")
        for chunk in response.iter_content(chunk_size=65536):
            if time.monotonic() - started > 25:
                raise TimeoutError(
                    "Snapshot download exceeded its maintenance deadline."
                )
            chunks.extend(chunk)
            if len(chunks) > MAX_BYTES:
                raise ValueError("The upstream response exceeds 2 MiB.")
    upstream = bytes(chunks)
    domains = validate_snapshot(upstream)
    normalized = ("\n".join(domains) + "\n").encode("utf-8")
    metadata = {
        "source_url": SOURCE_URL,
        "retrieved_at_utc": datetime.now(timezone.utc).isoformat(),
        "domain_count": len(domains),
        "sha256": hashlib.sha256(normalized).hexdigest(),
        "upstream_sha256": hashlib.sha256(upstream).hexdigest(),
        "license": "CC0 permission declared in upstream README; see DISPOSABLE-SOURCE.md",
    }
    atomic_write(target, normalized)
    atomic_write(
        DATA_DIR / "SNAPSHOT.json",
        (json.dumps(metadata, indent=2) + "\n").encode("utf-8"),
    )
    print(
        f"Updated {len(domains):,} domains. Restart Streamlit and review the diff before committing."
    )
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, ValueError, requests.RequestException) as error:
        raise SystemExit(f"Snapshot update failed: {error}") from error
