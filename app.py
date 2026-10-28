"""Disk-backed, batch-oriented email syntax and domain verification engine."""

from __future__ import annotations

import codecs
import csv
import json
import math
import os
import sqlite3
import threading
import time
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Callable, Iterable

import dns.exception
import dns.resolver
import pandas as pd
from email_validator import EmailNotValidError, validate_email
from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell

RESULT_COLUMNS = [
    "Verification_Status",
    "Reason",
    "Is_Role",
    "Is_Disposable",
    "Is_Duplicate",
    "Normalized_Email",
    "Email_Domain",
    "DNS_Check",
]
ROLE_ACCOUNTS = frozenset(
    {
        "admin",
        "administrator",
        "billing",
        "bookings",
        "care",
        "careers",
        "contact",
        "enquiries",
        "enquiry",
        "hello",
        "help",
        "hr",
        "info",
        "jobs",
        "legal",
        "marketing",
        "noreply",
        "no-reply",
        "office",
        "operations",
        "postmaster",
        "privacy",
        "reception",
        "sales",
        "security",
        "service",
        "support",
        "team",
        "webmaster",
    }
)
_DNS_LOCAL = threading.local()


@dataclass(frozen=True)
class DomainResult:
    """A domain result, never a claim that an individual mailbox receives mail."""

    status: str
    reason: str
    dns_check: str


def _build_resolver():
    """Honor explicit/system DNS; use public DNS only if the OS has none."""
    try:
        resolver = dns.resolver.Resolver()
    except dns.resolver.NoResolverConfiguration:
        resolver = dns.resolver.Resolver(configure=False)
    configured = os.getenv("VERIFIER_DNS_NAMESERVERS", "").strip()
    if configured:
        servers = [server.strip() for server in configured.split(",") if server.strip()]
        if not servers:
            raise ValueError(
                "VERIFIER_DNS_NAMESERVERS must contain comma-separated DNS server IP addresses."
            )
        try:
            # Assignment asks dnspython to validate each supplied server.
            resolver.nameservers = servers
        except (ValueError, TypeError) as error:
            raise ValueError(
                "Invalid VERIFIER_DNS_NAMESERVERS. Use comma-separated DNS server IP addresses."
            ) from error
    elif not resolver.nameservers:
        # Some Windows/container environments expose no OS DNS servers. Use
        # Cloudflare/Google only in that case; never replace working system DNS.
        resolver.nameservers = ["1.1.1.1", "8.8.8.8"]
    resolver.timeout = 0.65
    resolver.lifetime = 2.0
    resolver.cache = dns.resolver.LRUCache(max_size=2048)
    return resolver


def validate_dns_configuration():
    """Fail before starting a run when an explicit DNS override is malformed."""
    return tuple(_build_resolver().nameservers)


def _thread_resolver():
    """Each worker owns its resolver; the dnspython LRU handles record TTLs."""
    signature = os.getenv("VERIFIER_DNS_NAMESERVERS", "").strip()
    if (
        not hasattr(_DNS_LOCAL, "resolver")
        or getattr(_DNS_LOCAL, "signature", None) != signature
    ):
        _DNS_LOCAL.resolver = _build_resolver()
        _DNS_LOCAL.signature = signature
    return _DNS_LOCAL.resolver


def lookup_domain(
    domain: str, timeout: float = 2.0, resolver=None, clock: Callable = time.monotonic
) -> DomainResult:
    """Check MX then implicit A/AAAA, sharing a <=2s DNS lifetime budget.

    Timeout/SERVFAIL is inconclusive (Risky). Null MX explicitly declines mail.
    A/AAAA fallback is Risky because no explicit mail service was found.
    """
    timeout = min(2.0, max(0.05, float(timeout)))
    deadline = clock() + timeout
    resolver = resolver if resolver is not None else _thread_resolver()
    for record_type in ("MX", "A", "AAAA"):
        remaining = deadline - clock()
        if remaining <= 0:
            return DomainResult(
                "Risky", "DNS deadline exceeded; retry later", "Timeout"
            )
        try:
            records = list(
                resolver.resolve(domain, record_type, lifetime=remaining, search=False)
            )
            if not records:
                continue
            if record_type == "MX":
                if any(str(record.exchange).strip() == "." for record in records):
                    return DomainResult(
                        "Invalid",
                        "Domain explicitly declines email (null MX)",
                        "Null MX",
                    )
                return DomainResult(
                    "Valid", "MX records present; mailbox not tested", "MX"
                )
            return DomainResult(
                "Risky",
                f"No MX; {record_type} fallback exists, mail service unconfirmed",
                record_type,
            )
        except dns.resolver.NXDOMAIN:
            return DomainResult("Invalid", "Email domain does not exist", "NXDOMAIN")
        except dns.resolver.NoAnswer:
            continue
        except (dns.exception.Timeout, dns.resolver.LifetimeTimeout):
            return DomainResult("Risky", "DNS timed out; retry later", "Timeout")
        except dns.resolver.NoNameservers:
            return DomainResult(
                "Risky", "DNS servers unavailable or SERVFAIL; retry later", "DNS error"
            )
        except dns.exception.DNSException:
            return DomainResult(
                "Risky", "DNS response inconclusive; retry later", "DNS error"
            )
        except OSError:
            return DomainResult(
                "Risky", "Network unavailable; retry later", "Network error"
            )
    return DomainResult("Invalid", "No MX, A or AAAA mail route found", "No route")


def _safe_columns(headers: Iterable) -> tuple[list[str], list[dict]]:
    """Keep names where possible; make blank/duplicate/reserved names explicit."""
    used, result, renames = set(RESULT_COLUMNS), [], []
    for index, raw in enumerate(headers):
        original = "" if raw is None else str(raw)
        base = original if original.strip() else f"Column_{index + 1}"
        candidate = f"{base}_Original" if base in RESULT_COLUMNS else base
        stem, suffix = candidate, 2
        while candidate in used:
            candidate = f"{stem}_{suffix}"
            suffix += 1
        used.add(candidate)
        result.append(candidate)
        if candidate != original:
            renames.append(
                {"position": index + 1, "original": original, "renamed": candidate}
            )
    if not result:
        raise ValueError("The file has no column headers.")
    return result, renames


def detect_email_column(columns: Iterable[str]) -> int | None:
    names = [
        "".join(ch for ch in str(column).strip().casefold() if ch.isalnum())
        for column in columns
    ]
    for aliases in ({"email", "emailaddress", "emailid"}, {"mail", "contactemail"}):
        for index, name in enumerate(names):
            if name in aliases:
                return index
    for index, name in enumerate(names):
        if "email" in name:
            return index
    for index, name in enumerate(names):
        if name in {"contact", "mailaddress"}:
            return index
    return None


def _csv_format(path: Path) -> tuple[str, str]:
    """Validate UTF-8 incrementally, then fall back to lossless Latin-1."""
    decoder = codecs.getincrementaldecoder("utf-8-sig")("strict")
    encoding = "utf-8-sig"
    try:
        with path.open("rb") as handle:
            for block in iter(lambda: handle.read(1024 * 1024), b""):
                decoder.decode(block)
            decoder.decode(b"", final=True)
    except UnicodeDecodeError:
        encoding = "latin-1"
    with path.open("r", encoding=encoding, newline="") as handle:
        sample = handle.read(65536)
    try:
        delimiter = csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
    except csv.Error:
        delimiter = ","
    return encoding, delimiter


def _is_blank(row: Iterable) -> bool:
    return all(
        value is None or (isinstance(value, str) and not value.strip()) for value in row
    )


@contextmanager
def _xlsx_reader(path):
    """Use a file handle so temporary filenames need not end in .xlsx."""
    with Path(path).open("rb") as handle:
        workbook = load_workbook(handle, read_only=True, data_only=False)
        try:
            yield workbook
        finally:
            workbook.close()


def _raw_rows(path: Path, info: dict):
    """Skip only all-empty rows; keep all other cells, padding short CSV rows."""
    extension = info["extension"]
    if extension == ".csv":
        with path.open("r", encoding=info["encoding"], newline="") as handle:
            reader = csv.reader(handle, delimiter=info["delimiter"])
            next(reader, None)
            for source_row, row in enumerate(reader, 2):
                if _is_blank(row):
                    continue
                if len(row) > len(info["columns"]):
                    raise ValueError(
                        f"CSV row {source_row} has more cells than the header. Check the delimiter."
                    )
                yield row + [""] * (len(info["columns"]) - len(row))
    elif extension == ".xlsx":
        with _xlsx_reader(path) as workbook:
            sheet = workbook[info["sheet_name"]]
            sheet.reset_dimensions()
            rows = sheet.iter_rows(values_only=True)
            next(rows, None)
            for row in rows:
                row = list(row)
                if _is_blank(row):
                    continue
                if len(row) > len(info["columns"]):
                    if any(value is not None for value in row[len(info["columns"]) :]):
                        raise ValueError(
                            "An Excel data row extends beyond the header. Add names to all data columns."
                        )
                    row = row[: len(info["columns"])]
                yield row + [None] * (len(info["columns"]) - len(row))
    else:
        import xlrd

        workbook = xlrd.open_workbook(str(path), on_demand=True)
        try:
            sheet = workbook.sheet_by_name(info["sheet_name"])
            for row_index in range(1, sheet.nrows):
                row = []
                for column in range(sheet.ncols):
                    cell = sheet.cell(row_index, column)
                    value = cell.value
                    if cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                        value = None
                    elif cell.ctype == xlrd.XL_CELL_DATE:
                        value = xlrd.xldate.xldate_as_datetime(value, workbook.datemode)
                    elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                        value = bool(value)
                    row.append(value)
                if not _is_blank(row):
                    yield row
        finally:
            workbook.release_resources()


def inspect_input(
    path, filename: str | None = None, sheet_name: str | None = None
) -> dict:
    """Inspect headers/sheets and count rows without loading a giant frame."""
    path = Path(path)
    extension = Path(filename or path.name).suffix.lower()
    if extension not in {".csv", ".xlsx", ".xls"}:
        raise ValueError("Choose a CSV, XLSX or XLS file.")
    info = {
        "extension": extension,
        "sheets": [],
        "sheet_name": None,
        "encoding": None,
        "delimiter": None,
    }
    if extension == ".csv":
        info["encoding"], info["delimiter"] = _csv_format(path)
        with path.open("r", encoding=info["encoding"], newline="") as handle:
            headers = next(csv.reader(handle, delimiter=info["delimiter"]), [])
    elif extension == ".xlsx":
        with _xlsx_reader(path) as workbook:
            info["sheets"] = workbook.sheetnames
            selected = sheet_name or workbook.sheetnames[0]
            if selected not in workbook.sheetnames:
                raise ValueError("The selected sheet no longer exists.")
            info["sheet_name"] = selected
            sheet = workbook[selected]
            sheet.reset_dimensions()
            headers = next(sheet.iter_rows(values_only=True), ())
    else:
        import xlrd

        workbook = xlrd.open_workbook(str(path), on_demand=True)
        try:
            info["sheets"] = workbook.sheet_names()
            info["sheet_name"] = sheet_name or info["sheets"][0]
            sheet = workbook.sheet_by_name(info["sheet_name"])
            headers = sheet.row_values(0) if sheet.nrows else []
        finally:
            workbook.release_resources()
    info["columns"], info["column_renames"] = _safe_columns(headers)
    info["email_index"] = detect_email_column(info["columns"])
    info["row_count"] = sum(1 for _ in _raw_rows(path, info))
    return info


def _json_value(value):
    if isinstance(value, datetime):
        return {"__ohfx_type": "datetime", "value": value.isoformat()}
    if isinstance(value, date):
        return {"__ohfx_type": "date", "value": value.isoformat()}
    if isinstance(value, float) and not math.isfinite(value):
        return None
    if value is None or isinstance(value, (str, bool, int, float)):
        return value
    return str(value)


def _decode_value(value):
    if isinstance(value, dict) and value.get("__ohfx_type") == "datetime":
        return datetime.fromisoformat(value["value"])
    if isinstance(value, dict) and value.get("__ohfx_type") == "date":
        return date.fromisoformat(value["value"])
    return value


def _disposable_match(domain: str, domains: frozenset) -> bool:
    parts = domain.casefold().split(".")
    return any(".".join(parts[index:]) in domains for index in range(len(parts) - 1))


def _syntax(value) -> tuple[str, str, str | None]:
    """Validate a supported RFC email address; retain original cell separately."""
    if not isinstance(value, str) or not value.strip():
        return "", "", "Missing or non-text email address"
    try:
        parsed = validate_email(
            value.strip(),
            check_deliverability=False,
            allow_smtputf8=False,
            strict=True,
            allow_quoted_local=True,
            allow_domain_literal=False,
        )
        return parsed.normalized, parsed.ascii_domain.casefold(), None
    except EmailNotValidError as error:
        return "", "", f"Invalid syntax: {error}"


@contextmanager
def _open_db(path):
    """Always release Windows file handles, including on failed runs."""
    connection = sqlite3.connect(str(path))
    try:
        connection.execute("PRAGMA journal_mode=DELETE")
        with connection:
            yield connection
    finally:
        connection.close()


def _metadata(connection) -> dict:
    row = connection.execute(
        "SELECT value FROM metadata WHERE key='summary'"
    ).fetchone()
    if row is None:
        raise ValueError("This verification did not finish. Start a new run.")
    return json.loads(row[0])


def _emit_progress(
    callback, start, stage, processed, total, checked, unique, counts=None
):
    if callback is None:
        return
    elapsed = time.monotonic() - start
    eta = (elapsed / processed) * (total - processed) if processed else None
    callback(
        {
            "stage": stage,
            "rows_processed": processed,
            "total_rows": total,
            "domains_checked": checked,
            "unique_domains": unique,
            "elapsed": elapsed,
            "eta": eta,
            "counts": dict(counts or {}),
        }
    )


def run_verification(
    input_path,
    output_dir,
    email_index: int,
    workers: int = 35,
    batch_size: int = 2000,
    sheet_name: str | None = None,
    disposable_domains: frozenset = frozenset(),
    progress: Callable | None = None,
    domain_lookup: Callable | None = None,
    filename: str | None = None,
) -> dict:
    """Verify bounded batches with disk-backed deduplication and domain caching.

    At most 2*workers futures are outstanding. Per-run domain results and seen
    email keys live in SQLite. Duplicate rows remain in full exports as Invalid.
    """
    if not 1 <= int(workers) <= 50:
        raise ValueError("Workers must be between 1 and 50.")
    if not 1 <= int(batch_size) <= 5000:
        raise ValueError("Batch size must be between 1 and 5,000.")
    start = time.monotonic()
    info = inspect_input(input_path, filename=filename, sheet_name=sheet_name)
    if not 0 <= int(email_index) < len(info["columns"]):
        raise ValueError("Choose a valid email column.")
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    db_path = output_dir / "results.sqlite3"
    if db_path.exists():
        raise ValueError(
            "This result folder already contains a run. Choose a new folder."
        )
    batch_id = str(uuid.uuid4())
    lookup = domain_lookup or lookup_domain
    if domain_lookup is None:
        validate_dns_configuration()
    counts = dict.fromkeys(
        (
            "total",
            "valid",
            "invalid",
            "risky",
            "invalid_domains",
            "role",
            "disposable",
            "duplicates",
        ),
        0,
    )
    summary = {
        "batch_id": batch_id,
        "db_path": str(db_path),
        "filename": filename or Path(input_path).name,
        "created_at": datetime.now(timezone.utc).isoformat(),
        "sheet_name": info["sheet_name"],
        "original_columns": info["columns"],
        "column_renames": info["column_renames"],
        "email_column": info["columns"][email_index],
        "row_count": info["row_count"],
        "counts": counts,
        "exports": {},
        "workers": int(workers),
        "batch_size": int(batch_size),
    }
    checked = 0
    with _open_db(db_path) as connection:
        connection.executescript("""
            CREATE TABLE metadata (key TEXT PRIMARY KEY, value TEXT NOT NULL);
            CREATE TABLE seen_emails (email TEXT PRIMARY KEY);
            CREATE TABLE domains (domain TEXT PRIMARY KEY, status TEXT, reason TEXT, dns_check TEXT);
            CREATE TABLE results (
                row_index INTEGER PRIMARY KEY, payload TEXT NOT NULL, status TEXT NOT NULL,
                reason TEXT NOT NULL, is_role INTEGER NOT NULL, is_disposable INTEGER NOT NULL,
                is_duplicate INTEGER NOT NULL, normalized_email TEXT, domain TEXT, dns_check TEXT,
                search_text TEXT NOT NULL
            );
            CREATE INDEX results_status ON results(status);
        """)
        _emit_progress(progress, start, "Preparing", 0, info["row_count"], 0, 0, counts)
        with ThreadPoolExecutor(
            max_workers=int(workers), thread_name_prefix="ohfx-dns"
        ) as pool:
            source = iter(_raw_rows(Path(input_path), info))
            while True:
                batch = []
                for _ in range(int(batch_size)):
                    try:
                        batch.append(next(source))
                    except StopIteration:
                        break
                if not batch:
                    break
                prepared, uncached = [], set()
                for row in batch:
                    email, domain, syntax_error = _syntax(row[email_index])
                    duplicate = False
                    raw_email = row[email_index]
                    duplicate_key = (
                        email.casefold()
                        if email
                        else (
                            raw_email.strip().casefold()
                            if isinstance(raw_email, str) and raw_email.strip()
                            else None
                        )
                    )
                    if duplicate_key:
                        cursor = connection.execute(
                            "INSERT OR IGNORE INTO seen_emails(email) VALUES (?)",
                            (duplicate_key,),
                        )
                        duplicate = cursor.rowcount == 0
                    local = (
                        email.rsplit("@", 1)[0].strip('"').casefold().split("+", 1)[0]
                        if email
                        else ""
                    )
                    is_role = local in ROLE_ACCOUNTS
                    is_disposable = bool(
                        domain and _disposable_match(domain, disposable_domains)
                    )
                    if domain and not syntax_error:
                        if (
                            connection.execute(
                                "SELECT 1 FROM domains WHERE domain=?", (domain,)
                            ).fetchone()
                            is None
                        ):
                            uncached.add(domain)
                    prepared.append(
                        (
                            row,
                            email,
                            domain,
                            syntax_error,
                            duplicate,
                            is_role,
                            is_disposable,
                        )
                    )
                # Submission is bounded even if a batch contains many unique domains.
                domain_iter, pending = iter(sorted(uncached)), {}
                target_domains = checked + len(uncached)

                def submit_more():
                    while len(pending) < 2 * int(workers):
                        try:
                            domain = next(domain_iter)
                        except StopIteration:
                            break
                        pending[pool.submit(lookup, domain)] = domain

                submit_more()
                while pending:
                    future = next(as_completed(tuple(pending)))
                    domain = pending.pop(future)
                    try:
                        result = future.result()
                        if result.status not in {"Valid", "Invalid", "Risky"}:
                            raise ValueError("Unknown domain status")
                    except Exception:
                        result = DomainResult(
                            "Risky", "Domain lookup failed; retry later", "DNS error"
                        )
                    connection.execute(
                        "INSERT INTO domains VALUES (?,?,?,?)",
                        (domain, result.status, result.reason, result.dns_check),
                    )
                    checked += 1
                    if checked % 10 == 0 or not pending:
                        _emit_progress(
                            progress,
                            start,
                            "Checking domains",
                            counts["total"],
                            info["row_count"],
                            checked,
                            target_domains,
                            counts,
                        )
                    submit_more()
                inserts = []
                for (
                    row,
                    email,
                    domain,
                    syntax_error,
                    duplicate,
                    is_role,
                    is_disposable,
                ) in prepared:
                    dns_status, dns_reason, dns_check = "", "", "Not checked"
                    if domain:
                        dns_status, dns_reason, dns_check = connection.execute(
                            "SELECT status,reason,dns_check FROM domains WHERE domain=?",
                            (domain,),
                        ).fetchone()
                    if syntax_error:
                        status, reason = "Invalid", syntax_error
                        if duplicate:
                            reason += "; duplicate email entry"
                    elif duplicate:
                        status, reason = (
                            "Invalid",
                            "Duplicate email; see first occurrence",
                        )
                        if dns_status == "Invalid":
                            reason += "; " + dns_reason
                    elif dns_status == "Invalid":
                        status, reason = "Invalid", dns_reason
                    elif is_disposable or is_role or dns_status == "Risky":
                        flags = []
                        if is_role:
                            flags.append("Role-based business account")
                        if is_disposable:
                            flags.append("Disposable email provider")
                        if dns_status == "Risky":
                            flags.append(dns_reason)
                        status, reason = "Risky", "; ".join(flags)
                    else:
                        status, reason = "Valid", dns_reason
                    counts["total"] += 1
                    counts[status.lower()] += 1
                    counts["invalid_domains"] += int(dns_status == "Invalid")
                    counts["role"] += int(is_role)
                    counts["disposable"] += int(is_disposable)
                    counts["duplicates"] += int(duplicate)
                    payload = json.dumps(
                        [_json_value(value) for value in row], ensure_ascii=False
                    )
                    search_text = " ".join(
                        [
                            *(str(value) for value in row if value is not None),
                            status,
                            reason,
                            email,
                            domain,
                            dns_check,
                        ]
                    ).casefold()
                    inserts.append(
                        (
                            counts["total"],
                            payload,
                            status,
                            reason,
                            int(is_role),
                            int(is_disposable),
                            int(duplicate),
                            email,
                            domain,
                            dns_check,
                            search_text,
                        )
                    )
                connection.executemany(
                    "INSERT INTO results VALUES (?,?,?,?,?,?,?,?,?,?,?)", inserts
                )
                connection.commit()
                _emit_progress(
                    progress,
                    start,
                    "Verifying rows",
                    counts["total"],
                    info["row_count"],
                    checked,
                    checked,
                    counts,
                )
        summary["unique_domains"] = checked
        summary["duration_seconds"] = round(time.monotonic() - start, 3)
        summary["exports"] = {
            "full_csv": str(output_dir / "openhousefx_all_results.csv"),
            "valid_csv": str(output_dir / "openhousefx_valid_leads.csv"),
            "invalid_risky_csv": str(output_dir / "openhousefx_invalid_risky.csv"),
            "role_disposable_csv": str(output_dir / "openhousefx_role_disposable.csv"),
        }
    _emit_progress(
        progress,
        start,
        "Preparing exports",
        counts["total"],
        info["row_count"],
        checked,
        checked,
        counts,
    )
    _generate_csv_exports(db_path, summary)
    # Only completed reports are recoverable; interrupted/failed exports have no summary.
    with _open_db(db_path) as connection:
        connection.execute(
            "INSERT INTO metadata VALUES ('summary',?)", (json.dumps(summary),)
        )
    _emit_progress(
        progress,
        start,
        "Complete",
        counts["total"],
        info["row_count"],
        checked,
        checked,
        counts,
    )
    return summary


def _result_values(record):
    values = [_decode_value(value) for value in json.loads(record[1])]
    return values + [
        record[2],
        record[3],
        bool(record[4]),
        bool(record[5]),
        bool(record[6]),
        record[7],
        record[8],
        record[9],
    ]


def _generate_csv_exports(db_path, summary):
    """Write all category CSVs in a single streaming pass."""
    handles = {}
    try:
        writers = {}
        for key, path in summary["exports"].items():
            handles[key] = open(path, "w", encoding="utf-8-sig", newline="")
            writers[key] = csv.writer(handles[key])
            writers[key].writerow(summary["original_columns"] + RESULT_COLUMNS)
        with _open_db(db_path) as connection:
            for record in connection.execute(
                "SELECT * FROM results ORDER BY row_index"
            ):
                values = _result_values(record)
                writers["full_csv"].writerow(values)
                writers[
                    "valid_csv" if record[2] == "Valid" else "invalid_risky_csv"
                ].writerow(values)
                if record[4] or record[5]:
                    writers["role_disposable_csv"].writerow(values)
    finally:
        for handle in handles.values():
            handle.close()


def read_results(
    db_path, status: str = "All", search: str = "", limit: int = 100, offset: int = 0
) -> tuple[pd.DataFrame, int]:
    """Fetch a bounded page using parameterized status/search filters."""
    limit, offset = max(1, min(1000, int(limit))), max(0, int(offset))
    where, params = [], []
    if status in {"Valid", "Invalid", "Risky"}:
        where.append("status=?")
        params.append(status)
    elif status == "Invalid & Risky":
        where.append("status IN ('Invalid','Risky')")
    elif status == "Role / Disposable":
        where.append("(is_role=1 OR is_disposable=1)")
    if search.strip():
        where.append("instr(search_text,?)>0")
        params.append(search.strip().casefold())
    clause = " WHERE " + " AND ".join(where) if where else ""
    with _open_db(db_path) as connection:
        summary = _metadata(connection)
        count = connection.execute(
            "SELECT COUNT(*) FROM results" + clause, params
        ).fetchone()[0]
        records = connection.execute(
            "SELECT * FROM results" + clause + " ORDER BY row_index LIMIT ? OFFSET ?",
            [*params, limit, offset],
        ).fetchall()
    frame = pd.DataFrame(
        [_result_values(record) for record in records],
        columns=summary["original_columns"] + RESULT_COLUMNS,
    )
    return frame, count


def generate_xlsx(db_path, path) -> str:
    """Stream an XLSX report; source formula-like text remains literal text.

    Reject oversize cells rather than allowing Excel's 32,767-character cell
    limit to silently truncate them. Full CSV remains available for such input.
    """
    path = Path(path)
    with _open_db(db_path) as connection:
        summary = _metadata(connection)
        if summary["counts"]["total"] > 1_048_575:
            raise ValueError(
                "This report exceeds Excel's row limit. Use the full CSV export."
            )
        headers = summary["original_columns"] + RESULT_COLUMNS
        if len(headers) > 16_384:
            raise ValueError(
                "This report exceeds Excel's column limit. Use the CSV export."
            )
        workbook = Workbook(write_only=True)
        sheet = workbook.create_sheet("Verification results")
        sheet.freeze_panes = "A2"

        def append_literal(values):
            cells = []
            for value in values:
                if isinstance(value, str):
                    if len(value) > 32767:
                        raise ValueError(
                            "A cell exceeds Excel's 32,767-character limit. Use CSV to preserve it."
                        )
                    cell = WriteOnlyCell(sheet, value=value)
                    cell.data_type = "s"
                    cells.append(cell)
                else:
                    cells.append(value)
            sheet.append(cells)

        try:
            append_literal(headers)
            for record in connection.execute(
                "SELECT * FROM results ORDER BY row_index"
            ):
                append_literal(_result_values(record))
            workbook.save(path)
        except Exception:
            workbook.close()
            if path.exists():
                path.unlink()
            raise
    return str(path)


def load_summary(db_path) -> dict:
    """Read a completed disk-backed run after browser or process restart."""
    with _open_db(db_path) as connection:
        return _metadata(connection)


# ---------------------------------------------------------------------------
# OpenhouseFX UI. The engine above is independent of Streamlit and testable.
# ---------------------------------------------------------------------------
import base64
import hashlib
import html
import os
import shutil
import tempfile
from functools import lru_cache
from urllib.parse import urlparse

import requests
import streamlit as st

APP_ROOT = Path(__file__).resolve().parent
APP_VERSION = "2.0.0"
_xlsx_lock = threading.Lock()


@lru_cache(maxsize=1)
def brand_css():
    """Embed the original Pocket Studio fonts, with no third-party font calls."""
    faces = []
    for family, filename in [
        ("Inter", "inter-latin.woff2"),
        ("Plus Jakarta Sans", "jakarta-latin.woff2"),
    ]:
        encoded = base64.b64encode(
            (APP_ROOT / "static" / filename).read_bytes()
        ).decode("ascii")
        faces.append(
            f"@font-face{{font-family:'{family}';src:url(data:font/woff2;base64,{encoded}) format('woff2');font-weight:100 900;font-style:normal;font-display:swap;}}"
        )
    return "<style>" + "".join(faces) + """
    :root {--navy:#061521;--card:#0B1F30;--panel:#0F2A40;--gold:#D3AA4E;--text:#F0F4F8;--muted:#94A3B8;--line:rgba(211,170,78,.19)}
    html,body,[data-testid="stApp"] {font-family:Inter,sans-serif;color:var(--text);background:var(--navy)}
    h1,h2,h3,h4 {font-family:'Plus Jakarta Sans',Inter,sans-serif!important;letter-spacing:-.04em}
    [data-testid="stHeader"] {background:rgba(6,21,33,.94);border-bottom:1px solid rgba(211,170,78,.08)}
    [data-testid="stSidebar"] {background:var(--card);border-right:1px solid var(--line);min-width:290px;max-width:290px}
    [data-testid="stSidebarUserContent"] {padding:1.5rem 1.5rem 2rem}
    [data-testid="stMainBlockContainer"] {max-width:1500px;padding:2rem 3.2rem}
    [data-testid="stVerticalBlock"] {gap:1rem}
    [data-testid="stSidebar"] [data-testid="stVerticalBlock"] {gap:.85rem}
    .brand {display:flex;align-items:center;gap:12px;padding:8px 0 28px;border-bottom:1px solid var(--line);margin-bottom:18px}
    .brand svg {width:29px;height:36px;flex:none}
    .wordmark {font:600 16px/1.3 Inter,sans-serif;letter-spacing:1.7px;white-space:nowrap}
    .wordmark span {color:var(--gold)}
    .product {font:500 7.7px/1.5 Inter,sans-serif;letter-spacing:3px;color:var(--gold);margin-top:5px}
    .eyebrow {color:var(--gold);font:600 9px/1.5 Inter,sans-serif;letter-spacing:2.2px;margin:8px 0 14px;text-transform:uppercase}
    .sidebar-title {font:600 18px/1.4 'Plus Jakarta Sans',sans-serif;letter-spacing:-.6px;margin-bottom:8px}
    .small-copy {font:400 11px/1.8 Inter,sans-serif;color:var(--muted)}
    .crumbs {display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid rgba(211,170,78,.10);padding:0 0 22px;margin-bottom:32px;gap:12px}
    .crumbs-label {font-size:11px;color:var(--muted)}
    .crumbs-label b {color:#d9e2ec;font-weight:500;margin-left:16px}
    .local-pill {font-size:10px;color:var(--muted);border:1px solid rgba(148,163,184,.14);border-radius:30px;padding:7px 12px;white-space:nowrap}
    .local-pill i {display:inline-block;width:6px;height:6px;background:#14B8A6;border-radius:50%;margin-right:7px;box-shadow:0 0 9px #14b8a64d}
    .hero-title {font:500 clamp(28px,3vw,43px)/1.22 'Plus Jakarta Sans',sans-serif;letter-spacing:-.035em;word-spacing:.08em;margin:0 0 14px}
    .hero-title span {color:var(--gold)}
    .hero-copy {font:400 12px/1.8 Inter,sans-serif;color:var(--muted);margin:0 0 16px}
    .section-title {font:600 19px/1.4 'Plus Jakarta Sans',sans-serif;letter-spacing:-.6px;margin:2px 0 4px}
    .section-note {font-size:11px;line-height:1.8;color:var(--muted);margin-bottom:12px}
    .metric-grid {display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:12px;margin:12px 0 22px}
    .kpi {background:var(--card);border:1px solid var(--line);border-radius:10px;padding:20px 17px 17px;min-height:117px}
    .kpi-label {font-size:10px;font-weight:400;color:#a9bacb;white-space:nowrap}
    .kpi-value {font:600 28px/1.5 'Plus Jakarta Sans',sans-serif;color:var(--gold);letter-spacing:-.6px}
    .kpi-note {font-size:9px;color:var(--muted);line-height:1.5}
    .workspace-empty {background:linear-gradient(115deg,#0B1F30,#081a29);border:1px solid var(--line);border-radius:12px;padding:34px 28px;text-align:center;margin:3px 0 18px}
    .empty-icon {display:inline-flex;align-items:center;justify-content:center;width:42px;height:42px;border:1px solid rgba(211,170,78,.25);border-radius:12px;color:var(--gold);margin-bottom:16px}
    .empty-icon svg {width:21px;height:21px}
    .empty-title {font:500 20px/1.5 'Plus Jakarta Sans',sans-serif;letter-spacing:-.6px;margin-bottom:7px}
    .empty-copy {font-size:11px;line-height:1.9;color:var(--muted);max-width:410px;margin:auto}
    .step {display:flex;gap:12px;align-items:flex-start;margin:18px 0}
    .step-num {color:var(--gold);font-size:10px;border:1px solid rgba(211,170,78,.2);border-radius:6px;padding:5px}
    .step strong {font-size:11px;font-weight:500;display:block;margin-bottom:4px}
    .step p {font-size:10px;color:var(--muted);margin:0;line-height:1.7}
    .run-label {font-size:10px;color:var(--muted);margin-top:8px;line-height:1.8;word-break:break-word}
    .footer-note {border-top:1px solid rgba(211,170,78,.10);margin-top:25px;padding-top:18px;display:flex;justify-content:space-between;font-size:9px;letter-spacing:.3px;color:var(--muted);gap:10px}
    .sidebar-footer {padding-top:22px;border-top:1px solid var(--line);margin-top:24px}
    [data-testid="stFileUploaderDropzone"] {background:rgba(15,42,64,.4);border:1px dashed rgba(211,170,78,.38);border-radius:9px;padding:20px 18px}
    [data-testid="stFileUploaderDropzone"] p {font-size:12px}
    [data-testid="stFileUploaderDropzone"] small {color:var(--muted);font-size:10px}
    [data-testid="stFileUploaderDropzone"] svg {color:var(--gold)}
    [data-testid="stVerticalBlockBorderWrapper"]>div {border-color:var(--line)!important;border-radius:12px!important}
    [data-testid="stWidgetLabel"] p {font-size:11px!important;font-weight:400!important;color:#a9bacb}
    [data-testid="stCaptionContainer"] p {font-size:10px!important;line-height:1.8;color:var(--muted)}
    [data-testid="stBaseButton-primary"] {background:linear-gradient(120deg,#D3AA4E,#c79b39)!important;color:#061521!important;border:1px solid #d3aa4e!important;min-height:42px;border-radius:8px!important;font-weight:600;box-shadow:0 4px 18px #d3aa4e1c}
    [data-testid="stBaseButton-primary"]:hover {background:#C2993D!important;border-color:#e2bc65!important;box-shadow:0 4px 22px #d3aa4e30}
    [data-testid="stBaseButton-primary"] p {color:#061521!important;font-size:11px!important;font-weight:600}
    button {font-family:Inter,sans-serif!important;transition:background .16s ease,border-color .16s ease}
    [data-testid="stBaseButton-secondary"] {border-color:rgba(211,170,78,.22);background:rgba(15,42,64,.3);border-radius:8px;min-height:39px}
    [data-testid="stBaseButton-secondary"] p {font-size:11px}
    [data-testid="stBaseButton-secondary"]:hover {color:var(--gold);border-color:var(--gold);background:#0F2A40}
    input,[data-baseweb="select"]>div {font-family:Inter,sans-serif!important;font-size:12px!important}
    [data-testid="stAlert"] {border-radius:8px;font-size:11px}
    [data-testid="stAlert"] p {font-size:11px;line-height:1.8}
    [data-testid="stExpander"] {border-color:var(--line);border-radius:8px}
    [data-testid="stExpander"] summary p {font-size:11px}
    .stAppDeployButton {display:none}
    button:focus-visible,input:focus-visible {outline:2px solid var(--gold)!important;outline-offset:3px}
    @media(max-width:1100px){[data-testid="stMainBlockContainer"]{padding:2rem 1.5rem}.kpi{padding:15px 11px}.kpi-label{font-size:9px}.kpi-value{font-size:25px}}
    @media(max-width:1000px){.metric-grid{grid-template-columns:repeat(3,minmax(0,1fr))}.kpi{padding:16px}.kpi-label{font-size:10px}}
    @media(max-width:700px){[data-testid="stMainBlockContainer"]{padding:1.4rem 1rem}.metric-grid{grid-template-columns:repeat(2,minmax(0,1fr));gap:9px}.kpi:last-child{grid-column:1/-1}.kpi{min-height:100px}.hero-title{letter-spacing:-1px}.crumbs{margin-bottom:22px}.crumbs-label b{margin-left:8px}.local-pill{font-size:9px;padding:6px 9px}.workspace-empty{padding:26px 18px}.footer-note{flex-direction:column}}
    </style>"""


def show_metrics(counts):
    cards = [
        ("Total processed", "total", "Rows in this run"),
        ("Valid emails", "valid", "Syntax + MX checks passed"),
        ("Invalid domains", "invalid_domains", "No usable mail route"),
        ("Role accounts", "role", "Generic business inboxes"),
        ("Disposable / risky", "risky", "Review before use"),
    ]
    st.markdown(
        '<div class="metric-grid">'
        + "".join(
            f'<div class="kpi"><div class="kpi-label">{label}</div><div class="kpi-value">{counts.get(key,0):,}</div><div class="kpi-note">{note}</div></div>'
            for label, key, note in cards
        )
        + "</div>",
        unsafe_allow_html=True,
    )


def get_workspace():
    """Separate temporary storage per browser session. Never cache lead datasets globally."""
    if "workspace" not in st.session_state:
        holder = tempfile.TemporaryDirectory(prefix="openhousefx-")
        st.session_state.workspace_holder = holder
        st.session_state.workspace = holder.name
    return Path(st.session_state.workspace)


@lru_cache(maxsize=1)
def load_disposable_domains():
    path = APP_ROOT / "data" / "disposable_domains.txt"
    return frozenset(
        line.strip().lower()
        for line in path.read_text(encoding="utf-8").splitlines()
        if line.strip() and not line.startswith("#")
    )


def supabase_ready():
    return bool(os.environ.get("SUPABASE_URL") and os.environ.get("SUPABASE_KEY"))


def log_run_metadata(summary):
    """Opt-in metadata only; no lead rows or filenames leave the server."""
    url = os.environ.get("SUPABASE_URL", "").rstrip("/")
    key = os.environ.get("SUPABASE_KEY", "")
    if urlparse(url).scheme != "https" or not urlparse(url).hostname or not key:
        raise ValueError(
            "Supabase requires an HTTPS project URL and a server-side key."
        )
    payload = {
        "batch_id": summary["batch_id"],
        "total_validated": summary["counts"]["total"],
        "counts": summary["counts"],
        "created_at": summary["created_at"],
    }
    response = requests.post(
        url + "/rest/v1/verifier_runs",
        json=payload,
        headers={
            "apikey": key,
            "Authorization": "Bearer " + key,
            "Prefer": "resolution=merge-duplicates,return=minimal",
        },
        timeout=(3, 7),
    )
    response.raise_for_status()


def save_completed_run(summary):
    destination = Path(os.environ["VERIFIER_DATA_DIR"]).expanduser().resolve()
    destination.mkdir(parents=True, exist_ok=True)
    token = uuid.UUID(summary["batch_id"]).hex
    target = destination / token
    shutil.copytree(Path(summary["db_path"]).parent, target)
    saved = dict(summary)
    saved["db_path"] = str(target / Path(summary["db_path"]).name)
    saved["exports"] = {
        key: str(target / Path(value).name) for key, value in summary["exports"].items()
    }
    (target / "manifest.json").write_text(
        json.dumps(saved, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return token


def restore_completed_run(token):
    token = uuid.UUID(token.strip()).hex
    directory = Path(os.environ["VERIFIER_DATA_DIR"]).expanduser().resolve() / token
    saved = json.loads((directory / "manifest.json").read_text(encoding="utf-8"))
    saved["db_path"] = str(directory / Path(saved["db_path"]).name)
    saved["exports"] = {
        key: str(directory / Path(value).name)
        for key, value in saved["exports"].items()
    }
    if not Path(saved["db_path"]).is_file() or not all(
        Path(value).is_file() for value in saved["exports"].values()
    ):
        raise ValueError("The saved results could not be found.")
    try:
        completed = load_summary(saved["db_path"])
        if (
            completed["batch_id"] != saved["batch_id"]
            or completed["counts"] != saved["counts"]
        ):
            raise ValueError("Saved run metadata does not match its results.")
    except (sqlite3.Error, KeyError) as exc:
        raise ValueError(
            "The saved result database is incomplete or unreadable."
        ) from exc
    return saved


def render_sidebar():
    with st.sidebar:
        house = (APP_ROOT / "assets" / "openhousefx-house.svg").read_text(
            encoding="utf-8"
        )
        st.markdown(
            f'<div class="brand">{house}<div><div class="wordmark">OPENHOUSE<span>FX</span></div><div class="product">LEAD VERIFIER</div></div></div>',
            unsafe_allow_html=True,
        )
        st.markdown(
            '<div class="eyebrow">01 / CONFIGURE YOUR RUN</div><div class="sidebar-title">Verification settings</div><div class="small-copy">A considered check. A cleaner list.</div>',
            unsafe_allow_html=True,
        )
        workers = st.select_slider(
            "Concurrent domain checks",
            options=[20, 25, 30, 35, 40, 45, 50],
            value=35,
            key="workers",
            help="Each unique domain is checked once per run. DNS queries share a two-second budget per domain.",
        )
        batch_size = st.selectbox(
            "Rows per batch",
            [1000, 2000, 5000],
            index=1,
            key="batch_size",
            format_func=lambda n: f"{n:,} rows",
        )
        st.caption(
            "Duplicates are tagged automatically. Role and disposable addresses are set aside for review."
        )
        with st.expander("How results are classified"):
            st.markdown(
                "**Valid** · Syntax and MX records pass, with no list hygiene flags.\n\n**Invalid** · Invalid syntax, duplicate entry, nonexistent domain or no mail route.\n\n**Risky** · Role inbox, disposable provider, DNS uncertainty or A/AAAA fallback."
            )
            st.caption(
                "Valid describes these checks. Mailbox existence, catch-all status and actual delivery are not tested."
            )
        st.markdown(
            '<div class="eyebrow" style="margin-top:22px">YOUR WORKFLOW</div><div class="step"><div class="step-num">01</div><div><strong>Bring your lead sheet</strong><p>CSV or Excel. Original columns stay with you.</p></div></div><div class="step"><div class="step-num">02</div><div><strong>Check with confidence</strong><p>Syntax, duplicates and domain records.</p></div></div><div class="step"><div class="step-num">03</div><div><strong>Review. Then export.</strong><p>A full report or just the leads you need.</p></div></div>',
            unsafe_allow_html=True,
        )
        persist = log_metadata = False
        if os.environ.get("VERIFIER_DATA_DIR") or supabase_ready():
            with st.expander("Save & recovery"):
                if os.environ.get("VERIFIER_DATA_DIR"):
                    persist = st.checkbox(
                        "Save this run on the server",
                        value=False,
                        help="Saves lead rows and reports to the configured disk. Your private recovery code restores them after a reload. Keep it private.",
                    )
                    with st.form("restore_run"):
                        token = st.text_input("Recovery code", type="password")
                        restore = st.form_submit_button("Restore saved run")
                    if restore:
                        try:
                            st.session_state.result = restore_completed_run(token)
                            st.session_state.result_source = None
                            st.session_state.restored = True
                            st.session_state.pop("run_messages", None)
                            st.session_state.recovery_code = uuid.UUID(
                                token.strip()
                            ).hex
                            st.rerun()
                        except (ValueError, OSError, KeyError, json.JSONDecodeError):
                            st.error("No completed run found for that recovery code.")
                if supabase_ready():
                    log_metadata = st.checkbox(
                        "Log run totals to Supabase",
                        value=False,
                        help="Sends only the batch ID, timestamp and counts. Lead rows and filenames are not sent.",
                    )
        st.markdown(
            f'<div class="sidebar-footer"><div class="small-copy">OPENHOUSEFX · POCKET STUDIO</div><div class="small-copy">Lead Verifier {APP_VERSION}</div></div>',
            unsafe_allow_html=True,
        )
    return workers, batch_size, persist, log_metadata


def stage_upload(uploaded, workspace):
    """Hash a memoryview without allocating a second upload-sized buffer."""
    buffer = uploaded.getbuffer()
    try:
        signature = (hashlib.sha256(buffer).hexdigest(), uploaded.name)
        if st.session_state.get("upload_signature") != signature:
            target = workspace / ("source" + Path(uploaded.name).suffix.lower())
            with target.open("wb") as handle:
                handle.write(buffer)
            for key in (
                "result",
                "input_info",
                "sheet_choice",
                "restored",
                "run_messages",
                "recovery_code",
            ):
                st.session_state.pop(key, None)
            st.session_state.upload_signature = signature
            st.session_state.input_path = str(target)
        return signature
    finally:
        buffer.release()


def render_input(workspace):
    st.markdown(
        '<div class="section-title">Your source file</div><div class="section-note">Start with your lead sheet. We’ll keep its original columns in every report.</div>',
        unsafe_allow_html=True,
    )
    uploaded = st.file_uploader(
        "Upload a lead sheet",
        type=["csv", "xlsx", "xls"],
        key="source_upload",
        label_visibility="collapsed",
        help="Maximum 50 MB. CSV is read in chunks; XLSX is streamed. Legacy XLS uses a sheet reader.",
    )
    if uploaded is None:
        if st.session_state.pop("upload_signature", None) is not None:
            st.session_state.pop("input_info", None)
            if not st.session_state.get("restored"):
                st.session_state.pop("result", None)
        return None
    signature = stage_upload(uploaded, workspace)
    path = st.session_state.input_path
    try:
        if "input_info" not in st.session_state:
            with st.spinner("Reading your file…"):
                st.session_state.input_info = inspect_input(
                    path, filename=uploaded.name
                )
        info = st.session_state.input_info
        sheets = info.get("sheets", [])
        sheet = None
        if sheets:
            sheet = st.selectbox("Worksheet", sheets, key="sheet_choice")
            if info.get("sheet_name") != sheet:
                with st.spinner("Reading worksheet…"):
                    info = inspect_input(path, filename=uploaded.name, sheet_name=sheet)
                    st.session_state.input_info = info
        columns = info["columns"]
        default = info.get("email_index")
        mapping_key = "mapping_" + signature[0][:12] + "_" + str(sheet)
        left, right = st.columns([1.5, 1])
        with left:
            email_index = st.selectbox(
                "Email column",
                [None] + list(range(len(columns))),
                index=(int(default) + 1 if default is not None else 0),
                format_func=lambda i: (
                    "Choose an email column…" if i is None else columns[i]
                ),
                key=mapping_key,
            )
        with right:
            st.markdown(
                f'<div class="run-label">FILE READY<br><span style="color:#F0F4F8">{html.escape(uploaded.name)}</span><br>{info.get("row_count",0):,} rows · {len(columns)} columns</div>',
                unsafe_allow_html=True,
            )
        if default is None:
            st.caption("Choose the column containing email addresses to continue.")
        return {
            "path": path,
            "filename": uploaded.name,
            "sheet": sheet,
            "email_index": email_index,
            "signature": (signature, sheet, email_index),
            "info": info,
        }
    except Exception as exc:
        st.error(f"Could not read this file: {exc}")
        return None


def execute_run(source, workers, batch_size, persist, log_metadata, metrics_slot):
    run_dir = get_workspace() / uuid.uuid4().hex
    progress_bar = st.progress(0, text="Preparing your list…")
    detail = st.empty()
    last_update = [0.0]

    def report(event):
        # Engine callbacks run on the main Streamlit thread, never DNS workers.
        now = time.monotonic()
        stage = event.get("stage", "")
        if now - last_update[0] < 0.15 and stage not in {
            "Complete",
            "Preparing exports",
        }:
            return
        last_update[0] = now
        done = event.get("rows_processed", 0)
        total = event.get("total_rows") or source["info"].get("row_count") or 0
        fraction = min(done / total, 1) if total else 0
        eta = event.get("eta")
        eta_text = (
            f" · About {max(1,round(eta)):,}s remaining"
            if eta is not None and done < total
            else ""
        )
        label = (
            "Preparing your exports…"
            if stage == "Preparing exports"
            else f"{fraction:.0%} · {done:,} of {total:,} rows{eta_text}"
        )
        progress_bar.progress(fraction, text=label)
        detail.caption(
            f'{event.get("domains_checked",0):,} unique domains checked · {event.get("elapsed",0):.1f}s elapsed'
        )
        if event.get("counts"):
            with metrics_slot.container():
                show_metrics(event["counts"])

    try:
        summary = run_verification(
            source["path"],
            run_dir,
            source["email_index"],
            workers=workers,
            batch_size=batch_size,
            sheet_name=source["sheet"],
            filename=source["filename"],
            disposable_domains=load_disposable_domains(),
            progress=report,
        )
    except Exception as exc:
        progress_bar.empty()
        detail.empty()
        st.error(
            f"Verification could not finish: {exc}. Any previous completed report remains available."
        )
        return
    summary["source_name"] = source["filename"]
    st.session_state.result = summary
    st.session_state.result_source = source["signature"]
    st.session_state.pop("restored", None)
    st.session_state.pop("recovery_code", None)
    messages = []
    if persist:
        try:
            st.session_state.recovery_code = save_completed_run(summary)
            messages.append("Saved on the configured server disk.")
        except (OSError, ValueError):
            messages.append(
                "Server save failed; download your reports before leaving this session."
            )
    if log_metadata:
        try:
            log_run_metadata(summary)
            messages.append("Run totals logged to Supabase.")
        except (requests.RequestException, ValueError):
            messages.append(
                "Supabase logging failed; your verification results and downloads are ready."
            )
    st.session_state.run_messages = messages
    progress_bar.empty()
    detail.empty()
    st.rerun()


def download_file(path):
    """Read only the requested export, on click, rather than all files on rerun."""
    return Path(path).read_bytes()


def render_results(summary):
    counts = summary["counts"]
    st.success(
        f'Run complete · {counts["total"]:,} rows reviewed · {counts["valid"]:,} valid · {counts["risky"]:,} risky · {counts["invalid"]:,} invalid'
    )
    for message in st.session_state.get("run_messages", []):
        st.caption(message)
    st.markdown(
        '<div class="section-title">Your verification report</div><div class="section-note">Every row has a clear outcome. Filter the view, then choose the export you need.</div>',
        unsafe_allow_html=True,
    )
    left, middle, right = st.columns([1, 1.8, 0.75])
    with left:
        status = st.selectbox(
            "Validation status",
            ["All", "Valid", "Invalid", "Risky"],
            key="result_status",
        )
    with middle:
        search = st.text_input(
            "Search this report",
            placeholder="Email, company, reason…",
            key="result_search",
        )
    with right:
        limit = st.selectbox("Rows per page", [50, 100, 250], index=1, key="page_size")
    filter_key = (summary["batch_id"], status, search, limit)
    if st.session_state.get("last_filter") != filter_key:
        st.session_state.report_page = 1
        st.session_state.last_filter = filter_key
    frame, total = read_results(
        summary["db_path"], status=status, search=search, limit=limit, offset=0
    )
    pages = max(1, math.ceil(total / limit))
    pagination, info_col = st.columns([1, 3])
    with pagination:
        page = st.number_input(
            "Page", min_value=1, max_value=pages, step=1, key="report_page"
        )
    if page > 1:
        frame, total = read_results(
            summary["db_path"],
            status=status,
            search=search,
            limit=limit,
            offset=(page - 1) * limit,
        )
    with info_col:
        st.caption(
            f"{total:,} matching rows · Page {page:,} of {pages:,}. Exports include the complete category, independent of this view."
        )
    if total:
        leading = ["Verification_Status", summary["email_column"], "Reason"]
        display_order = list(dict.fromkeys(leading + list(frame.columns)))
        st.dataframe(
            frame,
            column_order=display_order,
            column_config={
                "Verification_Status": st.column_config.TextColumn(
                    "Status", width="small"
                ),
                "Reason": st.column_config.TextColumn("Reason", width="large"),
            },
            hide_index=True,
            width="stretch",
            height=min(460, max(180, len(frame) * 35 + 40)),
        )
    else:
        st.info("No rows match this filter. Try another status or search.")
    st.caption(
        f'{counts["duplicates"]:,} duplicate rows · {counts["role"]:,} role accounts · {counts["disposable"]:,} disposable addresses. Flags can overlap.'
    )
    st.markdown(
        '<div class="section-title" style="margin-top:15px">Take your list forward</div><div class="section-note">All original columns, with verification status and reasons appended.</div>',
        unsafe_allow_html=True,
    )
    downloads = [
        (
            "Full report · CSV",
            "full_csv",
            "openhousefx_full_report.csv",
            "All rows and all outcomes.",
        ),
        (
            "Valid only · CSV",
            "valid_csv",
            "mx_verified_leads.csv",
            "Unique addresses with passing syntax and MX checks; no risky flags.",
        ),
        (
            "Invalid & risky · CSV",
            "invalid_risky_csv",
            "openhousefx_review_report.csv",
            "Rows to remove or review.",
        ),
        (
            "Role & disposable · CSV",
            "role_disposable_csv",
            "openhousefx_role_disposable.csv",
            "All role or disposable rows, including duplicates.",
        ),
    ]
    for column, (label, key, filename, help_text) in zip(st.columns(4), downloads):
        with column:
            path = summary["exports"][key]
            st.download_button(
                label,
                data=lambda path=path: download_file(path),
                file_name=filename,
                mime="text/csv",
                key="export_" + key + summary["batch_id"],
                on_click="ignore",
                type="primary" if key == "valid_csv" else "secondary",
                width="stretch",
                help=help_text,
            )
    xlsx_path = str(Path(summary["db_path"]).parent / "openhousefx_full_report.xlsx")
    if not Path(xlsx_path).exists():
        if st.button(
            "Prepare full Excel report",
            key="prepare_xlsx",
            help="Generated on demand to keep normal verification fast.",
        ):
            try:
                with st.spinner("Preparing Excel report…"):
                    with _xlsx_lock:
                        generate_xlsx(summary["db_path"], xlsx_path)
                st.rerun()
            except Exception as exc:
                st.error(
                    f"Excel export could not be prepared: {exc}. Your CSV reports are still available."
                )
    else:
        st.download_button(
            "Full report · Excel (.xlsx)",
            data=lambda: download_file(xlsx_path),
            file_name="openhousefx_full_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="export_xlsx_" + summary["batch_id"],
            on_click="ignore",
        )
    with st.expander("Run details & interpretation"):
        st.caption(
            f'Batch {summary["batch_id"]} · {summary.get("source_name",summary.get("filename","Lead sheet"))}'
        )
        st.write(
            "**Valid** means syntax and MX checks passed. It does not confirm mailbox existence or guarantee delivery. Timeouts and temporary DNS failures are risky; retry them later. An A/AAAA fallback is also risky because it does not confirm an email service."
        )
        st.caption(
            "Duplicate comparison trims surrounding spaces and ignores case. Plus tags and dots are retained. All original rows remain in the full report."
        )
        if summary.get("column_renames"):
            st.write(
                "Original columns renamed to avoid duplicate names or verification-field collisions:"
            )
            st.json(summary["column_renames"])
        if st.session_state.get("recovery_code"):
            st.write(
                "Private recovery code — treat it as access to the saved lead sheet:"
            )
            st.code(st.session_state.recovery_code, language=None)


def main():
    st.set_page_config(
        page_title="OpenhouseFX · Lead Verifier",
        page_icon=str(APP_ROOT / "assets" / "openhousefx-house.svg"),
        layout="wide",
        initial_sidebar_state="expanded",
    )
    st.markdown(brand_css(), unsafe_allow_html=True)
    workspace = get_workspace()
    workers, batch_size, persist, log_metadata = render_sidebar()
    st.markdown(
        '<div class="crumbs"><div class="crumbs-label">Openhousefx <b>/ &nbsp; Verification workspace</b></div><div class="local-pill"><i></i>Private session</div></div><div class="eyebrow">THE CONFIDENCE BEFORE CONTACT</div><div class="hero-title" role="heading" aria-level="1">Every lead. <span>A clearer picture.</span></div><p class="hero-copy">Bring your list into focus. Check email syntax, domain records and list hygiene in one considered workspace.</p>',
        unsafe_allow_html=True,
    )
    metrics_slot = st.empty()
    with st.container(border=True):
        source = render_input(workspace)
        if source:
            completed = st.session_state.get("result")
            same_run = (
                completed is not None
                and st.session_state.get("result_source") == source["signature"]
            )
            if completed and not same_run and not st.session_state.get("restored"):
                st.session_state.pop("result", None)
            if st.button(
                "Verify again" if same_run else "Verify lead list",
                type="primary",
                width="stretch",
                disabled=source["email_index"] is None
                or not source["info"].get("row_count", 0),
            ):
                execute_run(
                    source, workers, batch_size, persist, log_metadata, metrics_slot
                )
    summary = st.session_state.get("result")
    with metrics_slot.container():
        show_metrics(summary["counts"] if summary else {})
    if summary:
        render_results(summary)
    else:
        st.markdown(
            '<div class="section-title">Live workspace</div><div class="section-note">Your report will appear here when the check is complete.</div><div class="workspace-empty"><div class="empty-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.3"><path d="M4 5h16v14H4z M4 6l8 6 8-6"/><path d="m8 16 2 2 5-5"/></svg></div><div class="empty-title">A cleaner list starts here.</div><div class="empty-copy">Upload your lead sheet above, confirm the email column, and run your first check. Every original row stays in your full report.</div></div>',
            unsafe_allow_html=True,
        )
        st.caption(
            "Checks confirm syntax and domain records. Mailbox existence and delivery are not guaranteed."
        )
    st.markdown(
        '<div class="footer-note"><span>OPENHOUSEFX · LEAD VERIFIER</span><span>Thoughtfully checked. Clearly organized.</span></div>',
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
