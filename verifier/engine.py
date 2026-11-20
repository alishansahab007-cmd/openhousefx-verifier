"""Disk-backed, batch-oriented lead verification engine (no Streamlit dependency)."""

from __future__ import annotations

import csv
import json
import math
import os
import re
import sqlite3
import threading
import time
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed
from contextlib import contextmanager
from dataclasses import asdict, dataclass
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Callable, Iterable

import pandas as pd
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell

from . import email_checks as ec
from . import enrich
from . import phone as ph
from . import smtp_check
from .columns import BOOLEAN_COLUMNS, FIELD_LABELS, NUMERIC_COLUMNS, RESULT_COLUMNS
from .email_checks import DomainResult
from .io import inspect_input, raw_rows

EMAIL_STATUSES = ("Valid", "Risky", "Invalid", "Duplicate", "Not checked")
VERDICTS = ("Good", "Review", "Bad")
STATUS_FILTERS = (
    "All", "Good", "Review", "Bad", "Valid", "Risky", "Invalid", "Duplicate",
    "Invalid & Risky", "Role / Disposable", "Catch-all", "Mailbox confirmed",
    "Phone valid", "Phone invalid", "Active", "Lapsing", "Inactive",
)
COUNT_KEYS = (
    "total", "good", "review", "bad", "valid", "risky", "invalid", "duplicates",
    "not_checked", "invalid_domains", "role", "disposable", "free", "typos",
    "mailbox_exists", "mailbox_missing", "catch_all", "mailbox_unverifiable",
    "phone_valid", "phone_invalid", "phone_missing", "phone_duplicates",
    "line_active", "line_inactive", "active", "lapsing", "inactive",
)


@dataclass
class VerifyOptions:
    smtp_check: bool = False
    domain_age: bool = False
    line_check: bool = False
    default_region: str = "US"
    dayfirst: bool = False
    active_days: int = 90
    inactive_days: int = 365
    smtp_workers: int = 10
    strict: bool = False  # only mailbox-confirmed emails can be "Good"
    today: str | None = None  # ISO date override (tests / reproducible runs)


# --------------------------------------------------------------------------- storage
@contextmanager
def _open_db(path):
    """Always release file handles, including on failed runs."""
    connection = sqlite3.connect(str(path))
    try:
        connection.execute("PRAGMA journal_mode=DELETE")
        with connection:
            yield connection
    finally:
        connection.close()


def _sql_type(column):
    if column in BOOLEAN_COLUMNS:
        return "INTEGER"
    if column in NUMERIC_COLUMNS:
        return "REAL"
    return "TEXT"


_SCHEMA = (
    """
    CREATE TABLE metadata (key TEXT PRIMARY KEY, value TEXT NOT NULL);
    CREATE TABLE seen_emails (email TEXT PRIMARY KEY);
    CREATE TABLE seen_phones (phone TEXT PRIMARY KEY);
    CREATE TABLE domains (domain TEXT PRIMARY KEY, status TEXT, reason TEXT, dns_check TEXT, mx_hosts TEXT);
    CREATE TABLE mailboxes (email TEXT PRIMARY KEY, state TEXT, detail TEXT);
    CREATE TABLE smtp_domains (domain TEXT PRIMARY KEY, catch_all INTEGER);
    CREATE TABLE domain_ages (domain TEXT PRIMARY KEY, created TEXT, status TEXT);
    CREATE TABLE lines (phone TEXT PRIMARY KEY, status TEXT, line_type TEXT, carrier TEXT);
    CREATE TABLE results (row_index INTEGER PRIMARY KEY, payload TEXT NOT NULL, search_text TEXT NOT NULL, """
    + ", ".join(f'"{column}" {_sql_type(column)}' for column in RESULT_COLUMNS)
    + """);
    CREATE INDEX results_verdict ON results(Lead_Verdict);
    CREATE INDEX results_status ON results(Verification_Status);
    CREATE INDEX results_country ON results(Lead_Country);
    CREATE INDEX results_state ON results(Lead_State);
    """
)


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


def _metadata(connection) -> dict:
    row = connection.execute("SELECT value FROM metadata WHERE key='summary'").fetchone()
    if row is None:
        raise ValueError("This verification did not finish. Start a new run.")
    return json.loads(row[0])


def load_summary(db_path) -> dict:
    with _open_db(db_path) as connection:
        return _metadata(connection)


# --------------------------------------------------------------------------- scoring
_RISK_SCORES = {
    "disposable": 10, "junk": 15, "typo": 20, "full": 30, "dns": 40, "fallback": 40,
    "role": 55, "catch_all": 60,
}
_PHONE_TYPE_SCORES = {
    "Mobile": 92, "Landline or mobile": 86, "Landline": 80, "Company (UAN)": 75,
    "VoIP": 62, "Toll-free": 60, "Personal number": 70, "Shared cost": 55,
    "Premium rate": 20, "Pager": 20, "Voicemail": 30, "Unknown": 60,
}


def _verdict(email: dict | None, phone: dict | None, activity: str) -> tuple[str, int, list[str]]:
    """Combine email and phone evidence into Good / Review / Bad with a 0-100 score."""
    issues = []
    email_score = phone_score = None
    email_hard_bad = email_dup = email_missing = False
    if email is not None:
        status = email["status"]
        if email["missing"]:
            email_missing = True
            issues.append("No email address")
        elif status == "Duplicate":
            email_dup, email_score = True, 0
            issues.append("Duplicate email (already earlier in the list)")
        elif status == "Invalid":
            email_hard_bad, email_score = True, 0
            issues.append("Email: " + email["reason"])
        elif status == "Risky":
            risks = email["risks"]
            email_score = min((_RISK_SCORES[risk] for risk in risks), default=40)
            email_hard_bad = bool({"disposable", "junk"} & set(risks))
            issues.append("Email: " + email["reason"])
        else:
            email_score = 100 if email["mailbox"] == smtp_check.EXISTS else 85
    phone_present = phone is not None and phone["status"] != "Missing"
    if phone is not None:
        if phone["status"] == "Missing":
            if email is None:
                issues.append("No phone number")
        elif phone["status"] == "Invalid":
            phone_score = 0
            issues.append("Phone: " + (phone["reason"] or "invalid number"))
        else:
            line = phone["line"]
            if line in ("Inactive", "Invalid"):
                phone_score = 0
                issues.append(f"Phone line {line.lower()} (carrier lookup)")
            elif line == "Unreachable":
                phone_score = 40
                issues.append("Phone currently unreachable (carrier lookup)")
            else:
                phone_score = 100 if line == "Active" else _PHONE_TYPE_SCORES.get(phone["type"], 60)
            if phone["duplicate"]:
                phone_score = max(0, phone_score - 15)
                issues.append("Phone number shared with an earlier lead")
    parts = [(score, weight) for score, weight in ((email_score, 0.7), (phone_score, 0.3)) if score is not None]
    score = round(sum(s * w for s, w in parts) / sum(w for _, w in parts)) if parts else 0
    if activity == "Lapsing":
        score -= 5
    elif activity == "Inactive":
        score -= 15
        issues.append("Client inactive (no recent activity)")
    score = max(0, min(100, score))
    phone_good = phone_score is not None and phone_score >= 75
    if email is None:
        if phone is None or not phone_present or phone_score == 0:
            verdict = "Bad"
        elif phone["duplicate"]:
            verdict = "Bad"
            issues.append("Duplicate lead (same phone)")
        else:
            verdict = "Good" if phone_good else "Review"
    elif email_dup:
        verdict = "Bad"
    elif email_missing or email_hard_bad:
        verdict = "Review" if phone_good else "Bad"
        if phone_good:
            issues.append("Reachable by phone only")
    elif email["status"] == "Risky":
        verdict = "Review"
    else:
        verdict = "Review" if phone_present and phone_score == 0 else "Good"
        if verdict == "Good" and email.get("strict") and email["mailbox"] != smtp_check.EXISTS:
            verdict = "Review"
            issues.append("Strict mode: mailbox not confirmed by the mail server")
    if email_dup or (email_missing and not phone_present) or (email is None and phone is not None and phone["duplicate"]):
        score = 0  # a duplicate lead adds nothing new
    return verdict, score, issues


# --------------------------------------------------------------------------- helpers
def _emit(callback, start, stage, processed, total, checked, unique, counts=None, note=""):
    if callback is None:
        return
    elapsed = time.monotonic() - start
    eta = (elapsed / processed) * (total - processed) if processed else None
    callback({
        "stage": stage, "rows_processed": processed, "total_rows": total,
        "domains_checked": checked, "unique_domains": unique, "elapsed": elapsed,
        "eta": eta, "counts": dict(counts or {}), "note": note,
    })


def _parallel(pool, items, fn, on_result, limit, tick=None):
    """Bounded submission: at most `limit` futures outstanding."""
    iterator, pending = iter(items), {}

    def fill():
        while len(pending) < limit:
            try:
                item = next(iterator)
            except StopIteration:
                return
            pending[pool.submit(fn, item)] = item

    fill()
    while pending:
        future = next(as_completed(tuple(pending)))
        item = pending.pop(future)
        try:
            value, error = future.result(), None
        except Exception as exc:  # a failed lookup must never sink the run
            value, error = None, exc
        on_result(item, value, error)
        if tick:
            tick()
        fill()


_HTTP_LOCAL = threading.local()


def _default_domain_age(domain):
    import requests

    if not hasattr(_HTTP_LOCAL, "session"):
        _HTTP_LOCAL.session = requests.Session()
        _HTTP_LOCAL.session.headers["User-Agent"] = "lead-verifier/3 (+rdap)"
    return enrich.rdap_creation_date(domain, _HTTP_LOCAL.session)


def _default_line(e164):
    import requests

    if not hasattr(_HTTP_LOCAL, "twilio"):
        _HTTP_LOCAL.twilio = requests.Session()
    return ph.twilio_line_lookup(e164, _HTTP_LOCAL.twilio)


def _cell(row, mapping, field):
    index = mapping.get(field)
    return row[index] if index is not None else None


# --------------------------------------------------------------------------- main run
def run_verification(
    input_path,
    output_dir,
    email_index: int | None = None,
    *,
    mapping: dict | None = None,
    workers: int = 35,
    batch_size: int = 2000,
    sheet_name: str | None = None,
    disposable_domains: frozenset = frozenset(),
    options: VerifyOptions | None = None,
    progress: Callable | None = None,
    domain_lookup: Callable | None = None,
    mailbox_lookup: Callable | None = None,
    domain_age_lookup: Callable | None = None,
    line_lookup: Callable | None = None,
    filename: str | None = None,
) -> dict:
    """Verify a lead sheet in bounded batches; every original row is kept."""
    options = options or VerifyOptions()
    if not 1 <= int(workers) <= 64:
        raise ValueError("Workers must be between 1 and 64.")
    if not 1 <= int(batch_size) <= 5000:
        raise ValueError("Batch size must be between 1 and 5,000.")
    start = time.monotonic()
    today = date.fromisoformat(options.today) if options.today else date.today()
    info = inspect_input(input_path, filename=filename, sheet_name=sheet_name)
    mapping = {k: v for k, v in dict(mapping or {}).items() if v is not None and k in FIELD_LABELS}
    if email_index is not None:
        mapping["email"] = email_index
    for field, index in list(mapping.items()):
        if not 0 <= int(index) < len(info["columns"]):
            raise ValueError(f"Choose a valid {FIELD_LABELS[field].lower()} column.")
        mapping[field] = int(index)
    if "email" not in mapping and "phone" not in mapping:
        raise ValueError("Choose an email column or a phone column to verify.")
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    db_path = output_dir / "results.sqlite3"
    if db_path.exists():
        raise ValueError("This result folder already contains a run. Choose a new folder.")

    notes = []
    has_email, has_phone = "email" in mapping, "phone" in mapping
    lookup = domain_lookup or ec.lookup_domain
    if domain_lookup is None and has_email:
        ec.validate_dns_configuration()
    smtp_active = False
    if options.smtp_check and has_email:
        if mailbox_lookup is not None:
            smtp_active = True
        elif smtp_check.api_provider_ready() and os.getenv("VERIFIER_MAILBOX_PROVIDER", "").lower() == "zerobounce":
            smtp_active, mailbox_lookup = True, smtp_check.zerobounce_mailboxes
        elif smtp_check.port25_available():
            smtp_active = True
            mailbox_lookup = smtp_check.check_domain_mailboxes
        elif smtp_check.api_provider_ready():
            smtp_active, mailbox_lookup = True, smtp_check.zerobounce_mailboxes
            notes.append("Port 25 is blocked on this server, so mailboxes were checked through the ZeroBounce API.")
        else:
            notes.append(
                "Mailbox check skipped: this server cannot reach mail servers on port 25 "
                "(most cloud hosts, including Render, block it). Results use syntax + DNS checks only. "
                "Add a ZEROBOUNCE_API_KEY or run on a host with port 25 open for mailbox-level accuracy."
                + (" Strict mode had no effect." if options.strict else "")
            )
    age_active = bool(options.domain_age)
    age_lookup = domain_age_lookup or _default_domain_age
    line_active = False
    if options.line_check and has_phone:
        if line_lookup is not None or ph.twilio_ready():
            line_active, line_lookup = True, line_lookup or _default_line
        else:
            notes.append("Live phone line check skipped: Twilio credentials are not configured.")

    counts = dict.fromkeys(COUNT_KEYS, 0)
    batch_id = str(uuid.uuid4())
    summary = {
        "batch_id": batch_id,
        "db_path": str(db_path),
        "filename": filename or Path(input_path).name,
        "created_at": datetime.now(timezone.utc).isoformat(),
        "sheet_name": info["sheet_name"],
        "original_columns": info["columns"],
        "column_renames": info["column_renames"],
        "mapping": {field: info["columns"][index] for field, index in mapping.items()},
        "email_column": info["columns"][mapping["email"]] if has_email else None,
        "row_count": info["row_count"],
        "counts": counts,
        "exports": {},
        "workers": int(workers),
        "batch_size": int(batch_size),
        "options": asdict(options),
        "checks": {
            "email": has_email, "phone": has_phone, "smtp": smtp_active,
            "domain_age": age_active, "line": line_active,
            "dob": "dob" in mapping or "age" in mapping, "created": "created" in mapping,
            "activity": "last_activity" in mapping,
        },
        "notes": notes,
    }
    checked = 0
    total_rows = info["row_count"]
    with _open_db(db_path) as connection:
        connection.executescript(_SCHEMA)
        _emit(progress, start, "Preparing", 0, total_rows, 0, 0, counts)
        dns_pool = ThreadPoolExecutor(max_workers=int(workers), thread_name_prefix="ohfx-dns")
        smtp_pool = ThreadPoolExecutor(max_workers=max(1, int(options.smtp_workers)), thread_name_prefix="ohfx-smtp")
        http_pool = ThreadPoolExecutor(max_workers=8, thread_name_prefix="ohfx-http")
        try:
            source = iter(raw_rows(Path(input_path), info))
            while True:
                batch = []
                for _ in range(int(batch_size)):
                    try:
                        batch.append(next(source))
                    except StopIteration:
                        break
                if not batch:
                    break
                prepared = []
                uncached_domains, age_domains, line_numbers = set(), set(), set()
                for row in batch:
                    item = {"row": row}
                    # ---- location first: it sets the phone region.
                    country_alpha, country_label = enrich.normalize_country(
                        enrich.clean_text(_cell(row, mapping, "country"))
                    )
                    region = country_alpha or (options.default_region or "US").upper()
                    # ---- email
                    if has_email:
                        email, domain, syntax_error, note = ec.check_syntax(_cell(row, mapping, "email"))
                        raw = _cell(row, mapping, "email")
                        key = email.casefold() if email else (
                            raw.strip().casefold() if isinstance(raw, str) and raw.strip() else None
                        )
                        duplicate = False
                        if key:
                            cursor = connection.execute("INSERT OR IGNORE INTO seen_emails(email) VALUES (?)", (key,))
                            duplicate = cursor.rowcount == 0
                        item.update(
                            email=email, domain=domain, syntax_error=syntax_error, note=note,
                            duplicate=duplicate,
                            role=bool(email) and ec.is_role(email),
                            no_reply=bool(email) and ec.is_no_reply(email),
                            junk=bool(email) and ec.is_junk(email),
                            disposable=bool(domain) and ec.disposable_match(domain, disposable_domains),
                            free=bool(domain) and ec.is_free_provider(domain),
                            typo=ec.typo_suggestion(domain) if domain else None,
                        )
                        if domain and not syntax_error and connection.execute(
                            "SELECT 1 FROM domains WHERE domain=?", (domain,)
                        ).fetchone() is None:
                            uncached_domains.add(domain)
                    # ---- phone
                    if has_phone:
                        result = ph.parse_phone(_cell(row, mapping, "phone"), region)
                        phone_dup = False
                        if result.status == "Valid":
                            cursor = connection.execute("INSERT OR IGNORE INTO seen_phones(phone) VALUES (?)", (result.e164,))
                            phone_dup = cursor.rowcount == 0
                            if line_active and connection.execute(
                                "SELECT 1 FROM lines WHERE phone=?", (result.e164,)
                            ).fetchone() is None:
                                line_numbers.add(result.e164)
                        item.update(phone=result, phone_dup=phone_dup)
                    # ---- location fallbacks
                    if not country_alpha and has_phone and item["phone"].country:
                        country_alpha = item["phone"].country
                        country_label = enrich.country_name(country_alpha)
                    if not country_alpha and not country_label and has_email and item.get("domain"):
                        hinted = enrich.country_from_domain(item["domain"])
                        if hinted:
                            country_alpha, country_label = hinted, enrich.country_name(hinted)
                    item["country"] = country_label
                    item["state"] = enrich.normalize_state(
                        enrich.clean_text(_cell(row, mapping, "state")), country_alpha
                    )
                    # ---- company domain for domain age
                    company_domain = ""
                    if "website" in mapping:
                        company_domain = enrich.website_domain(_cell(row, mapping, "website"))
                    if not company_domain and has_email and item.get("domain") and not item["free"] and not item["disposable"] and not item["typo"]:
                        company_domain = item["domain"]
                    if company_domain:
                        company_domain = enrich.registrable_domain(company_domain)
                    item["company_domain"] = company_domain
                    if age_active and company_domain and connection.execute(
                        "SELECT 1 FROM domain_ages WHERE domain=?", (company_domain,)
                    ).fetchone() is None:
                        age_domains.add(company_domain)
                    prepared.append(item)

                # ---- DNS (each domain once per run; one retry pass for timeouts)
                target = checked + len(uncached_domains)

                def store_domain(domain, result, error):
                    nonlocal checked
                    if error is not None or getattr(result, "status", None) not in {"Valid", "Invalid", "Risky"}:
                        result = DomainResult("Risky", "Domain lookup failed; retry later", "DNS error")
                    connection.execute(
                        "INSERT OR REPLACE INTO domains VALUES (?,?,?,?,?)",
                        (domain, result.status, result.reason, result.dns_check,
                         json.dumps(list(getattr(result, "mx_hosts", ()) or ()))),
                    )
                    checked += 1

                def tick_dns():
                    if checked % 10 == 0:
                        _emit(progress, start, "Checking email domains", counts["total"], total_rows, checked, target, counts)

                _parallel(dns_pool, sorted(uncached_domains), lookup, store_domain, 2 * int(workers), tick_dns)
                retry = [
                    domain for (domain,) in connection.execute(
                        "SELECT domain FROM domains WHERE dns_check IN ('Timeout','DNS error','Network error')"
                    ) if domain in uncached_domains
                ]
                if retry:
                    checked -= len(retry)
                    _parallel(dns_pool, retry, lookup, store_domain, 2 * int(workers))

                # ---- SMTP mailbox checks
                if smtp_active:
                    by_domain: dict[str, list[str]] = {}
                    for item in prepared:
                        if not item.get("email") or item["duplicate"] or item["disposable"] or item["no_reply"]:
                            continue
                        status, _, _, hosts = connection.execute(
                            "SELECT status, reason, dns_check, mx_hosts FROM domains WHERE domain=?", (item["domain"],)
                        ).fetchone()
                        if status != "Valid":
                            continue
                        if connection.execute("SELECT 1 FROM mailboxes WHERE email=?", (item["email"],)).fetchone():
                            continue
                        known = connection.execute("SELECT catch_all FROM smtp_domains WHERE domain=?", (item["domain"],)).fetchone()
                        if known and known[0] == 1:
                            connection.execute(
                                "INSERT OR REPLACE INTO mailboxes VALUES (?,?,?)",
                                (item["email"], smtp_check.CATCH_ALL, "Domain accepts any address"),
                            )
                            continue
                        by_domain.setdefault(item["domain"], []).append(item["email"])
                    smtp_done = [0]
                    # Read MX hosts here: SQLite connections stay on the main thread.
                    mx_for = {
                        domain: json.loads(connection.execute(
                            "SELECT mx_hosts FROM domains WHERE domain=?", (domain,)
                        ).fetchone()[0] or "[]")
                        for domain in by_domain
                    }

                    def probe(domain):
                        return mailbox_lookup(domain, mx_for[domain], by_domain[domain])

                    def store_mailboxes(domain, value, error):
                        if error is not None or value is None:
                            value = (None, {
                                address: smtp_check.MailboxResult(smtp_check.UNVERIFIABLE, "Check failed")
                                for address in by_domain[domain]
                            })
                        catch_all, results = value
                        if catch_all is not None:
                            connection.execute("INSERT OR REPLACE INTO smtp_domains VALUES (?,?)", (domain, int(catch_all)))
                        for address in by_domain[domain]:
                            result = results.get(address) or smtp_check.MailboxResult(smtp_check.UNVERIFIABLE, "No reply")
                            connection.execute(
                                "INSERT OR REPLACE INTO mailboxes VALUES (?,?,?)", (address, result.state, result.detail)
                            )
                        smtp_done[0] += 1
                        _emit(progress, start, "Checking mailboxes", counts["total"], total_rows, checked, checked, counts,
                              note=f"{smtp_done[0]:,} of {len(by_domain):,} mail domains in this batch")

                    _parallel(smtp_pool, sorted(by_domain), probe, store_mailboxes, 2 * max(1, int(options.smtp_workers)))

                # ---- company domain age (RDAP)
                if age_domains:
                    def store_age(domain, value, error):
                        created, status = (value if error is None and value else (None, "Lookup failed"))
                        connection.execute(
                            "INSERT OR REPLACE INTO domain_ages VALUES (?,?,?)",
                            (domain, created.isoformat() if created else None, status),
                        )

                    _emit(progress, start, "Looking up company domain ages", counts["total"], total_rows, checked, checked, counts)
                    _parallel(http_pool, sorted(age_domains), age_lookup, store_age, 16)

                # ---- live phone line status (Twilio, optional)
                if line_numbers:
                    def store_line(number, value, error):
                        value = value if error is None and value else ph.LineResult("Error")
                        connection.execute(
                            "INSERT OR REPLACE INTO lines VALUES (?,?,?,?)",
                            (number, value.status, value.line_type, value.carrier),
                        )

                    _emit(progress, start, "Checking phone lines", counts["total"], total_rows, checked, checked, counts)
                    _parallel(http_pool, sorted(line_numbers), line_lookup, store_line, 16)

                # ---- assemble rows
                inserts = []
                for item in prepared:
                    inserts.append(_assemble(item, mapping, connection, counts, options, today, has_email, has_phone,
                                             smtp_active, age_active))
                columns_sql = ", ".join(["row_index", "payload", "search_text"] + [f'"{c}"' for c in RESULT_COLUMNS])
                placeholders = ",".join("?" * (3 + len(RESULT_COLUMNS)))
                connection.executemany(f"INSERT INTO results ({columns_sql}) VALUES ({placeholders})", inserts)
                connection.commit()
                _emit(progress, start, "Verifying rows", counts["total"], total_rows, checked, checked, counts)
        finally:
            for pool in (dns_pool, smtp_pool, http_pool):
                pool.shutdown(wait=True, cancel_futures=True)
        if smtp_active and counts["mailbox_unverifiable"]:
            notes.append(
                f"{counts['mailbox_unverifiable']:,} email(s) could not be mailbox-checked (the mail server was "
                "unreachable from this network or refused the check). They keep their DNS-based result; "
                "see the Mailbox column for each reason."
            )
        summary["unique_domains"] = checked
        summary["duration_seconds"] = round(time.monotonic() - start, 3)
        summary["exports"] = {
            "full_csv": str(output_dir / "lead_report_all.csv"),
            "good_csv": str(output_dir / "lead_report_good.csv"),
            "review_csv": str(output_dir / "lead_report_review.csv"),
            "bad_csv": str(output_dir / "lead_report_bad.csv"),
            "valid_csv": str(output_dir / "lead_report_valid_emails.csv"),
        }
    _emit(progress, start, "Preparing exports", counts["total"], total_rows, checked, checked, counts)
    _generate_csv_exports(db_path, summary)
    with _open_db(db_path) as connection:
        connection.execute("INSERT INTO metadata VALUES ('summary',?)", (json.dumps(summary),))
    _emit(progress, start, "Complete", counts["total"], total_rows, checked, checked, counts)
    return summary


def _assemble(item, mapping, connection, counts, options, today, has_email, has_phone, smtp_active, age_active):
    row = item["row"]
    out = dict.fromkeys(RESULT_COLUMNS)
    email_info = None
    if has_email:
        email, domain = item["email"], item["domain"]
        dns_status = dns_reason = ""
        dns_check, mx_host = "Not checked", ""
        if domain and not item["syntax_error"]:
            dns_status, dns_reason, dns_check, hosts = connection.execute(
                "SELECT status, reason, dns_check, mx_hosts FROM domains WHERE domain=?", (domain,)
            ).fetchone()
            hosts = json.loads(hosts or "[]")
            mx_host = hosts[0] if hosts and dns_check in ("MX", "MX broken") else ""
        mailbox, mailbox_detail = smtp_check.NOT_CHECKED, ""
        if smtp_active and email:
            found = connection.execute("SELECT state, detail FROM mailboxes WHERE email=?", (email,)).fetchone()
            if found:
                mailbox, mailbox_detail = found
            elif item["duplicate"]:
                mailbox = "See first occurrence"
        risks, reasons = [], []
        if item["syntax_error"]:
            status, reason = "Invalid", item["syntax_error"]
            if item["typo"]:
                reason += f"; did you mean …@{item['typo']}?"
        elif item["duplicate"]:
            status, reason = "Duplicate", "Duplicate email; see first occurrence"
        elif item["no_reply"]:
            status, reason = "Invalid", "No-reply address; it does not receive mail"
        elif dns_status == "Invalid":
            status, reason = "Invalid", dns_reason
            if item["typo"]:
                reason += f"; did you mean {email.split('@')[0]}@{item['typo']}?"
        elif mailbox == smtp_check.DO_NOT_MAIL:
            status, reason = "Invalid", "Known spam trap or abuse address; never email it"
        elif mailbox in (smtp_check.NOT_EXISTS, smtp_check.DISABLED_MAILBOX):
            status = "Invalid"
            reason = ("Mailbox does not exist" if mailbox == smtp_check.NOT_EXISTS else "Mailbox is disabled") + (
                f" (server said: {mailbox_detail})" if mailbox_detail else ""
            )
        else:
            if item["disposable"]:
                risks.append("disposable"), reasons.append("Disposable / temporary email provider")
            if item["junk"]:
                risks.append("junk"), reasons.append("Looks like a placeholder or test address")
            if item["typo"]:
                risks.append("typo"), reasons.append(f"Possible typo; did you mean {email.split('@')[0]}@{item['typo']}?")
            if item["role"]:
                risks.append("role"), reasons.append("Role-based inbox (info@, sales@ …), not a person")
            if dns_status == "Risky":
                risks.append("fallback" if dns_check in ("A", "AAAA") else "dns"), reasons.append(dns_reason)
            if mailbox == smtp_check.CATCH_ALL:
                risks.append("catch_all"), reasons.append("Catch-all domain: accepts any address, so this mailbox can't be confirmed")
            if mailbox == smtp_check.FULL:
                risks.append("full"), reasons.append("Mailbox is full; mail will bounce until it is cleared")
            if risks:
                status, reason = "Risky", "; ".join(reasons)
            else:
                status = "Valid"
                reason = "Mailbox confirmed by the mail server" if mailbox == smtp_check.EXISTS else dns_reason
        if item["note"] and status != "Invalid":
            reason += f" ({item['note']})"
        email_info = {
            "status": status, "reason": reason, "risks": risks, "mailbox": mailbox,
            "missing": item["syntax_error"] == "Missing email address",
            "strict": bool(options.strict and smtp_active),
        }
        mailbox_label = mailbox
        if mailbox == smtp_check.UNVERIFIABLE and mailbox_detail:
            mailbox_label = f"Unverifiable: {mailbox_detail}"[:200]
        out.update({
            "Verification_Status": status, "Reason": reason, "Mailbox_Check": mailbox_label,
            "Normalized_Email": email, "Email_Domain": domain, "DNS_Check": dns_check, "MX_Host": mx_host,
            "Is_Role": int(item["role"]), "Is_Disposable": int(item["disposable"]),
            "Is_Free_Provider": int(item["free"]), "Is_Duplicate": int(item["duplicate"]),
            "Typo_Suggestion": f"{email.split('@')[0]}@{item['typo']}" if item["typo"] and email else (
                item["typo"] or ""),
        })
        counts[{"Valid": "valid", "Risky": "risky", "Invalid": "invalid", "Duplicate": "duplicates"}[status]] += 1
        counts["invalid_domains"] += int(dns_status == "Invalid")
        counts["role"] += int(item["role"])
        counts["disposable"] += int(item["disposable"])
        counts["free"] += int(item["free"])
        counts["typos"] += int(bool(item["typo"]))
        counts["mailbox_exists"] += int(mailbox == smtp_check.EXISTS)
        counts["mailbox_missing"] += int(mailbox in (smtp_check.NOT_EXISTS, smtp_check.DISABLED_MAILBOX, smtp_check.DO_NOT_MAIL))
        counts["catch_all"] += int(mailbox == smtp_check.CATCH_ALL)
        counts["mailbox_unverifiable"] += int(mailbox == smtp_check.UNVERIFIABLE)
    else:
        out.update({"Verification_Status": "Not checked", "Reason": "No email column mapped",
                    "Mailbox_Check": smtp_check.NOT_CHECKED, "Is_Role": 0, "Is_Disposable": 0,
                    "Is_Free_Provider": 0, "Is_Duplicate": 0})
        counts["not_checked"] += 1

    phone_info = None
    if has_phone:
        result = item["phone"]
        line_status, line_type, line_carrier = ("Not checked", "", "")
        if result.status == "Valid":
            found = connection.execute("SELECT status, line_type, carrier FROM lines WHERE phone=?", (result.e164,)).fetchone()
            if found:
                line_status, line_type, line_carrier = found
        phone_info = {"status": result.status, "reason": result.reason, "type": result.type,
                      "line": line_status, "duplicate": item["phone_dup"]}
        out.update({
            "Phone_Status": result.status, "Phone_E164": result.e164,
            "Phone_Type": line_type or result.type, "Phone_Country": enrich.country_name(result.country),
            "Phone_Carrier": line_carrier or result.carrier, "Phone_Location": result.location,
            "Phone_Line_Status": line_status if result.status == "Valid" else "", "Is_Duplicate_Phone": int(item["phone_dup"]),
        })
        if result.status != "Valid" and result.reason:
            out["Phone_Line_Status"] = result.reason
        counts["phone_" + result.status.lower()] += 1
        counts["phone_duplicates"] += int(item["phone_dup"])
        counts["line_active"] += int(line_status == "Active")
        counts["line_inactive"] += int(line_status in ("Inactive", "Invalid"))
    else:
        out["Is_Duplicate_Phone"] = 0

    # ---- people & activity
    age = None
    if "age" in mapping:
        age = enrich.parse_age(_cell(row, mapping, "age"))
    if age is None and "dob" in mapping:
        age = enrich.age_years(enrich.parse_date(_cell(row, mapping, "dob"), options.dayfirst, today), today)
    created_days = enrich.days_since(enrich.parse_date(_cell(row, mapping, "created"), options.dayfirst, today), today) if "created" in mapping else None
    activity_days = enrich.days_since(enrich.parse_date(_cell(row, mapping, "last_activity"), options.dayfirst, today), today) if "last_activity" in mapping else None
    activity = enrich.activity_status(activity_days, options.active_days, options.inactive_days) if "last_activity" in mapping else "Unknown"
    if activity != "Unknown":
        counts[activity.lower()] += 1
    out.update({
        "Lead_Name": enrich.build_name(_cell(row, mapping, "name"), _cell(row, mapping, "first_name"), _cell(row, mapping, "last_name")),
        "Lead_Company": enrich.clean_text(_cell(row, mapping, "company")),
        "Lead_Country": item["country"], "Lead_State": item["state"],
        "Age": age, "Record_Age_Days": created_days, "Record_Age": enrich.humanize_days(created_days),
        "Last_Activity_Days": activity_days,
        "Activity_Status": activity if "last_activity" in mapping else "",
    })
    if age_active and item["company_domain"]:
        found = connection.execute("SELECT created, status FROM domain_ages WHERE domain=?", (item["company_domain"],)).fetchone()
        if found and found[0]:
            created = date.fromisoformat(found[0])
            out["Domain_Created"] = f"{item['company_domain']} · {created.isoformat()}"
            out["Domain_Age_Years"] = round((today - created).days / 365.25, 1)
        elif found:
            out["Domain_Created"] = f"{item['company_domain']} · {found[1]}"

    verdict, score, issues = _verdict(email_info, phone_info, activity)
    out.update({"Lead_Verdict": verdict, "Lead_Score": score, "Lead_Issues": "; ".join(issues) or "No issues found"})
    counts["total"] += 1
    counts[verdict.lower()] += 1
    values = [out[column] for column in RESULT_COLUMNS]
    payload = json.dumps([_json_value(value) for value in row], ensure_ascii=False)
    search_text = " ".join(
        [*(str(value) for value in row if value is not None),
         *(str(out[c]) for c in ("Lead_Verdict", "Verification_Status", "Reason", "Lead_Issues", "Lead_Country",
                                 "Lead_State", "Phone_E164", "Activity_Status", "Mailbox_Check") if out[c])]
    ).casefold()
    return [counts["total"], payload, search_text, *values]


# --------------------------------------------------------------------------- queries
def _result_values(record):
    """record = (row_index, payload, *RESULT_COLUMNS)."""
    values = [_decode_value(value) for value in json.loads(record[1])]
    derived = []
    for column, value in zip(RESULT_COLUMNS, record[2:]):
        if column in BOOLEAN_COLUMNS:
            value = bool(value) if value is not None else False
        elif column in NUMERIC_COLUMNS and value is not None and float(value).is_integer() and column != "Domain_Age_Years":
            value = int(value)
        derived.append(value)
    return values + derived


_SELECT = "SELECT row_index, payload, " + ", ".join(f'"{c}"' for c in RESULT_COLUMNS) + " FROM results"


def _column_expr(column: str, original_columns: list[str]) -> str:
    if column in RESULT_COLUMNS:
        return f'"{column}"'
    if column in original_columns:
        return f"json_extract(payload, '$[{original_columns.index(column)}]')"
    raise ValueError(f"Unknown column: {column}")


def _where(status: str, search: str, filters: dict | None, original_columns) -> tuple[str, list]:
    where, params = [], []
    simple = {
        "Good": "Lead_Verdict='Good'", "Review": "Lead_Verdict='Review'", "Bad": "Lead_Verdict='Bad'",
        "Valid": "Verification_Status='Valid'", "Risky": "Verification_Status='Risky'",
        "Invalid": "Verification_Status='Invalid'", "Duplicate": "Verification_Status='Duplicate'",
        "Invalid & Risky": "Verification_Status IN ('Invalid','Risky','Duplicate')",
        "Role / Disposable": "(Is_Role=1 OR Is_Disposable=1)",
        "Catch-all": f"Mailbox_Check='{smtp_check.CATCH_ALL}'",
        "Mailbox confirmed": f"Mailbox_Check='{smtp_check.EXISTS}'",
        "Phone valid": "Phone_Status='Valid'", "Phone invalid": "Phone_Status='Invalid'",
        "Active": "Activity_Status='Active'", "Lapsing": "Activity_Status='Lapsing'",
        "Inactive": "Activity_Status='Inactive'",
    }
    if status in simple:
        where.append(simple[status])
    for column, value in (filters or {}).items():
        if value in (None, "All"):
            continue
        expr = _column_expr(column, original_columns)
        if value == "(blank)":
            where.append(f"({expr} IS NULL OR {expr} = '')")
        else:
            where.append(f"CAST({expr} AS TEXT) = ?")
            params.append(str(value))
    if search and search.strip():
        where.append("instr(search_text, ?) > 0")
        params.append(search.strip().casefold())
    return (" WHERE " + " AND ".join(where)) if where else "", params


def _order(sort_by, descending, original_columns) -> str:
    if not sort_by:
        return " ORDER BY row_index"
    expr = _column_expr(sort_by, original_columns)
    direction = "DESC" if descending else "ASC"
    numeric = sort_by in NUMERIC_COLUMNS
    collate = "" if numeric else " COLLATE NOCASE"
    return f" ORDER BY ({expr} IS NULL OR {expr} = ''), {expr}{collate} {direction}, row_index"


def read_results(
    db_path, status: str = "All", search: str = "", limit: int = 100, offset: int = 0,
    *, filters: dict | None = None, sort_by: str | None = None, descending: bool = False,
) -> tuple[pd.DataFrame, int]:
    """Fetch one bounded page with parameterized filters and whitelisted sorting."""
    limit, offset = max(1, min(10000, int(limit))), max(0, int(offset))
    with _open_db(db_path) as connection:
        summary = _metadata(connection)
        columns = summary["original_columns"]
        clause, params = _where(status, search, filters, columns)
        count = connection.execute("SELECT COUNT(*) FROM results" + clause, params).fetchone()[0]
        records = connection.execute(
            _SELECT + clause + _order(sort_by, descending, columns) + " LIMIT ? OFFSET ?",
            [*params, limit, offset],
        ).fetchall()
    frame = pd.DataFrame([_result_values(record) for record in records], columns=columns + RESULT_COLUMNS)
    return frame, count


def distinct_values(db_path, column: str, limit: int = 300) -> list:
    with _open_db(db_path) as connection:
        columns = _metadata(connection)["original_columns"]
        expr = _column_expr(column, columns)
        rows = connection.execute(
            f"SELECT CAST({expr} AS TEXT) AS v, COUNT(*) FROM results GROUP BY v ORDER BY COUNT(*) DESC LIMIT ?",
            (limit,),
        ).fetchall()
    return [("(blank)" if value in (None, "") else value) for value, _ in rows]


def group_counts(db_path, group_by: str, status="All", search="", filters=None) -> pd.DataFrame:
    """Lead counts per group (country, state, verdict, any column …)."""
    with _open_db(db_path) as connection:
        columns = _metadata(connection)["original_columns"]
        expr = _column_expr(group_by, columns)
        clause, params = _where(status, search, filters, columns)
        rows = connection.execute(
            f"""SELECT COALESCE(NULLIF(CAST({expr} AS TEXT), ''), '(blank)') AS g, COUNT(*),
                SUM(Lead_Verdict='Good'), SUM(Lead_Verdict='Review'), SUM(Lead_Verdict='Bad'),
                ROUND(AVG(Lead_Score), 1)
                FROM results{clause} GROUP BY g ORDER BY COUNT(*) DESC, g""",
            params,
        ).fetchall()
    return pd.DataFrame(rows, columns=[group_by, "Leads", "Good", "Review", "Bad", "Avg score"])


def iter_results(db_path, status="All", search="", filters=None, sort_by=None, descending=False) -> Iterable[list]:
    with _open_db(db_path) as connection:
        columns = _metadata(connection)["original_columns"]
        clause, params = _where(status, search, filters, columns)
        for record in connection.execute(_SELECT + clause + _order(sort_by, descending, columns), params):
            yield _result_values(record)


# --------------------------------------------------------------------------- exports
def _generate_csv_exports(db_path, summary):
    """Write all category CSVs in a single streaming pass."""
    handles = {}
    try:
        writers = {}
        for key, path in summary["exports"].items():
            handles[key] = open(path, "w", encoding="utf-8-sig", newline="")
            writers[key] = csv.writer(handles[key])
            writers[key].writerow(summary["original_columns"] + RESULT_COLUMNS)
        verdict_at = len(summary["original_columns"]) + RESULT_COLUMNS.index("Lead_Verdict")
        status_at = len(summary["original_columns"]) + RESULT_COLUMNS.index("Verification_Status")
        with _open_db(db_path) as connection:
            for record in connection.execute(_SELECT + " ORDER BY row_index"):
                values = _result_values(record)
                writers["full_csv"].writerow(values)
                writers[values[verdict_at].lower() + "_csv"].writerow(values)
                if values[status_at] == "Valid":
                    writers["valid_csv"].writerow(values)
    finally:
        for handle in handles.values():
            handle.close()


def export_view_csv(db_path, path, **query) -> str:
    """Export exactly the filtered + sorted view (all pages)."""
    summary = load_summary(db_path)
    with open(path, "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(summary["original_columns"] + RESULT_COLUMNS)
        for values in iter_results(db_path, **query):
            writer.writerow(values)
    return str(path)


_SHEET_BAD = re.compile(r"[\[\]:*?/\\]")


def _sheet_title(name, used):
    base = _SHEET_BAD.sub("-", str(name or "(blank)")).strip("'") or "(blank)"
    base = base[:31]
    title, n = base, 2
    while title.casefold() in used:
        suffix = f" ({n})"
        title, n = base[: 31 - len(suffix)] + suffix, n + 1
    used.add(title.casefold())
    return title


def _literal_appender(sheet):
    def append(values):
        cells = []
        for value in values:
            if isinstance(value, str):
                if len(value) > 32767:
                    raise ValueError("A cell exceeds Excel's 32,767-character limit. Use CSV to preserve it.")
                cell = WriteOnlyCell(sheet, value=value)
                cell.data_type = "s"  # formula-like text stays literal text
                cells.append(cell)
            else:
                cells.append(value)
        sheet.append(cells)

    return append


def _summary_rows(summary):
    counts = summary["counts"]
    rows = [
        ("Source file", summary.get("source_name") or summary.get("filename")),
        ("Verified at (UTC)", summary.get("created_at", "")[:19].replace("T", " ")),
        ("Total leads", counts["total"]),
        ("Good leads", counts["good"]), ("Review", counts["review"]), ("Bad leads", counts["bad"]),
    ]
    if summary["checks"].get("email"):
        rows += [("Valid emails", counts["valid"]), ("Risky emails", counts["risky"]),
                 ("Invalid emails", counts["invalid"]), ("Duplicate emails", counts["duplicates"]),
                 ("Mailboxes confirmed", counts["mailbox_exists"]), ("Catch-all domains", counts["catch_all"])]
    if summary["checks"].get("phone"):
        rows += [("Valid phones", counts["phone_valid"]), ("Invalid phones", counts["phone_invalid"]),
                 ("Missing phones", counts["phone_missing"])]
    if summary["checks"].get("activity"):
        rows += [("Active clients", counts["active"]), ("Lapsing clients", counts["lapsing"]),
                 ("Inactive clients", counts["inactive"])]
    return rows


def _check_excel_limits(summary):
    if summary["counts"]["total"] > 1_048_575:
        raise ValueError("This report exceeds Excel's row limit. Use the CSV export.")
    if len(summary["original_columns"]) + len(RESULT_COLUMNS) > 16_384:
        raise ValueError("This report exceeds Excel's column limit. Use the CSV export.")


def generate_xlsx(db_path, path, sort_by=None, descending=False) -> str:
    """Summary sheet + all leads; formula-like text is written as literal text."""
    path = Path(path)
    summary = load_summary(db_path)
    _check_excel_limits(summary)
    workbook = Workbook(write_only=True)
    try:
        sheet = workbook.create_sheet("Summary")
        append = _literal_appender(sheet)
        append(["Metric", "Value"])
        for row in _summary_rows(summary):
            append(list(row))
        sheet = workbook.create_sheet("All leads")
        sheet.freeze_panes = "A2"
        append = _literal_appender(sheet)
        append(summary["original_columns"] + RESULT_COLUMNS)
        for values in iter_results(db_path, sort_by=sort_by, descending=descending):
            append(values)
        workbook.save(path)
    except Exception:
        workbook.close()
        if path.exists():
            path.unlink()
        raise
    return str(path)


def generate_organized_xlsx(
    db_path, path, group_by: str, sort_by=None, descending=False, status="All", search="",
    filters=None, max_sheets: int = 120,
) -> str:
    """One worksheet per group (e.g. per country or state), each sorted, plus a summary."""
    path = Path(path)
    summary = load_summary(db_path)
    _check_excel_limits(summary)
    groups = group_counts(db_path, group_by, status=status, search=search, filters=filters)
    headers = summary["original_columns"] + RESULT_COLUMNS
    workbook = Workbook(write_only=True)
    used = {"summary", "other groups"}
    try:
        sheet = workbook.create_sheet("Summary")
        append = _literal_appender(sheet)
        append([f"Leads grouped by {group_by}"])
        append(list(groups.columns))
        for record in groups.itertuples(index=False):
            append(list(record))
        main_groups = list(groups[group_by])[:max_sheets]
        for value in main_groups:
            sheet = workbook.create_sheet(_sheet_title(value, used))
            sheet.freeze_panes = "A2"
            append = _literal_appender(sheet)
            append(headers)
            group_filter = dict(filters or {})
            group_filter[group_by] = value
            for values in iter_results(db_path, status=status, search=search, filters=group_filter,
                                       sort_by=sort_by, descending=descending):
                append(values)
        rest = list(groups[group_by])[max_sheets:]
        if rest:
            sheet = workbook.create_sheet("Other groups")
            append = _literal_appender(sheet)
            append(headers)
            for value in rest:
                group_filter = dict(filters or {})
                group_filter[group_by] = value
                for values in iter_results(db_path, status=status, search=search, filters=group_filter,
                                           sort_by=sort_by, descending=descending):
                    append(values)
        workbook.save(path)
    except Exception:
        workbook.close()
        if path.exists():
            path.unlink()
        raise
    return str(path)
