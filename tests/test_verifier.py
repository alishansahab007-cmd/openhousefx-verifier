"""Deterministic engine regressions; no live DNS or external services required."""

from __future__ import annotations
import csv
import importlib.util
import os
from pathlib import Path
import sys
from types import SimpleNamespace
import pytest
import dns.exception
import dns.name
import dns.resolver
from openpyxl import Workbook, load_workbook

DEFAULT_MODULE = Path(__file__).with_name("engine_section.py")
if not DEFAULT_MODULE.exists():
    DEFAULT_MODULE = Path(__file__).resolve().parents[1] / "app.py"
MODULE_PATH = Path(os.environ.get("VERIFIER_MODULE", DEFAULT_MODULE))
spec = importlib.util.spec_from_file_location("verifier_under_test", MODULE_PATH)
engine = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = engine
spec.loader.exec_module(engine)


def csv_input(
    tmp_path, rows, headers=("Email", "Contact ID", "Note"), encoding="utf-8-sig"
):
    source = tmp_path / "leads.csv"
    with source.open("w", encoding=encoding, newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(headers)
        writer.writerows(rows)
    return source


def good_lookup(domain):
    return engine.DomainResult("Valid", "MX records present", "MX")


def all_rows(summary):
    frame, count = engine.read_results(summary["db_path"], limit=10000)
    assert len(frame) == count
    return frame


def csv_rows(path):
    with Path(path).open(encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def test_chunked_pipeline_unique_dns_duplicates_flags_and_original_values(tmp_path):
    source = csv_input(
        tmp_path,
        [
            ("alice@goodmail.com", "0001", "Zoë, 東京"),
            ("bob@goodmail.com", "0002", "second"),
            (" alice@goodmail.com ", "0003", "duplicate"),
            ("sales@goodmail.com", "0004", "role"),
            ("temp@sub.throwaway.com", "0005", "disposable"),
            ("two@@goodmail.com", "0006", "syntax"),
            ("user@badmail.com", "0007", "invalid domain"),
            ("user@slowmail.com", "0008", "timeout"),
        ],
    )
    calls, updates = [], []

    def lookup(domain):
        calls.append(domain)
        if domain == "badmail.com":
            return engine.DomainResult("Invalid", "Domain does not exist", "NXDOMAIN")
        if domain == "slowmail.com":
            return engine.DomainResult("Risky", "DNS timed out; retry later", "TIMEOUT")
        return good_lookup(domain)

    result = engine.run_verification(
        source,
        tmp_path / "result",
        0,
        batch_size=2,
        disposable_domains={"throwaway.com"},
        domain_lookup=lookup,
        progress=updates.append,
    )
    frame = all_rows(result)
    assert len(frame) == 8
    assert frame["Contact ID"].tolist() == [f"{i:04d}" for i in range(1, 9)]
    assert frame.iloc[0]["Note"] == "Zoë, 東京"
    assert frame.iloc[2]["Email"] == " alice@goodmail.com "
    assert frame["Verification_Status"].tolist() == [
        "Valid",
        "Valid",
        "Invalid",
        "Risky",
        "Risky",
        "Invalid",
        "Invalid",
        "Risky",
    ]
    assert bool(frame.iloc[2]["Is_Duplicate"])
    assert bool(frame.iloc[3]["Is_Role"])
    assert bool(frame.iloc[4]["Is_Disposable"])
    assert calls.count("goodmail.com") == 1
    assert len(calls) == len(set(calls))
    assert result["counts"]["total"] == 8
    assert result["counts"]["valid"] == 2
    assert result["counts"]["duplicates"] == 1
    assert result["counts"]["invalid_domains"] == 1
    assert updates and updates[-1]["rows_processed"] == 8
    assert all(
        updates[i]["rows_processed"] <= updates[i + 1]["rows_processed"]
        for i in range(len(updates) - 1)
    )


@pytest.mark.parametrize(
    "header,expected",
    [
        ("email", 1),
        ("e-mail", 1),
        ("mail", 1),
        ("contact", 1),
        ("Contact Email Address", 1),
        ("Recipient", None),
    ],
)
def test_column_mapping_detection_and_manual_override(tmp_path, header, expected):
    source = csv_input(
        tmp_path, [("Nadia", "nadia@goodmail.com")], headers=("Name", header)
    )
    assert engine.inspect_input(source)["email_index"] == expected
    result = engine.run_verification(
        source, tmp_path / "result", 1, domain_lookup=good_lookup
    )
    assert result["counts"]["valid"] == 1
    assert all_rows(result).iloc[0][header] == "nadia@goodmail.com"


def test_invalid_syntax_never_reaches_dns(tmp_path):
    source = csv_input(
        tmp_path,
        [
            (value, str(i), "")
            for i, value in enumerate(
                [
                    "",
                    "not an address",
                    "name@@gmail.com",
                    "a..b@gmail.com",
                    "@gmail.com",
                    "x@-bad.com",
                    "x@gmail.com,other@gmail.com",
                ]
            )
        ],
    )

    def unexpected(_domain):
        pytest.fail("Invalid syntax must not trigger a DNS request")

    result = engine.run_verification(
        source, tmp_path / "result", 0, domain_lookup=unexpected
    )
    assert result["counts"]["valid"] == 0
    assert result["counts"]["invalid"] == 7


def test_csv_utf8_latin1_semicolon_and_header_collision(tmp_path):
    source = tmp_path / "legacy.csv"
    source.write_bytes(
        "e-mail;Verification_Status;Reason;Account\nzoe@goodmail.com;legacy;Müller;0012\n".encode(
            "latin1"
        )
    )
    inspected = engine.inspect_input(source)
    assert inspected["delimiter"] == ";"
    assert inspected["email_index"] == 0
    result = engine.run_verification(
        source, tmp_path / "result", 0, domain_lookup=good_lookup
    )
    frame = all_rows(result)
    assert frame.iloc[0]["Verification_Status"] == "Valid"
    assert "legacy" in frame.iloc[0].tolist()
    assert "Müller" in frame.iloc[0].tolist()
    assert "0012" in frame.iloc[0].tolist()
    assert len(frame.columns) == len(set(frame.columns))


def test_xlsx_sheet_selection_chunks_and_safe_generated_cells(tmp_path):
    source = tmp_path / "leads.xlsx"
    book = Workbook()
    book.active.title = "Instructions"
    book.active.append(["Read me"])
    sheet = book.create_sheet("Contacts")
    sheet.append(["Email", "ID", "Note"])
    sheet.append(["alice@goodmail.com", "0002", "=1+1"])
    sheet.cell(2, 3).data_type = "s"
    sheet.append(["bob@goodmail.com", "0003", "résumé"])
    book.save(source)
    inspected = engine.inspect_input(source, sheet_name="Contacts")
    assert inspected["sheets"] == ["Instructions", "Contacts"]
    assert inspected["columns"][:3] == ["Email", "ID", "Note"]
    result = engine.run_verification(
        source,
        tmp_path / "result",
        0,
        sheet_name="Contacts",
        batch_size=1,
        domain_lookup=good_lookup,
    )
    assert result["counts"]["total"] == 2
    export = tmp_path / "verified.xlsx"
    engine.generate_xlsx(result["db_path"], export)
    exported = load_workbook(export, data_only=False)
    rows = list(exported.active.iter_rows())
    headers = [cell.value for cell in rows[0]]
    note = rows[1][headers.index("Note")]
    assert note.value == "=1+1"
    assert note.data_type == "s"
    assert rows[1][headers.index("ID")].value == "0002"
    assert len(rows) == 3
    exported.close()


def test_real_legacy_xls_import(tmp_path):
    xlwt = pytest.importorskip(
        "xlwt", reason="Only needed to create a legacy XLS fixture"
    )
    source = tmp_path / "legacy.xls"
    book = xlwt.Workbook()
    sheet = book.add_sheet("Leads")
    for row_index, row in enumerate([["Email", "ID"], ["alice@goodmail.com", "0007"]]):
        for col_index, value in enumerate(row):
            sheet.write(row_index, col_index, value)
    book.save(str(source))
    assert engine.inspect_input(source)["email_index"] == 0
    result = engine.run_verification(
        source, tmp_path / "result", 0, domain_lookup=good_lookup
    )
    assert all_rows(result).iloc[0]["ID"] == "0007"


def test_exports_pagination_filters_and_result_reopen(tmp_path):
    source = csv_input(
        tmp_path,
        [
            ("alice@goodmail.com", "0001", "alpha"),
            ("sales@goodmail.com", "0002", "role"),
            ("bad", "0003", "invalid"),
            ("temp@throwaway.com", "0004", "disposable"),
        ],
    )
    result = engine.run_verification(
        source,
        tmp_path / "result",
        0,
        disposable_domains={"throwaway.com"},
        domain_lookup=good_lookup,
    )
    exports = result["exports"]
    assert len(csv_rows(exports["full_csv"])) == 4
    valid = csv_rows(exports["valid_csv"])
    assert len(valid) == 1 and valid[0]["Email"] == "alice@goodmail.com"
    assert len(csv_rows(exports["invalid_risky_csv"])) == 3
    assert len(csv_rows(exports["role_disposable_csv"])) == 2
    page, total = engine.read_results(result["db_path"], limit=2, offset=2)
    assert total == 4 and page["Contact ID"].tolist() == ["0003", "0004"]
    page, total = engine.read_results(
        result["db_path"], status="Risky", search="sales", limit=100
    )
    assert total == 1 and page.iloc[0]["Email"] == "sales@goodmail.com"
    page, total = engine.read_results(
        result["db_path"], search="%' OR 1=1 --", limit=100
    )
    assert total == 0
    reopened, total = engine.read_results(str(Path(result["db_path"])), status="Valid")
    assert total == 1 and len(reopened) == 1


class FakeResolver:
    def __init__(self, replies, clock_state=None, duration=0.0):
        self.replies, self.calls = replies, []
        self.clock_state, self.duration = clock_state, duration

    def resolve(self, qname, rdtype, lifetime=None, search=None):
        self.calls.append((str(qname), rdtype, lifetime))
        if self.clock_state is not None:
            self.clock_state[0] += self.duration
        value = self.replies[rdtype]
        if isinstance(value, BaseException):
            raise value
        return value


def mx_answer(host="mx.goodmail.com."):
    return [SimpleNamespace(exchange=dns.name.from_text(host), preference=10)]


def test_dns_mx_positive_and_null_mx_never_falls_back():
    resolver = FakeResolver({"MX": mx_answer()})
    assert engine.lookup_domain("goodmail.com", resolver=resolver).status == "Valid"
    assert [call[1] for call in resolver.calls] == ["MX"]
    resolver = FakeResolver({"MX": mx_answer(".")})
    assert engine.lookup_domain("nomail.com", resolver=resolver).status == "Invalid"
    assert [call[1] for call in resolver.calls] == ["MX"]


@pytest.mark.parametrize(
    "error,status",
    [
        (dns.resolver.NXDOMAIN(), "Invalid"),
        (dns.exception.Timeout(), "Risky"),
        (dns.resolver.NoNameservers(), "Risky"),
    ],
)
def test_dns_definitive_failure_and_transient_failure_separate(error, status):
    resolver = FakeResolver({"MX": error})
    assert engine.lookup_domain("unavailable.com", resolver=resolver).status == status
    assert len(resolver.calls) == 1


def test_dns_fallback_has_shared_two_second_budget_and_remains_risky():
    clock_state = [0.0]
    resolver = FakeResolver(
        {
            "MX": dns.resolver.NoAnswer(),
            "A": dns.resolver.NoAnswer(),
            "AAAA": [SimpleNamespace(address="2001:4860:4860::8888")],
        },
        clock_state,
        duration=0.5,
    )
    result = engine.lookup_domain(
        "fallback.com", resolver=resolver, timeout=2.0, clock=lambda: clock_state[0]
    )
    assert result.status == "Risky"
    assert [call[1] for call in resolver.calls] == ["MX", "A", "AAAA"]
    assert [call[2] for call in resolver.calls] == pytest.approx([2.0, 1.5, 1.0])
    assert clock_state[0] <= 2.0


def test_dns_fallback_budget_exhaustion_stops_new_queries():
    clock_state = [0.0]
    resolver = FakeResolver(
        {"MX": dns.resolver.NoAnswer(), "A": dns.resolver.NoAnswer()},
        clock_state,
        duration=1.0,
    )
    result = engine.lookup_domain(
        "fallback.com", resolver=resolver, timeout=2.0, clock=lambda: clock_state[0]
    )
    assert result.status == "Risky"
    assert [call[1] for call in resolver.calls] == ["MX", "A"]
    assert max(call[2] for call in resolver.calls) <= 2.0


@pytest.mark.skipif(
    os.environ.get("RUN_STRESS") != "1",
    reason="Set RUN_STRESS=1 for the 100,000-row regression",
)
def test_100k_rows_use_bounded_domain_cache_and_complete_exports(tmp_path):
    source = tmp_path / "100k.csv"
    with source.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["Email", "Lead ID"])
        for index in range(100_000):
            writer.writerow([f"person{index}@mail{index % 25}.com", f"{index:07d}"])
    calls = []

    def lookup(domain):
        calls.append(domain)
        return good_lookup(domain)

    result = engine.run_verification(
        source,
        tmp_path / "result",
        0,
        workers=20,
        batch_size=1000,
        domain_lookup=lookup,
    )
    assert result["counts"]["total"] == result["counts"]["valid"] == 100_000
    assert len(calls) == len(set(calls)) == 25
    with Path(result["exports"]["full_csv"]).open(
        encoding="utf-8-sig", newline=""
    ) as handle:
        assert sum(1 for _ in csv.reader(handle)) == 100_001
    page, total = engine.read_results(result["db_path"], limit=10, offset=99_990)
    assert total == 100_000 and len(page) == 10
    assert page.iloc[-1]["Lead ID"] == "0099999"


def test_strict_local_part_length_and_unicode_domain_normalization(tmp_path):
    source = csv_input(
        tmp_path,
        [
            ("a" * 65 + "@gmail.com", "0001", "overlong"),
            ("reader@bücher.de", "0002", "IDNA"),
        ],
    )
    calls = []

    def lookup(domain):
        calls.append(domain)
        return good_lookup(domain)

    result = engine.run_verification(
        source, tmp_path / "result", 0, domain_lookup=lookup
    )
    frame = all_rows(result)
    assert frame["Verification_Status"].tolist() == ["Invalid", "Valid"]
    assert calls == ["xn--bcher-kva.de"]


def test_dns_absent_routes_are_invalid_and_a_fallback_unconfirmed():
    resolver = FakeResolver(
        {
            "MX": dns.resolver.NoAnswer(),
            "A": dns.resolver.NoAnswer(),
            "AAAA": dns.resolver.NoAnswer(),
        }
    )
    assert engine.lookup_domain("nomail.com", resolver=resolver).status == "Invalid"
    resolver = FakeResolver(
        {"MX": dns.resolver.NoAnswer(), "A": [SimpleNamespace(address="8.8.8.8")]}
    )
    assert engine.lookup_domain("fallback.com", resolver=resolver).status == "Risky"
    assert [call[1] for call in resolver.calls] == ["MX", "A"]


def test_disposable_matching_does_not_flag_unrelated_suffix(tmp_path):
    source = csv_input(
        tmp_path,
        [
            ("a@throwaway.com", "01", "exact"),
            ("b@sub.throwaway.com", "02", "subdomain"),
            ("c@notthrowaway.com", "03", "unrelated"),
        ],
    )
    result = engine.run_verification(
        source,
        tmp_path / "result",
        0,
        disposable_domains={"throwaway.com"},
        domain_lookup=good_lookup,
    )
    assert all_rows(result)["Is_Disposable"].tolist() == [True, True, False]


def test_repeated_malformed_email_is_flagged_but_blank_cells_are_not(tmp_path):
    source = csv_input(
        tmp_path,
        [
            ("broken", "01", "first"),
            (" broken ", "02", "repeat"),
            ("", "03", "blank"),
            ("", "04", "second blank"),
        ],
    )
    result = engine.run_verification(
        source, tmp_path / "result", 0, batch_size=1, domain_lookup=good_lookup
    )
    frame = all_rows(result)
    assert frame["Verification_Status"].tolist() == ["Invalid"] * 4
    assert frame["Is_Duplicate"].tolist() == [False, True, False, False]


def test_failed_export_is_not_recoverable_as_completed_run(tmp_path, monkeypatch):
    source = csv_input(tmp_path, [("alice@goodmail.com", "0001", "test")])

    def failed_export(_db_path, _summary):
        raise OSError("Simulated export failure")

    monkeypatch.setattr(engine, "_generate_csv_exports", failed_export)
    destination = tmp_path / "failed"
    with pytest.raises(OSError, match="Simulated export failure"):
        engine.run_verification(source, destination, 0, domain_lookup=good_lookup)
    with pytest.raises(ValueError, match="did not finish"):
        engine.load_summary(destination / "results.sqlite3")


@pytest.mark.skipif(
    not hasattr(engine, "save_completed_run"),
    reason="Persistence UI helper belongs to final app",
)
def test_optional_saved_run_can_restore_without_temporary_source(tmp_path, monkeypatch):
    import shutil

    source = csv_input(tmp_path, [("alice@goodmail.com", "0001", "test")])
    result = engine.run_verification(
        source, tmp_path / "temporary", 0, domain_lookup=good_lookup
    )
    monkeypatch.setenv("VERIFIER_DATA_DIR", str(tmp_path / "persistent"))
    token = engine.save_completed_run(result)
    source.unlink()
    shutil.rmtree(tmp_path / "temporary")
    restored = engine.restore_completed_run(token)
    assert all(Path(path).is_file() for path in restored["exports"].values())
    assert all_rows(restored).iloc[0]["Contact ID"] == "0001"
    with pytest.raises(ValueError):
        engine.restore_completed_run("../../other-run")


@pytest.mark.skipif(
    not hasattr(engine, "log_run_metadata"),
    reason="Supabase helper belongs to final app",
)
def test_supabase_is_opt_in_metadata_only_and_https(tmp_path, monkeypatch):
    monkeypatch.setenv("SUPABASE_URL", "https://project.supabase.co")
    monkeypatch.setenv("SUPABASE_KEY", "test-server-key")
    captured = []

    def fake_post(url, **kwargs):
        captured.append((url, kwargs))
        return SimpleNamespace(raise_for_status=lambda: None)

    monkeypatch.setattr(engine.requests, "post", fake_post)
    summary = {
        "batch_id": "batch-1",
        "created_at": "2026-10-28T12:00:00+00:00",
        "counts": {"total": 4, "valid": 2},
        "filename": "private_customer_name.csv",
    }
    engine.log_run_metadata(summary)
    assert len(captured) == 1
    url, options = captured[0]
    assert url == "https://project.supabase.co/rest/v1/verifier_runs"
    assert options["json"] == {
        "batch_id": "batch-1",
        "created_at": summary["created_at"],
        "counts": summary["counts"],
        "total_validated": 4,
    }
    assert options["timeout"] == (3, 7)
    monkeypatch.setenv("SUPABASE_URL", "http://project.supabase.co")
    with pytest.raises(ValueError, match="HTTPS"):
        engine.log_run_metadata(summary)
    assert len(captured) == 1


def test_excel_original_filename_selects_reader_for_binary_staging_path(tmp_path):
    source = tmp_path / "upload.bin"
    book = Workbook()
    book.active.append(["Email", "ID"])
    book.active.append(["alice@goodmail.com", "0009"])
    book.save(source)
    info = engine.inspect_input(source, filename="leads.xlsx")
    assert info["email_index"] == 0
    result = engine.run_verification(
        source, tmp_path / "result", 0, filename="leads.xlsx", domain_lookup=good_lookup
    )
    assert all_rows(result).iloc[0]["ID"] == "0009"
