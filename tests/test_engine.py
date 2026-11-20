"""End-to-end engine tests with deterministic (injected) DNS, SMTP, RDAP and Twilio results."""

from __future__ import annotations

import csv
import os
from datetime import date
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from verifier import (
    DomainResult,
    MailboxResult,
    VerifyOptions,
    distinct_values,
    export_view_csv,
    generate_organized_xlsx,
    generate_xlsx,
    group_counts,
    inspect_input,
    load_summary,
    read_results,
    run_verification,
)
from verifier import smtp_check
from verifier.phone import LineResult

TODAY = "2026-11-20"
DISPOSABLE = frozenset({"mailinator.com", "throwaway.com"})


def write_csv(path, headers, rows, encoding="utf-8-sig"):
    with Path(path).open("w", encoding=encoding, newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(headers)
        writer.writerows(rows)
    return Path(path)


def dns_lookup(domain):
    if domain in {"badmail.com", "gmial.com"}:
        return DomainResult("Invalid", "Email domain does not exist", "NXDOMAIN")
    if domain == "slowmail.com":
        return DomainResult("Risky", "DNS lookup timed out; retry later", "Timeout")
    if domain == "webonly.com":
        return DomainResult("Risky", "No MX record; A fallback", "A", ("webonly.com",))
    return DomainResult("Valid", "Domain has working mail servers (MX)", "MX", ("mx." + domain,))


def mailbox_lookup(domain, hosts, addresses):
    if domain == "catchall.com":
        return True, {a: MailboxResult(smtp_check.CATCH_ALL, "") for a in addresses}
    results = {}
    for address in addresses:
        if address.startswith("ghost"):
            results[address] = MailboxResult(smtp_check.NOT_EXISTS, "550 5.1.1 User unknown")
        elif address.startswith("full"):
            results[address] = MailboxResult(smtp_check.FULL, "552 Mailbox full")
        elif address.startswith("blocked"):
            results[address] = MailboxResult(smtp_check.UNVERIFIABLE, "554 5.7.1 blocked")
        else:
            results[address] = MailboxResult(smtp_check.EXISTS, "250 OK")
    return False, results


def domain_age(domain):
    return (date(2010, 6, 1), "OK") if domain != "newco.com" else (date(2026, 9, 1), "OK")


HEADERS = ["First Name", "Last Name", "Email", "Phone", "Country", "State", "DOB", "Created", "Last Activity", "Company", "Note"]
ROWS = [
    ("alice", "ADAMS", "alice@acme.com", "(202) 555-0182", "USA", "ca", "1985-04-12", "2021-03-01", "2026-11-01", "Acme", "0001"),
    ("Bob", "Brown", "ghost@acme.com", "+44 20 7946 0958", "UK", "", "", "", "2024-01-01", "Acme", "0002"),
    ("Dup", "Alice", " ALICE@acme.com ", "", "United States", "California", "", "", "", "", "0003"),
    ("Sales", "Team", "sales@acme.com", "", "", "", "", "", "", "", "0004"),
    ("Temp", "User", "x1@mailinator.com", "", "", "", "", "", "", "", "0005"),
    ("Bad", "Syntax", "two@@acme.com", "", "", "", "", "", "", "", "0006"),
    ("No", "Domain", "user@badmail.com", "+61 2 9374 4000", "Australia", "NSW", "", "", "", "", "0007"),
    ("Slow", "Dns", "user@slowmail.com", "", "", "", "", "", "", "", "0008"),
    ("Catch", "All", "who@catchall.com", "", "", "", "", "", "", "", "0009"),
    ("Typo", "Person", "jo@gmial.com", "", "", "", "", "", "", "", "0010"),
    ("Full", "Box", "full@acme.com", "", "", "", "", "", "", "", "0011"),
    ("Blocked", "Check", "blocked@acme.com", "", "", "", "", "", "", "", "0012"),
    ("Phone", "Only", "", "+1 212 736 3100", "US", "NY", "", "", "", "", "0013"),
    ("No", "Reply", "noreply@acme.com", "", "", "", "", "", "", "", "0014"),
    ("Web", "Only", "hi.there@webonly.com", "", "", "", "", "", "", "", "0015"),
]


@pytest.fixture
def run(tmp_path):
    source = write_csv(tmp_path / "leads.csv", HEADERS, ROWS)
    info = inspect_input(source)
    calls = {"dns": [], "smtp": [], "age": []}

    def counting_dns(domain):
        calls["dns"].append(domain)
        return dns_lookup(domain)

    def counting_smtp(domain, hosts, addresses):
        calls["smtp"].append((domain, tuple(hosts), tuple(addresses)))
        return mailbox_lookup(domain, hosts, addresses)

    def counting_age(domain):
        calls["age"].append(domain)
        return domain_age(domain)

    summary = run_verification(
        source, tmp_path / "out", mapping=info["detected"], batch_size=4,
        disposable_domains=DISPOSABLE,
        options=VerifyOptions(smtp_check=True, domain_age=True, today=TODAY),
        domain_lookup=counting_dns, mailbox_lookup=counting_smtp, domain_age_lookup=counting_age,
    )
    frame, total = read_results(summary["db_path"], limit=1000)
    assert total == len(ROWS)
    rows = {row["Note"]: row for row in frame.to_dict("records")}
    return summary, rows, calls


def test_verdicts_and_email_statuses(run):
    summary, rows, _ = run
    expect = {
        "0001": ("Good", "Valid"), "0002": ("Review", "Invalid"), "0003": ("Bad", "Duplicate"),
        "0004": ("Review", "Risky"), "0005": ("Bad", "Risky"), "0006": ("Bad", "Invalid"),
        "0007": ("Review", "Invalid"), "0008": ("Review", "Risky"), "0009": ("Review", "Risky"),
        "0010": ("Bad", "Invalid"), "0011": ("Review", "Risky"), "0012": ("Good", "Valid"),
        "0013": ("Review", "Invalid"), "0014": ("Bad", "Invalid"), "0015": ("Review", "Risky"),
    }
    for note, (verdict, status) in expect.items():
        assert (rows[note]["Lead_Verdict"], rows[note]["Verification_Status"]) == (verdict, status), note
    assert rows["0001"]["Mailbox_Check"] == smtp_check.EXISTS
    assert "Mailbox does not exist" in rows["0002"]["Reason"]
    assert rows["0009"]["Mailbox_Check"] == smtp_check.CATCH_ALL
    assert "did you mean jo@gmail.com" in rows["0010"]["Reason"]
    assert rows["0010"]["Typo_Suggestion"] == "jo@gmail.com"
    assert rows["0012"]["Mailbox_Check"].startswith("Unverifiable")
    assert rows["0013"]["Lead_Issues"].startswith("No email address")
    assert rows["0001"]["Lead_Score"] == 100 - 0 or rows["0001"]["Lead_Score"] >= 90
    assert rows["0003"]["Lead_Score"] == 0
    counts = summary["counts"]
    assert counts["total"] == 15
    assert counts["good"] + counts["review"] + counts["bad"] == 15
    assert counts["duplicates"] == 1 and counts["mailbox_exists"] >= 2 and counts["catch_all"] == 1


def test_each_domain_and_mailbox_checked_once(run):
    _, _, calls = run
    dns_calls = [d for d in calls["dns"] if d != "slowmail.com"]
    assert len(dns_calls) == len(set(dns_calls))
    assert calls["dns"].count("slowmail.com") == 2  # one retry pass for timeouts
    probed = [a for _, _, addresses in calls["smtp"] for a in addresses]
    assert len(probed) == len(set(probed))
    assert "x1@mailinator.com" not in probed and "noreply@acme.com" not in probed
    assert all(hosts == ("mx.acme.com",) for domain, hosts, _ in calls["smtp"] if domain == "acme.com")


def test_people_location_and_activity(run):
    _, rows, calls = run
    alice = rows["0001"]
    assert alice["Lead_Name"] == "Alice Adams"
    assert (alice["Lead_Country"], alice["Lead_State"]) == ("United States", "California")
    assert alice["Age"] == 41
    assert alice["Record_Age"] == "5 years 8 mo" and alice["Record_Age_Days"] == 2090
    assert alice["Activity_Status"] == "Active"
    assert rows["0002"]["Activity_Status"] == "Inactive"
    assert rows["0002"]["Lead_Country"] == "United Kingdom"
    assert rows["0007"]["Lead_State"] == "New South Wales"
    assert alice["Phone_E164"] == "+12025550182" and alice["Phone_Status"] == "Valid"
    assert rows["0002"]["Phone_E164"] == "+442079460958"
    assert alice["Domain_Age_Years"] == pytest.approx(16.5, abs=0.1)
    assert "acme.com" in calls["age"] and "gmial.com" not in calls["age"]
    assert calls["age"].count("acme.com") == 1


def test_original_values_and_column_order_preserved(run, tmp_path):
    summary, rows, _ = run
    assert rows["0003"]["Email"] == " ALICE@acme.com "
    assert rows["0001"]["Note"] == "0001"  # leading zeros kept
    with open(summary["exports"]["full_csv"], encoding="utf-8-sig", newline="") as handle:
        reader = csv.reader(handle)
        header = next(reader)
        assert header[: len(HEADERS)] == HEADERS
        assert header[len(HEADERS)] == "Lead_Verdict"
        assert len(list(reader)) == len(ROWS)


def test_category_exports_partition_the_list(run):
    summary, _, _ = run

    def count(key):
        with open(summary["exports"][key], encoding="utf-8-sig", newline="") as handle:
            return sum(1 for _ in csv.DictReader(handle))

    counts = summary["counts"]
    assert count("good_csv") == counts["good"]
    assert count("review_csv") == counts["review"]
    assert count("bad_csv") == counts["bad"]
    assert count("valid_csv") == counts["valid"]
    assert count("full_csv") == counts["total"]


def test_sorting_filtering_grouping(run, tmp_path):
    summary, _, _ = run
    db = summary["db_path"]
    frame, _ = read_results(db, sort_by="Lead_Name", limit=100)
    names = [n for n in frame["Lead_Name"] if n]
    assert names == sorted(names, key=str.casefold)
    frame, _ = read_results(db, sort_by="Lead_Score", descending=True, limit=100)
    assert list(frame["Lead_Score"]) == sorted(frame["Lead_Score"], reverse=True)
    frame, total = read_results(db, filters={"Lead_Country": "United States"}, limit=100)
    assert total == 3 and set(frame["Lead_Country"]) == {"United States"}
    frame, total = read_results(db, status="Good", sort_by="Last Name", limit=100)
    assert total == summary["counts"]["good"]
    _, total = read_results(db, search="california")
    assert total == 2
    _, total = read_results(db, filters={"Lead_State": "(blank)"})
    assert total == 11
    groups = group_counts(db, "Lead_Country")
    assert groups["Leads"].sum() == 15
    us = groups[groups["Lead_Country"] == "United States"].iloc[0]
    assert (us["Leads"], us["Good"] + us["Review"] + us["Bad"]) == (3, 3)
    assert "United Kingdom" in distinct_values(db, "Lead_Country")
    view = export_view_csv(db, tmp_path / "view.csv", status="Bad", sort_by="Lead_Name")
    with open(view, encoding="utf-8-sig") as handle:
        assert sum(1 for _ in csv.DictReader(handle)) == summary["counts"]["bad"]
    with pytest.raises(ValueError):
        read_results(db, sort_by="Lead_Name; DROP TABLE results")


def test_excel_reports(run, tmp_path):
    summary, _, _ = run
    path = generate_xlsx(summary["db_path"], tmp_path / "full.xlsx")
    workbook = load_workbook(path)
    assert workbook.sheetnames == ["Summary", "All leads"]
    assert workbook["All leads"].max_row == len(ROWS) + 1
    organized = generate_organized_xlsx(summary["db_path"], tmp_path / "org.xlsx", "Lead_Country",
                                        sort_by="Lead_Name")
    workbook = load_workbook(organized)
    assert workbook.sheetnames[0] == "Summary"
    assert {"United States", "United Kingdom", "Australia", "(blank)"} <= set(workbook.sheetnames)
    total = sum(workbook[name].max_row - 1 for name in workbook.sheetnames[1:])
    assert total == len(ROWS)
    us = [row for row in workbook["United States"].iter_rows(min_row=2, values_only=True)]
    assert len(us) == 3


def test_formula_text_stays_literal_in_excel(tmp_path):
    source = write_csv(tmp_path / "f.csv", ["Email", "Note"], [("a@acme.com", "=HYPERLINK(\"x\")")])
    summary = run_verification(source, tmp_path / "out", 0, domain_lookup=dns_lookup, options=VerifyOptions(today=TODAY))
    path = generate_xlsx(summary["db_path"], tmp_path / "f.xlsx")
    sheet = load_workbook(path)["All leads"]
    assert sheet["B2"].value == "=HYPERLINK(\"x\")" and sheet["B2"].data_type == "s"


def test_phone_only_lists_and_duplicate_phones(tmp_path):
    source = write_csv(tmp_path / "p.csv", ["Name", "Mobile", "Country"], [
        ("A", "+1 212 736 3100", "US"), ("B", "212-736-3100", "US"), ("C", "12345", "US"),
        ("D", "", "US"), ("E", "+44 20 7946 0958", ""),
    ])
    lines = {"+12127363100": LineResult("Active", "Landline", "Carrier"), "+442079460958": LineResult("Inactive")}
    summary = run_verification(source, tmp_path / "out", mapping={"phone": 1, "name": 0, "country": 2},
                               options=VerifyOptions(line_check=True, today=TODAY),
                               line_lookup=lambda number: lines[number])
    frame, _ = read_results(summary["db_path"], limit=10)
    verdicts = dict(zip(frame["Name"], frame["Lead_Verdict"]))
    assert verdicts == {"A": "Good", "B": "Bad", "C": "Bad", "D": "Bad", "E": "Bad"}
    assert list(frame["Phone_Line_Status"])[:2] == ["Active", "Active"]
    assert frame.loc[frame["Name"] == "E", "Lead_Country"].item() == "United Kingdom"
    assert summary["checks"]["email"] is False
    assert set(frame["Verification_Status"]) == {"Not checked"}


def test_requires_email_or_phone(tmp_path):
    source = write_csv(tmp_path / "x.csv", ["Name"], [("A",)])
    with pytest.raises(ValueError, match="email column or a phone column"):
        run_verification(source, tmp_path / "out", mapping={"name": 0})


def test_smtp_skipped_when_port_blocked(tmp_path, monkeypatch):
    monkeypatch.setattr(smtp_check, "port25_available", lambda *a, **k: False)
    source = write_csv(tmp_path / "x.csv", ["Email"], [("a@acme.com",)])
    summary = run_verification(source, tmp_path / "out", 0, domain_lookup=dns_lookup,
                               options=VerifyOptions(smtp_check=True, today=TODAY))
    assert summary["checks"]["smtp"] is False and "port 25" in summary["notes"][0]
    frame, _ = read_results(summary["db_path"])
    assert (frame["Verification_Status"].item(), frame["Mailbox_Check"].item()) == ("Valid", "Not checked")


def test_xlsx_and_xls_inputs(tmp_path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Leads"
    sheet.append(["Email", "Phone", "Signup"])
    sheet.append(["a@acme.com", 2127363100, date(2020, 1, 1)])
    sheet.append([None, None, None])
    sheet.append(["b@acme.com", "+44 20 7946 0958", "2019-05-05"])
    workbook.create_sheet("Other").append(["Email"])
    path = tmp_path / "leads.xlsx"
    workbook.save(path)
    info = inspect_input(path)
    assert info["sheets"] == ["Leads", "Other"] and info["row_count"] == 2
    summary = run_verification(path, tmp_path / "out", mapping=info["detected"], domain_lookup=dns_lookup,
                               options=VerifyOptions(today=TODAY))
    frame, _ = read_results(summary["db_path"])
    assert list(frame["Phone_Status"]) == ["Valid", "Valid"]
    assert list(frame["Record_Age_Days"]) == [2515, 2756]

    xlwt = pytest.importorskip("xlwt")
    book = xlwt.Workbook()
    ws = book.add_sheet("Sheet1")
    for c, v in enumerate(["Email", "Zip"]):
        ws.write(0, c, v)
    ws.write(1, 0, "a@acme.com")
    ws.write(1, 1, "02134")
    xls = tmp_path / "leads.xls"
    book.save(str(xls))
    summary = run_verification(xls, tmp_path / "out2", 0, domain_lookup=dns_lookup, options=VerifyOptions(today=TODAY))
    frame, _ = read_results(summary["db_path"])
    assert frame["Zip"].item() == "02134"


def test_latin1_semicolon_csv(tmp_path):
    path = tmp_path / "l.csv"
    path.write_bytes("Email;Name\nzoe@acme.com;Zo\xe9\n".encode("cp1252"))
    info = inspect_input(path)
    assert info["delimiter"] == ";" and info["encoding"] == "cp1252"
    summary = run_verification(path, tmp_path / "out", 0, domain_lookup=dns_lookup, options=VerifyOptions(today=TODAY))
    frame, _ = read_results(summary["db_path"])
    assert frame["Name"].item() == "Zoé"


def test_incomplete_run_is_rejected(tmp_path):
    source = write_csv(tmp_path / "x.csv", ["Email"], [("a@acme.com",)])
    summary = run_verification(source, tmp_path / "out", 0, domain_lookup=dns_lookup, options=VerifyOptions(today=TODAY))
    assert load_summary(summary["db_path"])["batch_id"] == summary["batch_id"]
    with pytest.raises(ValueError, match="already contains a run"):
        run_verification(source, tmp_path / "out", 0, domain_lookup=dns_lookup)


def test_failing_lookups_never_sink_the_run(tmp_path):
    source = write_csv(tmp_path / "x.csv", ["Email"], [("a@acme.com",), ("b@acme.com",)])

    def broken(*args):
        raise RuntimeError("boom")

    summary = run_verification(source, tmp_path / "out", 0, domain_lookup=broken, mailbox_lookup=broken,
                               domain_age_lookup=broken,
                               options=VerifyOptions(smtp_check=True, domain_age=True, today=TODAY))
    frame, _ = read_results(summary["db_path"])
    assert set(frame["Verification_Status"]) == {"Risky"}


@pytest.mark.skipif(not os.environ.get("RUN_STRESS"), reason="set RUN_STRESS=1 for the 100k-row test")
def test_100k_rows(tmp_path):
    rows = [(f"user{i}@domain{i % 25}.com", f"+1 212 736 {3100 + i % 900:04d}") for i in range(100_000)]
    source = write_csv(tmp_path / "big.csv", ["Email", "Phone"], rows)
    calls = []

    def lookup(domain):
        calls.append(domain)
        return dns_lookup(domain)

    summary = run_verification(source, tmp_path / "out", mapping={"email": 0, "phone": 1}, batch_size=5000,
                               domain_lookup=lookup, options=VerifyOptions(today=TODAY))
    assert len(calls) == 25 and summary["counts"]["total"] == 100_000
    frame, total = read_results(summary["db_path"], limit=50, offset=99_950, sort_by="Lead_Score")
    assert total == 100_000 and len(frame) == 50


def test_strict_mode_requires_confirmed_mailbox(tmp_path):
    source = write_csv(tmp_path / "s.csv", ["Email"], [("alice@acme.com",), ("blocked@acme.com",)])
    summary = run_verification(source, tmp_path / "out", 0, domain_lookup=dns_lookup, mailbox_lookup=mailbox_lookup,
                               options=VerifyOptions(smtp_check=True, strict=True, today=TODAY))
    frame, _ = read_results(summary["db_path"])
    assert list(frame["Lead_Verdict"]) == ["Good", "Review"]
    assert "Strict mode" in frame["Lead_Issues"][1]


def test_smtp_breaker_skips_unreachable_provider(monkeypatch):
    smtp_check._UNREACHABLE.clear()
    smtp_check._FAILURES.clear()
    monkeypatch.setattr(smtp_check, "SMTP_PORT", 9)
    identity = ("h.test", "v@h.test")
    for _ in range(2):
        smtp_check.check_domain_mailboxes("a.test", ["127.0.0.1"], ["x@a.test"], identity=identity, timeout=2)
    _, results = smtp_check.check_domain_mailboxes("b.test", ["127.0.0.1"], ["y@b.test"], identity=identity)
    assert "blocked" in results["y@b.test"].detail
    smtp_check._UNREACHABLE.clear()
    smtp_check._FAILURES.clear()
