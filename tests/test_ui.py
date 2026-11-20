"""Streamlit AppTest regressions with synthetic data and no external services."""

from __future__ import annotations

import csv
import tempfile
from pathlib import Path

import pytest
from streamlit.testing.v1 import AppTest

from verifier import DomainResult, VerifyOptions, inspect_input, run_verification

APP_PATH = Path(__file__).resolve().parents[1] / "app.py"


@pytest.fixture(autouse=True)
def private_environment(tmp_path, monkeypatch):
    monkeypatch.setenv("TMPDIR", str(tmp_path))
    monkeypatch.setattr(tempfile, "tempdir", str(tmp_path))
    for name in ("VERIFIER_DATA_DIR", "SUPABASE_URL", "SUPABASE_KEY", "TWILIO_ACCOUNT_SID", "TWILIO_AUTH_TOKEN"):
        monkeypatch.delenv(name, raising=False)


@pytest.fixture
def completed_run(tmp_path):
    source = tmp_path / "leads.csv"
    with source.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["Name", "Email", "Phone", "Country", "State", "Last Activity"])
        writer.writerows([
            ("Zed Young", "zed@goodmail.com", "+1 212 736 3100", "US", "NY", "2026-11-01"),
            ("Amy Adams", "amy@goodmail.com", "+44 20 7946 0958", "UK", "", "2023-01-01"),
            ("Sales Team", "sales@goodmail.com", "", "US", "CA", ""),
            ("Broken", "malformed", "123", "", "", ""),
        ])
    info = inspect_input(source)
    return run_verification(
        source, tmp_path / "results", mapping=info["detected"],
        domain_lookup=lambda domain: DomainResult("Valid", "MX records present", "MX", ("mx." + domain,)),
        options=VerifyOptions(today="2026-11-20"),
    )


def start(summary=None):
    at = AppTest.from_file(str(APP_PATH), default_timeout=60)
    if summary:
        at.session_state["result"] = summary
    at.run()
    assert not at.exception, at.exception
    return at


def test_initial_render():
    at = start()
    text = " ".join(m.value for m in at.markdown)
    assert "A cleaner list starts here" in text
    assert at.toggle(key="opt_smtp").value is True


def test_results_render_with_tabs_and_metrics(completed_run):
    at = start(completed_run)
    assert "Run complete" in at.success[0].value
    assert [tab.label for tab in at.tabs] == ["Leads", "Organize & group", "Export", "Run details"]
    assert len(at.dataframe) >= 2  # leads table + group table


def test_sorting_and_filtering_keep_state(completed_run):
    at = start(completed_run)
    at.selectbox(key="leads_sort").set_value("Lead_Name").run()
    assert not at.exception
    frame = at.dataframe[0].value
    names = list(frame["Lead_Name"])
    assert names == sorted(names, key=str.casefold)
    at.selectbox(key="leads_country").set_value("United States").run()
    assert not at.exception
    assert set(at.dataframe[0].value["Lead_Country"]) == {"United States"}
    at.selectbox(key="leads_status").set_value("Good").run()
    assert not at.exception
    assert set(at.dataframe[0].value["Lead_Verdict"]) <= {"Good"}
    assert at.session_state["result"]["batch_id"] == completed_run["batch_id"]


def test_organize_group_by_state(completed_run):
    at = start(completed_run)
    at.selectbox(key="org_group").set_value("Lead_State").run()
    assert not at.exception
    groups = at.dataframe[1].value
    assert int(groups["Leads"].sum()) == 4
    assert "New York" in set(groups["Lead_State"])


def test_restore_rejects_unknown_code(tmp_path, monkeypatch):
    monkeypatch.setenv("VERIFIER_DATA_DIR", str(tmp_path / "saved"))
    at = start()
    at.text_input[0].input("00000000-0000-0000-0000-000000000000")
    at.button[0].click().run()
    assert any("No completed run found" in e.value for e in at.error)
