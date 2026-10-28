"""Streamlit rerun regressions using synthetic data and no external services."""

from __future__ import annotations

import csv
import importlib.util
import io
import os
from pathlib import Path
import sys
import tempfile

import pytest
import streamlit as st
from streamlit.testing.v1 import AppTest

APP_PATH = Path(__file__).resolve().parents[1] / "app.py"
spec = importlib.util.spec_from_file_location("verifier_ui_helpers", APP_PATH)
app = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = app
spec.loader.exec_module(app)


@pytest.fixture(autouse=True)
def private_test_environment(tmp_path, monkeypatch):
    # Keep every session's workspace within pytest's own temporary tree.
    monkeypatch.setenv("TEMP", str(tmp_path))
    monkeypatch.setenv("TMP", str(tmp_path))
    monkeypatch.setattr(tempfile, "tempdir", str(tmp_path))
    monkeypatch.delenv("VERIFIER_DATA_DIR", raising=False)
    monkeypatch.delenv("SUPABASE_URL", raising=False)
    monkeypatch.delenv("SUPABASE_KEY", raising=False)


@pytest.fixture
def completed_run(tmp_path):
    source = tmp_path / "leads.csv"
    with source.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["Email", "ID", "Company"])
        writer.writerows(
            [
                ("alice@goodmail.com", "0001", "Alpha"),
                ("bob@goodmail.com", "0002", "Beta"),
                ("sales@goodmail.com", "0003", "Sales team"),
                ("malformed", "0004", "Invalid record"),
            ]
        )
    return app.run_verification(
        source,
        tmp_path / "results",
        0,
        domain_lookup=lambda domain: app.DomainResult(
            "Valid", "MX records present", "MX"
        ),
    )


def start_with_result(summary):
    at = AppTest.from_file(str(APP_PATH), default_timeout=30)
    at.session_state["result"] = summary
    at.session_state["restored"] = True
    at.run()
    assert not at.exception
    return at


def assert_report_retained(at, batch_id):
    assert not at.exception
    assert at.session_state["result"]["batch_id"] == batch_id
    assert len(at.get("download_button")) >= 4
    assert len(at.success) >= 1


def test_initial_app_renders_without_exceptions():
    at = AppTest.from_file(str(APP_PATH), default_timeout=30).run()
    assert not at.exception
    assert len(at.select_slider) == 1
    assert at.select_slider(key="workers").value == 35
    assert at.selectbox(key="batch_size").value == 2000
    assert not at.get("download_button")


def test_results_survive_filters_and_configuration_reruns(completed_run):
    at = start_with_result(completed_run)
    batch = completed_run["batch_id"]
    assert len(at.dataframe) == 1
    assert len(at.dataframe[0].value) == 4

    at.selectbox(key="result_status").select("Risky").run()
    assert_report_retained(at, batch)
    assert at.dataframe[0].value["Email"].tolist() == ["sales@goodmail.com"]

    at.text_input(key="result_search").input("no matching company").run()
    assert_report_retained(at, batch)
    assert not at.dataframe
    assert any("No rows match" in element.value for element in at.info)

    at.text_input(key="result_search").input("").run()
    at.selectbox(key="result_status").select("All").run()
    at.selectbox(key="page_size").select(50).run()
    at.select_slider(key="workers").set_value(50).run()
    at.selectbox(key="batch_size").select(5000).run()
    assert_report_retained(at, batch)
    assert len(at.dataframe[0].value) == 4
    assert at.session_state["result"]["counts"] == completed_run["counts"]


def test_excel_preparation_and_later_reruns_preserve_downloads(completed_run):
    at = start_with_result(completed_run)
    assert len(at.get("download_button")) == 4
    at.button(key="prepare_xlsx").click().run()
    assert_report_retained(at, completed_run["batch_id"])
    assert len(at.get("download_button")) == 5
    export = Path(completed_run["db_path"]).parent / "openhousefx_full_report.xlsx"
    assert export.is_file()
    at.selectbox(key="result_status").select("Valid").run()
    assert_report_retained(at, completed_run["batch_id"])
    assert len(at.get("download_button")) == 5
    assert len(at.dataframe[0].value) == 2


def test_exports_are_not_read_into_memory_on_normal_reruns(completed_run, monkeypatch):
    paths = {Path(value).resolve() for value in completed_run["exports"].values()}
    original_read = Path.read_bytes

    def guard_export_read(path):
        if path.resolve() in paths:
            pytest.fail("An export was read before the user requested a download")
        return original_read(path)

    monkeypatch.setattr(Path, "read_bytes", guard_export_read)
    at = start_with_result(completed_run)
    at.selectbox(key="result_status").select("Invalid").run()
    assert_report_retained(at, completed_run["batch_id"])
    assert len(at.get("download_button")) == 4


def test_unmapped_upload_requires_explicit_email_column(monkeypatch):
    upload = io.BytesIO(b"Name,RecipientAddress\nAlice,alice@goodmail.com\n")
    upload.name = "unmapped.csv"
    # AppTest does not expose a file-uploader input API. Substitute only the
    # uploaded object; production parsing/mapping widgets still run unchanged.
    monkeypatch.setattr(st, "file_uploader", lambda *args, **kwargs: upload)
    at = AppTest.from_file(str(APP_PATH), default_timeout=30).run()
    assert not at.exception
    mapping = next(
        element for element in at.selectbox if element.label == "Email column"
    )
    assert mapping.value is None
    verify = next(
        element for element in at.button if element.label == "Verify lead list"
    )
    assert verify.disabled
    mapping.select(1).run()
    assert not at.exception
    mapping = next(
        element for element in at.selectbox if element.label == "Email column"
    )
    assert mapping.value == 1
    verify = next(
        element for element in at.button if element.label == "Verify lead list"
    )
    assert not verify.disabled


@pytest.mark.parametrize("damage", ["missing_export", "corrupt_database"])
def test_restore_form_rejects_incomplete_saved_artifacts_gracefully(
    completed_run, tmp_path, monkeypatch, damage
):
    monkeypatch.setenv("VERIFIER_DATA_DIR", str(tmp_path / "saved"))
    token = app.save_completed_run(completed_run)
    directory = tmp_path / "saved" / token
    if damage == "missing_export":
        (directory / Path(completed_run["exports"]["valid_csv"]).name).unlink()
    else:
        (directory / Path(completed_run["db_path"]).name).write_bytes(
            b"corrupt database"
        )
    at = AppTest.from_file(str(APP_PATH), default_timeout=30).run()
    recovery = next(
        element for element in at.text_input if element.label == "Recovery code"
    )
    recovery.input(token)
    submit = next(
        element for element in at.button if element.label == "Restore saved run"
    )
    submit.click().run()
    assert not at.exception
    assert any("No completed run found" in element.value for element in at.error)
    assert not at.get("download_button")
