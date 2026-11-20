"""OpenhouseFX Lead Verifier — Streamlit UI. The engine lives in the `verifier` package."""

from __future__ import annotations

import base64
import hashlib
import html
import json
import math
import os
import shutil
import sqlite3
import tempfile
import threading
import time
import uuid
from functools import lru_cache
from pathlib import Path
from urllib.parse import urlparse

import requests
import streamlit as st

from verifier import (
    FIELD_LABELS,
    RESULT_COLUMNS,
    DomainResult,
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
from verifier import __version__ as APP_VERSION
from verifier.enrich import country_name
from verifier.phone import supported_regions, twilio_ready

APP_ROOT = Path(__file__).resolve().parent
_file_lock = threading.Lock()

# Re-exported for scripts/tests that import the app module directly.
__all__ = ["run_verification", "read_results", "DomainResult", "generate_xlsx", "main"]


# --------------------------------------------------------------------------- styling
@lru_cache(maxsize=1)
def brand_css():
    """Embed the original Pocket Studio fonts, with no third-party font calls."""
    faces = []
    for family, filename in [("Inter", "inter-latin.woff2"), ("Plus Jakarta Sans", "jakarta-latin.woff2")]:
        encoded = base64.b64encode((APP_ROOT / "static" / filename).read_bytes()).decode("ascii")
        faces.append(
            f"@font-face{{font-family:'{family}';src:url(data:font/woff2;base64,{encoded}) format('woff2');"
            "font-weight:100 900;font-style:normal;font-display:swap;}"
        )
    return "<style>" + "".join(faces) + """
    :root {--navy:#061521;--card:#0B1F30;--panel:#0F2A40;--gold:#D3AA4E;--text:#F0F4F8;--muted:#94A3B8;--line:rgba(211,170,78,.19);--good:#14B8A6;--review:#D3AA4E;--bad:#F87171}
    html,body,[data-testid="stApp"] {font-family:Inter,sans-serif;color:var(--text);background:var(--navy)}
    h1,h2,h3,h4 {font-family:'Plus Jakarta Sans',Inter,sans-serif!important;letter-spacing:-.04em}
    [data-testid="stHeader"] {background:rgba(6,21,33,.94);border-bottom:1px solid rgba(211,170,78,.08)}
    [data-testid="stSidebar"] {background:var(--card);border-right:1px solid var(--line);min-width:300px;max-width:300px}
    [data-testid="stSidebarUserContent"] {padding:1.5rem 1.4rem 2rem}
    [data-testid="stMainBlockContainer"] {max-width:1500px;padding:2rem 3.2rem}
    [data-testid="stVerticalBlock"] {gap:1rem}
    [data-testid="stSidebar"] [data-testid="stVerticalBlock"] {gap:.8rem}
    .brand {display:flex;align-items:center;gap:12px;padding:8px 0 26px;border-bottom:1px solid var(--line);margin-bottom:16px}
    .brand svg {width:29px;height:36px;flex:none}
    .wordmark {font:600 16px/1.3 Inter,sans-serif;letter-spacing:1.7px;white-space:nowrap}
    .wordmark span {color:var(--gold)}
    .product {font:500 7.7px/1.5 Inter,sans-serif;letter-spacing:3px;color:var(--gold);margin-top:5px}
    .eyebrow {color:var(--gold);font:600 9px/1.5 Inter,sans-serif;letter-spacing:2.2px;margin:8px 0 10px;text-transform:uppercase}
    .sidebar-title {font:600 18px/1.4 'Plus Jakarta Sans',sans-serif;letter-spacing:-.6px;margin-bottom:6px}
    .small-copy {font:400 11px/1.8 Inter,sans-serif;color:var(--muted)}
    .crumbs {display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid rgba(211,170,78,.10);padding:0 0 22px;margin-bottom:30px;gap:12px}
    .crumbs-label {font-size:11px;color:var(--muted)}
    .crumbs-label b {color:#d9e2ec;font-weight:500;margin-left:16px}
    .local-pill {font-size:10px;color:var(--muted);border:1px solid rgba(148,163,184,.14);border-radius:30px;padding:7px 12px;white-space:nowrap}
    .local-pill i {display:inline-block;width:6px;height:6px;background:#14B8A6;border-radius:50%;margin-right:7px;box-shadow:0 0 9px #14b8a64d}
    .hero-title {font:500 clamp(28px,3vw,43px)/1.22 'Plus Jakarta Sans',sans-serif;letter-spacing:-.035em;word-spacing:.08em;margin:0 0 14px}
    .hero-title span {color:var(--gold)}
    .hero-copy {font:400 12px/1.8 Inter,sans-serif;color:var(--muted);margin:0 0 16px;max-width:760px}
    .section-title {font:600 19px/1.4 'Plus Jakarta Sans',sans-serif;letter-spacing:-.6px;margin:2px 0 4px}
    .section-note {font-size:11px;line-height:1.8;color:var(--muted);margin-bottom:12px}
    .metric-grid {display:grid;grid-template-columns:repeat(6,minmax(0,1fr));gap:12px;margin:12px 0 22px}
    .kpi {background:var(--card);border:1px solid var(--line);border-radius:10px;padding:18px 16px 15px;min-height:112px}
    .kpi-label {font-size:10px;font-weight:400;color:#a9bacb;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
    .kpi-value {font:600 27px/1.5 'Plus Jakarta Sans',sans-serif;color:var(--gold);letter-spacing:-.6px}
    .kpi-value.good {color:var(--good)} .kpi-value.bad {color:var(--bad)}
    .kpi-note {font-size:9px;color:var(--muted);line-height:1.5}
    .chips {display:flex;flex-wrap:wrap;gap:8px;margin:-8px 0 16px}
    .chip {font-size:10px;color:#cbd5e1;border:1px solid rgba(148,163,184,.18);border-radius:20px;padding:5px 11px;background:rgba(15,42,64,.35)}
    .chip b {color:var(--gold);font-weight:600;margin-right:4px}
    .workspace-empty {background:linear-gradient(115deg,#0B1F30,#081a29);border:1px solid var(--line);border-radius:12px;padding:34px 28px;text-align:center;margin:3px 0 18px}
    .empty-icon {display:inline-flex;align-items:center;justify-content:center;width:42px;height:42px;border:1px solid rgba(211,170,78,.25);border-radius:12px;color:var(--gold);margin-bottom:16px}
    .empty-icon svg {width:21px;height:21px}
    .empty-title {font:500 20px/1.5 'Plus Jakarta Sans',sans-serif;letter-spacing:-.6px;margin-bottom:7px}
    .empty-copy {font-size:11px;line-height:1.9;color:var(--muted);max-width:470px;margin:auto}
    .step {display:flex;gap:12px;align-items:flex-start;margin:14px 0}
    .step-num {color:var(--gold);font-size:10px;border:1px solid rgba(211,170,78,.2);border-radius:6px;padding:5px}
    .step strong {font-size:11px;font-weight:500;display:block;margin-bottom:4px}
    .step p {font-size:10px;color:var(--muted);margin:0;line-height:1.7}
    .run-label {font-size:10px;color:var(--muted);margin-top:8px;line-height:1.8;word-break:break-word}
    .footer-note {border-top:1px solid rgba(211,170,78,.10);margin-top:25px;padding-top:18px;display:flex;justify-content:space-between;font-size:9px;letter-spacing:.3px;color:var(--muted);gap:10px}
    .sidebar-footer {padding-top:20px;border-top:1px solid var(--line);margin-top:20px}
    [data-testid="stFileUploaderDropzone"] {background:rgba(15,42,64,.4);border:1px dashed rgba(211,170,78,.38);border-radius:9px;padding:20px 18px}
    [data-testid="stFileUploaderDropzone"] p {font-size:12px}
    [data-testid="stFileUploaderDropzone"] small {color:var(--muted);font-size:10px}
    [data-testid="stFileUploaderDropzone"] svg {color:var(--gold)}
    [data-testid="stVerticalBlockBorderWrapper"]>div {border-color:var(--line)!important;border-radius:12px!important}
    [data-testid="stWidgetLabel"] p {font-size:11px!important;font-weight:400!important;color:#a9bacb}
    [data-testid="stCaptionContainer"] p {font-size:10px!important;line-height:1.8;color:var(--muted)}
    [data-testid="stBaseButton-primary"] {background:linear-gradient(120deg,#D3AA4E,#c79b39)!important;color:#061521!important;border:1px solid #d3aa4e!important;min-height:42px;border-radius:8px!important;font-weight:600;box-shadow:0 4px 18px #d3aa4e1c}
    [data-testid="stBaseButton-primary"]:hover {background:#C2993D!important;border-color:#e2bc65!important}
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
    [data-baseweb="tab-list"] button p {font-size:12px}
    .stAppDeployButton {display:none}
    button:focus-visible,input:focus-visible {outline:2px solid var(--gold)!important;outline-offset:3px}
    @media(max-width:1200px){[data-testid="stMainBlockContainer"]{padding:2rem 1.5rem}.metric-grid{grid-template-columns:repeat(3,minmax(0,1fr))}}
    @media(max-width:700px){[data-testid="stMainBlockContainer"]{padding:1.4rem 1rem}.metric-grid{grid-template-columns:repeat(2,minmax(0,1fr));gap:9px}.kpi{min-height:96px}.crumbs{margin-bottom:22px}.crumbs-label b{margin-left:8px}.workspace-empty{padding:26px 18px}.footer-note{flex-direction:column}}
    </style>"""


# --------------------------------------------------------------------------- helpers
def show_metrics(counts, checks=None):
    checks = checks or {"email": True, "phone": False}
    total = counts.get("total", 0)

    def pct(key):
        return f"{counts.get(key, 0) / total:.0%} of leads" if total else ""

    cards = [
        ("Total leads", f"{total:,}", "", "Rows in this run"),
        ("Good leads", f"{counts.get('good', 0):,}", "good", pct("good") or "Ready to contact"),
        ("Needs review", f"{counts.get('review', 0):,}", "", pct("review") or "Check before use"),
        ("Bad leads", f"{counts.get('bad', 0):,}", "bad", pct("bad") or "Remove or fix"),
    ]
    if checks.get("email", True):
        cards.append(("Valid emails", f"{counts.get('valid', 0):,}",
                      "", f"{counts.get('mailbox_exists', 0):,} mailboxes confirmed"))
    if checks.get("phone"):
        cards.append(("Valid phones", f"{counts.get('phone_valid', 0):,}", "",
                      f"{counts.get('phone_invalid', 0):,} invalid · {counts.get('phone_missing', 0):,} missing"))
    elif len(cards) < 6:
        cards.append(("Duplicates", f"{counts.get('duplicates', 0):,}", "", "Repeat emails in the list"))
    st.markdown(
        '<div class="metric-grid">'
        + "".join(
            f'<div class="kpi"><div class="kpi-label">{label}</div><div class="kpi-value {cls}">{value}</div>'
            f'<div class="kpi-note">{note}</div></div>'
            for label, value, cls, note in cards
        )
        + "</div>",
        unsafe_allow_html=True,
    )


def show_chips(counts, checks):
    chips = []
    if checks.get("email"):
        chips += [("Duplicates", counts.get("duplicates", 0)), ("Role inboxes", counts.get("role", 0)),
                  ("Disposable", counts.get("disposable", 0)), ("Typos", counts.get("typos", 0)),
                  ("Free providers", counts.get("free", 0))]
        if checks.get("smtp"):
            chips += [("Mailbox missing", counts.get("mailbox_missing", 0)), ("Catch-all", counts.get("catch_all", 0)),
                      ("Unverifiable", counts.get("mailbox_unverifiable", 0))]
    if checks.get("phone"):
        chips += [("Duplicate phones", counts.get("phone_duplicates", 0))]
        if checks.get("line"):
            chips += [("Lines active", counts.get("line_active", 0)), ("Lines dead", counts.get("line_inactive", 0))]
    if checks.get("activity"):
        chips += [("Active", counts.get("active", 0)), ("Lapsing", counts.get("lapsing", 0)),
                  ("Inactive", counts.get("inactive", 0))]
    if chips:
        st.markdown(
            '<div class="chips">' + "".join(f'<span class="chip"><b>{n:,}</b>{html.escape(l)}</span>' for l, n in chips) + "</div>",
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


@lru_cache(maxsize=1)
def region_options():
    regions = supported_regions()
    names = {code: country_name(code) for code in regions}
    return sorted(regions, key=lambda code: names[code]), names


def _env_flag(name, default):
    value = os.environ.get(name, "").strip().lower()
    if not value:
        return default
    return value in {"1", "true", "yes", "on"}


# --------------------------------------------------------------------------- optional integrations
def supabase_ready():
    return bool(os.environ.get("SUPABASE_URL") and os.environ.get("SUPABASE_KEY"))


def log_run_metadata(summary):
    """Opt-in metadata only; no lead rows or filenames leave the server."""
    url = os.environ.get("SUPABASE_URL", "").rstrip("/")
    key = os.environ.get("SUPABASE_KEY", "")
    if urlparse(url).scheme != "https" or not urlparse(url).hostname or not key:
        raise ValueError("Supabase requires an HTTPS project URL and a server-side key.")
    payload = {
        "batch_id": summary["batch_id"],
        "total_validated": summary["counts"]["total"],
        "counts": summary["counts"],
        "created_at": summary["created_at"],
    }
    response = requests.post(
        url + "/rest/v1/verifier_runs",
        json=payload,
        headers={"apikey": key, "Authorization": "Bearer " + key,
                 "Prefer": "resolution=merge-duplicates,return=minimal"},
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
    saved["exports"] = {key: str(target / Path(value).name) for key, value in summary["exports"].items()}
    (target / "manifest.json").write_text(json.dumps(saved, ensure_ascii=False, indent=2), encoding="utf-8")
    return token


def restore_completed_run(token):
    token = uuid.UUID(token.strip()).hex
    directory = Path(os.environ["VERIFIER_DATA_DIR"]).expanduser().resolve() / token
    saved = json.loads((directory / "manifest.json").read_text(encoding="utf-8"))
    saved["db_path"] = str(directory / Path(saved["db_path"]).name)
    saved["exports"] = {key: str(directory / Path(value).name) for key, value in saved["exports"].items()}
    if not Path(saved["db_path"]).is_file() or not all(Path(v).is_file() for v in saved["exports"].values()):
        raise ValueError("The saved results could not be found.")
    try:
        completed = load_summary(saved["db_path"])
        if completed["batch_id"] != saved["batch_id"] or completed["counts"] != saved["counts"]:
            raise ValueError("Saved run metadata does not match its results.")
    except (sqlite3.Error, KeyError) as exc:
        raise ValueError("The saved result database is incomplete or unreadable.") from exc
    return saved


# --------------------------------------------------------------------------- sidebar
def render_sidebar():
    with st.sidebar:
        house = (APP_ROOT / "assets" / "openhousefx-house.svg").read_text(encoding="utf-8")
        st.markdown(
            f'<div class="brand">{house}<div><div class="wordmark">OPENHOUSE<span>FX</span></div>'
            '<div class="product">LEAD VERIFIER</div></div></div>',
            unsafe_allow_html=True,
        )
        st.markdown(
            '<div class="eyebrow">01 / CONFIGURE YOUR RUN</div><div class="sidebar-title">Verification settings</div>'
            '<div class="small-copy">Turn on the deep checks for the most accurate verdicts.</div>',
            unsafe_allow_html=True,
        )
        smtp = st.toggle(
            "Mailbox check (SMTP)",
            value=_env_flag("VERIFIER_SMTP_DEFAULT", True),
            key="opt_smtp",
            help="Asks each mail server whether the exact mailbox exists, without sending an email. "
            "Also detects catch-all domains. Needs outbound port 25; skipped automatically where the host blocks it.",
        )
        if smtp:
            st.caption(
                "Mailbox checks use the ZeroBounce API when port 25 is blocked."
                if os.environ.get("ZEROBOUNCE_API_KEY")
                else "Needs a server with outbound port 25 open (not Render). Skipped automatically if blocked."
            )
        strict = st.toggle(
            "Strict mode",
            value=_env_flag("VERIFIER_STRICT_DEFAULT", False),
            key="opt_strict",
            disabled=not smtp,
            help="Only emails whose mailbox the mail server confirmed can be graded Good. "
            "Valid-but-unconfirmed emails move to Review.",
        )
        domain_age = st.toggle(
            "Company domain age",
            value=_env_flag("VERIFIER_DOMAIN_AGE_DEFAULT", True),
            key="opt_age",
            help="Looks up when the lead's company website/email domain was registered (RDAP). Free providers like Gmail are skipped.",
        )
        line_ok = twilio_ready()
        line = st.toggle(
            "Live phone line status",
            value=False,
            key="opt_line",
            disabled=not line_ok,
            help="Uses Twilio Lookup to report whether each number is active, inactive or unreachable (billed per lookup by Twilio).",
        )
        if not line_ok:
            st.caption("Live line status needs TWILIO_ACCOUNT_SID and TWILIO_AUTH_TOKEN on the server. Offline phone validation always runs.")
        regions, names = region_options()
        default_region = os.environ.get("VERIFIER_DEFAULT_REGION", "US").upper()
        region = st.selectbox(
            "Default phone country",
            regions,
            index=regions.index(default_region) if default_region in regions else regions.index("US"),
            format_func=lambda code: f"{names[code]} ({code})",
            key="opt_region",
            help="Used for numbers written without a +country code when the row has no country.",
        )
        dayfirst = st.checkbox("Dates are day-first (DD/MM/YYYY)", value=False, key="opt_dayfirst")
        col1, col2 = st.columns(2)
        with col1:
            active_days = st.number_input("Active within (days)", 1, 3650, 90, key="opt_active")
        with col2:
            inactive_days = st.number_input("Inactive after (days)", 1, 7300, 365, key="opt_inactive")
        with st.expander("Speed & batch size"):
            workers = st.select_slider("Concurrent domain checks", options=[10, 20, 25, 30, 35, 40, 45, 50], value=35, key="workers")
            smtp_workers = st.select_slider("Concurrent mailbox checks", options=[2, 4, 6, 8, 10, 12, 16, 20], value=10, key="smtp_workers")
            batch_size = st.selectbox("Rows per batch", [1000, 2000, 5000], index=1, key="batch_size",
                                      format_func=lambda n: f"{n:,} rows")
        with st.expander("How leads are graded"):
            st.markdown(
                "**Good** · Email valid (mailbox confirmed or domain working) and no bad phone.\n\n"
                "**Review** · Risky email (catch-all, role inbox, typo, DNS doubt) or a bad phone next to a good email, "
                "or only the phone works.\n\n"
                "**Bad** · Invalid/duplicate email, mailbox doesn't exist, disposable or placeholder address, and no working phone.\n\n"
                "**Lead score** · 0–100 combining email (70%) and phone (30%) quality, minus points for inactive clients."
            )
        persist = log_metadata = False
        if os.environ.get("VERIFIER_DATA_DIR") or supabase_ready():
            with st.expander("Save & recovery"):
                if os.environ.get("VERIFIER_DATA_DIR"):
                    persist = st.checkbox("Save this run on the server", value=False,
                                          help="Saves lead rows and reports to the configured disk. Keep the recovery code private.")
                    with st.form("restore_run"):
                        token = st.text_input("Recovery code", type="password")
                        restore = st.form_submit_button("Restore saved run")
                    if restore:
                        try:
                            st.session_state.result = restore_completed_run(token)
                            st.session_state.result_source = None
                            st.session_state.restored = True
                            st.session_state.pop("run_messages", None)
                            st.session_state.recovery_code = uuid.UUID(token.strip()).hex
                            st.rerun()
                        except (ValueError, OSError, KeyError, json.JSONDecodeError):
                            st.error("No completed run found for that recovery code.")
                if supabase_ready():
                    log_metadata = st.checkbox("Log run totals to Supabase", value=False,
                                               help="Sends only the batch ID, timestamp and counts.")
        st.markdown(
            f'<div class="sidebar-footer"><div class="small-copy">OPENHOUSEFX · POCKET STUDIO</div>'
            f'<div class="small-copy">Lead Verifier {APP_VERSION}</div></div>',
            unsafe_allow_html=True,
        )
    options = VerifyOptions(
        smtp_check=bool(smtp), domain_age=bool(domain_age), line_check=bool(line and line_ok),
        default_region=region, dayfirst=bool(dayfirst), active_days=int(active_days),
        inactive_days=max(int(inactive_days), int(active_days)), smtp_workers=int(smtp_workers),
        strict=bool(smtp and strict),
    )
    return {"workers": workers, "batch_size": batch_size, "options": options,
            "persist": persist, "log_metadata": log_metadata}


# --------------------------------------------------------------------------- input
def stage_upload(uploaded, workspace):
    buffer = uploaded.getbuffer()
    try:
        signature = (hashlib.sha256(buffer).hexdigest(), uploaded.name)
        if st.session_state.get("upload_signature") != signature:
            target = workspace / ("source" + Path(uploaded.name).suffix.lower())
            with target.open("wb") as handle:
                handle.write(buffer)
            for key in ("result", "input_info", "sheet_choice", "restored", "run_messages", "recovery_code"):
                st.session_state.pop(key, None)
            st.session_state.upload_signature = signature
            st.session_state.input_path = str(target)
        return signature
    finally:
        buffer.release()


def _column_picker(label, field, columns, default, key):
    options = [None] + list(range(len(columns)))
    return st.selectbox(
        label,
        options,
        index=(int(default) + 1 if default is not None else 0),
        format_func=lambda i: "— Not in this file —" if i is None else columns[i],
        key=key,
    )


def render_input(workspace):
    st.markdown(
        '<div class="section-title">Your source file</div><div class="section-note">Upload a lead sheet. '
        "Every original column stays in every report.</div>",
        unsafe_allow_html=True,
    )
    uploaded = st.file_uploader(
        "Upload a lead sheet", type=["csv", "xlsx", "xls"], key="source_upload", label_visibility="collapsed",
        help="CSV, XLSX or XLS up to 50 MB. The first row must be the column headers.",
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
                st.session_state.input_info = inspect_input(path, filename=uploaded.name)
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
        detected = info.get("detected", {})
        prefix = "map_" + signature[0][:12] + "_" + str(sheet) + "_"
        left, middle, right = st.columns([1.2, 1.2, 1])
        with left:
            email_index = _column_picker("Email column", "email", columns, detected.get("email"), prefix + "email")
        with middle:
            phone_index = _column_picker("Phone column", "phone", columns, detected.get("phone"), prefix + "phone")
        with right:
            st.markdown(
                f'<div class="run-label">FILE READY<br><span style="color:#F0F4F8">{html.escape(uploaded.name)}</span>'
                f'<br>{info.get("row_count", 0):,} rows · {len(columns)} columns</div>',
                unsafe_allow_html=True,
            )
        mapping = {"email": email_index, "phone": phone_index}
        extra = [field for field in FIELD_LABELS if field not in ("email", "phone")]
        found = sum(1 for field in extra if detected.get(field) is not None)
        with st.expander(f"More columns · {found} auto-detected (name, location, dates, company)", expanded=False):
            st.caption("Map these to unlock name sorting, country/state grouping, client age and activity status.")
            grid = st.columns(3)
            for position, field in enumerate(extra):
                with grid[position % 3]:
                    mapping[field] = _column_picker(FIELD_LABELS[field], field, columns, detected.get(field), prefix + field)
        if email_index is None and phone_index is None:
            st.caption("Choose an email column or a phone column to continue.")
        return {
            "path": path, "filename": uploaded.name, "sheet": sheet, "mapping": mapping,
            "signature": (signature, sheet, tuple(sorted((k, v) for k, v in mapping.items() if v is not None))),
            "info": info,
        }
    except Exception as exc:
        st.error(f"Could not read this file: {exc}")
        return None


def execute_run(source, settings, metrics_slot):
    run_dir = get_workspace() / uuid.uuid4().hex
    progress_bar = st.progress(0, text="Preparing your list…")
    detail = st.empty()
    last_update = [0.0]
    stages = {
        "Checking email domains": "Checking email domains",
        "Checking mailboxes": "Asking mail servers if each mailbox exists",
        "Looking up company domain ages": "Looking up company domain ages",
        "Checking phone lines": "Checking phone lines",
        "Verifying rows": "Grading leads",
    }

    def report(event):
        now = time.monotonic()
        stage = event.get("stage", "")
        if now - last_update[0] < 0.2 and stage not in {"Complete", "Preparing exports"}:
            return
        last_update[0] = now
        done = event.get("rows_processed", 0)
        total = event.get("total_rows") or source["info"].get("row_count") or 0
        fraction = min(done / total, 1) if total else 0
        eta = event.get("eta")
        eta_text = f" · about {max(1, round(eta)):,}s left" if eta is not None and done < total else ""
        label = "Preparing your exports…" if stage == "Preparing exports" else (
            f"{fraction:.0%} · {done:,} of {total:,} rows{eta_text}"
        )
        progress_bar.progress(fraction, text=label)
        note = event.get("note") or ""
        detail.caption(
            f'{stages.get(stage, stage)} · {event.get("domains_checked", 0):,} domains checked · '
            f'{event.get("elapsed", 0):.1f}s elapsed' + (f" · {note}" if note else "")
        )
        if event.get("counts"):
            with metrics_slot.container():
                show_metrics(event["counts"], {"email": source["mapping"].get("email") is not None,
                                               "phone": source["mapping"].get("phone") is not None})

    try:
        summary = run_verification(
            source["path"], run_dir,
            mapping=source["mapping"],
            workers=settings["workers"],
            batch_size=settings["batch_size"],
            sheet_name=source["sheet"],
            filename=source["filename"],
            disposable_domains=load_disposable_domains(),
            options=settings["options"],
            progress=report,
        )
    except Exception as exc:
        progress_bar.empty()
        detail.empty()
        st.error(f"Verification could not finish: {exc}. Any previous completed report remains available.")
        return
    summary["source_name"] = source["filename"]
    st.session_state.result = summary
    st.session_state.result_source = source["signature"]
    st.session_state.pop("restored", None)
    st.session_state.pop("recovery_code", None)
    messages = []
    if settings["persist"]:
        try:
            st.session_state.recovery_code = save_completed_run(summary)
            messages.append("Saved on the configured server disk.")
        except (OSError, ValueError):
            messages.append("Server save failed; download your reports before leaving this session.")
    if settings["log_metadata"]:
        try:
            log_run_metadata(summary)
            messages.append("Run totals logged to Supabase.")
        except (requests.RequestException, ValueError):
            messages.append("Supabase logging failed; your verification results and downloads are ready.")
    st.session_state.run_messages = messages
    progress_bar.empty()
    detail.empty()
    st.rerun()


# --------------------------------------------------------------------------- results
def download_file(path):
    return Path(path).read_bytes()


def _relevant_columns(summary):
    checks = summary.get("checks", {"email": True})
    mapping = summary.get("mapping", {})
    columns = ["Lead_Verdict", "Lead_Score", "Lead_Issues"]
    if mapping.get("name") or mapping.get("first_name") or mapping.get("last_name"):
        columns.append("Lead_Name")
    if mapping.get("company"):
        columns.append("Lead_Company")
    columns += ["Lead_Country", "Lead_State"]
    if checks.get("email"):
        columns += ["Verification_Status", "Reason", "Mailbox_Check", "Normalized_Email", "Email_Domain",
                    "DNS_Check", "MX_Host", "Is_Role", "Is_Disposable", "Is_Free_Provider", "Is_Duplicate",
                    "Typo_Suggestion"]
    if checks.get("phone"):
        columns += ["Phone_Status", "Phone_E164", "Phone_Type", "Phone_Country", "Phone_Carrier",
                    "Phone_Location", "Phone_Line_Status", "Is_Duplicate_Phone"]
    if checks.get("dob"):
        columns.append("Age")
    if checks.get("created"):
        columns += ["Record_Age_Days", "Record_Age"]
    if checks.get("activity"):
        columns += ["Last_Activity_Days", "Activity_Status"]
    if checks.get("domain_age"):
        columns += ["Domain_Created", "Domain_Age_Years"]
    return [column for column in RESULT_COLUMNS if column in columns]


_PRETTY = {
    "Lead_Verdict": "Verdict", "Lead_Score": "Score", "Lead_Issues": "Issues", "Lead_Name": "Name",
    "Lead_Company": "Company", "Lead_Country": "Country", "Lead_State": "State",
    "Verification_Status": "Email status", "Mailbox_Check": "Mailbox", "Phone_Status": "Phone status",
    "Phone_E164": "Phone (intl.)", "Phone_Type": "Phone type", "Phone_Line_Status": "Line status",
    "Record_Age": "Lead age", "Record_Age_Days": "Lead age (days)", "Last_Activity_Days": "Days since activity",
    "Activity_Status": "Activity", "Domain_Age_Years": "Domain age (yrs)", "Domain_Created": "Domain registered",
    "Age": "Age (yrs)",
}


def _label(column):
    return _PRETTY.get(column, column.replace("_", " "))


def _choice_label(column):
    """Menu label; marks sheet columns whose name matches a verified column (e.g. 'State')."""
    if column is None:
        return "Original order"
    if column in RESULT_COLUMNS:
        label = _label(column)
        return label + " (standardized)" if column in ("Lead_Country", "Lead_State", "Lead_Name", "Lead_Company") else label
    return f"{column} (as in your sheet)"


def _sort_choices(summary):
    preferred = [c for c in ("Lead_Name", "Lead_Country", "Lead_State", "Lead_Score", "Lead_Verdict", "Lead_Company",
                             "Age", "Record_Age_Days", "Last_Activity_Days", "Domain_Age_Years",
                             "Verification_Status", "Phone_Status", "Email_Domain", "Activity_Status")
                 if c in _relevant_columns(summary)]
    return [None] + preferred + summary["original_columns"]


def _group_choices(summary):
    preferred = [c for c in ("Lead_Country", "Lead_State", "Lead_Verdict", "Verification_Status", "Activity_Status",
                             "Phone_Type", "Phone_Country", "Phone_Status", "Email_Domain", "Lead_Company",
                             "Mailbox_Check")
                 if c in _relevant_columns(summary)]
    return preferred + summary["original_columns"]


@st.cache_data(show_spinner=False, max_entries=64)
def _distinct(db_path, batch_id, column):
    return distinct_values(db_path, column, limit=400)


def _filter_widgets(summary, prefix):
    checks = summary.get("checks", {})
    statuses = ["All", "Good", "Review", "Bad"]
    if checks.get("email"):
        statuses += ["Valid", "Risky", "Invalid", "Duplicate", "Role / Disposable"]
        if checks.get("smtp"):
            statuses += ["Mailbox confirmed", "Catch-all"]
    if checks.get("phone"):
        statuses += ["Phone valid", "Phone invalid"]
    if checks.get("activity"):
        statuses += ["Active", "Lapsing", "Inactive"]
    a, b, c, d = st.columns([1, 1, 1, 1.6])
    with a:
        status = st.selectbox("Show", statuses, key=prefix + "status")
    with b:
        countries = ["All"] + _distinct(summary["db_path"], summary["batch_id"], "Lead_Country")
        country = st.selectbox("Country", countries, key=prefix + "country")
    with c:
        states = ["All"] + _distinct(summary["db_path"], summary["batch_id"], "Lead_State")
        state = st.selectbox("State / province", states, key=prefix + "state")
    with d:
        search = st.text_input("Search", placeholder="Name, email, company, reason…", key=prefix + "search")
    filters = {"Lead_Country": country, "Lead_State": state}
    return status, search, filters


def render_leads_tab(summary):
    status, search, filters = _filter_widgets(summary, "leads_")
    sort_options = _sort_choices(summary)
    a, b, c, d = st.columns([1.4, 1, 0.8, 0.8])
    with a:
        sort_by = st.selectbox("Sort by", sort_options, key="leads_sort", format_func=_choice_label)
    with b:
        descending = st.selectbox("Order", ["Ascending (A→Z, low→high)", "Descending (Z→A, high→low)"],
                                  key="leads_order").startswith("Desc")
    with c:
        limit = st.selectbox("Rows per page", [50, 100, 250, 500], index=1, key="page_size")
    filter_key = (summary["batch_id"], status, search, json.dumps(filters), sort_by, descending, limit)
    if st.session_state.get("last_filter") != filter_key:
        st.session_state.report_page = 1
        st.session_state.last_filter = filter_key
    query = {"status": status, "search": search, "filters": filters, "sort_by": sort_by, "descending": descending}
    _, total = read_results(summary["db_path"], limit=1, **query)
    pages = max(1, math.ceil(total / limit))
    with d:
        page = st.number_input("Page", min_value=1, max_value=pages, step=1, key="report_page")
    frame, total = read_results(summary["db_path"], limit=limit, offset=(page - 1) * limit, **query)
    st.caption(f"{total:,} matching leads · page {page:,} of {pages:,}")
    if total:
        relevant = _relevant_columns(summary)
        lead = ["Lead_Verdict", "Lead_Score", "Lead_Name"]
        if summary.get("email_column"):
            lead.append(summary["email_column"])
        lead += ["Verification_Status", "Mailbox_Check"]
        if summary.get("mapping", {}).get("phone"):
            lead.append(summary["mapping"]["phone"])
        lead += ["Phone_Status", "Phone_Type", "Lead_Country", "Lead_State", "Age", "Record_Age",
                 "Activity_Status", "Domain_Age_Years", "Lead_Issues", "Reason"]
        order = [c for c in dict.fromkeys(lead + summary["original_columns"] + relevant) if c in frame.columns
                 and (c in summary["original_columns"] or c in relevant)]
        config = {column: st.column_config.Column(_label(column)) for column in relevant}
        config["Lead_Score"] = st.column_config.ProgressColumn("Score", min_value=0, max_value=100, format="%d")
        config["Lead_Issues"] = st.column_config.TextColumn("Issues", width="large")
        config["Reason"] = st.column_config.TextColumn("Email reason", width="large")
        st.dataframe(frame, column_order=order, column_config=config, hide_index=True, width="stretch",
                     height=min(520, max(180, len(frame) * 35 + 40)))
    else:
        st.info("No leads match these filters. Try another status, country or search.")
    view_path = Path(summary["db_path"]).parent / "current_view.csv"

    def build_view():
        with _file_lock:
            export_view_csv(summary["db_path"], view_path, **query)
            return view_path.read_bytes()

    st.download_button(
        f"Download this view · {total:,} rows (CSV)", data=build_view, file_name="leads_current_view.csv",
        mime="text/csv", key="export_view_" + summary["batch_id"], on_click="ignore", disabled=not total,
        help="Exactly the filtered and sorted rows above, across all pages.",
    )


def render_organize_tab(summary):
    st.markdown(
        '<div class="section-note">Group leads by country, state, verdict or any column. '
        "Download an Excel workbook with one sorted sheet per group.</div>",
        unsafe_allow_html=True,
    )
    a, b, c = st.columns([1.2, 1.2, 1])
    with a:
        group_by = st.selectbox("Group by", _group_choices(summary), key="org_group", format_func=_choice_label)
    with b:
        sort_by = st.selectbox("Sort inside each group by", _sort_choices(summary), key="org_sort",
                               format_func=_choice_label)
    with c:
        descending = st.selectbox("Order", ["Ascending", "Descending"], key="org_order") == "Descending"
    status, search, filters = _filter_widgets(summary, "org_")
    groups = group_counts(summary["db_path"], group_by, status=status, search=search, filters=filters)
    if groups.empty:
        st.info("No leads match these filters.")
        return
    left, right = st.columns([1.1, 1])
    with left:
        st.dataframe(
            groups, hide_index=True, width="stretch", height=min(430, 38 + 35 * len(groups)),
            column_config={group_by: st.column_config.TextColumn(_label(group_by)),
                           "Avg score": st.column_config.ProgressColumn("Avg score", min_value=0, max_value=100, format="%.0f")},
        )
        st.caption(f"{len(groups):,} groups · {int(groups['Leads'].sum()):,} leads")
    with right:
        top = groups.head(15).set_index(group_by)[["Good", "Review", "Bad"]]
        st.bar_chart(top, horizontal=True, color=["#14B8A6", "#D3AA4E", "#F87171"], stack=True, height=430, sort=False)
    safe = "".join(ch if ch.isalnum() else "_" for ch in group_by).strip("_").lower() or "group"
    workbook_path = Path(summary["db_path"]).parent / f"organized_{safe}.xlsx"

    def build_workbook():
        with _file_lock:
            generate_organized_xlsx(summary["db_path"], workbook_path, group_by, sort_by=sort_by,
                                    descending=descending, status=status, search=search, filters=filters)
            return workbook_path.read_bytes()

    st.download_button(
        f"Download organized workbook · one sheet per {_label(group_by).lower()} (.xlsx)",
        data=build_workbook, file_name=f"leads_by_{safe}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="export_org_" + summary["batch_id"], on_click="ignore", type="primary",
        help="Up to 120 group sheets; any remaining groups go into an 'Other groups' sheet.",
    )


def render_export_tab(summary):
    st.markdown('<div class="section-note">All original columns, with verdicts, scores and every check appended. '
                "Exports always cover the full category.</div>", unsafe_allow_html=True)
    counts = summary["counts"]
    downloads = [
        ("Good leads", "good_csv", "good_leads.csv", f"{counts.get('good', 0):,} rows ready to contact", "primary"),
        ("Needs review", "review_csv", "review_leads.csv", f"{counts.get('review', 0):,} rows to double-check", "secondary"),
        ("Bad leads", "bad_csv", "bad_leads.csv", f"{counts.get('bad', 0):,} rows to remove or fix", "secondary"),
        ("Full report", "full_csv", "full_lead_report.csv", "Every row and every check", "secondary"),
    ]
    if summary.get("checks", {}).get("email", True):
        downloads.append(("Valid emails only", "valid_csv", "valid_email_leads.csv",
                          f"{counts.get('valid', 0):,} rows with a valid email", "secondary"))
    columns = st.columns(len(downloads))
    for column, (label, key, filename, help_text, kind) in zip(columns, downloads):
        with column:
            path = summary["exports"].get(key)
            if not path:
                continue
            st.download_button(f"{label} · CSV", data=lambda path=path: download_file(path), file_name=filename,
                               mime="text/csv", key="export_" + key + summary["batch_id"], on_click="ignore",
                               type=kind, width="stretch", help=help_text)
            st.caption(help_text)
    xlsx_path = Path(summary["db_path"]).parent / "full_lead_report.xlsx"

    def build_xlsx():
        with _file_lock:
            if not xlsx_path.exists():
                generate_xlsx(summary["db_path"], xlsx_path)
            return xlsx_path.read_bytes()

    st.download_button("Full report · Excel (.xlsx) with summary sheet", data=build_xlsx,
                       file_name="full_lead_report.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       key="export_xlsx_" + summary["batch_id"], on_click="ignore")


def render_details_tab(summary):
    checks = summary.get("checks", {})
    st.caption(f'Batch {summary["batch_id"]} · {summary.get("source_name", summary.get("filename", "Lead sheet"))} · '
               f'{summary.get("duration_seconds", 0):,.1f}s · {summary.get("unique_domains", 0):,} unique domains')
    ran = [name for name, on in (("Email syntax & duplicates", checks.get("email")), ("Domain DNS / MX", checks.get("email")),
                                 ("Mailbox (SMTP)", checks.get("smtp")), ("Phone validation", checks.get("phone")),
                                 ("Live phone line", checks.get("line")), ("Company domain age", checks.get("domain_age")),
                                 ("Client age", checks.get("dob")), ("Lead age", checks.get("created")),
                                 ("Activity status", checks.get("activity"))) if on]
    st.write("**Checks in this run:** " + ", ".join(ran))
    if summary.get("mapping"):
        st.write("**Columns used:** " + " · ".join(f"{FIELD_LABELS.get(k, k)} → `{v}`" for k, v in summary["mapping"].items()))
    st.markdown(
        "- **Mailbox: Exists** — the recipient's mail server confirmed the mailbox. This is the strongest signal available "
        "without sending an email.\n"
        "- **Catch-all** — the server accepts every address, so no checker can confirm the mailbox; graded Review.\n"
        "- **Unverifiable** — the server refused or delayed the check (greylisting, anti-abuse). The email keeps its DNS result.\n"
        "- **Phone status** validates the number against official numbering plans (country, length, type, carrier). "
        "Whether a line is switched on right now needs the optional live carrier lookup.\n"
        "- Duplicates compare emails ignoring case and spaces; phones compare in international (E.164) format."
    )
    if summary.get("column_renames"):
        st.write("Original columns renamed to avoid collisions:")
        st.json(summary["column_renames"])
    if st.session_state.get("recovery_code"):
        st.write("Private recovery code — treat it as access to the saved lead sheet:")
        st.code(st.session_state.recovery_code, language=None)


def render_results(summary):
    counts = summary["counts"]
    checks = summary.get("checks", {"email": True})
    st.success(
        f'Run complete · {counts["total"]:,} leads · {counts.get("good", 0):,} good · '
        f'{counts.get("review", 0):,} to review · {counts.get("bad", 0):,} bad'
    )
    for note in summary.get("notes", []):
        st.warning(note)
    for message in st.session_state.get("run_messages", []):
        st.caption(message)
    show_chips(counts, checks)
    leads, organize, export, details = st.tabs(["Leads", "Organize & group", "Export", "Run details"])
    with leads:
        render_leads_tab(summary)
    with organize:
        render_organize_tab(summary)
    with export:
        render_export_tab(summary)
    with details:
        render_details_tab(summary)


def main():
    st.set_page_config(
        page_title="OpenhouseFX · Lead Verifier",
        page_icon=str(APP_ROOT / "assets" / "openhousefx-house.svg"),
        layout="wide",
        initial_sidebar_state="expanded",
    )
    st.markdown(brand_css(), unsafe_allow_html=True)
    workspace = get_workspace()
    settings = render_sidebar()
    st.markdown(
        '<div class="crumbs"><div class="crumbs-label">Openhousefx <b>/ &nbsp; Verification workspace</b></div>'
        '<div class="local-pill"><i></i>Private session</div></div>'
        '<div class="eyebrow">THE CONFIDENCE BEFORE CONTACT</div>'
        '<div class="hero-title" role="heading" aria-level="1">Every lead. <span>A clearer picture.</span></div>'
        '<p class="hero-copy">Verify emails down to the mailbox, validate phone numbers, see each client\'s age and '
        "activity, then sort and group your list by name, state or country.</p>",
        unsafe_allow_html=True,
    )
    metrics_slot = st.empty()
    with st.container(border=True):
        source = render_input(workspace)
        if source:
            completed = st.session_state.get("result")
            same_run = completed is not None and st.session_state.get("result_source") == source["signature"]
            if completed and not same_run and not st.session_state.get("restored"):
                st.session_state.pop("result", None)
            mapping = source["mapping"]
            if st.button(
                "Verify again" if same_run else "Verify lead list",
                type="primary",
                width="stretch",
                disabled=(mapping.get("email") is None and mapping.get("phone") is None)
                or not source["info"].get("row_count", 0),
            ):
                execute_run(source, settings, metrics_slot)
    summary = st.session_state.get("result")
    with metrics_slot.container():
        show_metrics(summary["counts"] if summary else {}, summary.get("checks") if summary else None)
    if summary:
        render_results(summary)
    else:
        st.markdown(
            '<div class="section-title">Live workspace</div><div class="section-note">Your report will appear here when '
            'the check is complete.</div><div class="workspace-empty"><div class="empty-icon"><svg viewBox="0 0 24 24" '
            'fill="none" stroke="currentColor" stroke-width="1.3"><path d="M4 5h16v14H4z M4 6l8 6 8-6"/><path d="m8 16 2 2 5-5"/>'
            '</svg></div><div class="empty-title">A cleaner list starts here.</div><div class="empty-copy">Upload your lead '
            "sheet above, confirm the email and phone columns, and run your first check. Each lead gets a Good / Review / Bad "
            "verdict, a 0–100 score and a plain-English reason.</div></div>",
            unsafe_allow_html=True,
        )
    st.markdown(
        '<div class="footer-note"><span>OPENHOUSEFX · LEAD VERIFIER</span><span>Thoughtfully checked. Clearly organized.</span></div>',
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
