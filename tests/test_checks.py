"""Unit tests for the individual checks. No live network access."""

from __future__ import annotations

import socketserver
import threading
from datetime import date, datetime
from types import SimpleNamespace

import dns.exception
import dns.resolver
import pytest

from verifier import email_checks as ec
from verifier import enrich, smtp_check
from verifier import phone as ph
from verifier.columns import detect_columns


# --------------------------------------------------------------------------- columns
def test_detects_common_lead_sheet_headers():
    columns = ["First Name", "Last Name", "E-mail Address", "Mobile", "Country", "State/Province",
               "DOB", "Created At", "Last Login", "Company Name", "Website"]
    found = detect_columns(columns)
    assert found == {
        "email": 2, "phone": 3, "first_name": 0, "last_name": 1, "company": 9, "website": 10,
        "country": 4, "state": 5, "dob": 6, "created": 7, "last_activity": 8,
    }


def test_detection_does_not_reuse_a_column():
    found = detect_columns(["Email", "Email Status"])
    assert found["email"] == 0


# --------------------------------------------------------------------------- email syntax & flags
@pytest.mark.parametrize(
    "raw, expected",
    [
        ("  Alex@Gmail.COM ", "Alex@gmail.com"),
        ("mailto:alex@gmail.com?subject=hi", "alex@gmail.com"),
        ("Alex Doe <alex@gmail.com>", "alex@gmail.com"),
        ("alex@gmail.com; second@gmail.com", "alex@gmail.com"),
        ("'alex@gmail.com'.", "alex@gmail.com"),
    ],
)
def test_messy_cells_are_cleaned_before_validation(raw, expected):
    email, domain, error, _ = ec.check_syntax(raw)
    assert error is None
    assert email == expected and domain == "gmail.com"


@pytest.mark.parametrize("raw", ["not-an-email", "two@@x.com", "a@b", "a@b..com", ""])
def test_bad_syntax_is_rejected(raw):
    assert ec.check_syntax(raw)[2]


def test_address_is_pulled_out_of_surrounding_text():
    assert ec.check_syntax("Contact: john@acme.com")[0] == "john@acme.com"


def test_missing_and_non_text_cells():
    assert ec.check_syntax(None)[2] == "Missing email address"
    assert ec.check_syntax(12345)[2] == "Missing email address"


@pytest.mark.parametrize(
    "domain, suggestion",
    [("gmial.com", "gmail.com"), ("gmail.co", "gmail.com"), ("gmail.con", "gmail.com"),
     ("hotmial.com", "hotmail.com"), ("yahooo.com", "yahoo.com"), ("outlok.com", "outlook.com"),
     ("gmail.com", None), ("mail.com", None), ("ymail.com", None), ("openhousefx.com", None),
     ("live.com", None), ("aol.com", None), ("me.com", None)],
)
def test_typo_suggestions(domain, suggestion):
    assert ec.typo_suggestion(domain) == suggestion


def test_role_noreply_junk_free_disposable():
    assert ec.is_role("sales+leads@acme.com") and ec.is_role("Info@acme.com")
    assert not ec.is_role("sarah@acme.com")
    assert ec.is_no_reply("no-reply@acme.com") and ec.is_no_reply("donotreply@acme.com")
    assert ec.is_junk("test@acme.com") and ec.is_junk("aaaa@acme.com")
    assert not ec.is_junk("tess@acme.com")
    assert ec.is_free_provider("gmail.com") and not ec.is_free_provider("acme.com")
    assert ec.disposable_match("x.mailinator.com", frozenset({"mailinator.com"}))
    assert not ec.disposable_match("com", frozenset({"com"}))


# --------------------------------------------------------------------------- DNS
class FakeResolver:
    def __init__(self, answers):
        self.answers, self.calls = answers, []

    def resolve(self, name, record_type, lifetime=None, search=False):
        self.calls.append((name, record_type))
        value = self.answers.get((name, record_type), dns.resolver.NoAnswer())
        if isinstance(value, Exception):
            raise value
        return value


def mx(host, preference=10):
    return SimpleNamespace(exchange=host + ".", preference=preference)


def test_mx_with_resolvable_hosts_is_valid():
    resolver = FakeResolver({
        ("acme.com", "MX"): [mx("mx2.acme.com", 20), mx("mx1.acme.com", 10)],
        ("mx1.acme.com", "A"): ["1.2.3.4"],
    })
    result = ec.lookup_domain("acme.com", resolver=resolver)
    assert result.status == "Valid" and result.mx_hosts == ("mx1.acme.com", "mx2.acme.com")


def test_mx_pointing_to_nonexistent_hosts_is_invalid():
    resolver = FakeResolver({
        ("acme.com", "MX"): [mx("gone.acme.com")],
        ("gone.acme.com", "A"): dns.resolver.NXDOMAIN(),
    })
    assert ec.lookup_domain("acme.com", resolver=resolver).status == "Invalid"


def test_mx_host_timeout_does_not_penalize():
    resolver = FakeResolver({
        ("acme.com", "MX"): [mx("slow.acme.com")],
        ("slow.acme.com", "A"): dns.exception.Timeout(),
        ("slow.acme.com", "AAAA"): dns.exception.Timeout(),
    })
    assert ec.lookup_domain("acme.com", resolver=resolver).status == "Valid"


def test_null_mx_nxdomain_fallback_timeout_and_no_route():
    assert ec.lookup_domain("x.com", resolver=FakeResolver({("x.com", "MX"): [mx("")]})).dns_check == "Null MX"
    assert ec.lookup_domain("x.com", resolver=FakeResolver({("x.com", "MX"): dns.resolver.NXDOMAIN()})).status == "Invalid"
    fallback = ec.lookup_domain("x.com", resolver=FakeResolver({("x.com", "A"): ["1.1.1.1"]}))
    assert (fallback.status, fallback.dns_check) == ("Risky", "A")
    timeout = ec.lookup_domain("x.com", resolver=FakeResolver({("x.com", "MX"): dns.exception.Timeout()}))
    assert (timeout.status, timeout.dns_check) == ("Risky", "Timeout")
    assert ec.lookup_domain("x.com", resolver=FakeResolver({})).dns_check == "No route"


def test_dns_override_is_validated(monkeypatch):
    monkeypatch.setenv("VERIFIER_DNS_NAMESERVERS", "1.1.1.1, 8.8.8.8")
    assert ec.validate_dns_configuration() == ("1.1.1.1", "8.8.8.8")
    monkeypatch.setenv("VERIFIER_DNS_NAMESERVERS", "not-an-ip")
    with pytest.raises(ValueError):
        ec.validate_dns_configuration()


# --------------------------------------------------------------------------- SMTP
@pytest.mark.parametrize(
    "code, message, state",
    [
        (250, "2.1.5 OK", smtp_check.EXISTS),
        (550, "5.1.1 The email account that you tried to reach does not exist.", smtp_check.NOT_EXISTS),
        (550, "5.4.1 Recipient address rejected: Access denied", smtp_check.NOT_EXISTS),
        (550, "Requested action not taken: mailbox unavailable", smtp_check.NOT_EXISTS),
        (550, "5.2.1 The email account that you tried to reach is disabled.", smtp_check.DISABLED_MAILBOX),
        (552, "5.2.2 Mailbox full", smtp_check.FULL),
        (450, "4.2.1 Greylisted, please try again", smtp_check.UNVERIFIABLE),
        (554, "5.7.1 Service unavailable; client host blocked using Spamhaus", smtp_check.UNVERIFIABLE),
        (550, "Your IP is listed in a blocklist", smtp_check.UNVERIFIABLE),
        (550, "", smtp_check.NOT_EXISTS),
    ],
)
def test_classify_rcpt_replies(code, message, state):
    assert smtp_check.classify_reply(code, message.encode()).state == state


class FakeSMTPHandler(socketserver.StreamRequestHandler):
    mailboxes = {"real@acme.test"}
    catch_all_domains = {"catchall.test"}

    def reply(self, text):
        self.wfile.write((text + "\r\n").encode())

    def handle(self):
        self.reply("220 fake ESMTP")
        for raw in self.rfile:
            line = raw.decode().strip()
            verb = line.split(" ", 1)[0].upper()
            if verb in ("EHLO", "HELO"):
                self.reply("250 fake")
            elif verb == "MAIL":
                self.reply("250 2.1.0 OK")
            elif verb == "RCPT":
                address = line.split("<", 1)[1].split(">", 1)[0].lower()
                domain = address.rsplit("@", 1)[1]
                if domain in self.catch_all_domains or address in self.mailboxes:
                    self.reply("250 2.1.5 OK")
                else:
                    self.reply("550 5.1.1 User unknown")
            elif verb == "QUIT":
                self.reply("221 bye")
                return
            else:
                self.reply("500 unknown")


@pytest.fixture
def fake_smtp(monkeypatch):
    server = socketserver.ThreadingTCPServer(("127.0.0.1", 0), FakeSMTPHandler)
    server.daemon_threads = True
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    monkeypatch.setattr(smtp_check, "SMTP_PORT", server.server_address[1])
    yield
    server.shutdown()
    server.server_close()


def test_smtp_probe_exists_missing_and_catch_all(fake_smtp):
    identity = ("verifier.test", "verify@verifier.test")
    catch_all, results = smtp_check.check_domain_mailboxes(
        "acme.test", ["127.0.0.1"], ["real@acme.test", "ghost@acme.test"], identity=identity, timeout=5)
    assert catch_all is False
    assert results["real@acme.test"].state == smtp_check.EXISTS
    assert results["ghost@acme.test"].state == smtp_check.NOT_EXISTS
    catch_all, results = smtp_check.check_domain_mailboxes(
        "catchall.test", ["127.0.0.1"], ["anyone@catchall.test"], identity=identity, timeout=5)
    assert catch_all is True and results["anyone@catchall.test"].state == smtp_check.CATCH_ALL


def test_smtp_probe_unreachable_server_is_unverifiable(monkeypatch):
    monkeypatch.setattr(smtp_check, "SMTP_PORT", 9)  # discard port: nothing listens
    _, results = smtp_check.check_domain_mailboxes(
        "acme.test", ["127.0.0.1"], ["a@acme.test"], identity=("h.test", "v@h.test"), timeout=2)
    assert results["a@acme.test"].state == smtp_check.UNVERIFIABLE


def test_smtp_probe_many_addresses_reconnects(fake_smtp):
    addresses = [f"user{i}@acme.test" for i in range(45)] + ["real@acme.test"]
    _, results = smtp_check.check_domain_mailboxes(
        "acme.test", ["127.0.0.1"], addresses, identity=("h.test", "v@h.test"), timeout=5, per_session=10)
    assert len(results) == 46
    assert results["real@acme.test"].state == smtp_check.EXISTS
    assert sum(r.state == smtp_check.NOT_EXISTS for r in results.values()) == 45


# --------------------------------------------------------------------------- phone
@pytest.mark.parametrize(
    "value, region, status, e164",
    [
        ("(202) 555-0182", "US", "Valid", "+12025550182"),
        ("020 7946 0958", "GB", "Valid", "+442079460958"),
        ("+44 20 7946 0958", "US", "Valid", "+442079460958"),
        ("0044 20 7946 0958", "US", "Valid", "+442079460958"),
        ("447911123456", "US", "Valid", "+447911123456"),  # Excel dropped the '+'
        (4155552671.0, "US", "Valid", "+14155552671"),  # Excel number cell
        ("12345", "US", "Invalid", ""),
        ("1111111111", "US", "Invalid", ""),
        ("", "US", "Missing", ""),
        (None, "US", "Missing", ""),
        ("n/a", "US", "Missing", ""),
    ],
)
def test_phone_parsing(value, region, status, e164):
    result = ph.parse_phone(value, region)
    assert result.status == status
    if status == "Valid":
        assert result.e164 == e164


def test_phone_type_and_country():
    result = ph.parse_phone("+61 412 345 678", "US")
    assert result.type == "Mobile" and result.country == "AU"


def test_twilio_response_parsing(monkeypatch):
    monkeypatch.setenv("TWILIO_ACCOUNT_SID", "AC_test")
    monkeypatch.setenv("TWILIO_AUTH_TOKEN", "secret")

    class Response:
        status_code = 200

        def __init__(self, data):
            self.data = data

        def json(self):
            return self.data

        def raise_for_status(self):
            pass

    class Session:
        def __init__(self, data):
            self.data, self.calls = data, []

        def get(self, url, **kwargs):
            self.calls.append((url, kwargs))
            return Response(self.data)

    session = Session({"valid": True, "line_status": {"status": "inactive"},
                       "line_type_intelligence": {"type": "mobile", "carrier_name": "Carrier X"}})
    result = ph.twilio_line_lookup("+14155552671", session)
    assert (result.status, result.line_type, result.carrier) == ("Inactive", "Mobile", "Carrier X")
    assert session.calls[0][1]["auth"] == ("AC_test", "secret")
    assert ph.twilio_line_lookup("+1", Session({"valid": False})).status == "Invalid"


# --------------------------------------------------------------------------- enrichment
@pytest.mark.parametrize(
    "value, code, name",
    [("USA", "US", "United States"), ("u.s.a.", "US", "United States"), ("UK", "GB", "United Kingdom"),
     ("England", "GB", "United Kingdom"), ("de", "DE", "Germany"), ("Deutschland", "DE", "Germany"),
     ("UAE", "AE", "United Arab Emirates"), ("Pakistan", "PK", "Pakistan"), ("", "", "")],
)
def test_country_normalization(value, code, name):
    assert enrich.normalize_country(value) == (code, name)


def test_unknown_country_is_kept():
    assert enrich.normalize_country("Atlantis") == ("", "Atlantis")


@pytest.mark.parametrize(
    "value, country, expected",
    [("ca", "US", "California"), ("CA", "", "California"), ("New york", "US", "New York"),
     ("ON", "CA", "Ontario"), ("NSW", "AU", "New South Wales"), ("Punjab", "PK", "Punjab"),
     ("Somewhere", "US", "Somewhere")],
)
def test_state_normalization(value, country, expected):
    assert enrich.normalize_state(value, country) == expected


def test_date_parsing_and_ages():
    today = date(2026, 11, 20)
    assert enrich.parse_date("1990-05-12", today=today) == date(1990, 5, 12)
    assert enrich.parse_date("12/05/1990", dayfirst=True, today=today) == date(1990, 5, 12)
    assert enrich.parse_date("12/05/1990", dayfirst=False, today=today) == date(1990, 12, 5)
    assert enrich.parse_date(datetime(2020, 1, 2, 3, 4), today=today) == date(2020, 1, 2)
    assert enrich.parse_date(45000, today=today) == date(2023, 3, 15)  # Excel serial
    assert enrich.parse_date("n/a", today=today) is None
    assert enrich.parse_date("2099-01-01", today=today) is None
    assert enrich.age_years(date(1990, 11, 21), today) == 35
    assert enrich.age_years(date(1990, 11, 20), today) == 36
    assert enrich.parse_age("42 years") == 42 and enrich.parse_age("abc") is None
    assert enrich.activity_status(10) == "Active"
    assert enrich.activity_status(200) == "Lapsing"
    assert enrich.activity_status(800) == "Inactive"
    assert enrich.activity_status(None) == "Unknown"
    assert enrich.humanize_days(400) == "1 year 1 mo"
    assert enrich.humanize_days(45) == "1 month"


def test_names_and_domains():
    assert enrich.build_name(first="alex", last="SMITH") == "Alex Smith"
    assert enrich.build_name(full="Mary McDonald") == "Mary McDonald"
    assert enrich.registrable_domain("mail.acme.co.uk") == "acme.co.uk"
    assert enrich.website_domain("https://www.acme.com/about") == "acme.com"
    assert enrich.website_domain("acme.com") == "acme.com"
    assert enrich.country_from_domain("acme.co.uk") == "GB"
    assert enrich.country_from_domain("acme.io") == ""


def test_zerobounce_mapping(monkeypatch):
    monkeypatch.setenv("ZEROBOUNCE_API_KEY", "k")
    answers = {"a@x.com": {"status": "valid"}, "b@x.com": {"status": "invalid", "sub_status": "mailbox_not_found"},
               "c@x.com": {"status": "catch-all"}, "d@x.com": {"status": "spamtrap"}, "e@x.com": {"error": "Invalid API key"}}

    class Response:
        def __init__(self, data):
            self.data = data

        def raise_for_status(self):
            pass

        def json(self):
            return self.data

    class Session:
        def get(self, url, params, timeout):
            assert params["api_key"] == "k"
            return Response(answers[params["email"]])

    catch_all, results = smtp_check.zerobounce_mailboxes("x.com", [], list(answers), session=Session())
    assert catch_all is True
    assert [results[a].state for a in answers] == [
        smtp_check.EXISTS, smtp_check.NOT_EXISTS, smtp_check.CATCH_ALL, smtp_check.DO_NOT_MAIL, smtp_check.UNVERIFIABLE]
