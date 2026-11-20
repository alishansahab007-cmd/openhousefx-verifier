"""Optional SMTP mailbox probe (RCPT TO) with catch-all detection.

No message is ever sent: the session stops after RCPT TO and issues QUIT.
Outbound TCP port 25 must be open on the host; many cloud providers block it,
in which case the engine reports "Not checked" instead of guessing.
"""

from __future__ import annotations

import os
import re
import secrets
import smtplib
import socket
import threading
import time
from dataclasses import dataclass
from functools import lru_cache

EXISTS = "Exists"
NOT_EXISTS = "Does not exist"
CATCH_ALL = "Catch-all (accepts any address)"
FULL = "Mailbox full"
DISABLED_MAILBOX = "Mailbox disabled"
UNVERIFIABLE = "Unverifiable"
NOT_CHECKED = "Not checked"

_BLOCK_WORDS = (
    "spamhaus", "blocked", "blacklist", "block list", "blocklist", "spam", "reputation",
    "policy", "not permitted", "relay", "rbl", "barracuda", "too many", "rate limit",
    "client host", "your ip", "listed", "authentication required", "tls required",
    "must issue a starttls", "sender address rejected", "helo", "reverse dns", "ptr",
)
_MISSING_WORDS = (
    "does not exist", "doesn't exist", "no such user", "user unknown", "unknown user",
    "mailbox unavailable", "mailbox not found", "recipient not found", "invalid recipient",
    "invalid mailbox", "address rejected", "recipient rejected", "no mailbox",
    "not a valid", "unknown recipient", "user not found", "account disabled",
    "account has been disabled", "inactive", "unrouteable address", "no such recipient",
    "undeliverable", "is not a known user", "mailbox is disabled",
)
_ENHANCED = re.compile(r"\b([245])\.(\d{1,3})\.(\d{1,3})\b")
SMTP_PORT = 25
_MX_LOCKS: dict[str, threading.BoundedSemaphore] = {}
_MX_LOCKS_GUARD = threading.Lock()


@dataclass(frozen=True)
class MailboxResult:
    state: str
    detail: str = ""


@lru_cache(maxsize=1)
def _fqdn() -> str:
    return socket.getfqdn()  # can take seconds on some hosts: resolve once


def smtp_identity() -> tuple[str, str]:
    """HELO name and MAIL FROM address, configurable for better acceptance."""
    helo = os.getenv("VERIFIER_SMTP_HELO", "").strip()
    if not helo:
        fqdn = _fqdn()
        helo = fqdn if "." in fqdn and not fqdn.endswith(".local") else "verifier.localdomain"
    sender = os.getenv("VERIFIER_SMTP_FROM", "").strip() or f"verify@{helo}"
    return helo, sender


def port25_available(hosts=("gmail-smtp-in.l.google.com", "alt1.gmail-smtp-in.l.google.com", "mx.zoho.com"), timeout=6.0) -> bool:
    """True if this machine can open an outbound TCP connection on port 25."""
    if os.getenv("VERIFIER_SMTP_DISABLED", "").strip() in {"1", "true", "yes"}:
        return False
    for host in hosts:
        try:
            with socket.create_connection((host, 25), timeout=timeout):
                return True
        except OSError:
            continue
    return False


def short_reply(message) -> str:
    """First line of a server reply, without URLs, capped for display."""
    if isinstance(message, bytes):
        message = message.decode("utf-8", "replace")
    first = str(message).splitlines()[0] if str(message).strip() else ""
    first = re.sub(r"https?://\S+", "", first)
    first = " ".join(first.split())
    return first[:160]


def classify_reply(code: int, message) -> MailboxResult:
    """Map an RCPT TO reply to a mailbox state."""
    if isinstance(message, bytes):
        message = message.decode("utf-8", "replace")
    full = " ".join(str(message).split())
    lower = full.lower()
    detail = f"{code} {short_reply(message)}"
    if 200 <= code < 300:
        return MailboxResult(EXISTS, detail)
    enhanced = _ENHANCED.search(lower)
    klass = enhanced.groups() if enhanced else None
    if code == 552 or (klass and klass[1:] == ("2", "2")) or any(
        word in lower for word in ("mailbox full", "over quota", "quota exceeded", "mailbox is full")
    ):
        return MailboxResult(FULL, detail)
    if 400 <= code < 500:
        return MailboxResult(UNVERIFIABLE, "Temporary reply (greylisting or rate limit): " + detail)
    if code >= 500:
        if klass and klass[0] == "5" and klass[1:] == ("2", "1"):
            return MailboxResult(DISABLED_MAILBOX, detail)
        if klass and klass[0] == "5" and (klass[1] == "1" or klass[1:] == ("4", "1")):
            return MailboxResult(NOT_EXISTS, detail)
        if klass and klass[1] == "7":
            return MailboxResult(UNVERIFIABLE, "Server refused the check: " + detail)
        if any(word in lower for word in _MISSING_WORDS):
            return MailboxResult(NOT_EXISTS, detail)
        if any(word in lower for word in _BLOCK_WORDS):
            return MailboxResult(UNVERIFIABLE, "Server refused the check: " + detail)
        if code in (550, 551, 553):
            return MailboxResult(NOT_EXISTS, detail)
    return MailboxResult(UNVERIFIABLE, detail)


def _mx_lock(host: str, per_host: int) -> threading.BoundedSemaphore:
    with _MX_LOCKS_GUARD:
        if host not in _MX_LOCKS:
            _MX_LOCKS[host] = threading.BoundedSemaphore(per_host)
        return _MX_LOCKS[host]


CONNECT_TIMEOUT = 6.0
BLOCK_SECONDS = 900
_UNREACHABLE: dict[str, float] = {}
_FAILURES: dict[str, int] = {}
_BREAKER_GUARD = threading.Lock()


def _provider(host: str) -> str:
    """Group MX hosts by provider (outlook.com, yahoodns.net, google.com …)."""
    host = host.casefold().rstrip(".")
    if ":" in host or host.replace(".", "").isdigit():
        return host  # IP address
    parts = host.split(".")
    return ".".join(parts[-2:]) if len(parts) >= 2 else host


def _blocked(host: str) -> bool:
    with _BREAKER_GUARD:
        return _UNREACHABLE.get(_provider(host), 0) > time.monotonic()


def _record_connect(host: str, ok: bool):
    key = _provider(host)
    with _BREAKER_GUARD:
        if ok:
            _FAILURES[key] = 0
            _UNREACHABLE.pop(key, None)
            return
        _FAILURES[key] = _FAILURES.get(key, 0) + 1
        if _FAILURES[key] >= 2:  # stop waiting on a provider this network cannot reach
            _UNREACHABLE[key] = time.monotonic() + BLOCK_SECONDS


class _FastSMTP(smtplib.SMTP):
    """SMTP client with a short connect timeout that tries IPv4 first, two addresses max."""

    def _get_socket(self, host, port, timeout):
        infos = socket.getaddrinfo(host, port, type=socket.SOCK_STREAM)
        infos.sort(key=lambda info: info[0] != socket.AF_INET)
        last_error = OSError(f"No address for {host}")
        for family, kind, proto, _, address in infos[:2]:
            sock = socket.socket(family, kind, proto)
            sock.settimeout(min(CONNECT_TIMEOUT, timeout))
            try:
                sock.connect(address)
                sock.settimeout(timeout)
                return sock
            except OSError as error:
                last_error = error
                sock.close()
        raise last_error


class _Session:
    def __init__(self, hosts, helo, sender, timeout):
        self.hosts, self.helo, self.sender, self.timeout = hosts, helo, sender, timeout
        self.smtp, self.host = None, None

    def open(self):
        last_error = "No mail server reachable"
        for host in self.hosts[:2]:
            if _blocked(host):
                last_error = f"Port 25 to {_provider(host)} is blocked from this server's network"
                continue
            smtp = _FastSMTP(local_hostname=self.helo, timeout=self.timeout)
            try:
                code, _ = smtp.connect(host, SMTP_PORT)
            except (OSError, smtplib.SMTPException) as error:
                _record_connect(host, False)
                last_error = f"Could not connect to {host} on port 25 ({type(error).__name__})"
                continue
            _record_connect(host, True)
            try:
                if code != 220:
                    smtp.close()
                    last_error = f"Server greeting {code}"
                    continue
                code, _ = smtp.ehlo(self.helo)
                if not 200 <= code < 300:
                    code, _ = smtp.helo(self.helo)
                code, message = smtp.mail(self.sender)
                if not 200 <= code < 300:
                    smtp.close()
                    return None, "Sender refused: " + classify_reply(code, message).detail
                self.smtp, self.host = smtp, host
                return smtp, None
            except (OSError, smtplib.SMTPException) as error:
                last_error = f"{host} dropped the connection ({type(error).__name__})"
                try:
                    smtp.close()
                except OSError:
                    pass
                continue
        return None, last_error

    def rcpt(self, address) -> MailboxResult:
        code, message = self.smtp.rcpt(address)
        return classify_reply(code, message)

    def close(self):
        if self.smtp is not None:
            try:
                self.smtp.quit()
            except (OSError, smtplib.SMTPException):
                try:
                    self.smtp.close()
                except OSError:
                    pass
        self.smtp = None


def check_domain_mailboxes(
    domain: str,
    mx_hosts,
    addresses,
    timeout: float = 12.0,
    per_session: int = 20,
    per_host: int = 4,
    identity: tuple[str, str] | None = None,
) -> tuple[bool | None, dict]:
    """Probe each address with RCPT TO. Returns (catch_all, {address: MailboxResult})."""
    helo, sender = identity or smtp_identity()
    hosts = [host for host in mx_hosts if host] or [domain]
    results: dict[str, MailboxResult] = {}
    pending = [address for address in dict.fromkeys(addresses) if address.isascii()]
    for address in addresses:
        if not address.isascii():
            results[address] = MailboxResult(UNVERIFIABLE, "International address; SMTPUTF8 probe not attempted")
    catch_all: bool | None = None
    lock = _mx_lock(hosts[0], per_host)
    attempts, max_attempts = 0, 2 + 2 * (-(-len(pending) // per_session))
    with lock:
        while pending and attempts < max_attempts:
            attempts += 1
            session = _Session(hosts, helo, sender, timeout)
            smtp, error = session.open()
            if smtp is None:
                for address in pending:
                    results[address] = MailboxResult(UNVERIFIABLE, error)
                return catch_all, results
            try:
                if catch_all is None:
                    probe = f"ohfx-{secrets.token_hex(6)}-nobody@{domain}"
                    probe_result = session.rcpt(probe)
                    if probe_result.state == EXISTS:
                        catch_all = True
                        for address in pending:
                            results[address] = MailboxResult(CATCH_ALL, "Server accepted a random address: " + probe_result.detail)
                        return catch_all, results
                    if probe_result.state in (NOT_EXISTS, DISABLED_MAILBOX):
                        catch_all = False
                chunk, pending = pending[:per_session], pending[per_session:]
                done = []
                for address in chunk:
                    try:
                        results[address] = session.rcpt(address)
                        done.append(address)
                    except smtplib.SMTPServerDisconnected:
                        break
                remaining = [address for address in chunk if address not in done]
                pending = remaining + pending
            except (OSError, smtplib.SMTPException):
                pass
            finally:
                session.close()
    for address in pending:
        results.setdefault(address, MailboxResult(UNVERIFIABLE, "Server disconnected during the check"))
    return catch_all, results


# --------------------------------------------------------------------------- API fallback
DO_NOT_MAIL = "Do not mail (spam trap / abuse)"


def api_provider_ready() -> bool:
    return bool(os.getenv("ZEROBOUNCE_API_KEY"))


def zerobounce_mailboxes(domain, mx_hosts, addresses, session=None, timeout=(5, 25)):
    """Mailbox check through ZeroBounce's HTTPS API (works where port 25 is blocked; billed per email)."""
    import requests

    http = session or requests
    key = os.getenv("ZEROBOUNCE_API_KEY", "")
    mapping = {"valid": EXISTS, "invalid": NOT_EXISTS, "catch-all": CATCH_ALL, "unknown": UNVERIFIABLE,
               "spamtrap": DO_NOT_MAIL, "abuse": DO_NOT_MAIL, "do_not_mail": DO_NOT_MAIL}
    results, catch_all = {}, None
    for address in addresses:
        try:
            response = http.get("https://api.zerobounce.net/v2/validate",
                                params={"api_key": key, "email": address, "ip_address": ""}, timeout=timeout)
            response.raise_for_status()
            data = response.json()
        except Exception as error:
            results[address] = MailboxResult(UNVERIFIABLE, f"Verification API error ({type(error).__name__})")
            continue
        if data.get("error"):
            results[address] = MailboxResult(UNVERIFIABLE, "Verification API: " + str(data["error"])[:120])
            continue
        status = str(data.get("status", "unknown")).lower()
        state = mapping.get(status, UNVERIFIABLE)
        detail = "ZeroBounce: " + status + (f" / {data['sub_status']}" if data.get("sub_status") else "")
        results[address] = MailboxResult(state, detail)
        if state == CATCH_ALL:
            catch_all = True
    return catch_all, results
