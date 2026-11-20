"""Email syntax, list-hygiene flags and DNS mail-route checks."""

from __future__ import annotations

import ipaddress
import os
import re
import threading
import time
from dataclasses import dataclass, field
from typing import Callable

import dns.exception
import dns.resolver
from email_validator import EmailNotValidError, validate_email

ROLE_ACCOUNTS = frozenset(
    """abuse accounting accounts admin administrator admissions all billing bookings
    booking care careers ceo compliance contact contactus customercare customerservice
    customersupport dev development enquiries enquiry everyone events feedback finance
    frontdesk hello help helpdesk hostmaster hr inbox info information inquiries inquiry
    investor investors it jobs legal list mail marketing media news newsletter office
    operations orders partners partnership payments postmaster press privacy purchasing
    reception recruiting recruitment registrar remove reservations root sales security
    service services shop staff studio subscribe support team tech hiring web webmaster
    welcome""".split()
)
NO_REPLY = frozenset(
    {"noreply", "no-reply", "no_reply", "donotreply", "do-not-reply", "do_not_reply",
     "mailer-daemon", "bounce", "bounces", "unsubscribe"}
)
JUNK_LOCALS = frozenset(
    {"test", "testing", "tester", "test1", "test123", "asdf", "asdfgh", "qwerty", "abc",
     "abc123", "xyz", "xxx", "aaa", "fake", "fakeemail", "none", "noemail", "nomail",
     "na", "n/a", "null", "nobody", "example", "sample", "demo", "dummy", "spam", "junk",
     "anonymous", "123", "1234", "12345", "123456"}
)
FREE_PROVIDERS = frozenset(
    """gmail.com googlemail.com yahoo.com yahoo.co.uk yahoo.ca yahoo.com.au yahoo.co.in
    yahoo.fr yahoo.de yahoo.es yahoo.it yahoo.co.jp ymail.com rocketmail.com hotmail.com
    hotmail.co.uk hotmail.fr hotmail.de hotmail.it hotmail.es hotmail.ca outlook.com
    outlook.co.uk outlook.fr outlook.de outlook.es outlook.it live.com live.co.uk live.ca
    live.fr live.de live.it live.nl msn.com aol.com aol.co.uk aim.com icloud.com me.com
    mac.com protonmail.com protonmail.ch proton.me pm.me gmx.com gmx.de gmx.net gmx.at
    gmx.ch gmx.co.uk web.de mail.com email.com zoho.com zohomail.com yandex.com yandex.ru
    ya.ru mail.ru inbox.ru list.ru bk.ru tutanota.com tuta.io fastmail.com hushmail.com
    qq.com 163.com 126.com sina.com naver.com daum.net hanmail.net rediffmail.com
    comcast.net verizon.net att.net sbcglobal.net bellsouth.net cox.net charter.net
    earthlink.net optonline.net btinternet.com sky.com virginmedia.com talktalk.net
    orange.fr wanadoo.fr free.fr laposte.net sfr.fr libero.it virgilio.it t-online.de
    freenet.de bigpond.com optusnet.com.au shaw.ca rogers.com sympatico.ca telus.net
    uol.com.br bol.com.br terra.com.br""".split()
)
# Typo targets: the most common consumer domains.
POPULAR_DOMAINS = (
    "gmail.com", "yahoo.com", "hotmail.com", "outlook.com", "aol.com", "icloud.com",
    "live.com", "msn.com", "comcast.net", "googlemail.com", "ymail.com", "protonmail.com",
    "yahoo.co.uk", "hotmail.co.uk", "btinternet.com", "verizon.net", "att.net",
    "sbcglobal.net", "bellsouth.net", "me.com", "mac.com", "gmx.com", "mail.com",
)
_TLD_TYPOS = {"co", "cm", "con", "cmo", "comm", "om", "vom", "xom", "cpm", "c0m", "ocm",
              "net.com", "nte", "ent", "nett", "cim", "coom", "comcom"}

_DNS_LOCAL = threading.local()


@dataclass(frozen=True)
class DomainResult:
    """Domain mail-route result. Never a claim that an individual mailbox exists."""

    status: str
    reason: str
    dns_check: str
    mx_hosts: tuple = field(default=())


# --------------------------------------------------------------------------- syntax
_ANGLE = re.compile(r"<\s*([^<>\s]+@[^<>\s]+)\s*>")
_SPLIT = re.compile(r"[;,\s]+")


def extract_email(value) -> tuple[str | None, str | None]:
    """Pull one address out of common messy cell formats.

    Returns (candidate, note). Handles mailto:, "Name <a@b.com>", surrounding
    quotes/punctuation and cells holding several addresses (first is used).
    """
    if value is None:
        return None, None
    if not isinstance(value, str):
        return None, None
    text = value.strip().replace("​", "").replace("﻿", "")
    if not text:
        return None, None
    note = None
    match = _ANGLE.search(text)
    if match:
        text, note = match.group(1), "Extracted address from display name"
    if text.lower().startswith("mailto:"):
        text = text[7:].split("?", 1)[0]
        note = note or "Removed mailto: prefix"
    parts = [part for part in _SPLIT.split(text) if part]
    with_at = [part for part in parts if "@" in part]
    if len(with_at) > 1:
        text, note = with_at[0], f"Cell had {len(with_at)} addresses; checked the first"
    elif len(parts) > 1 and len(with_at) == 1:
        text = with_at[0]
    text = text.strip("\"'()[]{}<>.,;: ")
    return text or None, note


def check_syntax(value) -> tuple[str, str, str | None, str | None]:
    """Return (normalized_email, ascii_domain, error, note)."""
    candidate, note = extract_email(value)
    if not candidate:
        return "", "", "Missing email address", None
    try:
        parsed = validate_email(
            candidate,
            check_deliverability=False,
            allow_smtputf8=True,
            strict=True,
            allow_quoted_local=True,
            allow_domain_literal=False,
        )
    except EmailNotValidError as error:
        return "", "", f"Invalid syntax: {error}", note
    domain = parsed.ascii_domain.casefold()
    local = parsed.local_part
    return f"{local}@{domain}", domain, None, note


def local_part(email: str) -> str:
    return email.rsplit("@", 1)[0].strip('"').casefold() if email else ""


def base_local(email: str) -> str:
    """Local part without +tag, used for role/junk matching."""
    return local_part(email).split("+", 1)[0]


def is_role(email: str) -> bool:
    local = base_local(email)
    return local in ROLE_ACCOUNTS or local.replace(".", "").replace("-", "") in ROLE_ACCOUNTS


def is_no_reply(email: str) -> bool:
    return base_local(email) in NO_REPLY


def is_junk(email: str) -> bool:
    local = base_local(email)
    return local in JUNK_LOCALS or bool(re.fullmatch(r"(.)\1{3,}", local))


def disposable_match(domain: str, domains: frozenset) -> bool:
    parts = domain.casefold().split(".")
    return any(".".join(parts[index:]) in domains for index in range(len(parts) - 1))


def is_free_provider(domain: str) -> bool:
    return domain.casefold() in FREE_PROVIDERS


def _distance(a: str, b: str, limit: int = 2) -> int:
    """Damerau-Levenshtein (optimal string alignment) with an early exit."""
    if abs(len(a) - len(b)) > limit:
        return limit + 1
    previous2, previous = None, list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        current = [i] + [0] * len(b)
        for j, cb in enumerate(b, 1):
            cost = 0 if ca == cb else 1
            current[j] = min(previous[j] + 1, current[j - 1] + 1, previous[j - 1] + cost)
            if i > 1 and j > 1 and ca == b[j - 2] and a[i - 2] == cb:
                current[j] = min(current[j], previous2[j - 2] + 1)
        if min(current) > limit:
            return limit + 1
        previous2, previous = previous, current
    return previous[-1]


def typo_suggestion(domain: str) -> str | None:
    """Suggest a popular provider when the domain looks like a misspelling of it."""
    domain = domain.casefold()
    if domain in FREE_PROVIDERS or domain in POPULAR_DOMAINS:
        return None
    name, _, tld = domain.partition(".")
    for popular in POPULAR_DOMAINS:
        pop_name, _, pop_tld = popular.partition(".")
        if name == pop_name and tld != pop_tld and tld in _TLD_TYPOS:
            return popular
    best = None
    for popular in POPULAR_DOMAINS:
        if len(popular) < 9:  # aol.com / msn.com / live.com: too short to judge safely
            continue
        if _distance(domain, popular, 1) <= 1:
            best = popular
            break
    return best


# --------------------------------------------------------------------------- DNS
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
            raise ValueError("VERIFIER_DNS_NAMESERVERS must contain comma-separated DNS server IP addresses.")
        try:
            resolver.nameservers = servers
        except (ValueError, TypeError) as error:
            raise ValueError(
                "Invalid VERIFIER_DNS_NAMESERVERS. Use comma-separated DNS server IP addresses."
            ) from error
    elif not resolver.nameservers:
        resolver.nameservers = ["1.1.1.1", "8.8.8.8"]
    resolver.timeout = 1.5
    resolver.lifetime = 4.0
    resolver.cache = dns.resolver.LRUCache(max_size=4096)
    return resolver


def validate_dns_configuration():
    """Fail before starting a run when an explicit DNS override is malformed."""
    return tuple(_build_resolver().nameservers)


def _thread_resolver():
    signature = os.getenv("VERIFIER_DNS_NAMESERVERS", "").strip()
    if not hasattr(_DNS_LOCAL, "resolver") or getattr(_DNS_LOCAL, "signature", None) != signature:
        _DNS_LOCAL.resolver = _build_resolver()
        _DNS_LOCAL.signature = signature
    return _DNS_LOCAL.resolver


def _is_ip(host: str) -> bool:
    try:
        ipaddress.ip_address(host)
        return True
    except ValueError:
        return False


def _host_resolves(resolver, host, deadline, clock) -> bool | None:
    """True if the MX host has an address, False if it definitely does not, None if unknown."""
    if _is_ip(host):
        return True
    answered_empty = True
    for record_type in ("A", "AAAA"):
        remaining = deadline - clock()
        if remaining <= 0:
            return None
        try:
            if list(resolver.resolve(host, record_type, lifetime=remaining, search=False)):
                return True
        except dns.resolver.NXDOMAIN:
            return False
        except dns.resolver.NoAnswer:
            continue
        except Exception:  # timeouts, SERVFAIL: do not penalize the domain
            answered_empty = False
    return False if answered_empty else None


def lookup_domain(
    domain: str, timeout: float = 4.0, resolver=None, clock: Callable = time.monotonic
) -> DomainResult:
    """Check MX (and that the MX hosts exist), then A/AAAA fallback.

    Timeout/SERVFAIL is inconclusive (Risky). Null MX explicitly declines mail.
    A/AAAA fallback is Risky because no explicit mail service was published.
    """
    timeout = min(10.0, max(0.05, float(timeout)))
    deadline = clock() + timeout
    resolver = resolver if resolver is not None else _thread_resolver()
    for record_type in ("MX", "A", "AAAA"):
        remaining = deadline - clock()
        if remaining <= 0:
            return DomainResult("Risky", "DNS lookup timed out; retry later", "Timeout")
        try:
            records = list(resolver.resolve(domain, record_type, lifetime=remaining, search=False))
            if not records:
                continue
            if record_type == "MX":
                if any(str(record.exchange).strip() == "." for record in records):
                    return DomainResult(
                        "Invalid", "Domain explicitly does not accept email (null MX)", "Null MX"
                    )
                ordered = sorted(records, key=lambda record: getattr(record, "preference", 0))
                hosts = tuple(
                    dict.fromkeys(str(record.exchange).rstrip(".").casefold() for record in ordered)
                )
                states = [_host_resolves(resolver, host, deadline, clock) for host in hosts[:3]]
                if states and all(state is False for state in states):
                    return DomainResult(
                        "Invalid", "Mail server (MX) hosts do not exist", "MX broken", hosts
                    )
                return DomainResult("Valid", "Domain has working mail servers (MX)", "MX", hosts)
            return DomainResult(
                "Risky",
                f"No MX record; mail would fall back to the {record_type} record, mail service unconfirmed",
                record_type,
                (domain,),
            )
        except dns.resolver.NXDOMAIN:
            return DomainResult("Invalid", "Email domain does not exist", "NXDOMAIN")
        except dns.resolver.NoAnswer:
            continue
        except (dns.exception.Timeout, dns.resolver.LifetimeTimeout):
            return DomainResult("Risky", "DNS lookup timed out; retry later", "Timeout")
        except dns.resolver.NoNameservers:
            return DomainResult("Risky", "DNS servers failed (SERVFAIL); retry later", "DNS error")
        except dns.exception.DNSException:
            return DomainResult("Risky", "DNS response inconclusive; retry later", "DNS error")
        except OSError:
            return DomainResult("Risky", "Network unavailable; retry later", "Network error")
    return DomainResult("Invalid", "Domain has no mail server (no MX, A or AAAA record)", "No route")


RETRYABLE_DNS = frozenset({"Timeout", "DNS error", "Network error"})
