"""Lead enrichment: names, location normalization, dates, age, activity, domain age."""

from __future__ import annotations

import re
import threading
from datetime import date, datetime, timedelta, timezone
from functools import lru_cache
from urllib.parse import urlparse

import pycountry

# --------------------------------------------------------------------------- names
def build_name(full=None, first=None, last=None) -> str:
    def clean(value):
        if value is None or isinstance(value, bool):
            return ""
        return " ".join(str(value).split())

    def tidy(part):
        # Fix ALL-CAPS / all-lowercase entries; leave mixed case (McDonald) alone.
        if len(part) > 1 and (part.isupper() or part.islower()):
            return part.title()
        return part

    full, first, last = tidy(clean(full)), tidy(clean(first)), tidy(clean(last))
    return full or " ".join(part for part in (first, last) if part)


def clean_text(value) -> str:
    if value is None or isinstance(value, bool):
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return " ".join(str(value).split())


# --------------------------------------------------------------------------- location
_COUNTRY_ALIASES = {
    "uk": "GB", "u.k.": "GB", "england": "GB", "scotland": "GB", "wales": "GB",
    "northern ireland": "GB", "great britain": "GB", "britain": "GB",
    "us": "US", "u.s.": "US", "u.s.a.": "US", "usa": "US", "america": "US",
    "united states of america": "US", "the united states": "US",
    "uae": "AE", "u.a.e.": "AE", "emirates": "AE", "dubai": "AE", "abu dhabi": "AE",
    "russia": "RU", "turkey": "TR", "turkiye": "TR", "türkiye": "TR",
    "ivory coast": "CI", "cote d'ivoire": "CI", "macedonia": "MK", "north macedonia": "MK",
    "holland": "NL", "the netherlands": "NL", "deutschland": "DE", "korea": "KR",
    "south korea": "KR", "north korea": "KP", "vatican": "VA", "palestine": "PS",
    "ksa": "SA", "saudi": "SA", "hong kong": "HK", "macau": "MO", "macao": "MO",
    "czechia": "CZ", "burma": "MM", "swaziland": "SZ", "east timor": "TL",
    "cape verde": "CV", "brunei": "BN", "laos": "LA", "syria": "SY", "iran": "IR",
    "vietnam": "VN", "venezuela": "VE", "bolivia": "BO", "tanzania": "TZ", "moldova": "MD",
    "micronesia": "FM", "congo": "CG", "drc": "CD", "dr congo": "CD",
    "democratic republic of the congo": "CD", "espana": "ES", "españa": "ES",
    "brasil": "BR", "méxico": "MX", "mexico": "MX", "sverige": "SE", "italia": "IT",
    "nederland": "NL", "schweiz": "CH", "suisse": "CH", "österreich": "AT",
    "polska": "PL", "nippon": "JP", "pakistan": "PK", "bharat": "IN",
}
_DISPLAY = {
    "US": "United States", "GB": "United Kingdom", "RU": "Russia", "KR": "South Korea",
    "KP": "North Korea", "IR": "Iran", "SY": "Syria", "VN": "Vietnam", "TZ": "Tanzania",
    "VE": "Venezuela", "BO": "Bolivia", "MD": "Moldova", "LA": "Laos", "TW": "Taiwan",
    "TR": "Turkey", "PS": "Palestine", "CD": "DR Congo", "CG": "Congo", "CZ": "Czechia",
    "AE": "United Arab Emirates", "VA": "Vatican City", "FM": "Micronesia",
    "MK": "North Macedonia", "BN": "Brunei", "CV": "Cape Verde", "CI": "Côte d'Ivoire",
}


def country_name(alpha2: str) -> str:
    if not alpha2:
        return ""
    if alpha2 in _DISPLAY:
        return _DISPLAY[alpha2]
    record = pycountry.countries.get(alpha_2=alpha2)
    if record is None:
        return alpha2
    return getattr(record, "common_name", None) or record.name


@lru_cache(maxsize=4096)
def normalize_country(value: str) -> tuple[str, str]:
    """Return (alpha2, display name); ('', original) when unrecognized."""
    text = " ".join(str(value or "").split()).strip(" .,")
    if not text:
        return "", ""
    key = text.casefold()
    for candidate in (key, key.replace(".", "").replace(" ", "") if len(key) <= 8 else key):
        if candidate in _COUNTRY_ALIASES:
            code = _COUNTRY_ALIASES[candidate]
            return code, country_name(code)
    if key.startswith("the "):
        key = key[4:]
    try:
        record = pycountry.countries.lookup(key)
        return record.alpha_2, country_name(record.alpha_2)
    except LookupError:
        pass
    try:
        matches = pycountry.countries.search_fuzzy(text)
        if len(matches) == 1 or (matches and len(text) > 5):
            return matches[0].alpha_2, country_name(matches[0].alpha_2)
    except LookupError:
        pass
    return "", text.title() if text.islower() or text.isupper() else text


@lru_cache(maxsize=256)
def _subdivisions(alpha2: str) -> dict:
    table = {}
    for record in pycountry.subdivisions.get(country_code=alpha2) or []:
        name = record.name
        suffix = record.code.split("-", 1)[1]
        table[name.casefold()] = name
        table[suffix.casefold()] = name
        ascii_name = name.encode("ascii", "ignore").decode()
        table.setdefault(ascii_name.casefold(), name)
    if alpha2 == "US":
        table.update({"d.c.": "District of Columbia", "washington dc": "District of Columbia",
                      "washington d.c.": "District of Columbia"})
    return table


@lru_cache(maxsize=8192)
def normalize_state(value: str, alpha2: str = "") -> str:
    """Full subdivision name for a state/province name or code (e.g. 'ca' -> 'California')."""
    text = " ".join(str(value or "").split()).strip(" .,")
    if not text:
        return ""
    key = text.casefold()
    for code in [alpha2] if alpha2 else ["US", "CA", "AU", "IN"]:
        found = _subdivisions(code).get(key)
        if found:
            return found
    return text.title() if (text.isupper() and len(text) > 3) or text.islower() else text


_CCTLD_SKIP = {"co", "io", "ai", "me", "tv", "fm", "ly", "to", "cc", "ws", "gg", "app", "dev"}


def country_from_domain(domain: str) -> str:
    """Country hint from a country-code TLD (e.g. .uk, .de); generic TLDs give ''."""
    tld = domain.rsplit(".", 1)[-1].casefold() if domain else ""
    if len(tld) != 2 or tld in _CCTLD_SKIP:
        return ""
    code = "GB" if tld == "uk" else tld.upper()
    return code if pycountry.countries.get(alpha_2=code) else ""


# --------------------------------------------------------------------------- dates
_ISO = re.compile(r"^\d{4}-\d{1,2}-\d{1,2}")


def parse_date(value, dayfirst: bool = False, today: date | None = None) -> date | None:
    today = today or date.today()
    result = None
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, datetime):
        result = value.date()
    elif isinstance(value, date):
        result = value
    elif isinstance(value, (int, float)):
        if value != value:
            return None
        if 1900 <= value <= today.year and float(value).is_integer():
            result = date(int(value), 7, 1)  # year only: assume mid-year
        elif 20000 <= value <= 80000:  # Excel serial date
            result = (datetime(1899, 12, 30) + timedelta(days=float(value))).date()
        elif value > 10**9:  # unix timestamp (s or ms)
            seconds = value / 1000 if value > 10**11 else value
            try:
                result = datetime.fromtimestamp(seconds, timezone.utc).date()
            except (OverflowError, OSError, ValueError):
                return None
    else:
        text = str(value).strip()
        if not text or text.casefold() in {"na", "n/a", "none", "null", "-", "unknown"}:
            return None
        if text.isdigit():
            return parse_date(int(text), dayfirst, today)
        try:
            if _ISO.match(text):
                result = datetime.fromisoformat(text.replace("Z", "+00:00").split("T")[0].split(" ")[0]).date()
            else:
                from dateutil import parser as date_parser

                result = date_parser.parse(text, dayfirst=dayfirst, default=datetime(today.year, 1, 1)).date()
        except (ValueError, OverflowError, TypeError):
            return None
    if result is None or result.year < 1900 or result > today + timedelta(days=366):
        return None
    return result


def age_years(dob: date | None, today: date | None = None) -> int | None:
    if dob is None:
        return None
    today = today or date.today()
    years = today.year - dob.year - ((today.month, today.day) < (dob.month, dob.day))
    return years if 0 <= years <= 120 else None


def parse_age(value) -> int | None:
    if value is None or isinstance(value, bool):
        return None
    try:
        number = float(str(value).strip().split()[0])
    except (ValueError, IndexError):
        return None
    return int(number) if 0 <= number <= 120 else None


def days_since(when: date | None, today: date | None = None) -> int | None:
    if when is None:
        return None
    return max(0, ((today or date.today()) - when).days)


def humanize_days(days: int | None) -> str:
    if days is None:
        return ""
    if days < 31:
        return f"{days} day{'s' if days != 1 else ''}"
    if days < 365:
        months = days // 30
        return f"{months} month{'s' if months != 1 else ''}"
    years, months = days // 365, (days % 365) // 30
    text = f"{years} year{'s' if years != 1 else ''}"
    return text + (f" {months} mo" if months else "")


def activity_status(days: int | None, active_days: int = 90, inactive_days: int = 365) -> str:
    if days is None:
        return "Unknown"
    if days <= active_days:
        return "Active"
    if days <= inactive_days:
        return "Lapsing"
    return "Inactive"


# --------------------------------------------------------------------------- domains
_EXTRACT = None
_EXTRACT_LOCK = threading.Lock()


def registrable_domain(host: str) -> str:
    """Company domain without subdomains, using the bundled Public Suffix List."""
    global _EXTRACT
    host = (host or "").strip().casefold().rstrip(".")
    if not host:
        return ""
    with _EXTRACT_LOCK:
        if _EXTRACT is None:
            import tldextract

            _EXTRACT = tldextract.TLDExtract(suffix_list_urls=(), cache_dir=None)
    parts = _EXTRACT(host)
    registered = getattr(parts, "top_domain_under_public_suffix", None) or parts.registered_domain
    return registered or host


def website_domain(value) -> str:
    text = clean_text(value).casefold()
    if not text or " " in text or "." not in text:
        return ""
    if "@" in text:
        text = text.rsplit("@", 1)[1]
    if "://" not in text:
        text = "http://" + text
    host = urlparse(text).hostname or ""
    return host[4:] if host.startswith("www.") else host


_BOOTSTRAP: dict | None = None
_BOOTSTRAP_LOCK = threading.Lock()


def _rdap_base(tld: str, session) -> str | None:
    global _BOOTSTRAP
    with _BOOTSTRAP_LOCK:
        if _BOOTSTRAP is None:
            try:
                data = session.get("https://data.iana.org/rdap/dns.json", timeout=(4, 10)).json()
                _BOOTSTRAP = {
                    suffix.casefold(): urls[0]
                    for suffixes, urls in data.get("services", [])
                    for suffix in suffixes
                    if urls
                }
            except Exception:
                _BOOTSTRAP = {}
    base = _BOOTSTRAP.get(tld.casefold())
    return base.rstrip("/") + "/" if base else None


def rdap_creation_date(domain: str, session=None) -> tuple[date | None, str]:
    """Registration date from RDAP (the WHOIS successor). Returns (date, status)."""
    import requests

    session = session or requests.Session()
    tld = domain.rsplit(".", 1)[-1]
    urls = []
    base = _rdap_base(tld, session)
    if base:
        urls.append(base + "domain/" + domain)
    urls.append("https://rdap.org/domain/" + domain)
    for url in urls:
        try:
            response = session.get(url, timeout=(4, 10), headers={"Accept": "application/rdap+json"})
        except Exception:
            continue
        if response.status_code == 404:
            return None, "Not registered"
        if response.status_code != 200:
            continue
        try:
            events = response.json().get("events", [])
        except ValueError:
            continue
        for event in events:
            if str(event.get("eventAction", "")).casefold() == "registration" and event.get("eventDate"):
                try:
                    return datetime.fromisoformat(event["eventDate"].replace("Z", "+00:00")).date(), "OK"
                except ValueError:
                    break
        return None, "Registration date not published"
    return None, "Lookup unavailable for this domain"
