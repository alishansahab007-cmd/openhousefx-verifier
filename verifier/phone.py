"""Phone number validation (offline) and optional live line-status lookup (Twilio)."""

from __future__ import annotations

import os
import re
from dataclasses import dataclass
from functools import lru_cache
from urllib.parse import quote

import phonenumbers
from phonenumbers import PhoneNumberType, carrier, geocoder

TYPE_LABELS = {
    PhoneNumberType.MOBILE: "Mobile",
    PhoneNumberType.FIXED_LINE: "Landline",
    PhoneNumberType.FIXED_LINE_OR_MOBILE: "Landline or mobile",
    PhoneNumberType.TOLL_FREE: "Toll-free",
    PhoneNumberType.PREMIUM_RATE: "Premium rate",
    PhoneNumberType.SHARED_COST: "Shared cost",
    PhoneNumberType.VOIP: "VoIP",
    PhoneNumberType.PERSONAL_NUMBER: "Personal number",
    PhoneNumberType.PAGER: "Pager",
    PhoneNumberType.UAN: "Company (UAN)",
    PhoneNumberType.VOICEMAIL: "Voicemail",
    PhoneNumberType.UNKNOWN: "Unknown",
}
_JUNK = re.compile(r"^(\d)\1{6,}$|^0123456789|^1234567890|^9876543210")


@dataclass(frozen=True)
class PhoneResult:
    status: str  # Valid / Invalid / Missing
    e164: str = ""
    type: str = ""
    country: str = ""  # ISO alpha-2
    carrier: str = ""
    location: str = ""
    reason: str = ""


def _text(value) -> str:
    if value is None or isinstance(value, bool):
        return ""
    if isinstance(value, float):
        if value != value:  # NaN
            return ""
        return str(int(value)) if value.is_integer() else str(value)
    return str(value).strip()


def supported_regions() -> list[str]:
    return sorted(phonenumbers.SUPPORTED_REGIONS)


@lru_cache(maxsize=65536)
def _parse_cached(text: str, region: str) -> PhoneResult:
    candidate = text
    if re.match(r"^\s*00[1-9]", candidate):
        candidate = "+" + candidate.lstrip()[2:]
    number = None
    try:
        number = phonenumbers.parse(candidate, region or None)
    except phonenumbers.NumberParseException:
        match = next(iter(phonenumbers.PhoneNumberMatcher(candidate, region or "US")), None)
        number = match.number if match else None
    digits = re.sub(r"\D", "", candidate)
    if number is not None and not phonenumbers.is_valid_number(number) and not candidate.lstrip().startswith("+"):
        # Excel often drops the leading "+" from international numbers.
        if len(digits) >= 11:
            try:
                alternative = phonenumbers.parse("+" + digits, None)
                if phonenumbers.is_valid_number(alternative):
                    number = alternative
            except phonenumbers.NumberParseException:
                pass
    if number is None:
        return PhoneResult("Invalid", reason="Not a recognizable phone number")
    national = str(number.national_number)
    if _JUNK.match(national):
        return PhoneResult("Invalid", reason="Placeholder / repeated-digit number")
    if not phonenumbers.is_valid_number(number):
        reason = (
            "Number has the right length but is not an assigned number range"
            if phonenumbers.is_possible_number(number)
            else "Wrong length for this country"
        )
        possible = phonenumbers.is_possible_number(number)
        return PhoneResult(
            "Invalid",
            e164=phonenumbers.format_number(number, phonenumbers.PhoneNumberFormat.E164) if possible else "",
            country=(phonenumbers.region_code_for_number(number) or "") if possible else "",
            reason=reason,
        )
    kind = phonenumbers.number_type(number)
    return PhoneResult(
        "Valid",
        e164=phonenumbers.format_number(number, phonenumbers.PhoneNumberFormat.E164),
        type=TYPE_LABELS.get(kind, "Unknown"),
        country=phonenumbers.region_code_for_number(number) or "",
        carrier=carrier.name_for_number(number, "en") or "",
        location=geocoder.description_for_number(number, "en") or "",
    )


def parse_phone(value, region: str = "US") -> PhoneResult:
    text = _text(value)
    if not text or not re.search(r"\d", text):
        return PhoneResult("Missing", reason="No phone number")
    return _parse_cached(text[:64], (region or "").upper())


# --------------------------------------------------------------------------- Twilio
def twilio_ready() -> bool:
    return bool(os.getenv("TWILIO_ACCOUNT_SID") and os.getenv("TWILIO_AUTH_TOKEN"))


@dataclass(frozen=True)
class LineResult:
    status: str  # Active / Inactive / Unreachable / Invalid / Unknown / Error
    line_type: str = ""
    carrier: str = ""


def twilio_line_lookup(e164: str, session=None, timeout=(4, 10)) -> LineResult:
    """Query Twilio Lookup v2 (paid per lookup) for line status and line type."""
    import requests

    sid, token = os.getenv("TWILIO_ACCOUNT_SID", ""), os.getenv("TWILIO_AUTH_TOKEN", "")
    fields = os.getenv("TWILIO_LOOKUP_FIELDS", "line_type_intelligence,line_status")
    http = session or requests
    url = "https://lookups.twilio.com/v2/PhoneNumbers/" + quote(e164, safe="+")
    try:
        response = http.get(url, params={"Fields": fields}, auth=(sid, token), timeout=timeout)
        if response.status_code in (400, 403) and "line_status" in fields:
            response = http.get(url, params={"Fields": "line_type_intelligence"}, auth=(sid, token), timeout=timeout)
        if response.status_code == 404:
            return LineResult("Invalid")
        response.raise_for_status()
        data = response.json()
    except Exception as error:  # network, auth, JSON
        return LineResult("Error", carrier=type(error).__name__)
    if data.get("valid") is False:
        return LineResult("Invalid")
    intel = data.get("line_type_intelligence") or {}
    line_type = (intel.get("type") or "").replace("nonFixedVoip", "VoIP (non-fixed)").replace("fixedVoip", "VoIP (fixed)")
    status_block = data.get("line_status") or {}
    status = (status_block.get("status") or "").strip().lower()
    label = {"active": "Active", "inactive": "Inactive", "unreachable": "Unreachable"}.get(status)
    if label is None:
        label = "Valid (status not reported)" if data.get("valid") else "Unknown"
    return LineResult(label, line_type.title() if line_type else "", intel.get("carrier_name") or "")
