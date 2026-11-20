"""Output schema and source-column auto-detection."""

from __future__ import annotations

from typing import Iterable

# Columns appended to every exported row, in this order.
RESULT_COLUMNS = [
    "Lead_Verdict",
    "Lead_Score",
    "Lead_Issues",
    "Lead_Name",
    "Lead_Company",
    "Lead_Country",
    "Lead_State",
    "Verification_Status",
    "Reason",
    "Mailbox_Check",
    "Normalized_Email",
    "Email_Domain",
    "DNS_Check",
    "MX_Host",
    "Is_Role",
    "Is_Disposable",
    "Is_Free_Provider",
    "Is_Duplicate",
    "Typo_Suggestion",
    "Phone_Status",
    "Phone_E164",
    "Phone_Type",
    "Phone_Country",
    "Phone_Carrier",
    "Phone_Location",
    "Phone_Line_Status",
    "Is_Duplicate_Phone",
    "Age",
    "Record_Age_Days",
    "Record_Age",
    "Last_Activity_Days",
    "Activity_Status",
    "Domain_Created",
    "Domain_Age_Years",
]

BOOLEAN_COLUMNS = frozenset(
    {"Is_Role", "Is_Disposable", "Is_Free_Provider", "Is_Duplicate", "Is_Duplicate_Phone"}
)
NUMERIC_COLUMNS = frozenset(
    {"Lead_Score", "Age", "Record_Age_Days", "Last_Activity_Days", "Domain_Age_Years"}
)

# Optional source fields the user can map. Order drives the mapping UI.
FIELD_LABELS = {
    "email": "Email",
    "phone": "Phone",
    "name": "Full name",
    "first_name": "First name",
    "last_name": "Last name",
    "company": "Company",
    "website": "Website / company domain",
    "country": "Country",
    "state": "State / province",
    "dob": "Date of birth",
    "age": "Age (already in sheet)",
    "created": "Lead created / signup date",
    "last_activity": "Last activity / contact date",
}

_EXACT = {
    "email": [
        {"email", "emailaddress", "emailid", "email1", "primaryemail", "workemail"},
        {"mail", "contactemail", "businessemail", "personalemail"},
    ],
    "phone": [
        {"phone", "phonenumber", "mobile", "mobilenumber", "mobilephone", "cell",
         "cellphone", "cellnumber", "telephone", "tel", "telephonenumber",
         "contactnumber", "phoneno", "mobileno", "whatsapp", "whatsappnumber",
         "primaryphone", "workphone", "directphone", "phone1"},
    ],
    "name": [
        {"name", "fullname", "contactname", "leadname", "clientname", "customername",
         "contactfullname", "personname"},
    ],
    "first_name": [{"firstname", "fname", "givenname", "first", "forename"}],
    "last_name": [{"lastname", "lname", "surname", "familyname", "last"}],
    "company": [
        {"company", "companyname", "organization", "organisation", "business",
         "businessname", "employer", "account", "accountname", "firm", "agency",
         "brokerage"},
    ],
    "website": [
        {"website", "web", "url", "site", "domain", "companywebsite", "companydomain",
         "homepage", "websiteurl", "companyurl", "weburl"},
    ],
    "country": [
        {"country", "countryname", "countrycode", "nation", "countryregion",
         "mailingcountry", "billingcountry"},
    ],
    "state": [
        {"state", "province", "stateprovince", "region", "statecode", "stateregion",
         "county", "mailingstate", "billingstate", "provincestate"},
    ],
    "dob": [{"dob", "dateofbirth", "birthdate", "birthday", "birthdt"}],
    "age": [{"age", "ageyears", "clientage", "customerage"}],
    "created": [
        {"created", "createdat", "createddate", "datecreated", "createdon",
         "signupdate", "signedup", "signupat", "registered", "registrationdate",
         "dateadded", "added", "addedon", "leaddate", "leadcreated", "joindate",
         "joined", "joinedat", "optindate", "subscribed", "subscribedat",
         "customersince", "clientsince", "firstseen", "datejoined", "date"},
    ],
    "last_activity": [
        {"lastactivity", "lastactivitydate", "lastactive", "lastlogin", "lastseen",
         "lastcontact", "lastcontacted", "lastcontactdate", "lastpurchase",
         "lastpurchasedate", "lastorder", "lastorderdate", "lastvisit", "lastengaged",
         "lastopened", "lastinteraction", "lasttouch", "lasttrade", "lasttradedate",
         "lastdeposit", "lastdepositdate", "updated", "updatedat", "lastmodified",
         "lastupdated", "modified", "modifiedat"},
    ],
}

# Substring rules used only when no exact alias matched.
_CONTAINS = {
    "email": ("email",),
    "phone": ("phone", "mobile"),
    "country": ("country",),
    "dob": ("birth",),
    "created": ("created", "signup", "registered"),
    "last_activity": ("lastactiv", "lastlogin", "lastcontact", "lastseen"),
    "website": ("website",),
    "company": ("company",),
}


def _norm(name) -> str:
    return "".join(ch for ch in str(name).strip().casefold() if ch.isalnum())


def detect_columns(columns: Iterable[str]) -> dict:
    """Return {field: column index} for every confidently detected field."""
    names = [_norm(column) for column in columns]
    found, used = {}, set()
    for field in FIELD_LABELS:
        index = None
        for aliases in _EXACT.get(field, []):
            index = next(
                (i for i, n in enumerate(names) if n in aliases and i not in used), None
            )
            if index is not None:
                break
        if index is None:
            for fragment in _CONTAINS.get(field, ()):
                index = next(
                    (i for i, n in enumerate(names) if fragment in n and i not in used),
                    None,
                )
                if index is not None:
                    break
        if index is not None:
            found[field] = index
            used.add(index)
    # A "Name" column next to first/last columns is usually the full name; keep all.
    return found


def detect_email_column(columns: Iterable[str]) -> int | None:
    return detect_columns(columns).get("email")
