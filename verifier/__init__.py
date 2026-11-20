"""OpenhouseFX lead verification package (engine is independent of the UI)."""

from .columns import FIELD_LABELS, RESULT_COLUMNS, detect_columns, detect_email_column
from .email_checks import DomainResult, lookup_domain, validate_dns_configuration
from .engine import (
    COUNT_KEYS,
    STATUS_FILTERS,
    VerifyOptions,
    distinct_values,
    export_view_csv,
    generate_organized_xlsx,
    generate_xlsx,
    group_counts,
    iter_results,
    load_summary,
    read_results,
    run_verification,
)
from .io import inspect_input
from .smtp_check import MailboxResult

__version__ = "3.0.0"
