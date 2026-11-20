import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import pytest


@pytest.fixture(autouse=True)
def reset_smtp_breaker():
    from verifier import smtp_check

    smtp_check._UNREACHABLE.clear()
    smtp_check._FAILURES.clear()
    yield
    smtp_check._UNREACHABLE.clear()
    smtp_check._FAILURES.clear()
