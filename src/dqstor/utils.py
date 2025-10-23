"""Utility helpers for the DQ/STOR pipeline."""
from __future__ import annotations

from datetime import datetime
from hashlib import sha256
from typing import Iterable


_DATE_FORMATS = [
    "%Y-%m-%d",
    "%d/%m/%Y",
    "%m/%d/%Y",
]


def parse_date(value: str):
    value = (value or "").strip()
    if not value:
        raise ValueError("blank date")
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"could not parse date '{value}'")


def sha256_hexdigest(rows: Iterable[str]) -> str:
    digest = sha256()
    for row in rows:
        digest.update(row.encode("utf-8"))
    return digest.hexdigest()


class PipelineError(RuntimeError):
    """Raised when the pipeline encounters a fatal error."""

