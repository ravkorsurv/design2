"""Configuration models and defaults for the DQ/STOR pipeline."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import timedelta
from typing import Dict, List


@dataclass(frozen=True)
class LikelihoodBanding:
    """Thresholds for mapping numeric likelihood indicators into textual bands."""

    almost_certain: float = 10.0
    likely: float = 5.0
    possible: float = 1.0

    def band_for(self, value: float) -> str:
        if value >= self.almost_certain:
            return "Almost certain"
        if value >= self.likely:
            return "Likely"
        if value >= self.possible:
            return "Possible"
        return "Unlikely"


@dataclass(frozen=True)
class SeverityBanding:
    """Thresholds for translating percentage volume impact into severity labels."""

    major: float = 50.0
    significant: float = 25.0
    moderate: float = 10.0
    minor: float = 5.0

    def severity_for(self, pct_volume_impacted: float) -> str:
        if pct_volume_impacted >= self.major:
            return "major"
        if pct_volume_impacted >= self.significant:
            return "significant"
        if pct_volume_impacted >= self.moderate:
            return "moderate"
        if pct_volume_impacted >= self.minor:
            return "minor"
        return "minor"


@dataclass(frozen=True)
class Config:
    """Runtime configuration for the DQ/STOR pipeline."""

    lookback_days: int = 90
    dq_matrix: Dict[str, Dict[str, str]] = field(default_factory=dict)
    likelihood_thresholds: LikelihoodBanding = field(default_factory=LikelihoodBanding)
    severity_thresholds: SeverityBanding = field(default_factory=SeverityBanding)
    version: str = "1.0.0"

    def lookback_delta(self) -> timedelta:
        return timedelta(days=self.lookback_days)


DEFAULT_DQ_MATRIX: Dict[str, Dict[str, str]] = {
    "major": {
        "Almost certain": "High",
        "Likely": "High",
        "Possible": "High",
        "Unlikely": "Medium",
    },
    "significant": {
        "Almost certain": "High",
        "Likely": "High",
        "Possible": "High",
        "Unlikely": "Medium",
    },
    "moderate": {
        "Almost certain": "Medium",
        "Likely": "Medium",
        "Possible": "Medium",
        "Unlikely": "Medium",
    },
    "minor": {
        "Almost certain": "Medium",
        "Likely": "Medium",
        "Possible": "Medium",
        "Unlikely": "Low",
    },
}


def default_config() -> Config:
    return Config(dq_matrix=DEFAULT_DQ_MATRIX)
