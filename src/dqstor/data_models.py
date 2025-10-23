"""Core data structures for the DQ/STOR calculations."""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Dict, Iterable, List, Optional


@dataclass(frozen=True)
class IncidentInput:
    incident_id: str
    model_scope: str
    incident_date: date
    records_impacted: int
    volume_pct_impacted: float
    alert_impact: float
    optional_notes: str = ""


@dataclass(frozen=True)
class HistoryRow:
    model_scope: str
    period_start: date
    period_end: date
    records_observed: int
    alerts_investigated: int
    stors_filed: int


@dataclass
class IncidentResult:
    incident_id: str
    model_scope: str
    incident_date: date
    severity: str
    records_impacted: int
    baseline_alert_rate: float
    missed_alerts: float
    likelihood_band: str
    dq_final_risk: str
    jeffreys_alpha: float
    jeffreys_beta: float
    stor_rate_mean: float
    stor_rate_95ucb: float
    expected_missed_stors_mean: float
    expected_missed_stors_95ucb: float
    p_at_least_one_missed_stor_95ucb: float
    run_timestamp: str
    run_user: str
    workbook_version: str
    notes: Optional[str] = None

    def as_dict(self) -> Dict[str, object]:
        return {
            "Incident_ID": self.incident_id,
            "Model_Scope": self.model_scope,
            "Incident_Date": self.incident_date.isoformat(),
            "Severity": self.severity,
            "Records_Impacted": self.records_impacted,
            "Baseline_AlertRate": self.baseline_alert_rate,
            "Missed_Alerts": self.missed_alerts,
            "Likelihood_Band": self.likelihood_band,
            "DQ_Final_Risk": self.dq_final_risk,
            "Jeffreys_alpha": self.jeffreys_alpha,
            "Jeffreys_beta": self.jeffreys_beta,
            "STOR_Rate_Mean": self.stor_rate_mean,
            "STOR_Rate_95UCB": self.stor_rate_95ucb,
            "Expected_Missed_STORs_Mean": self.expected_missed_stors_mean,
            "Expected_Missed_STORs_95UCB": self.expected_missed_stors_95ucb,
            "P_AtLeast_One_Missed_STOR_95UCB": self.p_at_least_one_missed_stor_95ucb,
            "Run_Timestamp": self.run_timestamp,
            "Run_User": self.run_user,
            "Workbook_Version": self.workbook_version,
            "Notes": self.notes or "",
        }

