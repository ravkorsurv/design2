from datetime import date, datetime

import pytest

from dqstor.calculations import HistoryRollup, rollup_history, run_stage
from dqstor.config import default_config
from dqstor.data_models import HistoryRow, IncidentInput


def make_incident(**kwargs):
    defaults = dict(
        incident_id="INC1:Spoofing",
        model_scope="Spoofing",
        incident_date=date(2024, 1, 1),
        records_impacted=120,
        volume_pct_impacted=55.0,
        alert_impact=12.0,
        optional_notes="",
    )
    defaults.update(kwargs)
    return IncidentInput(**defaults)


def test_rollup_history_filters_window():
    history = [
        HistoryRow("Spoofing", date(2023, 9, 1), date(2023, 9, 30), 1000, 200, 5),
        HistoryRow("Spoofing", date(2022, 1, 1), date(2022, 1, 31), 500, 100, 2),
    ]
    window_start = date(2023, 7, 1)
    rollup = rollup_history(history, window_start)
    assert rollup["Spoofing"].total_records == 1000
    assert rollup["Spoofing"].total_alerts == 200
    assert rollup["Spoofing"].total_stors == 5


def test_run_stage_computes_expected_fields():
    config = default_config()
    history_rollup = {
        "Spoofing": HistoryRollup(total_records=1000, total_alerts=200, total_stors=5)
    }
    incident = make_incident()
    results = run_stage([incident], history_rollup, config, datetime(2024, 1, 2), "alice")
    result = results[0]
    assert result.baseline_alert_rate == 0.2
    assert result.missed_alerts == 24.0
    assert result.likelihood_band == "Almost certain"
    assert result.dq_final_risk == "High"
    assert result.jeffreys_alpha == 5.5
    assert result.jeffreys_beta == 195.5
    assert result.stor_rate_mean == pytest.approx(5.5 / (5.5 + 195.5))
    assert result.notes is None

