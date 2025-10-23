from datetime import date
from pathlib import Path

import pytest

from dqstor.loader import read_history, read_incidents


def write_tmp(tmp_path: Path, content: str) -> Path:
    path = tmp_path / "incidents.csv"
    path.write_text(content, encoding="utf-8")
    return path


def test_parses_multiple_alerts(tmp_path: Path):
    csv_content = """Incident_ID,Incident_Date,Records_Impacted,% Volume Impacted,Alert Impacted,Optional_Notes\n"""
    csv_content += """INC1,2024-01-01,100,27.5,\"Spoofing (10.2), Order Cancellations (4.3)\",Check\n"""
    path = write_tmp(tmp_path, csv_content)
    incidents = read_incidents(path)
    assert {i.model_scope for i in incidents} == {"Spoofing", "Order Cancellations"}
    spoofing = next(i for i in incidents if i.model_scope == "Spoofing")
    assert spoofing.incident_id == "INC1:Spoofing"
    assert spoofing.alert_impact == pytest.approx(10.2)
    assert spoofing.volume_pct_impacted == pytest.approx(27.5)
    assert spoofing.incident_date == date(2024, 1, 1)


def test_parses_alert_impacts_with_commas(tmp_path: Path):
    csv_content = """Incident_ID,Incident_Date,Records_Impacted,% Volume Impacted,Alert Impacted,Optional_Notes\n"""
    csv_content += """INC2,2024-02-01,10,5.0,\"Spoofing (1,200), Wash Trades (-3.5%)\",\n"""
    path = write_tmp(tmp_path, csv_content)
    incidents = read_incidents(path)
    assert len(incidents) == 2
    spoofing = next(i for i in incidents if i.model_scope == "Spoofing")
    assert spoofing.alert_impact == pytest.approx(1200)
    wash = next(i for i in incidents if i.model_scope == "Wash Trades")
    assert wash.alert_impact == pytest.approx(-3.5)


def test_history_validates_period_and_counts(tmp_path: Path):
    path = tmp_path / "history.csv"
    path.write_text(
        """Model_Scope,Period_Start,Period_End,Records_Observed,Alerts_Investigated,Stors_Filed\n"""
        """ModelA,2024-01-10,2024-01-01,100,5,1\n""",
        encoding="utf-8",
    )
    with pytest.raises(ValueError):
        read_history(path)

    path.write_text(
        """Model_Scope,Period_Start,Period_End,Records_Observed,Alerts_Investigated,Stors_Filed\n"""
        """ModelA,2024-01-01,2024-01-10,100,1,3\n""",
        encoding="utf-8",
    )
    with pytest.raises(ValueError):
        read_history(path)

