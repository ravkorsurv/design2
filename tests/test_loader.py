from datetime import date
from pathlib import Path

import pytest

from dqstor.loader import read_incidents


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

