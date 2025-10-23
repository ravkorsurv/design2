"""Loading and parsing utilities for incidents and historical data."""
from __future__ import annotations

import csv
import re
from pathlib import Path
from typing import Dict, Iterator, List

from .data_models import HistoryRow, IncidentInput
from .utils import parse_date


_ALERT_PATTERN = re.compile(
    r"(?:^|,)\s*(?P<model>[^(),]+?)\s*\((?P<value>[-+]?[0-9]*\.?[0-9]+)\)\s*"
)


def _parse_alert_impacts(cell: str) -> Iterator[Dict[str, str]]:
    for match in _ALERT_PATTERN.finditer(cell or ""):
        yield {
            "model": match.group("model").strip(),
            "value": match.group("value"),
        }

def read_incidents(csv_path: Path) -> List[IncidentInput]:
    incidents: List[IncidentInput] = []
    with csv_path.open(newline="", encoding="utf-8-sig") as handle:
        reader = csv.DictReader(handle)
        required = {"Incident_ID", "Incident_Date", "Records_Impacted", "% Volume Impacted", "Alert Impacted"}
        missing = required.difference(reader.fieldnames or [])
        if missing:
            raise ValueError(f"Incident file missing columns: {sorted(missing)}")
        for idx, row in enumerate(reader, start=2):
            try:
                incident_date = parse_date(row["Incident_Date"])
            except ValueError as exc:
                raise ValueError(f"Row {idx}: invalid Incident_Date '{row['Incident_Date']}': {exc}") from exc
            try:
                records_impacted = int(row["Records_Impacted"] or 0)
            except ValueError as exc:
                raise ValueError(f"Row {idx}: invalid Records_Impacted '{row['Records_Impacted']}'") from exc
            try:
                volume_pct = float(row["% Volume Impacted"] or 0.0)
            except ValueError as exc:
                raise ValueError(f"Row {idx}: invalid % Volume Impacted '{row['% Volume Impacted']}'") from exc
            alert_impacts = list(_parse_alert_impacts(row.get("Alert Impacted", "")))
            if not alert_impacts:
                raise ValueError(f"Row {idx}: no parsable alert impacts in '{row.get('Alert Impacted', '')}'")
            notes = row.get("Optional_Notes", "").strip()
            base_id = row["Incident_ID"].strip()
            if not base_id:
                raise ValueError(f"Row {idx}: missing Incident_ID")
            for alert in alert_impacts:
                try:
                    value = float(alert["value"])
                except ValueError as exc:
                    raise ValueError(f"Row {idx}: invalid alert impact value '{alert['value']}'") from exc
                incidents.append(
                    IncidentInput(
                        incident_id=f"{base_id}:{alert['model'].strip()}",
                        model_scope=alert["model"].strip(),
                        incident_date=incident_date,
                        records_impacted=records_impacted,
                        volume_pct_impacted=volume_pct,
                        alert_impact=value,
                        optional_notes=notes,
                    )
                )
    return incidents


def read_history(csv_path: Path) -> List[HistoryRow]:
    history: List[HistoryRow] = []
    with csv_path.open(newline="", encoding="utf-8-sig") as handle:
        reader = csv.DictReader(handle)
        required = {
            "Model_Scope",
            "Period_Start",
            "Period_End",
            "Records_Observed",
            "Alerts_Investigated",
            "Stors_Filed",
        }
        missing = required.difference(reader.fieldnames or [])
        if missing:
            raise ValueError(f"History file missing columns: {sorted(missing)}")
        for idx, row in enumerate(reader, start=2):
            try:
                period_start = parse_date(row["Period_Start"])
                period_end = parse_date(row["Period_End"])
            except ValueError as exc:
                raise ValueError(f"Row {idx}: invalid period dates") from exc
            try:
                records_observed = int(row["Records_Observed"] or 0)
                alerts_investigated = int(row["Alerts_Investigated"] or 0)
                stors_filed = int(row["Stors_Filed"] or 0)
            except ValueError as exc:
                raise ValueError(f"Row {idx}: invalid numeric history values") from exc
            history.append(
                HistoryRow(
                    model_scope=row["Model_Scope"].strip(),
                    period_start=period_start,
                    period_end=period_end,
                    records_observed=records_observed,
                    alerts_investigated=alerts_investigated,
                    stors_filed=stors_filed,
                )
            )
    return history

