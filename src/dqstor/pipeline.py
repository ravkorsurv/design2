"""High level orchestration for the DQ/STOR pipeline."""
from __future__ import annotations

import csv
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Iterable, List

from .calculations import HistoryRollup, rollup_history, run_stage
from .config import Config, default_config
from .data_models import IncidentInput, IncidentResult
from .loader import read_history, read_incidents
from .utils import sha256_hexdigest


@dataclass
class RunContext:
    config: Config
    run_user: str
    run_timestamp: datetime
    history_rollup: dict


OUTPUT_HEADERS = [
    "Incident_ID",
    "Model_Scope",
    "Incident_Date",
    "Severity",
    "Records_Impacted",
    "Baseline_AlertRate",
    "Missed_Alerts",
    "Likelihood_Band",
    "DQ_Final_Risk",
    "Jeffreys_alpha",
    "Jeffreys_beta",
    "STOR_Rate_Mean",
    "STOR_Rate_95UCB",
    "Expected_Missed_STORs_Mean",
    "Expected_Missed_STORs_95UCB",
    "P_AtLeast_One_Missed_STOR_95UCB",
    "Run_Timestamp",
    "Run_User",
    "Workbook_Version",
    "Notes",
]


AUDIT_HEADERS = [
    "Run_Timestamp",
    "Run_User",
    "Workbook_Version",
    "Action",
    "Incidents_Count",
    "Output_Hash",
]


class Pipeline:
    def __init__(self, config: Config | None = None):
        self.config = config or default_config()

    def build_context(self, history_rows: Iterable, run_user: str) -> RunContext:
        run_timestamp = datetime.utcnow()
        window_start = run_timestamp.date() - self.config.lookback_delta()
        history_rollup = rollup_history(history_rows, window_start)
        return RunContext(
            config=self.config,
            run_user=run_user,
            run_timestamp=run_timestamp,
            history_rollup=history_rollup,
        )

    def run(self, incidents: List[IncidentInput], context: RunContext) -> List[IncidentResult]:
        return run_stage(
            incidents,
            context.history_rollup,
            context.config,
            context.run_timestamp,
            context.run_user,
        )


class OutputWriter:
    def __init__(self, output_path: Path, audit_path: Path):
        self.output_path = output_path
        self.audit_path = audit_path

    def _read_existing(self) -> List[dict]:
        if not self.output_path.exists():
            return []
        with self.output_path.open(newline="", encoding="utf-8") as handle:
            reader = csv.DictReader(handle)
            return list(reader)

    def _write_all(self, rows: List[dict]) -> None:
        with self.output_path.open("w", newline="", encoding="utf-8") as handle:
            writer = csv.DictWriter(handle, OUTPUT_HEADERS)
            writer.writeheader()
            for row in rows:
                writer.writerow(row)

    def write_results(self, results: List[IncidentResult], action: str, context: RunContext) -> None:
        existing = self._read_existing()
        index = {row["Incident_ID"]: row for row in existing}
        for result in results:
            index[result.incident_id] = result.as_dict()
        merged = list(index.values())
        merged.sort(key=lambda row: (row["Incident_Date"], row["Incident_ID"]))
        self._write_all(merged)
        self._append_audit(results, action, context)

    def _append_audit(self, results: List[IncidentResult], action: str, context: RunContext) -> None:
        hash_input = [result.incident_id + repr(result.as_dict()) for result in results]
        digest = sha256_hexdigest(hash_input)
        exists = self.audit_path.exists()
        with self.audit_path.open("a", newline="", encoding="utf-8") as handle:
            writer = csv.DictWriter(handle, AUDIT_HEADERS)
            if not exists:
                writer.writeheader()
            writer.writerow(
                {
                    "Run_Timestamp": context.run_timestamp.isoformat(),
                    "Run_User": context.run_user,
                    "Workbook_Version": context.config.version,
                    "Action": action,
                    "Incidents_Count": len(results),
                    "Output_Hash": digest,
                }
            )


def run_pipeline(
    incidents_csv: Path,
    history_csv: Path,
    output_csv: Path,
    audit_csv: Path,
    run_user: str,
    action: str = "RUN_ALL",
    config: Config | None = None,
) -> List[IncidentResult]:
    pipeline = Pipeline(config)
    incidents = read_incidents(incidents_csv)
    history_rows = read_history(history_csv)
    context = pipeline.build_context(history_rows, run_user)
    results = pipeline.run(incidents, context)
    writer = OutputWriter(output_csv, audit_csv)
    writer.write_results(results, action, context)
    return results

