"""High level orchestration for the DQ/STOR pipeline."""
from __future__ import annotations

import csv
import json
import os
import tempfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, Iterator, List, MutableMapping

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

    def run(
        self,
        incidents: List[IncidentInput],
        context: RunContext,
        action: str = "RUN_ALL",
    ) -> List[IncidentResult]:
        normalized = (action or "RUN_ALL").upper()
        if normalized not in {"DQ_ONLY", "STOR_ONLY", "RUN_ALL"}:
            raise ValueError(f"Unknown action '{action}'")
        return run_stage(
            incidents,
            context.history_rollup,
            context.config,
            context.run_timestamp,
            context.run_user,
        )


class OutputWriter:
    _GENERAL_COLUMNS = {
        "Incident_ID",
        "Model_Scope",
        "Incident_Date",
        "Records_Impacted",
        "Baseline_AlertRate",
        "Missed_Alerts",
        "Run_Timestamp",
        "Run_User",
        "Workbook_Version",
        "Notes",
    }
    _DQ_COLUMNS = {"Severity", "Likelihood_Band", "DQ_Final_Risk"}
    _STOR_COLUMNS = {
        "Jeffreys_alpha",
        "Jeffreys_beta",
        "STOR_Rate_Mean",
        "STOR_Rate_95UCB",
        "Expected_Missed_STORs_Mean",
        "Expected_Missed_STORs_95UCB",
        "P_AtLeast_One_Missed_STOR_95UCB",
    }

    def __init__(self, output_path: Path, audit_path: Path):
        self.output_path = output_path
        self.audit_path = audit_path

    def _columns_for_action(self, action: str) -> set[str]:
        normalized = (action or "RUN_ALL").upper()
        if normalized == "DQ_ONLY":
            return self._GENERAL_COLUMNS | self._DQ_COLUMNS
        if normalized == "STOR_ONLY":
            return self._GENERAL_COLUMNS | self._STOR_COLUMNS
        return self._GENERAL_COLUMNS | self._DQ_COLUMNS | self._STOR_COLUMNS

    def _iter_existing(self) -> Iterator[dict]:
        if not self.output_path.exists():
            return
        with self.output_path.open(newline="", encoding="utf-8") as handle:
            reader = csv.DictReader(handle)
            yield from reader

    def _merge_row(
        self,
        base_row: MutableMapping[str, object],
        updates: Dict[str, object],
        columns_to_update: set[str],
    ) -> Dict[str, object]:
        merged: Dict[str, object] = {header: base_row.get(header, "") for header in OUTPUT_HEADERS}
        for column in columns_to_update:
            if column in updates:
                merged[column] = updates[column]
        return merged

    def write_results(self, results: List[IncidentResult], action: str, context: RunContext) -> None:
        columns_to_update = self._columns_for_action(action)
        pending: Dict[str, Dict[str, object]] = {
            result.incident_id: result.as_dict() for result in results
        }
        output_dir = self.output_path.parent
        if output_dir and not output_dir.exists():
            output_dir.mkdir(parents=True, exist_ok=True)
        temp_dir = output_dir or Path(".")
        temp_file = tempfile.NamedTemporaryFile(
            "w",
            newline="",
            encoding="utf-8",
            delete=False,
            dir=temp_dir,
        )
        try:
            with temp_file:
                writer = csv.DictWriter(temp_file, OUTPUT_HEADERS)
                writer.writeheader()
                for existing_row in self._iter_existing():
                    incident_id = existing_row.get("Incident_ID", "")
                    if incident_id in pending:
                        updated = self._merge_row(existing_row, pending.pop(incident_id), columns_to_update)
                    else:
                        updated = {header: existing_row.get(header, "") for header in OUTPUT_HEADERS}
                    writer.writerow(updated)
                def _sort_key(item: tuple[str, Dict[str, object]]) -> tuple[str, str]:
                    incident_updates = item[1]
                    return (
                        str(incident_updates.get("Incident_Date", "")),
                        item[0],
                    )

                for incident_id, updates in sorted(pending.items(), key=_sort_key):
                    new_row = {header: "" for header in OUTPUT_HEADERS}
                    new_row["Incident_ID"] = incident_id
                    merged = self._merge_row(new_row, updates, columns_to_update)
                    writer.writerow(merged)
            os.replace(temp_file.name, self.output_path)
        finally:
            if os.path.exists(temp_file.name):
                os.unlink(temp_file.name)
        self._append_audit(results, action, context)

    def _append_audit(self, results: List[IncidentResult], action: str, context: RunContext) -> None:
        hash_input = [
            json.dumps(result.as_dict(), sort_keys=True, ensure_ascii=False, separators=(",", ":"))
            for result in results
        ]
        digest = sha256_hexdigest(hash_input)
        exists = self.audit_path.exists()
        audit_dir = self.audit_path.parent
        if audit_dir and not audit_dir.exists():
            audit_dir.mkdir(parents=True, exist_ok=True)
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
    results = pipeline.run(incidents, context, action)
    writer = OutputWriter(output_csv, audit_csv)
    writer.write_results(results, action, context)
    return results

