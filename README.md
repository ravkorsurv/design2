# DQ/STOR Pipeline

This repository provides a Python implementation of the data-quality (DQ) final risk
and suspicious transaction/ordering report (STOR) materiality workflow described in
the functional specification. The solution ingests weekly incident extracts and
90-day surveillance funnel history, computes the DQ final risk using deterministic
severity and likelihood rules, and derives Jeffreys posterior statistics for
missed STOR expectations. Results are persisted to append-only output and audit
CSV files mirroring the original Excel workbook design.

## Features

- Parses the weekly incidents CSV, including the composite `Alert Impacted` column
  (for example `"Spoofing (10.2), Order Cancellations (4.3)"`) into one incident row per
  model/scope.
- Derives severity from the `% Volume Impacted` column using the agreed thresholds:
  `>=50` → major, `>=25` → significant, `>=10` → moderate, `>=5` → minor, otherwise minor.
- Maps the parsed alert impact values straight into likelihood bands:
  `>=10` almost certain, `>=5` likely, `>=1` possible, `<1` unlikely.
- Computes baseline alert rates, missed alerts, Jeffreys posterior mean and 95% upper
  confidence bound, and the Poisson approximation of missing at least one STOR.
- Maintains append-only output and audit CSVs including a SHA-256 hash of each run’s
  written rows for tamper evidence.
- Offers a command-line interface mirroring the three macro entry points via the
  `--action` flag (`DQ_ONLY`, `STOR_ONLY`, `RUN_ALL`).

## Getting started

1. Install the runtime dependencies:

   ```bash
   pip install -r requirements.txt
   ```

2. Prepare the input files:
   - `incidents.csv` with headers `Incident_ID`, `Incident_Date`, `Records_Impacted`,
     `% Volume Impacted`, `Alert Impacted`, and an optional `Optional_Notes` column.
   - `history.csv` with headers `Model_Scope`, `Period_Start`, `Period_End`,
     `Records_Observed`, `Alerts_Investigated`, `Stors_Filed`.

3. Execute the pipeline:

   ```bash
   python -m dqstor.cli incidents.csv history.csv output.csv audit.csv --user alice --action RUN_ALL
   ```

   The command writes/updates `output.csv` with one row per incident/model scope and
   appends a run entry to `audit.csv`.

## Project structure

- `src/dqstor/config.py` – default configuration, severity and likelihood thresholds.
- `src/dqstor/loader.py` – CSV parsing, including the alert-impact extractor.
- `src/dqstor/calculations.py` – DQ and STOR calculation functions.
- `src/dqstor/pipeline.py` – orchestration, persistence, and audit logging.
- `src/dqstor/cli.py` – simple CLI for batch execution.

## Testing

Unit tests can be added under `tests/` using `pytest`.

## License

MIT
