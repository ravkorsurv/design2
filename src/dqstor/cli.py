"""Command line interface for executing the DQ/STOR pipeline."""
from __future__ import annotations

import argparse
from pathlib import Path

from .config import default_config
from .pipeline import run_pipeline


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Compute DQ risk and STOR impact metrics")
    parser.add_argument("incidents", type=Path, help="Path to incidents CSV extract")
    parser.add_argument("history", type=Path, help="Path to history 90-day CSV extract")
    parser.add_argument("output", type=Path, help="Path to write the output table CSV")
    parser.add_argument("audit", type=Path, help="Path to append the audit log CSV")
    parser.add_argument("--user", default="auto", help="Run user for audit logging")
    parser.add_argument("--action", default="RUN_ALL", choices=["DQ_ONLY", "STOR_ONLY", "RUN_ALL"], help="Action code to log")
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    run_pipeline(
        incidents_csv=args.incidents,
        history_csv=args.history,
        output_csv=args.output,
        audit_csv=args.audit,
        run_user=args.user,
        action=args.action,
        config=default_config(),
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
