from __future__ import annotations

import json
import re
from collections import OrderedDict
from pathlib import Path
from typing import Dict, List

import pytest

ROOT = Path(__file__).resolve().parent.parent
BAS_PATH = ROOT / "excel" / "DQMateriality.bas"
DATA_PATH = Path(__file__).resolve().parent / "data" / "dqmateriality_scenarios.json"


def _load_regex_pattern() -> re.Pattern[str]:
    text = BAS_PATH.read_text(encoding="utf-8")
    match = re.search(r'scenarioRegex\.Pattern = "([^"]+)"', text)
    if not match:
        raise AssertionError("Could not locate scenario regex pattern in DQMateriality.bas")
    literal = match.group(1)
    return re.compile(literal, flags=re.IGNORECASE)


SCENARIO_REGEX = _load_regex_pattern()


def _normalize(value: str) -> str:
    text = str(value).replace("\r\n", " ").replace("\t", " ")
    text = text.strip()
    while "  " in text:
        text = text.replace("  ", " ")
    return text


def _split_scenarios(value: str) -> List[str]:
    trimmed = value.strip()
    if not trimmed:
        return []
    results: List[str] = []
    for part in trimmed.split(","):
        candidate = _normalize(part)
        if candidate:
            results.append(candidate)
    return results


def _nz_double(value: str) -> float:
    candidate = value.strip()
    if not candidate:
        return 0.0
    candidate = candidate.replace(",", "").replace("%", "")
    candidate = candidate.replace("\u2212", "-")  # minus sign
    try:
        return float(candidate)
    except ValueError as exc:  # pragma: no cover - helps debugging unexpected formats
        raise AssertionError(f"Unable to convert '{value}' to float") from exc


def parse_scenario_materiality(
    scenarios_text: str, potential_text: str, fallback_scenario: str
) -> List[Dict[str, float]]:
    scenario_set: "OrderedDict[str, str]" = OrderedDict()
    potential_dict: Dict[str, float] = {}

    for match in SCENARIO_REGEX.finditer(potential_text):
        scenario = _normalize(match.group(1))
        missing = _nz_double(match.group(2))
        key = scenario.lower()
        potential_dict[key] = potential_dict.get(key, 0.0) + missing
        if key not in scenario_set:
            scenario_set[key] = scenario

    for scenario in _split_scenarios(scenarios_text):
        key = scenario.lower()
        if key not in scenario_set:
            scenario_set[key] = scenario

    if not scenario_set:
        fallback = _normalize(fallback_scenario)
        if not fallback:
            fallback = "Unspecified"
        scenario_set[fallback.lower()] = fallback

    results = []
    for key, name in scenario_set.items():
        results.append({"Scenario": name, "MissingAlerts": potential_dict.get(key, 0.0)})
    return results


@pytest.mark.parametrize("case", json.loads(DATA_PATH.read_text(encoding="utf-8")), ids=lambda c: c["name"])
def test_parse_scenario_materiality_cases(case: Dict[str, object]) -> None:
    actual = parse_scenario_materiality(
        str(case["scenarios_text"]), str(case["potential_text"]), str(case["fallback_scenario"])
    )
    expected: List[Dict[str, float]] = list(case["expected"])
    assert len(actual) == len(expected)

    actual_sorted = sorted(actual, key=lambda item: item["Scenario"].lower())
    expected_sorted = sorted(expected, key=lambda item: item["Scenario"].lower())

    for result, exp in zip(actual_sorted, expected_sorted):
        assert result["Scenario"] == exp["Scenario"]
        assert result["MissingAlerts"] == pytest.approx(exp["MissingAlerts"], rel=1e-9, abs=1e-9)


def test_regex_pattern_matches_expected_structure() -> None:
    assert SCENARIO_REGEX.pattern == r"([^\(]+?)\s*\(([^\)]+)\)"
