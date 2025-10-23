"""Computation primitives for the DQ and STOR stages."""
from __future__ import annotations

import math
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from typing import Dict, Iterable, List, Tuple

from .config import Config
from .data_models import HistoryRow, IncidentInput, IncidentResult


@dataclass
class HistoryRollup:
    total_records: int = 0
    total_alerts: int = 0
    total_stors: int = 0

    def baseline_rate(self) -> float:
        return self.total_alerts / self.total_records if self.total_records else 0.0


def rollup_history(history: Iterable[HistoryRow], window_start: date) -> Dict[str, HistoryRollup]:
    rollup: Dict[str, HistoryRollup] = defaultdict(HistoryRollup)
    for row in history:
        if row.period_end < window_start:
            continue
        bucket = rollup[row.model_scope]
        bucket.total_records += row.records_observed
        bucket.total_alerts += row.alerts_investigated
        bucket.total_stors += row.stors_filed
    return rollup


def compute_dq(incident: IncidentInput, history_rollup: HistoryRollup, config: Config) -> Tuple[float, str, str]:
    baseline_alert_rate = history_rollup.baseline_rate()
    missed_alerts = incident.records_impacted * baseline_alert_rate
    likelihood_band = config.likelihood_thresholds.band_for(incident.alert_impact)
    severity = config.severity_thresholds.severity_for(incident.volume_pct_impacted)
    dq_matrix = config.dq_matrix or {}
    severity_row = dq_matrix.get(severity, {})
    dq_final = severity_row.get(likelihood_band, "Medium")
    return missed_alerts, likelihood_band, dq_final


def jeffreys_posterior(k: int, n: int) -> Tuple[float, float]:
    alpha = k + 0.5
    beta = (n - k) + 0.5
    p_mean = alpha / (alpha + beta)
    p95 = beta_inverse(alpha, beta, 0.95)
    return p_mean, p95


def beta_inverse(alpha: float, beta: float, quantile: float) -> float:
    try:
        import mpmath  # type: ignore
    except ModuleNotFoundError:
        return _beta_inverse_numeric(alpha, beta, quantile)
    return float(mpmath.betaincinv(alpha, beta, 0, quantile))


def _beta_inverse_numeric(alpha: float, beta: float, quantile: float) -> float:
    if not 0.0 <= quantile <= 1.0:
        raise ValueError("quantile must be within [0, 1]")
    if quantile == 0.0:
        return 0.0
    if quantile == 1.0:
        return 1.0

    def regularized_beta(x: float) -> float:
        return _regularized_incomplete_beta(alpha, beta, x)

    lo, hi = 0.0, 1.0
    for _ in range(80):
        mid = (lo + hi) / 2.0
        cdf = regularized_beta(mid)
        if abs(cdf - quantile) < 1e-8:
            return mid
        if cdf < quantile:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2.0


def _regularized_incomplete_beta(a: float, b: float, x: float) -> float:
    if x <= 0.0:
        return 0.0
    if x >= 1.0:
        return 1.0
    bt = math.exp(
        math.lgamma(a + b)
        - math.lgamma(a)
        - math.lgamma(b)
        + a * math.log(x)
        + b * math.log1p(-x)
    )
    if x < (a + 1.0) / (a + b + 2.0):
        return bt * _betacf(a, b, x) / a
    return 1.0 - bt * _betacf(b, a, 1.0 - x) / b


def _betacf(a: float, b: float, x: float) -> float:
    MAX_ITER = 200
    EPS = 3e-8
    FPMIN = 1e-30
    c = 1.0
    d = 1.0 - (a + b) * x / (a + 1.0)
    if abs(d) < FPMIN:
        d = FPMIN
    d = 1.0 / d
    h = d
    for m in range(1, MAX_ITER + 1):
        m2 = 2 * m
        aa = m * (b - m) * x / ((a - 1.0 + m2) * (a + m2))
        d = 1.0 + aa * d
        if abs(d) < FPMIN:
            d = FPMIN
        c = 1.0 + aa / c
        if abs(c) < FPMIN:
            c = FPMIN
        d = 1.0 / d
        h *= d * c
        aa = -(a + m) * (a + b + m) * x / ((a + m2) * (a + 1.0 + m2))
        d = 1.0 + aa * d
        if abs(d) < FPMIN:
            d = FPMIN
        c = 1.0 + aa / c
        if abs(c) < FPMIN:
            c = FPMIN
        d = 1.0 / d
        delta = d * c
        h *= delta
        if abs(delta - 1.0) < EPS:
            break
    return h


def poisson_prob_at_least_one(lmbda: float) -> float:
    return 1.0 - math.exp(-lmbda)


def run_stage(
    incidents: Iterable[IncidentInput],
    history_rollup: Dict[str, HistoryRollup],
    config: Config,
    run_timestamp: datetime,
    run_user: str,
) -> List[IncidentResult]:
    results: List[IncidentResult] = []
    for incident in incidents:
        rollup = history_rollup.get(incident.model_scope, HistoryRollup())
        missed_alerts, likelihood_band, dq_final = compute_dq(incident, rollup, config)
        baseline_rate = rollup.baseline_rate()
        stor_mean, stor_95 = jeffreys_posterior(rollup.total_stors, rollup.total_alerts)
        expected_mean = missed_alerts * stor_mean
        expected_95 = missed_alerts * stor_95
        prob_at_least_one = poisson_prob_at_least_one(expected_95)
        results.append(
            IncidentResult(
                incident_id=incident.incident_id,
                model_scope=incident.model_scope,
                incident_date=incident.incident_date,
                severity=config.severity_thresholds.severity_for(incident.volume_pct_impacted),
                records_impacted=incident.records_impacted,
                baseline_alert_rate=baseline_rate,
                missed_alerts=missed_alerts,
                likelihood_band=likelihood_band,
                dq_final_risk=dq_final,
                jeffreys_alpha=rollup.total_stors + 0.5,
                jeffreys_beta=(rollup.total_alerts - rollup.total_stors) + 0.5,
                stor_rate_mean=stor_mean,
                stor_rate_95ucb=stor_95,
                expected_missed_stors_mean=expected_mean,
                expected_missed_stors_95ucb=expected_95,
                p_at_least_one_missed_stor_95ucb=prob_at_least_one,
                run_timestamp=run_timestamp.isoformat(),
                run_user=run_user,
                workbook_version=config.version,
                notes=(
                    "No lookback history available"
                    if rollup.total_alerts == 0 and rollup.total_records == 0
                    else None
                ),
            )
        )
    return results

