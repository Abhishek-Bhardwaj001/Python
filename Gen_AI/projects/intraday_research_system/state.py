"""
Shared workflow state for the intraday research graph.

Educational and research use only.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Any

import pandas as pd

from .settings import AppConfig


@dataclass
class IntradayResearchState:
    config: AppConfig
    run_date: date
    market_session: dict[str, Any] = field(default_factory=dict)
    universe_symbols: list[str] = field(default_factory=list)
    historical_data: dict[str, pd.DataFrame] = field(default_factory=dict)
    daily_data: dict[str, pd.DataFrame] = field(default_factory=dict)
    liquidity_metrics: dict[str, dict[str, float]] = field(default_factory=dict)
    technical_features: dict[str, dict[str, Any]] = field(default_factory=dict)
    technical_score: dict[str, float] = field(default_factory=dict)
    sentiment_score: dict[str, float] = field(default_factory=dict)
    news_flags: dict[str, list[str]] = field(default_factory=dict)
    sentiment_summary: dict[str, str] = field(default_factory=dict)
    learning_insights: dict[str, Any] = field(default_factory=dict)
    learned_thresholds: dict[str, float] = field(default_factory=dict)
    learned_score_adjustments: dict[str, Any] = field(default_factory=dict)
    composite_scores: dict[str, float] = field(default_factory=dict)
    selected_symbols: list[str] = field(default_factory=list)
    per_symbol_plan: dict[str, dict[str, Any]] = field(default_factory=dict)
    final_plan: dict[str, dict[str, Any]] = field(default_factory=dict)
    risk_summary: str = ""
    explanations: dict[str, str] = field(default_factory=dict)
    daily_summary: dict[str, Any] = field(default_factory=dict)
    backtest_results: list[dict[str, Any]] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)

    def log_warning(self, message: str) -> None:
        self.warnings.append(message)

    def log_error(self, message: str) -> None:
        self.errors.append(message)
