from __future__ import annotations

from datetime import date
from pathlib import Path

import pandas as pd
import pytest

from Gen_AI.projects.intraday_research_system.agents import (
    ExplanationReportingAgent,
    IntradayStrategyScoringAgent,
    LearningAgent,
    NewsSentimentAgent,
    RiskManagementAgent,
    TechnicalAnalysisAgent,
    UniverseDataAgent,
)
from Gen_AI.projects.intraday_research_system.backtest import simulate_symbol_outcome
from Gen_AI.projects.intraday_research_system.reporting import write_excel_report
from Gen_AI.projects.intraday_research_system.settings import AppConfig
from Gen_AI.projects.intraday_research_system.state import IntradayResearchState


class FakeProvider:
    def get_universe_symbols(self, config, run_date):
        return list(config.universe_symbols)

    def get_recent_daily_bars(self, symbols, lookback_days, run_date):
        return {symbol: sample_daily_frame(symbol) for symbol in symbols}

    def get_intraday_bars(self, symbols, interval, lookback_days, run_date):
        return {symbol: sample_intraday_frame(symbol) for symbol in symbols}


def build_config(tmp_path: Path, risk_profile: str = "balanced") -> AppConfig:
    return AppConfig.model_validate(
        {
            "markets": ["NSE"],
            "universe_name": "NIFTY_50",
            "universe_symbols": ["AAA.NS", "BBB.NS", "CCC.NS"],
            "sector_map": {"AAA.NS": "Financials", "BBB.NS": "Financials", "CCC.NS": "Technology"},
            "top_k_stocks": 2,
            "daily_profit_target": {"mode": "absolute", "value": 10_000},
            "max_daily_loss": {"mode": "absolute", "value": 8_000},
            "capital_allocation": {
                "total_capital": 200_000,
                "max_capital_per_name": 100_000,
                "max_risk_per_trade": 4_000,
            },
            "time_window_intraday": {"start": "09:15", "end": "15:30"},
            "data_provider": {"name": "yfinance", "credentials": {}, "interval": "15m"},
            "timezone": "Asia/Kolkata",
            "trading_calendar": "NSE",
            "report_output_path": str(tmp_path),
            "backtest_mode": False,
            "risk_profile": risk_profile,
            "reporting": {"overwrite_mode": "overwrite"},
            "sentiment": {"enabled": False, "top_n_symbols": 2},
            "filters": {
                "min_avg_volume": 1000,
                "min_price": 10,
                "max_price": 10000,
                "min_composite_score": 50,
                "min_atr_pct": 0.5,
                "max_stop_distance_pct": 3.0,
                "max_stocks_per_sector": 2,
            },
            "logging": {"level": "INFO"},
        }
    )


def sample_daily_frame(symbol: str) -> pd.DataFrame:
    closes = {
        "AAA.NS": [100, 101, 102, 103, 104, 105],
        "BBB.NS": [100, 100.5, 101, 101.5, 102, 102.5],
        "CCC.NS": [100, 103, 98, 106, 95, 108],
    }[symbol]
    idx = pd.date_range("2026-04-01", periods=len(closes), freq="D", tz="Asia/Kolkata")
    return pd.DataFrame(
        {
            "Open": closes,
            "High": [c + 1 for c in closes],
            "Low": [c - 1 for c in closes],
            "Close": closes,
            "Volume": [250000, 260000, 255000, 270000, 280000, 290000],
        },
        index=idx,
    )


def sample_intraday_frame(symbol: str) -> pd.DataFrame:
    base = {"AAA.NS": 105, "BBB.NS": 102.5, "CCC.NS": 108}[symbol]
    drift = {"AAA.NS": 0.35, "BBB.NS": 0.12, "CCC.NS": 0.6}[symbol]
    rows = []
    price = base
    for i in range(24):
        open_price = price
        high = open_price + drift + 0.2
        low = open_price - 0.15
        close = open_price + drift
        rows.append(
            {
                "Open": round(open_price, 2),
                "High": round(high, 2),
                "Low": round(low, 2),
                "Close": round(close, 2),
                "Volume": 150000 + (i * 2000),
            }
        )
        price = close
    idx = pd.date_range("2026-04-11 09:15", periods=24, freq="15min", tz="Asia/Kolkata")
    return pd.DataFrame(rows, index=idx)


def run_pipeline(tmp_path: Path, risk_profile: str = "balanced") -> IntradayResearchState:
    config = build_config(tmp_path, risk_profile=risk_profile)
    state = IntradayResearchState(config=config, run_date=date(2026, 4, 11))
    state = UniverseDataAgent(FakeProvider()).run(state)
    state = TechnicalAnalysisAgent().run(state)
    state = NewsSentimentAgent().run(state)
    state = LearningAgent().run(state)
    state = IntradayStrategyScoringAgent().run(state)
    state = RiskManagementAgent().run(state)
    state = ExplanationReportingAgent().run(state)
    return state


def seed_prior_reports(report_dir: Path) -> None:
    report_dir.mkdir(parents=True, exist_ok=True)
    datasets = [
        pd.DataFrame(
            [
                {
                    "symbol": "OLD1.NS",
                    "sector": "Technology",
                    "atr_pct": 1.8,
                    "risk_percent_of_entry": 1.2,
                    "composite_score": 82.0,
                    "target_achieved": True,
                },
                {
                    "symbol": "OLD2.NS",
                    "sector": "Technology",
                    "atr_pct": 1.9,
                    "risk_percent_of_entry": 1.1,
                    "composite_score": 78.0,
                    "target_achieved": True,
                },
                {
                    "symbol": "OLD3.NS",
                    "sector": "Financials",
                    "atr_pct": 0.6,
                    "risk_percent_of_entry": 2.8,
                    "composite_score": 72.0,
                    "target_achieved": False,
                },
            ]
        ),
        pd.DataFrame(
            [
                {
                    "symbol": "OLD4.NS",
                    "sector": "Technology",
                    "atr_pct": 1.7,
                    "risk_percent_of_entry": 1.3,
                    "composite_score": 80.0,
                    "target_achieved": True,
                },
                {
                    "symbol": "OLD5.NS",
                    "sector": "Financials",
                    "atr_pct": 0.5,
                    "risk_percent_of_entry": 2.9,
                    "composite_score": 71.0,
                    "target_achieved": False,
                },
            ]
        ),
        pd.DataFrame(
            [
                {
                    "symbol": "OLD6.NS",
                    "sector": "Technology",
                    "atr_pct": 2.1,
                    "risk_percent_of_entry": 1.0,
                    "composite_score": 85.0,
                    "target_achieved": True,
                },
                {
                    "symbol": "OLD7.NS",
                    "sector": "Financials",
                    "atr_pct": 0.4,
                    "risk_percent_of_entry": 3.0,
                    "composite_score": 69.0,
                    "target_achieved": False,
                },
            ]
        ),
    ]
    for idx, frame in enumerate(datasets, start=1):
        path = report_dir / f"intraday_report_2026030{idx}.xlsx"
        with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
            frame.to_excel(writer, sheet_name="Intraday Candidates", index=False)
            pd.DataFrame([{"note": "seed"}]).to_excel(writer, sheet_name="Daily Summary", index=False)


def test_invalid_risk_profile_rejected(tmp_path: Path):
    with pytest.raises(Exception):
        build_config(tmp_path, risk_profile="invalid")


def test_invalid_time_window_rejected(tmp_path: Path):
    with pytest.raises(Exception):
        AppConfig.model_validate(
            {
                **build_config(tmp_path).model_dump(),
                "time_window_intraday": {"start": "15:30", "end": "09:15"},
            }
        )


def test_technical_features_and_scores_exist(tmp_path: Path):
    state = run_pipeline(tmp_path)
    assert "AAA.NS" in state.technical_features
    assert state.technical_features["AAA.NS"]["rsi_14"] >= 0
    assert state.technical_features["AAA.NS"]["atr_pct"] > 0
    assert state.technical_score["AAA.NS"] > 0


def test_conservative_and_aggressive_profiles_rank_differently(tmp_path: Path):
    conservative = run_pipeline(tmp_path / "conservative", risk_profile="conservative")
    aggressive = run_pipeline(tmp_path / "aggressive", risk_profile="aggressive")
    assert conservative.selected_symbols
    assert aggressive.selected_symbols
    assert conservative.composite_scores != aggressive.composite_scores


def test_learning_agent_adapts_thresholds_from_prior_reports(tmp_path: Path):
    seed_prior_reports(tmp_path)
    config = build_config(tmp_path)
    state = IntradayResearchState(config=config, run_date=date(2026, 4, 11))
    state = LearningAgent().run(state)
    assert state.learning_insights["applied"] is True
    assert state.learned_thresholds["min_atr_pct"] > config.filters.min_atr_pct
    assert state.learned_thresholds["max_stop_distance_pct"] < config.filters.max_stop_distance_pct
    assert state.learned_score_adjustments["sector_bias"]["Technology"] > 0


def test_risk_limits_respected(tmp_path: Path):
    state = run_pipeline(tmp_path)
    total_loss = sum(plan["estimated_loss_if_stop_hit"] for plan in state.final_plan.values())
    assert total_loss <= state.config.max_daily_loss.value
    for plan in state.final_plan.values():
        assert plan["notional_exposure"] <= state.config.capital_allocation.max_capital_per_name
        assert plan["risk_percent_of_entry"] <= state.config.filters.max_stop_distance_pct
    sector_counts: dict[str, int] = {}
    for plan in state.final_plan.values():
        sector = plan["sector"]
        sector_counts[sector] = sector_counts.get(sector, 0) + 1
    assert all(count <= state.config.filters.max_stocks_per_sector for count in sector_counts.values())


def test_report_contains_required_columns_and_writes_excel(tmp_path: Path):
    state = run_pipeline(tmp_path)
    path = write_excel_report(state)
    assert path.exists()
    candidates = pd.DataFrame(state.market_session["candidates_table"])
    expected = {
        "date",
        "symbol",
        "open_price",
        "close_price",
        "suggested_entry_price",
        "suggested_target_price",
        "suggested_stop_price",
        "position_size",
        "estimated_profit_if_target_hit",
        "estimated_loss_if_stop_hit",
        "composite_score",
        "technical_score",
        "sentiment_score",
        "sector",
        "atr_pct",
        "risk_percent_of_entry",
        "target_achieved",
        "achieved_price",
        "achieved_move_from_entry",
        "actual_outcome_reason",
        "possible_reason_for_failure",
        "reason_for_selection",
    }
    assert expected.issubset(candidates.columns)


def test_backtest_target_stop_and_eod_logic(tmp_path: Path):
    state = run_pipeline(tmp_path)
    symbol = next(iter(state.final_plan))
    result = simulate_symbol_outcome(symbol, state.final_plan[symbol], state.historical_data[symbol])
    assert result["exit_reason"] in {"target", "eod", "stop", "no_fill"}
    assert "target_achieved" in result
    assert "failure_reason" in result


def test_end_to_end_smoke_summary_matches_plan(tmp_path: Path):
    state = run_pipeline(tmp_path)
    expected_total = round(sum(plan["estimated_profit_if_target_hit"] for plan in state.final_plan.values()), 2)
    assert state.daily_summary["num_candidates"] == len(state.final_plan)
    assert state.daily_summary["total_expected_profit_scenario"] == expected_total
    assert "learning_summary" in state.daily_summary
