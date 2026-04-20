"""
Agent implementations for the intraday research workflow.

Educational and research use only. Outputs are scenario estimates, not trade
recommendations or guarantees.
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from collections import defaultdict
from logging import getLogger
from math import floor
from typing import Any

import pandas as pd

from .backtest import simulate_symbol_outcome
from .indicators import compute_rsi, compute_vwap, normalize_score
from .learning import list_prior_report_paths, load_candidate_history
from .settings import AppConfig
from .state import IntradayResearchState

LOGGER = getLogger(__name__)


class BaseAgent(ABC):
    @abstractmethod
    def run(self, state: IntradayResearchState) -> IntradayResearchState:
        raise NotImplementedError


class UniverseDataAgent(BaseAgent):
    def __init__(self, provider) -> None:
        self.provider = provider

    def run(self, state: IntradayResearchState) -> IntradayResearchState:
        config = state.config
        symbols = self.provider.get_universe_symbols(config, state.run_date)
        LOGGER.info("Universe before filtering: %s", len(symbols))
        daily_frames = self.provider.get_recent_daily_bars(
            symbols, config.data_provider.daily_lookback_days, state.run_date
        )
        intraday_frames = self.provider.get_intraday_bars(
            symbols,
            config.data_provider.interval,
            config.data_provider.intraday_lookback_days,
            state.run_date,
        )

        filtered: list[str] = []
        for symbol in symbols:
            daily_df = daily_frames.get(symbol, pd.DataFrame())
            intraday_df = intraday_frames.get(symbol, pd.DataFrame())
            if daily_df.empty or intraday_df.empty:
                state.log_warning(f"Skipping {symbol}: missing daily or intraday data")
                continue

            avg_volume = float(daily_df["Volume"].tail(10).fillna(0).mean())
            last_close = float(daily_df["Close"].dropna().iloc[-1])
            realized_vol = float(daily_df["Close"].pct_change().tail(10).std(ddof=0) or 0.0)
            if avg_volume < config.filters.min_avg_volume:
                continue
            if not (config.filters.min_price <= last_close <= config.filters.max_price):
                continue

            filtered.append(symbol)
            state.daily_data[symbol] = daily_df
            state.historical_data[symbol] = intraday_df
            state.liquidity_metrics[symbol] = {
                "avg_volume_10d": avg_volume,
                "last_close": last_close,
                "realized_vol_10d": realized_vol,
            }

        state.universe_symbols = filtered
        LOGGER.info("Universe after filtering: %s", len(filtered))
        return state


class TechnicalAnalysisAgent(BaseAgent):
    def run(self, state: IntradayResearchState) -> IntradayResearchState:
        for symbol in state.universe_symbols:
            intraday_df = state.historical_data[symbol].copy()
            daily_df = state.daily_data[symbol].copy()
            intraday_df["ema_fast"] = intraday_df["Close"].ewm(span=5, adjust=False).mean()
            intraday_df["ema_slow"] = intraday_df["Close"].ewm(span=13, adjust=False).mean()
            intraday_df["rsi_14"] = compute_rsi(intraday_df["Close"], period=14)
            intraday_df["vwap"] = compute_vwap(intraday_df)

            latest = intraday_df.iloc[-1]
            prev_close = float(daily_df["Close"].iloc[-2]) if len(daily_df) > 1 else float(daily_df["Close"].iloc[-1])
            session_open = float(intraday_df["Open"].iloc[0])
            opening_range = intraday_df.head(min(3, len(intraday_df)))
            support = float(intraday_df["Low"].tail(10).min())
            resistance = float(intraday_df["High"].tail(10).max())
            daily_range = daily_df["High"] - daily_df["Low"]
            prev_close_series = daily_df["Close"].shift(1)
            true_range = pd.concat(
                [
                    daily_range,
                    (daily_df["High"] - prev_close_series).abs(),
                    (daily_df["Low"] - prev_close_series).abs(),
                ],
                axis=1,
            ).max(axis=1)
            atr_value = float(true_range.tail(14).mean()) if not true_range.dropna().empty else 0.0
            gap_pct = ((session_open - prev_close) / prev_close) if prev_close else 0.0
            intraday_vol = float(intraday_df["Close"].pct_change().tail(12).std(ddof=0) or 0.0)
            ema_spread = float((latest["ema_fast"] - latest["ema_slow"]) / latest["Close"]) if latest["Close"] else 0.0
            vwap_deviation = float((latest["Close"] - latest["vwap"]) / latest["vwap"]) if latest["vwap"] else 0.0
            rsi_value = float(latest["rsi_14"]) if pd.notna(latest["rsi_14"]) else 50.0
            breakout = float(latest["Close"] > opening_range["High"].max())
            avg_intraday_volume = float(intraday_df["Volume"].tail(12).mean())
            atr_pct = (atr_value / prev_close) * 100 if prev_close else 0.0

            score_components = {
                "momentum": normalize_score(ema_spread, -0.02, 0.02),
                "rsi_bias": normalize_score(abs(rsi_value - 50), 0, 30),
                "vwap_bias": normalize_score(vwap_deviation, -0.02, 0.02),
                "gap": normalize_score(abs(gap_pct), 0, 0.05),
                "breakout": breakout,
                "liquidity": normalize_score(avg_intraday_volume, 100_000, 5_000_000),
                "volatility": normalize_score(intraday_vol, 0.001, 0.03),
                "atr": normalize_score(atr_pct, 0.25, 4.0),
            }
            technical_score = (
                0.2 * score_components["momentum"]
                + 0.15 * score_components["rsi_bias"]
                + 0.15 * score_components["vwap_bias"]
                + 0.1 * score_components["gap"]
                + 0.15 * score_components["breakout"]
                + 0.15 * score_components["liquidity"]
                + 0.05 * score_components["volatility"]
                + 0.05 * score_components["atr"]
            )
            state.technical_features[symbol] = {
                "gap_pct": gap_pct,
                "intraday_volatility": intraday_vol,
                "atr_value": atr_value,
                "atr_pct": atr_pct,
                "ema_spread": ema_spread,
                "rsi_14": rsi_value,
                "vwap_deviation": vwap_deviation,
                "support": support,
                "resistance": resistance,
                "opening_range_high": float(opening_range["High"].max()),
                "opening_range_low": float(opening_range["Low"].min()),
                "avg_intraday_volume": avg_intraday_volume,
            }
            state.technical_score[symbol] = round(min(technical_score * 130, 100), 2)
        return state


class NewsSentimentAgent(BaseAgent):
    def run(self, state: IntradayResearchState) -> IntradayResearchState:
        enabled = state.config.sentiment.enabled
        ranked = sorted(
            state.universe_symbols,
            key=lambda sym: state.liquidity_metrics.get(sym, {}).get("avg_volume_10d", 0.0),
            reverse=True,
        )
        covered = set(ranked[: state.config.sentiment.top_n_symbols])
        for symbol in state.universe_symbols:
            if not enabled:
                state.sentiment_score[symbol] = 0.0
                state.news_flags[symbol] = ["sentiment_unavailable"]
                state.sentiment_summary[symbol] = "Sentiment module disabled; neutral score applied."
            elif symbol not in covered:
                state.sentiment_score[symbol] = 0.0
                state.news_flags[symbol] = ["outside_sentiment_coverage"]
                state.sentiment_summary[symbol] = "Outside configured sentiment coverage set."
            else:
                state.sentiment_score[symbol] = 0.0
                state.news_flags[symbol] = ["neutral_default"]
                state.sentiment_summary[symbol] = "No news adapter configured; neutral score applied."
        return state


class LearningAgent(BaseAgent):
    def run(self, state: IntradayResearchState) -> IntradayResearchState:
        config = state.config
        state.learned_thresholds = {
            "min_composite_score": config.filters.min_composite_score,
            "min_atr_pct": config.filters.min_atr_pct,
            "max_stop_distance_pct": config.filters.max_stop_distance_pct,
        }
        state.learned_score_adjustments = {"sector_bias": {}}
        state.learning_insights = {
            "enabled": config.learning.enabled,
            "reports_scanned": 0,
            "applied": False,
            "summary": "Learning agent disabled or insufficient history.",
        }
        if not config.learning.enabled:
            return state

        report_paths = list_prior_report_paths(
            config.report_output_path, state.run_date.isoformat(), config.learning.max_reports_to_scan
        )
        history = load_candidate_history(report_paths)
        state.learning_insights["reports_scanned"] = len(report_paths)
        if len(report_paths) < config.learning.min_reports_required or history.empty:
            state.learning_insights["summary"] = "Not enough prior reports with outcome columns to adapt thresholds."
            return state

        success_rate = float(history["target_achieved"].mean())
        success_rows = history[history["target_achieved"]]
        failure_rows = history[~history["target_achieved"]]
        lr = config.learning.learning_rate

        min_score = config.filters.min_composite_score
        if success_rate < 0.35:
            min_score = min(90.0, min_score + (1 - success_rate) * 10 * lr)
        elif success_rate > 0.6:
            min_score = max(50.0, min_score - success_rate * 5 * lr)

        min_atr = config.filters.min_atr_pct
        if not success_rows.empty and not failure_rows.empty:
            atr_shift = float(success_rows["atr_pct"].mean() - failure_rows["atr_pct"].mean())
            min_atr = max(0.2, min_atr + atr_shift * 0.15 * lr)

        max_risk = config.filters.max_stop_distance_pct
        if not success_rows.empty and not failure_rows.empty:
            risk_shift = float(success_rows["risk_percent_of_entry"].mean() - failure_rows["risk_percent_of_entry"].mean())
            max_risk = min(5.0, max(0.5, max_risk + risk_shift * 0.25 * lr))

        sector_bias: dict[str, float] = {}
        sector_stats = history.groupby("sector")["target_achieved"].agg(["mean", "count"]).reset_index()
        for _, row in sector_stats.iterrows():
            if int(row["count"]) < 2:
                continue
            diff = float(row["mean"] - success_rate)
            if diff > 0.10:
                sector_bias[str(row["sector"])] = round(min(8.0, diff * 20 * lr), 2)
            elif diff < -0.10:
                sector_bias[str(row["sector"])] = round(max(-8.0, diff * 20 * lr), 2)

        state.learned_thresholds = {
            "min_composite_score": round(min_score, 2),
            "min_atr_pct": round(min_atr, 2),
            "max_stop_distance_pct": round(max_risk, 2),
        }
        state.learned_score_adjustments = {"sector_bias": sector_bias}
        state.learning_insights = {
            "enabled": True,
            "reports_scanned": len(report_paths),
            "applied": True,
            "historical_success_rate": round(success_rate, 3),
            "summary": (
                f"Adapted using {len(report_paths)} prior reports. "
                f"Composite score floor={min_score:.2f}, ATR floor={min_atr:.2f}%, "
                f"max risk={max_risk:.2f}%."
            ),
        }
        return state


class IntradayStrategyScoringAgent(BaseAgent):
    def run(self, state: IntradayResearchState) -> IntradayResearchState:
        profile_weights = {
            "conservative": {"technical": 0.65, "sentiment": 0.05, "liquidity": 0.2, "vol_penalty": 0.2},
            "balanced": {"technical": 0.75, "sentiment": 0.1, "liquidity": 0.15, "vol_penalty": 0.12},
            "aggressive": {"technical": 0.8, "sentiment": 0.1, "liquidity": 0.1, "vol_penalty": 0.05},
        }
        weights = profile_weights[state.config.risk_profile]
        eligible_symbols: list[str] = []
        sector_counts: dict[str, int] = defaultdict(int)
        sector_bias = state.learned_score_adjustments.get("sector_bias", {})
        min_composite_score = state.learned_thresholds.get(
            "min_composite_score", state.config.filters.min_composite_score
        )
        min_atr_pct = state.learned_thresholds.get("min_atr_pct", state.config.filters.min_atr_pct)
        max_stop_distance_pct = state.learned_thresholds.get(
            "max_stop_distance_pct", state.config.filters.max_stop_distance_pct
        )

        for symbol in state.universe_symbols:
            features = state.technical_features.get(symbol, {})
            tech_score = state.technical_score.get(symbol, 0.0)
            sentiment = state.sentiment_score.get(symbol, 0.0)
            liquidity = normalize_score(
                state.liquidity_metrics.get(symbol, {}).get("avg_volume_10d", 0.0), 100_000, 5_000_000
            )
            vol_penalty = normalize_score(features.get("intraday_volatility", 0.0), 0.001, 0.05)
            composite = (
                weights["technical"] * (tech_score / 100.0)
                + weights["sentiment"] * normalize_score(sentiment, -1, 1)
                + weights["liquidity"] * liquidity
                - weights["vol_penalty"] * vol_penalty
            )
            state.composite_scores[symbol] = round(min(composite * 130, 100), 2)

            reference_price = state.liquidity_metrics[symbol]["last_close"]
            support = features.get("support", reference_price * 0.99)
            resistance = features.get("resistance", reference_price * 1.01)
            intraday_vol = max(features.get("intraday_volatility", 0.005), 0.0025)
            breakout_buffer = reference_price * intraday_vol * 0.35
            entry = max(reference_price, features.get("opening_range_high", reference_price)) + breakout_buffer
            stop_buffer = entry * intraday_vol * 1.25
            stop = min(entry - stop_buffer, support * 0.995, reference_price * 0.995, entry * 0.998)
            stop = max(stop, 0.01)
            target = min(max(entry + (entry * intraday_vol * 1.9), resistance * 1.003), entry * 1.05)
            risk_per_share = max(entry - stop, entry * 0.002)
            reward_per_share = max(target - entry, risk_per_share * 1.2)
            stop_distance_pct = ((entry - stop) / entry) * 100 if entry else 0.0
            sector = state.config.sector_map.get(symbol, "Unknown")
            learned_sector_bias = sector_bias.get(sector, 0.0)
            adjusted_composite = min(100.0, max(0.0, state.composite_scores[symbol] + learned_sector_bias))
            state.composite_scores[symbol] = round(adjusted_composite, 2)

            state.per_symbol_plan[symbol] = {
                "reference_price": round(reference_price, 2),
                "suggested_entry_price": round(entry, 2),
                "suggested_stop_price": round(stop, 2),
                "suggested_target_price": round(target, 2),
                "risk_per_share": round(risk_per_share, 2),
                "reward_per_share": round(reward_per_share, 2),
                "risk_percent_of_entry": round(stop_distance_pct, 2),
                "atr_pct": round(features.get("atr_pct", 0.0), 2),
                "sector": sector,
                "learned_sector_bias": learned_sector_bias,
                "estimated_profit_if_target_hit": round(reward_per_share, 2),
                "estimated_loss_if_stop_hit": round(risk_per_share, 2),
            }
        ranked = sorted(
            state.universe_symbols,
            key=lambda sym: (state.composite_scores.get(sym, 0.0), state.technical_score.get(sym, 0.0)),
            reverse=True,
        )
        for symbol in ranked:
            plan = state.per_symbol_plan[symbol]
            if state.composite_scores.get(symbol, 0.0) <= min_composite_score:
                continue
            if plan["atr_pct"] < min_atr_pct:
                continue
            if plan["risk_percent_of_entry"] > max_stop_distance_pct:
                continue
            sector = plan["sector"]
            if sector_counts[sector] >= state.config.filters.max_stocks_per_sector:
                continue
            eligible_symbols.append(symbol)
            sector_counts[sector] += 1
            if len(eligible_symbols) >= state.config.top_k_stocks:
                break
        state.selected_symbols = eligible_symbols
        LOGGER.info("Top score table: %s", [(sym, state.composite_scores[sym]) for sym in state.selected_symbols])
        return state


class RiskManagementAgent(BaseAgent):
    def run(self, state: IntradayResearchState) -> IntradayResearchState:
        cfg = state.config.capital_allocation
        total_profit = 0.0
        total_loss = 0.0
        final_plan: dict[str, dict[str, Any]] = {}
        pruned: list[str] = []

        for symbol in state.selected_symbols:
            plan = dict(state.per_symbol_plan[symbol])
            entry = plan["suggested_entry_price"]
            risk_per_share = max(plan["risk_per_share"], 0.01)
            max_units_by_risk = floor(cfg.max_risk_per_trade / risk_per_share)
            max_units_by_capital = floor(min(cfg.max_capital_per_name, cfg.total_capital) / entry)
            position_size = max(min(max_units_by_risk, max_units_by_capital), 0)
            if position_size <= 0:
                pruned.append(symbol)
                continue

            projected_loss = position_size * plan["estimated_loss_if_stop_hit"]
            if total_loss + projected_loss > self._max_loss_budget(state.config):
                pruned.append(symbol)
                continue

            projected_profit = position_size * plan["estimated_profit_if_target_hit"]
            total_profit += projected_profit
            total_loss += projected_loss
            plan["position_size"] = position_size
            plan["notional_exposure"] = round(position_size * entry, 2)
            plan["estimated_profit_if_target_hit"] = round(projected_profit, 2)
            plan["estimated_loss_if_stop_hit"] = round(projected_loss, 2)
            plan["composite_score"] = state.composite_scores[symbol]
            plan["technical_score"] = state.technical_score[symbol]
            plan["sentiment_score"] = state.sentiment_score[symbol]
            final_plan[symbol] = plan

        state.final_plan = final_plan
        state.selected_symbols = list(final_plan.keys())
        target_budget = self._profit_target_budget(state.config)
        target_status = "within target envelope" if total_profit >= target_budget else "below target envelope"
        state.risk_summary = (
            f"Final plan contains {len(final_plan)} symbols. "
            f"Estimated target-hit profit is {total_profit:,.2f}; aggregate stop-loss exposure is {total_loss:,.2f}. "
            f"Daily profit target status: {target_status}."
        )
        if pruned:
            LOGGER.warning("Pruned symbols due to sizing/risk limits: %s", pruned)
        return state

    def _max_loss_budget(self, config: AppConfig) -> float:
        if config.max_daily_loss.mode == "absolute":
            return config.max_daily_loss.value
        return config.capital_allocation.total_capital * config.max_daily_loss.value / 100.0

    def _profit_target_budget(self, config: AppConfig) -> float:
        if config.daily_profit_target.mode == "absolute":
            return config.daily_profit_target.value
        return config.capital_allocation.total_capital * config.daily_profit_target.value / 100.0


class ExplanationReportingAgent(BaseAgent):
    def run(self, state: IntradayResearchState) -> IntradayResearchState:
        rows: list[dict[str, Any]] = []
        total_profit = 0.0
        total_loss = 0.0
        for symbol, plan in state.final_plan.items():
            intraday_df = state.historical_data.get(symbol, pd.DataFrame())
            open_price = float(intraday_df["Open"].iloc[0]) if not intraday_df.empty else None
            close_price = float(intraday_df["Close"].iloc[-1]) if not intraday_df.empty else None
            outcome = simulate_symbol_outcome(symbol, plan, intraday_df)
            explanation = self._build_explanation(symbol, plan, state)
            state.explanations[symbol] = explanation
            rows.append(
                {
                    "date": state.run_date.isoformat(),
                    "symbol": symbol,
                    "open_price": round(open_price, 2) if open_price is not None else None,
                    "close_price": round(close_price, 2) if close_price is not None else None,
                    "suggested_entry_price": plan["suggested_entry_price"],
                    "suggested_target_price": plan["suggested_target_price"],
                    "suggested_stop_price": plan["suggested_stop_price"],
                    "position_size": plan["position_size"],
                    "estimated_profit_if_target_hit": plan["estimated_profit_if_target_hit"],
                    "estimated_loss_if_stop_hit": plan["estimated_loss_if_stop_hit"],
                    "composite_score": plan["composite_score"],
                    "technical_score": plan["technical_score"],
                    "sentiment_score": plan["sentiment_score"],
                    "sector": plan["sector"],
                    "atr_pct": plan["atr_pct"],
                    "risk_percent_of_entry": plan["risk_percent_of_entry"],
                    "target_achieved": outcome["target_achieved"],
                    "achieved_price": outcome["achieved_price"],
                    "achieved_move_from_entry": outcome["achieved_move_from_entry"],
                    "actual_outcome_reason": outcome["exit_reason"],
                    "possible_reason_for_failure": outcome["failure_reason"],
                    "reason_for_selection": explanation,
                }
            )
            total_profit += plan["estimated_profit_if_target_hit"]
            total_loss += plan["estimated_loss_if_stop_hit"]

        state.market_session = {"candidates_table": rows}
        state.daily_summary = {
            "date": state.run_date.isoformat(),
            "num_candidates": len(rows),
            "total_expected_profit_scenario": round(total_profit, 2),
            "max_possible_loss_scenario": round(total_loss, 2),
            "daily_profit_target": state.config.daily_profit_target.value,
            "max_daily_loss": state.config.max_daily_loss.value,
            "risk_profile": state.config.risk_profile,
            "note": (
                "Research-only intraday candidate sheet. Not financial advice. "
                "Backtest and validate before any real-capital use."
            ),
            "learning_summary": state.learning_insights.get("summary", ""),
        }
        return state

    def _build_explanation(self, symbol: str, plan: dict[str, Any], state: IntradayResearchState) -> str:
        features = state.technical_features.get(symbol, {})
        volume = state.liquidity_metrics.get(symbol, {}).get("avg_volume_10d", 0.0)
        return (
            f"{symbol} scored well for the {state.config.risk_profile} profile with technical score "
            f"{plan['technical_score']:.2f} and composite score {plan['composite_score']:.2f}. "
            f"RSI={features.get('rsi_14', 50.0):.1f}, gap={features.get('gap_pct', 0.0) * 100:.2f}%, "
            f"ATR={features.get('atr_pct', 0.0):.2f}%, "
            f"VWAP deviation={features.get('vwap_deviation', 0.0) * 100:.2f}%, "
            f"10-day average volume={volume:,.0f}, sector={plan.get('sector', 'Unknown')}, "
            f"learned sector bias={plan.get('learned_sector_bias', 0.0):+.2f}, "
            f"and sizing was constrained by the daily loss budget."
        )
