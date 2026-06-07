"""
Backtesting helpers for the intraday research workflow.

Educational and research use only. This is a lightweight simulation and not a
substitute for institutional-grade execution or market impact analysis.
"""

from __future__ import annotations

from datetime import date, timedelta
from typing import Any

import pandas as pd


def simulate_symbol_outcome(symbol: str, plan: dict[str, Any], intraday_df: pd.DataFrame) -> dict[str, Any]:
    entry = plan["suggested_entry_price"]
    target = plan["suggested_target_price"]
    stop = plan["suggested_stop_price"]
    position_size = plan.get("position_size", 0)
    entered = False
    exit_price = None
    exit_reason = "no_fill"
    max_price_after_entry = None

    for _, row in intraday_df.iterrows():
        if not entered and row["Low"] <= entry <= row["High"]:
            entered = True
        if not entered:
            continue
        max_price_after_entry = row["High"] if max_price_after_entry is None else max(max_price_after_entry, row["High"])
        if row["Low"] <= stop:
            exit_price = stop
            exit_reason = "stop"
            break
        if row["High"] >= target:
            exit_price = target
            exit_reason = "target"
            break

    if entered and exit_price is None and not intraday_df.empty:
        exit_price = float(intraday_df["Close"].iloc[-1])
        exit_reason = "eod"

    realized = 0.0
    if entered and exit_price is not None:
        realized = (exit_price - entry) * position_size

    achieved_price = round(max_price_after_entry, 2) if max_price_after_entry is not None else None
    target_achieved = exit_reason == "target"
    if target_achieved:
        failure_reason = ""
    elif not entered:
        failure_reason = "Entry price was never reached."
    elif exit_reason == "stop":
        failure_reason = "Price reversed to the stop before reaching the target."
    elif exit_reason == "eod":
        failure_reason = "Session ended before the target was reached."
    else:
        failure_reason = "Target was not achieved."

    return {
        "symbol": symbol,
        "entered": entered,
        "target_achieved": target_achieved,
        "exit_reason": exit_reason,
        "exit_price": round(exit_price, 2) if exit_price is not None else None,
        "achieved_price": achieved_price,
        "achieved_move_from_entry": round((achieved_price - entry), 2) if achieved_price is not None else None,
        "failure_reason": failure_reason,
        "realized_pnl": round(realized, 2),
        "estimated_profit_if_target_hit": plan.get("estimated_profit_if_target_hit", 0.0),
        "estimated_loss_if_stop_hit": plan.get("estimated_loss_if_stop_hit", 0.0),
    }


def daterange(start: date, end: date) -> list[date]:
    current = start
    dates: list[date] = []
    while current <= end:
        dates.append(current)
        current += timedelta(days=1)
    return dates
