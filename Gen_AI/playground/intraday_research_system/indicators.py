from __future__ import annotations

import numpy as np
import pandas as pd


def compute_rsi(series: pd.Series, period: int = 14) -> pd.Series:
    delta = series.diff()
    gain = delta.clip(lower=0)
    loss = -delta.clip(upper=0)
    avg_gain = gain.ewm(alpha=1 / period, min_periods=period, adjust=False).mean()
    avg_loss = loss.ewm(alpha=1 / period, min_periods=period, adjust=False).mean()
    rs = avg_gain / avg_loss.replace(0, np.nan)
    return 100 - (100 / (1 + rs))


def compute_vwap(frame: pd.DataFrame) -> pd.Series:
    typical_price = (frame["High"] + frame["Low"] + frame["Close"]) / 3.0
    cumulative_volume = frame["Volume"].replace(0, np.nan).cumsum()
    cumulative_value = (typical_price * frame["Volume"]).cumsum()
    return cumulative_value / cumulative_volume


def normalize_score(value: float, low: float, high: float) -> float:
    if high == low:
        return 0.5
    clipped = min(max(value, low), high)
    return (clipped - low) / (high - low)
