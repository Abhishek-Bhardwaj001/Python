"""
Market data provider abstractions.

Educational and research use only. Data quality varies by provider and must be
validated before any production or capital allocation usage.
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from datetime import date, timedelta

import pandas as pd

from .settings import AppConfig

try:
    import yfinance as yf
except ImportError:  # pragma: no cover
    yf = None


class MarketDataProvider(ABC):
    @abstractmethod
    def get_universe_symbols(self, config: AppConfig, run_date: date) -> list[str]:
        raise NotImplementedError

    @abstractmethod
    def get_recent_daily_bars(
        self, symbols: list[str], lookback_days: int, run_date: date
    ) -> dict[str, pd.DataFrame]:
        raise NotImplementedError

    @abstractmethod
    def get_intraday_bars(
        self, symbols: list[str], interval: str, lookback_days: int, run_date: date
    ) -> dict[str, pd.DataFrame]:
        raise NotImplementedError


class YFinanceProvider(MarketDataProvider):
    def __init__(self, timezone: str) -> None:
        self.timezone = timezone
        if yf is None:
            raise ImportError(
                "yfinance is required for YFinanceProvider. Install it before running this system."
            )

    def get_universe_symbols(self, config: AppConfig, run_date: date) -> list[str]:
        return list(config.universe_symbols)

    def get_recent_daily_bars(
        self, symbols: list[str], lookback_days: int, run_date: date
    ) -> dict[str, pd.DataFrame]:
        start = run_date - timedelta(days=lookback_days * 3)
        end = run_date + timedelta(days=1)
        return {
            symbol: self._clean_frame(
                yf.download(
                    tickers=symbol,
                    start=start.isoformat(),
                    end=end.isoformat(),
                    interval="1d",
                    auto_adjust=False,
                    progress=False,
                    threads=False,
                )
            )
            for symbol in symbols
        }

    def get_intraday_bars(
        self, symbols: list[str], interval: str, lookback_days: int, run_date: date
    ) -> dict[str, pd.DataFrame]:
        period = f"{lookback_days}d"
        return {
            symbol: self._clean_frame(
                yf.download(
                    tickers=symbol,
                    period=period,
                    interval=interval,
                    auto_adjust=False,
                    progress=False,
                    threads=False,
                )
            )
            for symbol in symbols
        }

    def _clean_frame(self, frame: pd.DataFrame) -> pd.DataFrame:
        if frame is None or frame.empty:
            return pd.DataFrame(columns=["Open", "High", "Low", "Close", "Volume"])
        clean = frame.copy()
        if isinstance(clean.columns, pd.MultiIndex):
            clean.columns = clean.columns.get_level_values(0)
        clean.columns = [str(col).title() for col in clean.columns]
        clean = clean.sort_index().dropna(how="all")
        if not isinstance(clean.index, pd.DatetimeIndex):
            clean.index = pd.to_datetime(clean.index)
        if clean.index.tz is None:
            clean.index = clean.index.tz_localize("UTC").tz_convert(self.timezone)
        else:
            clean.index = clean.index.tz_convert(self.timezone)
        return clean


def build_provider(config: AppConfig) -> MarketDataProvider:
    if config.data_provider.name.lower() == "yfinance":
        return YFinanceProvider(timezone=config.timezone)
    raise ValueError(f"Unsupported data provider: {config.data_provider.name}")
