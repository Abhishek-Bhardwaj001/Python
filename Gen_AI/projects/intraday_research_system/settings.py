"""
Configuration models for the intraday research system.

Educational and research use only. This module does not support live trading.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Literal

import yaml
from pydantic import BaseModel, Field, field_validator, model_validator


class TargetLimit(BaseModel):
    mode: Literal["percent", "absolute"] = "percent"
    value: float = Field(gt=0)


class CapitalAllocation(BaseModel):
    total_capital: float = Field(gt=0)
    max_capital_per_name: float = Field(gt=0)
    max_risk_per_trade: float = Field(gt=0)

    @model_validator(mode="after")
    def validate_caps(self) -> "CapitalAllocation":
        if self.max_capital_per_name > self.total_capital:
            raise ValueError("max_capital_per_name cannot exceed total_capital")
        return self


class TimeWindow(BaseModel):
    start: str = "09:15"
    end: str = "15:30"

    @model_validator(mode="after")
    def validate_window(self) -> "TimeWindow":
        if self.start >= self.end:
            raise ValueError("time_window_intraday.start must be before end")
        return self


class DataProviderConfig(BaseModel):
    name: str = "yfinance"
    credentials: dict[str, str] = Field(default_factory=dict)
    interval: str = "15m"
    daily_lookback_days: int = Field(default=30, ge=5)
    intraday_lookback_days: int = Field(default=5, ge=1)


class ReportingConfig(BaseModel):
    overwrite_mode: Literal["overwrite", "version"] = "version"


class SentimentConfig(BaseModel):
    enabled: bool = False
    top_n_symbols: int = Field(default=10, ge=1)


class FilterConfig(BaseModel):
    min_avg_volume: int = Field(default=100_000, ge=0)
    min_price: float = Field(default=100.0, ge=0)
    max_price: float = Field(default=10_000.0, gt=0)
    min_composite_score: float = Field(default=70.0, ge=0, le=100)
    min_atr_pct: float = Field(default=0.5, ge=0)
    max_stop_distance_pct: float = Field(default=2.5, gt=0)
    max_stocks_per_sector: int = Field(default=2, ge=1)

    @model_validator(mode="after")
    def validate_price_band(self) -> "FilterConfig":
        if self.min_price > self.max_price:
            raise ValueError("min_price cannot exceed max_price")
        return self


class LoggingConfig(BaseModel):
    level: str = "INFO"


class LearningConfig(BaseModel):
    enabled: bool = True
    min_reports_required: int = Field(default=3, ge=1)
    max_reports_to_scan: int = Field(default=30, ge=1)
    learning_rate: float = Field(default=0.5, gt=0, le=1)


class AppConfig(BaseModel):
    markets: list[str] = Field(default_factory=lambda: ["NSE"])
    universe_name: str = "NIFTY_50"
    universe_symbols: list[str] = Field(default_factory=list)
    sector_map: dict[str, str] = Field(default_factory=dict)
    top_k_stocks: int = Field(default=10, ge=1)
    daily_profit_target: TargetLimit
    max_daily_loss: TargetLimit
    capital_allocation: CapitalAllocation
    time_window_intraday: TimeWindow
    data_provider: DataProviderConfig
    timezone: str = "Asia/Kolkata"
    trading_calendar: str = "NSE"
    report_output_path: Path
    backtest_mode: bool = False
    risk_profile: Literal["conservative", "balanced", "aggressive"] = "balanced"
    reporting: ReportingConfig = Field(default_factory=ReportingConfig)
    sentiment: SentimentConfig = Field(default_factory=SentimentConfig)
    learning: LearningConfig = Field(default_factory=LearningConfig)
    filters: FilterConfig = Field(default_factory=FilterConfig)
    logging: LoggingConfig = Field(default_factory=LoggingConfig)

    @field_validator("report_output_path", mode="before")
    @classmethod
    def cast_report_path(cls, value: Any) -> Path:
        return Path(value)


def load_config(path: str | Path | None = None) -> AppConfig:
    config_path = Path(path) if path else Path(__file__).with_name("config.yaml")
    with config_path.open("r", encoding="utf-8") as handle:
        raw = yaml.safe_load(handle) or {}
    config = AppConfig.model_validate(raw)
    if not config.report_output_path.is_absolute():
        config.report_output_path = (config_path.parent / config.report_output_path).resolve()
    return config
