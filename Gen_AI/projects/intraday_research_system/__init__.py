"""
Intraday trading research system.

Educational and research use only. This package does not place live trades,
does not provide financial advice, and must be backtested and validated before
any real capital deployment.
"""

from .settings import AppConfig, load_config

__all__ = ["AppConfig", "load_config"]
