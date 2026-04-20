"""
CLI entrypoint for the intraday research system.

Educational and research use only. Not financial advice and not live trading.
"""

from __future__ import annotations

import argparse
from datetime import date, datetime

from .backtest import daterange, simulate_symbol_outcome
from .graph import run_daily_research
from .logging_utils import configure_logging
from .settings import load_config


def parse_date(value: str) -> date:
    return datetime.strptime(value, "%Y-%m-%d").date()


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Intraday multi-agent research system")
    parser.add_argument("--config", default=None, help="Path to config.yaml")
    subparsers = parser.add_subparsers(dest="command", required=True)

    run_parser = subparsers.add_parser("run", help="Run single-day analysis")
    run_parser.add_argument("--date", dest="run_date", required=False, help="Run date in YYYY-MM-DD")

    backtest_parser = subparsers.add_parser("backtest", help="Run backtest over date range")
    backtest_parser.add_argument("--start", required=True, help="Start date in YYYY-MM-DD")
    backtest_parser.add_argument("--end", required=True, help="End date in YYYY-MM-DD")
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    config = load_config(args.config)
    configure_logging(config.logging.level)

    if args.command == "run":
        run_date = parse_date(args.run_date) if args.run_date else date.today()
        result = run_daily_research(config, run_date)
        print(result.daily_summary)
        return

    start = parse_date(args.start)
    end = parse_date(args.end)
    all_results = []
    for run_date in daterange(start, end):
        result = run_daily_research(config, run_date)
        for symbol, plan in result.final_plan.items():
            intraday_df = result.historical_data.get(symbol)
            if intraday_df is None or intraday_df.empty:
                continue
            outcome = simulate_symbol_outcome(symbol, plan, intraday_df)
            outcome["date"] = run_date.isoformat()
            all_results.append(outcome)
    print({"days": len(daterange(start, end)), "results": len(all_results)})


if __name__ == "__main__":
    main()
