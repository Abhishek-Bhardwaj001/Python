# Python
This Repository hold all of my python language projects in Software Engineering, Gen AI and Data Science.

## Intraday Research System
Path: `Gen_AI/projects/intraday_research_system`

This package implements a multi-agent intraday stock research workflow for educational and paper-trading use only. It is not financial advice, does not place live trades, and should be backtested and validated before any real-money usage.

Run instructions:
- Activate the project environment.
- Install any missing packages from `Gen_AI/projects/intraday_research_system/requirements.txt` plus the repo requirements you already use.
- Single day: `python -m Gen_AI.projects.intraday_research_system run --date 2026-04-11`
- Backtest: `python -m Gen_AI.projects.intraday_research_system backtest --start 2026-04-01 --end 2026-04-11`

Default config:
- `Gen_AI/projects/intraday_research_system/config.yaml`

Manual verification checklist:
- Confirm the Excel report is created under the configured reports directory.
- Confirm the `Intraday Candidates` sheet contains prices, score columns, position sizes, and reasons.
- Confirm the `Daily Summary` sheet reflects the configured profit target, max loss, and risk profile.
- Confirm symbols are pruned or resized when aggregate stop-loss exposure exceeds the configured daily max loss.
- Confirm backtest runs produce realized-vs-estimated outcomes without crashing on missing data.
