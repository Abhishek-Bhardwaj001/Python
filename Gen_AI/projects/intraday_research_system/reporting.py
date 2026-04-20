"""
Excel reporting utilities.

Educational and research use only. Reports summarize scenario analysis and do
not guarantee future performance.
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from .state import IntradayResearchState


def build_report_path(base_dir: Path, run_date: str, overwrite_mode: str) -> Path:
    base_dir.mkdir(parents=True, exist_ok=True)
    candidate = base_dir / f"intraday_report_{run_date.replace('-', '')}.xlsx"
    if overwrite_mode == "overwrite" or not candidate.exists():
        return candidate
    version = 1
    while True:
        versioned = base_dir / f"intraday_report_{run_date.replace('-', '')}_v{version}.xlsx"
        if not versioned.exists():
            return versioned
        version += 1


def write_excel_report(state: IntradayResearchState) -> Path:
    candidates = pd.DataFrame(state.market_session.get("candidates_table", []))
    summary = pd.DataFrame([state.daily_summary])
    report_path = build_report_path(
        state.config.report_output_path,
        state.run_date.isoformat(),
        state.config.reporting.overwrite_mode,
    )
    with pd.ExcelWriter(report_path, engine="xlsxwriter") as writer:
        candidates.to_excel(writer, sheet_name="Intraday Candidates", index=False)
        summary.to_excel(writer, sheet_name="Daily Summary", index=False)
        workbook = writer.book
        candidates_sheet = writer.sheets["Intraday Candidates"]
        summary_sheet = writer.sheets["Daily Summary"]
        wrap = workbook.add_format({"text_wrap": True, "valign": "top"})
        risk_fill = workbook.add_format({"bg_color": "#FDE9D9"})
        candidates_sheet.set_column("A:T", 18)
        candidates_sheet.set_column("U:U", 60, wrap)
        summary_sheet.set_column("A:H", 22)
        summary_sheet.set_column("H:H", 70, wrap)
        if not candidates.empty:
            candidates_sheet.conditional_format(
                f"J2:J{len(candidates) + 1}",
                {"type": "cell", "criteria": ">", "value": 10000, "format": risk_fill},
            )
    return report_path
