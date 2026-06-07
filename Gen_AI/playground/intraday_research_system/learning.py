"""
Learning utilities for adaptive research-only selection rules.

Educational and research use only. This module adjusts heuristic thresholds
based on prior report outcomes; it does not guarantee predictive performance.
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd


def list_prior_report_paths(report_dir: Path, run_date: str, max_reports: int) -> list[Path]:
    if not report_dir.exists():
        return []
    candidates = sorted(report_dir.glob("intraday_report_*.xlsx"), reverse=True)
    filtered: list[Path] = []
    run_token = run_date.replace("-", "")
    for path in candidates:
        if run_token in path.stem:
            continue
        filtered.append(path)
        if len(filtered) >= max_reports:
            break
    return filtered


def load_candidate_history(paths: list[Path]) -> pd.DataFrame:
    frames: list[pd.DataFrame] = []
    required = {"symbol", "sector", "atr_pct", "risk_percent_of_entry", "composite_score", "target_achieved"}
    for path in paths:
        try:
            frame = pd.read_excel(path, sheet_name="Intraday Candidates")
        except Exception:
            continue
        if required.issubset(frame.columns):
            frame["source_file"] = path.name
            frames.append(frame)
    if not frames:
        return pd.DataFrame()
    history = pd.concat(frames, ignore_index=True)
    history["target_achieved"] = history["target_achieved"].fillna(False).astype(bool)
    return history
