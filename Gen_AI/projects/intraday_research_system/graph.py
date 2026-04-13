"""
LangGraph orchestration for the intraday research system.

Educational and research use only. This graph produces analysis artifacts and
does not submit orders.
"""

from __future__ import annotations

from datetime import date
from logging import getLogger

from langgraph.graph import END, START, StateGraph

from .agents import (
    ExplanationReportingAgent,
    IntradayStrategyScoringAgent,
    LearningAgent,
    NewsSentimentAgent,
    RiskManagementAgent,
    TechnicalAnalysisAgent,
    UniverseDataAgent,
)
from .providers import build_provider
from .reporting import write_excel_report
from .settings import AppConfig
from .state import IntradayResearchState

LOGGER = getLogger(__name__)


def build_research_graph(config: AppConfig):
    provider = build_provider(config)
    graph = StateGraph(IntradayResearchState)
    graph.add_node("universe_data", UniverseDataAgent(provider).run)
    graph.add_node("technical", TechnicalAnalysisAgent().run)
    graph.add_node("sentiment", NewsSentimentAgent().run)
    graph.add_node("learning", LearningAgent().run)
    graph.add_node("strategy", IntradayStrategyScoringAgent().run)
    graph.add_node("risk", RiskManagementAgent().run)
    graph.add_node("explain", ExplanationReportingAgent().run)
    graph.add_edge(START, "universe_data")
    graph.add_edge("universe_data", "technical")
    graph.add_edge("technical", "sentiment")
    graph.add_edge("sentiment", "learning")
    graph.add_edge("learning", "strategy")
    graph.add_edge("strategy", "risk")
    graph.add_edge("risk", "explain")
    graph.add_edge("explain", END)
    return graph.compile()


def run_daily_research(config: AppConfig, run_date: date) -> IntradayResearchState:
    state = IntradayResearchState(config=config, run_date=run_date)
    graph = build_research_graph(config)
    result = graph.invoke(state)
    if isinstance(result, dict):
        normalized = IntradayResearchState(**result)
    else:
        normalized = result
    report_path = write_excel_report(normalized)
    LOGGER.info("Report output path: %s", report_path)
    normalized.daily_summary["report_path"] = str(report_path)
    return normalized
