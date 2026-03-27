"""VaR business logic."""

from src.logger import logger

from src.utils import (
    fetch_prices,
    compute_returns,
    plot_distribution,
)
from src.historical import (
    calculate_historical_var,
    calculate_historical_es,
    compute_historical_var_es,
    compute_stressed_historical_var_es,
    historical_var_es_pipeline,
)
from src.parametric import (
    estimate_distribution,
    calculate_parametric_var,
    calculate_parametric_es,
    compute_parametric_var_es,
    compute_stressed_parametric_var_es,
    parametric_var_es_pipeline,
)
from src.excel_export import export_historical_var_report, export_parametric_var_report
from src.config import (
    TICKERS,
    LOOKBACK_DAYS,
    STRESS_LABEL,
    STRESS_START_DATE,
    STRESS_END_DATE,
)

__all__ = [
    "logger",
    "fetch_prices",
    "compute_returns",
    "plot_distribution",
    "calculate_historical_var",
    "calculate_historical_es",
    "compute_historical_var_es",
    "compute_stressed_historical_var_es",
    "historical_var_es_pipeline",
    "estimate_distribution",
    "calculate_parametric_var",
    "calculate_parametric_es",
    "compute_parametric_var_es",
    "compute_stressed_parametric_var_es",
    "parametric_var_es_pipeline",
    "export_historical_var_report",
    "export_parametric_var_report",
    "TICKERS",
    "LOOKBACK_DAYS",
    "STRESS_LABEL",
    "STRESS_START_DATE",
    "STRESS_END_DATE",
]
