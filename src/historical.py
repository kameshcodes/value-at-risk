"""
historical_var.py -- Historical Value at Risk calculation and analysis pipeline.
"""

import numpy as np
import pandas as pd
from loguru import logger
from src.utils import fetch_prices, compute_returns, plot_distribution
from src.excel_export import export_historical_var_report


def calculate_historical_var(returns: pd.Series, confidence: float) -> float:
    """Return VaR as a positive loss value.

    VaR = (1 - confidence) percentile of returns, negated to express as loss.
    """
    vals = returns.values
    return -float(np.percentile(np.asarray(vals), (1.0 - confidence) * 100))


def calculate_historical_es(returns: pd.Series, confidence: float) -> float:
    """Return ES as a positive loss value.

    ES = E[loss | loss > VaR], the mean of losses exceeding VaR.
    """
    var = calculate_historical_var(returns, confidence)
    losses = -np.asarray(returns.values)
    tail = losses[losses > var]
    return float(np.mean(tail)) if len(tail) > 0 else var



def compute_historical_var_es(
    returns: pd.Series,
    var_confidence: float,
    es_confidence: float,
    n_days: int,
    portfolio_value: float,
) -> dict:
    """Compute VaR and ES from returns and scale to n-day horizon.

    Returns a dict with 1-day and n-day dollar VaR and ES.
    """
    var_1d_pct = calculate_historical_var(returns, var_confidence)
    es_1d_pct = calculate_historical_es(returns, es_confidence)
    var_1d = var_1d_pct * portfolio_value
    es_1d = es_1d_pct * portfolio_value

    scaling_factor = np.sqrt(n_days)
    var_nd = var_1d * scaling_factor
    es_nd = es_1d * scaling_factor

    return {
        "var_1d": var_1d,
        "var_nd": var_nd,
        "es_1d": es_1d,
        "es_nd": es_nd,
    }


def compute_stressed_historical_var_es(
    ticker: str,
    var_confidence: float,
    es_confidence: float,
    n_days: int,
    portfolio_value: float,
    stress_start: str,
    stress_end: str,
    stress_label: str,
) -> dict:
    """Compute Stressed Historical VaR and ES over a defined stress window."""
    prices = fetch_prices(ticker, start_date=stress_start, end_date=stress_end)
    daily_returns = compute_returns(prices, kind="arithmetic")

    result = compute_historical_var_es(daily_returns, var_confidence, es_confidence, n_days, portfolio_value)

    logger.debug(
        f"Stressed VaR: 1d=${result['var_1d']:,.2f}, {n_days}d=${result['var_nd']:,.2f} | "
        f"Stressed ES: 1d=${result['es_1d']:,.2f}, {n_days}d=${result['es_nd']:,.2f}"
    )

    return {
        **result,
        "stress_start": stress_start,
        "stress_end": stress_end,
        "stress_label": stress_label,
        "prices": prices,
    }


def historical_var_es_pipeline(
    ticker: str,
    var_confidence: float,
    es_confidence: float,
    lookback: int,
    n_days: int,
    portfolio_value: float,
    end_date: pd.Timestamp | None = None,
    stress_start: str = "2008-01-01",
    stress_end: str = "2008-12-31",
    stress_label: str = "Global Financial Crisis (2008)",
):
    """Execute the full Historical VaR pipeline.

    Returns a dict with all computed results.
    PnL and VaR values are expressed in dollars based on *portfolio_value*.
    If end_date is None, defaults to current date.
    """
    # 1. Fetch data and compute returns
    prices = fetch_prices(ticker, lookback, end_date)
    daily_returns = compute_returns(prices, kind="arithmetic")

    # 2. Compute normal VaR and ES
    normal = compute_historical_var_es(
        daily_returns, var_confidence, es_confidence, n_days, portfolio_value,
    )

    # 3. Compute Stressed VaR/ES
    stressed = compute_stressed_historical_var_es(
        ticker, var_confidence, es_confidence, n_days, portfolio_value,
        stress_start, stress_end, stress_label,
    )

    # 4. Generate Excel report (normal + stressed sheets)
    excel_path = export_historical_var_report(
        prices=prices,
        ticker=ticker,
        n_days=n_days,
        portfolio_value=portfolio_value,
        var_date=end_date,
        lookback=lookback,
        stressed_prices=stressed["prices"],
        stress_start=stressed["stress_start"],
        stress_end=stressed["stress_end"],
        stress_label=stressed["stress_label"],
        var_confidence=var_confidence,
        es_confidence=es_confidence,
    )

    # 5. Generate distribution plot
    var_date_str = end_date.strftime("%Y-%m-%d") if end_date else ""
    var_conf_pct = f"{var_confidence * 100:g}"
    es_conf_pct = f"{es_confidence * 100:g}"
    fig_dist = plot_distribution(
        returns=daily_returns * portfolio_value,
        var_cutoff=-normal["var_nd"],
        var_label=f"VaR ({var_conf_pct}%, {n_days}d)",
        es_cutoff=-normal["es_nd"],
        es_label=f"ES ({es_conf_pct}%, {n_days}d)",
        var_date=var_date_str,
        method="Historical",
        ticker=ticker,
    )

    logger.debug(
        f"VaR: 1d=${normal['var_1d']:,.2f}, {n_days}d=${normal['var_nd']:,.2f} | "
        f"ES: 1d=${normal['es_1d']:,.2f}, {n_days}d=${normal['es_nd']:,.2f}"
    )

    return {
        **normal,
        "stressed_var_nd": stressed["var_nd"],
        "stressed_es_nd": stressed["es_nd"],
        "prices": prices,
        "daily_returns": daily_returns,
        "excel_path": excel_path,
        "fig_dist": fig_dist,
    }
