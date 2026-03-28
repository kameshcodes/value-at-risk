"""
utils.py -- Shared utilities: data fetching, return computation, and plotting.
"""

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import yfinance as yf
from loguru import logger


def fetch_prices(
    ticker: str,
    lookback: int | None = None,
    var_date: pd.Timestamp | None = None,
    start_date: str | None = None,
    end_date: str | None = None,
) -> pd.Series:
    """Download close prices for *ticker*.

    Two modes of operation:

    **Lookback mode** (default):  Supply *lookback* and optionally *var_date*.
    Fetches the last *lookback* trading days ending before *var_date*.

    **Date-range mode**:  Supply *start_date* and *end_date* (YYYY-MM-DD strings).
    Fetches all trading days in that window, plus one prior day so the
    first daily return falls on or near *start_date*.
    """
    if start_date and end_date:
        # Date-range mode (stress periods)
        start = pd.to_datetime(start_date) - pd.Timedelta(days=10)
        end = pd.to_datetime(end_date) + pd.Timedelta(days=1)  # yfinance 'end' is exclusive

        logger.debug(
            f"Fetching {ticker}: {start.strftime('%Y-%m-%d')} to {end_date}"
        )

        try:
            df = yf.download(
                ticker,
                start=start.strftime("%Y-%m-%d"),
                end=end.strftime("%Y-%m-%d"),
                progress=False,
                interval="1d",
                auto_adjust=True,
            )
        except Exception:
            raise ValueError(
                f"No data returned for ticker '{ticker}' ({start_date} to {end_date})."
            )
        if not isinstance(df, pd.DataFrame) or df.empty:
            raise ValueError(
                f"No data returned for ticker '{ticker}' ({start_date} to {end_date})."
            )

        prices = pd.Series(df["Close"].squeeze())
        prices.name = ticker

        # Trim to one trading day before start_date through end_date
        start_ts = pd.to_datetime(start_date)
        start_idx = prices.index.searchsorted(start_ts)
        start_idx = max(0, start_idx - 1)
        prices = prices.iloc[start_idx:]
        prices = prices.loc[:end_date]

        logger.info(
            f"Fetched {len(prices)} trading days for {ticker} "
            f"({prices.index[0].strftime('%Y-%m-%d')} to {prices.index[-1].strftime('%Y-%m-%d')})"
        )
        return prices

    # Lookback mode (historical VaR)
    if var_date is None:
        var_date = pd.Timestamp((pd.Timestamp.today() - pd.offsets.BDay()).date())
    
    if lookback is None:
        raise ValueError("lookback is required when start_date/end_date are not provided.")
    calendar_days = int(lookback * 1.6)
    # yfinance 'end' is exclusive, so passing var_date fetches up to the day before
    start = var_date - pd.Timedelta(days=calendar_days)
    logger.debug(
        f"Fetching {ticker}: {start.strftime('%Y-%m-%d')} to {var_date.strftime('%Y-%m-%d')} (lookback={lookback})"
    )

    try:
        df = yf.download(
            ticker,
            start=start.strftime("%Y-%m-%d"),
            end=var_date.strftime("%Y-%m-%d"),
            progress=False,
            interval="1d",
            auto_adjust=True
        )
    except Exception:
        raise ValueError(f"No data returned for ticker '{ticker}'.")
    if not isinstance(df, pd.DataFrame) or df.empty:
        raise ValueError(f"No data returned for ticker '{ticker}'.")

    prices = pd.Series(df["Close"].squeeze())
    prices.name = ticker
    result = prices.tail(lookback)
    logger.info(
        f"Fetched {len(result)} trading days for {ticker} (last date: {result.index[-1].strftime('%Y-%m-%d')})"
    )
    return result


# ------------------------------------------------------------------
# Return computation
# ------------------------------------------------------------------


def compute_returns(prices: pd.Series, kind: str = "arithmetic") -> pd.Series:
    """Compute daily returns from a price series.

    Parameters
    ----------
    kind : "arithmetic" or "log"
        arithmetic  ->  (P_t - P_{t-1}) / P_{t-1}
        log         ->  log(P_t) - log(P_{t-1})
    """
    if kind == "log":
        log_prices = pd.Series(np.log(prices))
        returns = log_prices - log_prices.shift(1)
        name = "Daily Log Return"
    else:
        returns = (prices - prices.shift(1)) / prices.shift(1)
        name = "Daily Return"
    returns = pd.Series(returns, name=name)
    return returns.dropna()



# ------------------------------------------------------------------
# Plotting (Plotly)
# ------------------------------------------------------------------


def plot_distribution(
    returns: pd.Series,
    var_cutoff: float,
    var_label: str = "VaR",
    es_cutoff: float | None = None,
    es_label: str = "ES",
    var_date: str = "",
    method: str = "",
    ticker: str = "",
) -> go.Figure:
    """Return a histogram of the daily P&L distribution highlighting VaR and ES tail risk."""
    fig = go.Figure()

    # Split the distribution at the VaR cutoff (P&L below VaR are in the left tail)
    normal_returns = returns[returns >= var_cutoff]
    tail_returns = returns[returns < var_cutoff]

    fig.add_trace(
        go.Histogram(
            x=normal_returns.values,
            marker_color="steelblue",
            opacity=0.8,
        )
    )
    fig.add_trace(
        go.Histogram(
            x=tail_returns.values,
            marker_color="darkorange",
            opacity=0.8,
        )
    )

    if var_cutoff is not None:
        fig.add_vline(x=var_cutoff, line_width=1.5, line_dash="dot", line_color="black")
        fig.add_annotation(
            x=var_cutoff, xref="x",
            y=0.5, yref="paper",
            text=f"{var_label}<br>= ${abs(var_cutoff):,.2f}",
            xanchor="left", yanchor="middle",
            xshift=6,
            showarrow=False,
            font=dict(size=9, color="#444444"),
        )

    if es_cutoff is not None:
        fig.add_vline(x=es_cutoff, line_width=1.5, line_dash="dash", line_color="darkred")
        fig.add_annotation(
            x=es_cutoff, xref="x",
            y=0.5, yref="paper",
            text=f"{es_label}<br>= ${abs(es_cutoff):,.2f}",
            xanchor="right", yanchor="middle",
            xshift=-6,
            showarrow=False,
            font=dict(size=9, color="darkred"),
        )

    title = "Daily Portfolio P&L Distribution with VaR & ES Thresholds"

    fig.update_layout(
        title=dict(text=title, font=dict(size=14)),
        xaxis_title=dict(text="P&L ($)", font=dict(size=12)),
        yaxis_title=dict(text="Frequency", font=dict(size=12)),
        barmode="stack",
        template="plotly_white",
        yaxis=dict(showgrid=False),
        margin=dict(t=80, b=40),
        height=388.5,
        showlegend=False,
    )

    if var_date:
        fig.add_annotation(
            text=f"VaR Date: {var_date}",
            xref="paper", yref="paper",
            x=1.08, y=1.22,
            xanchor="right", yanchor="top",
            showarrow=False,
            font=dict(size=9, color="#444444"),
        )

    if method:
        fig.add_annotation(
            text=f"Method: {method}",
            xref="paper", yref="paper",
            x=1.08, y=1.16,
            xanchor="right", yanchor="top",
            showarrow=False,
            font=dict(size=9, color="#444444"),
        )

    if ticker:
        fig.add_annotation(
            text=f"Ticker: {ticker}",
            xref="paper", yref="paper",
            x=1.08, y=1.10,
            xanchor="right", yanchor="top",
            showarrow=False,
            font=dict(size=9, color="#444444"),
        )

    return fig
