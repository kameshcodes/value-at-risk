"""
app.py -- Gradio UI for Value at Risk Analysis.
"""

import pandas as pd
import gradio as gr
from src.logger import logger
from src.config import TICKERS, LOOKBACK_DAYS, STRESS_START_DATE, STRESS_END_DATE, STRESS_LABEL
from src.historical import historical_var_es_pipeline
from src.parametric import parametric_var_es_pipeline


def calculate_var_analysis(
    ticker: str,
    end_date_str: str,
    portfolio_value: float,
    n_days: int,
    var_confidence_label: str,
    es_confidence_label: str,
    method: str,
):
    """Calculate Value at Risk analysis based on Gradio inputs and delegate to the analysis pipeline."""
    logger.info(
        f"Analysis requested: {ticker} | VaR={var_confidence_label} ES={es_confidence_label} | {method} | N={n_days} | Date={end_date_str} | PV=${portfolio_value:,.0f}"
    )

    var_confidence = float(var_confidence_label.strip().replace("%", "")) / 100.0
    es_confidence = float(es_confidence_label.strip().replace("%", "")) / 100.0
    var_conf_pct = var_confidence_label.strip()
    es_conf_pct = es_confidence_label.strip()

    today = pd.Timestamp.today().normalize()

    try:
        end_date = pd.to_datetime(end_date_str, errors="raise").normalize()
    except Exception:
        gr.Warning("Invalid date selection. Please try again.")
        return reset_analysis_results(n_days, var_confidence_label, es_confidence_label, method)

    if end_date >= today:
        gr.Warning(
            "Invalid date selection. VaR estimation requires historical data, so please choose a date prior to today."
        )
        return reset_analysis_results(n_days, var_confidence_label, es_confidence_label, method)

    fy_start = pd.Timestamp(year=today.year - 1, month=4, day=1)

    if end_date < fy_start:
        gr.Warning(f"VaR Date must be after {fy_start.strftime('%Y-%m-%d')}")
        return reset_analysis_results(n_days, var_confidence_label, es_confidence_label, method)

    if portfolio_value <= 0:
        gr.Warning("Portfolio value must be positive.")
        return reset_analysis_results(n_days, var_confidence_label, es_confidence_label, method)

    pipeline = historical_var_es_pipeline if method == "Historical VaR" else parametric_var_es_pipeline
    result = pipeline(
        ticker,
        var_confidence,
        es_confidence,
        LOOKBACK_DAYS,
        int(n_days),
        portfolio_value,
        end_date,
        stress_start=STRESS_START_DATE,
        stress_end=STRESS_END_DATE,
        stress_label=STRESS_LABEL,
    )

    logger.success(
        f"Analysis complete: VaR=${result['var_nd']:,.2f}, ES=${result['es_nd']:,.2f}, Excel={result['excel_path']}\n"
    )

    return (
        gr.update(
            label=f"{int(n_days)}-day {var_conf_pct} VaR",
            value=f"${result['var_nd']:,.2f}",
        ),
        gr.update(
            label=f"{int(n_days)}-day {es_conf_pct} ES",
            value=f"${result['es_nd']:,.2f}",
        ),
        gr.update(
            label=f"{int(n_days)}-day {var_conf_pct} Stressed VaR",
            value=f"${result['stressed_var_nd']:,.2f}",
        ),
        gr.update(
            label=f"{int(n_days)}-day {es_conf_pct} Stressed ES",
            value=f"${result['stressed_es_nd']:,.2f}",
        ),
        gr.update(value=result["fig_dist"], visible=True),
        gr.update(value=result["excel_path"], visible=True),
    )


def reset_analysis_results(n_days: float, var_confidence_label: str, es_confidence_label: str, method: str):
    """Reset and hide analysis results when input parameters are modified."""
    var_conf_pct = var_confidence_label.strip()
    es_conf_pct = es_confidence_label.strip()
    method_short = "Historical" if method == "Historical VaR" else "Parametric"

    return (
        gr.update(value="", label=f"{int(n_days)}-day {var_conf_pct} VaR"),
        gr.update(value="", label=f"{int(n_days)}-day {es_conf_pct} ES"),
        gr.update(value="", label=f"{int(n_days)}-day {var_conf_pct} Stressed VaR"),
        gr.update(value="", label=f"{int(n_days)}-day {es_conf_pct} Stressed ES"),
        gr.update(value=None, visible=False),
        gr.update(visible=False),
    )


def enable_run_button_for_method(method: str):
    """Enable the Run Analysis button only for fully implemented VaR methods."""
    return gr.update(interactive=(method in ("Historical VaR", "Parametric VaR")))


# ------------------------------------------------------------------
# UI builder
# ------------------------------------------------------------------


def build_app() -> gr.Blocks:
    """Construct and return the Gradio Blocks application."""

    with gr.Blocks(
        title="VaR Engine",
    ) as app:
        with gr.Row():
            with gr.Column(scale=3):
                gr.Markdown("# VaR Engine")
            with gr.Column(scale=1, min_width=260):
                download_file = gr.DownloadButton(
                    "Download Excel Report",
                    variant="primary",
                    visible=False,
                    elem_id="excel-btn",
                )

        with gr.Row():
            # Sidebar
            with gr.Column(scale=1, min_width=260):
                gr.Markdown("### Inputs")
                with gr.Group():
                    ticker_dd = gr.Dropdown(
                        choices=TICKERS,
                        value=TICKERS[0],
                        label="Ticker",
                    )
                    end_date_input = gr.DateTime(
                        include_time=False,
                        type="datetime",
                        label="VaR Date",
                    )
                    portfolio_val_input = gr.Number(
                        value=1000000,
                        label="Portfolio Value ($)",
                    )
                    n_days_slider = gr.Slider(
                        minimum=1,
                        maximum=15,
                        step=1,
                        value=10,
                        label="N Days for VaR",
                    )
                    with gr.Row():
                        confidence_dd = gr.Dropdown(
                            choices=["99%", "97.5%", "95%"],
                            value="99%",
                            label="VaR Confidence",
                            min_width=100,
                        )
                        es_confidence_dd = gr.Dropdown(
                            choices=["99%", "97.5%", "95%"],
                            value="99%",
                            label="ES Confidence",
                            min_width=100,
                        )
                    method_radio = gr.Radio(
                        choices=[
                            "Historical VaR",
                            "Parametric VaR",
                        ],
                        value="Historical VaR",
                        label="Method",
                    )
                run_btn = gr.Button("Run Analysis", variant="primary")

                # Enable/disable run button based on method availability
                method_radio.change(
                    fn=enable_run_button_for_method,
                    inputs=method_radio,
                    outputs=run_btn,
                )
            with gr.Column(scale=3):
                gr.Markdown("### Results")
                with gr.Row():
                    with gr.Column(scale=1):
                        var_box = gr.Textbox(
                            label="10-day 99% VaR",
                            interactive=False,
                        )
                    with gr.Column(scale=1):
                        es_box = gr.Textbox(
                            label="10-day 99% ES",
                            interactive=False,
                        )
                with gr.Row():
                    with gr.Column(scale=1):
                        stressed_var_box = gr.Textbox(
                            label="10-day 99% Stressed VaR",
                            interactive=False,
                        )
                    with gr.Column(scale=1):
                        stressed_es_box = gr.Textbox(
                            label="10-day 99% Stressed ES",
                            interactive=False,
                        )

                plot_dist = gr.Plot(show_label=False, visible=False)

        # Wiring
        all_outputs = [var_box, es_box, stressed_var_box, stressed_es_box, plot_dist, download_file]

        run_btn.click(
            fn=calculate_var_analysis,
            inputs=[
                ticker_dd,
                end_date_input,
                portfolio_val_input,
                n_days_slider,
                confidence_dd,
                es_confidence_dd,
                method_radio,
            ],
            outputs=all_outputs,
        )

        all_inputs = [
            ticker_dd,
            end_date_input,
            portfolio_val_input,
            n_days_slider,
            confidence_dd,
            es_confidence_dd,
            method_radio,
        ]

        label_inputs = [n_days_slider, confidence_dd, es_confidence_dd, method_radio]

        for comp in all_inputs:
            if comp is n_days_slider:
                comp.release(
                    fn=reset_analysis_results,
                    inputs=label_inputs,
                    outputs=all_outputs,
                )
            else:
                comp.change(
                    fn=reset_analysis_results,
                    inputs=label_inputs,
                    outputs=all_outputs,
                )

        method_radio.change(
            fn=enable_run_button_for_method, inputs=method_radio, outputs=run_btn
        )

    return app


# ------------------------------------------------------------------
# Entry point
# ------------------------------------------------------------------

if __name__ == "__main__":
    custom_css = """
    .form { border: none !important; box-shadow: none !important; gap: 0 !important; }
    .form .block, .form .row, .form > * { border: none !important; box-shadow: none !important; }
    #excel-btn, #excel-btn.primary {
        background: #ea580c !important;
        background-color: #ea580c !important;
        color: white !important;
        border-color: #7c2d12 !important;
        border: 1px solid #7c2d12 !important;
    }
    #excel-btn:hover, #excel-btn.primary:hover {
        background: #c2410c !important;
        background-color: #c2410c !important;
        border-color: #451a03 !important;
        border: 1px solid #451a03 !important;
    }
    """

    application = build_app()
    application.launch(share=True, theme=gr.themes.Base(), css=custom_css)  
