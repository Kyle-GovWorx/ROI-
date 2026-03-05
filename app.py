import io
import math
from datetime import datetime

import pandas as pd
import streamlit as st


st.set_page_config(page_title="CommsCoach ROI Calculator", layout="wide")


def money(x: float) -> str:
    try:
        return f"${x:,.0f}"
    except Exception:
        return "$0"


def pct_from_ratio(x: float) -> str:
    try:
        return f"{x * 100:,.1f}%"
    except Exception:
        return "0.0%"


def safe_div(a: float, b: float) -> float:
    return a / b if b else 0.0


def build_excel_export(inputs: dict, results: dict, breakdown_df: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        pd.DataFrame(list(inputs.items()), columns=["Input", "Value"]).to_excel(
            writer, sheet_name="Inputs", index=False
        )
        pd.DataFrame(list(results.items()), columns=["Metric", "Value"]).to_excel(
            writer, sheet_name="Results", index=False
        )
        breakdown_df.to_excel(writer, sheet_name="Savings Breakdown", index=False)
    buffer.seek(0)
    return buffer.read()


st.title("CommsCoach ROI Calculator")
st.caption("Estimate annual value, net benefit, ROI, and payback period for CommsCoach.")

with st.sidebar:
    st.header("Investment")
    annual_investment = st.number_input(
        "Annual CommsCoach cost ($)", min_value=0.0, value=225000.0, step=5000.0
    )

    st.header("Agency size")
    annual_calls = st.number_input(
        "Annual calls eligible for QA", min_value=0, value=1200000, step=10000
    )
    dispatchers = st.number_input(
        "Number of dispatchers / telecommunicators", min_value=0, value=300, step=5
    )

    st.header("Manual QA baseline")
    qa_specialists_fte = st.number_input(
        "Dedicated QA specialists (FTE)", min_value=0.0, value=1.0, step=0.25
    )
    qa_fte_fully_loaded_cost = st.number_input(
        "QA specialist fully loaded annual cost ($)", min_value=0.0, value=95000.0, step=5000.0
    )
    supervisor_hours_per_week_on_qa = st.number_input(
        "Supervisor hours per week spent on QA", min_value=0.0, value=20.0, step=1.0
    )
    supervisor_hourly_fully_loaded = st.number_input(
        "Supervisor fully loaded hourly rate ($/hr)", min_value=0.0, value=110.0, step=1.0
    )
    manual_qa_coverage_pct = st.number_input(
        "Manual QA coverage (% of calls)", min_value=0.0, max_value=100.0, value=5.0, step=1.0
    )

    st.header("CommsCoach impact")
    qa_labor_reduction_pct = st.number_input(
        "Reduction in manual QA labor (%)", min_value=0.0, max_value=100.0, value=70.0, step=5.0
    )
    supervisor_time_reduction_pct = st.number_input(
        "Reduction in supervisor QA time (%)", min_value=0.0, max_value=100.0, value=70.0, step=5.0
    )

    st.header("Turnover")
    annual_turnover_pct = st.number_input(
        "Annual dispatcher turnover (%)", min_value=0.0, max_value=100.0, value=15.0, step=1.0
    )
    replacement_cost_per_dispatcher = st.number_input(
        "Replacement cost per dispatcher ($)", min_value=0.0, value=35000.0, step=1000.0
    )
    turnover_reduction_pct = st.number_input(
        "Turnover reduction due to CommsCoach (%)", min_value=0.0, max_value=100.0, value=10.0, step=1.0
    )

    st.header("Training")
    new_hires_per_year = st.number_input(
        "New hires per year", min_value=0, value=60, step=1
    )
    training_hours_per_new_hire = st.number_input(
        "Training hours per new hire", min_value=0.0, value=80.0, step=1.0
    )
    trainer_hourly_fully_loaded = st.number_input(
        "Trainer fully loaded hourly rate ($/hr)", min_value=0.0, value=95.0, step=1.0
    )
    training_time_reduction_pct = st.number_input(
        "Training time reduction (%)", min_value=0.0, max_value=100.0, value=25.0, step=1.0
    )

    st.header("Optional value buckets")
    annual_productivity_value = st.number_input(
        "Annual productivity / effectiveness value ($)", min_value=0.0, value=0.0, step=10000.0
    )
    annual_risk_reduction_value = st.number_input(
        "Annual risk reduction value ($)", min_value=0.0, value=0.0, step=10000.0
    )


weeks_per_year = 52.0

manual_qa_labor_cost = qa_specialists_fte * qa_fte_fully_loaded_cost
manual_supervisor_qa_cost = (
    supervisor_hours_per_week_on_qa * weeks_per_year * supervisor_hourly_fully_loaded
)
manual_total_cost = manual_qa_labor_cost + manual_supervisor_qa_cost

qa_labor_savings = manual_qa_labor_cost * (qa_labor_reduction_pct / 100.0)
supervisor_time_savings = manual_supervisor_qa_cost * (supervisor_time_reduction_pct / 100.0)

annual_turnovers = dispatchers * (annual_turnover_pct / 100.0)
turnover_cost_baseline = annual_turnovers * replacement_cost_per_dispatcher
turnover_savings = turnover_cost_baseline * (turnover_reduction_pct / 100.0)

baseline_training_cost = (
    new_hires_per_year * training_hours_per_new_hire * trainer_hourly_fully_loaded
)
training_savings = baseline_training_cost * (training_time_reduction_pct / 100.0)

manual_cov = max(0.0, min(1.0, manual_qa_coverage_pct / 100.0))
coverage_uplift_pct = max(0.0, 1.0 - manual_cov) * 100.0

total_gross_savings = (
    qa_labor_savings
    + supervisor_time_savings
    + turnover_savings
    + training_savings
    + annual_productivity_value
    + annual_risk_reduction_value
)
net_benefit = total_gross_savings - annual_investment
roi_ratio = safe_div(net_benefit, annual_investment)
payback_months = (annual_investment / total_gross_savings) * 12.0 if total_gross_savings else math.inf

inputs = {
    "Annual CommsCoach cost ($)": annual_investment,
    "Annual calls eligible for QA": annual_calls,
    "Dispatchers / telecommunicators": dispatchers,
    "Dedicated QA specialists (FTE)": qa_specialists_fte,
    "QA specialist fully loaded annual cost ($)": qa_fte_fully_loaded_cost,
    "Supervisor hours per week on QA": supervisor_hours_per_week_on_qa,
    "Supervisor hourly fully loaded rate ($/hr)": supervisor_hourly_fully_loaded,
    "Manual QA coverage (%)": manual_qa_coverage_pct,
    "QA labor reduction (%)": qa_labor_reduction_pct,
    "Supervisor QA time reduction (%)": supervisor_time_reduction_pct,
    "Annual dispatcher turnover (%)": annual_turnover_pct,
    "Replacement cost per dispatcher ($)": replacement_cost_per_dispatcher,
    "Turnover reduction (%)": turnover_reduction_pct,
    "New hires per year": new_hires_per_year,
    "Training hours per new hire": training_hours_per_new_hire,
    "Trainer hourly fully loaded rate ($/hr)": trainer_hourly_fully_loaded,
    "Training time reduction (%)": training_time_reduction_pct,
    "Annual productivity / effectiveness value ($)": annual_productivity_value,
    "Annual risk reduction value ($)": annual_risk_reduction_value,
}

results = {
    "Manual QA labor cost baseline ($)": manual_qa_labor_cost,
    "Manual supervisor QA cost baseline ($)": manual_supervisor_qa_cost,
    "Manual total QA cost baseline ($)": manual_total_cost,
    "QA labor savings ($)": qa_labor_savings,
    "Supervisor time savings ($)": supervisor_time_savings,
    "Turnover savings ($)": turnover_savings,
    "Training savings ($)": training_savings,
    "Coverage uplift (%)": coverage_uplift_pct,
    "Annual gross savings ($)": total_gross_savings,
    "Net annual benefit ($)": net_benefit,
    "ROI ratio": roi_ratio,
    "Payback period (months)": None if not math.isfinite(payback_months) else payback_months,
}

breakdown_df = pd.DataFrame(
    [
        {"Bucket": "QA specialist labor savings", "Annual Value ($)": qa_labor_savings},
        {"Bucket": "Supervisor time savings", "Annual Value ($)": supervisor_time_savings},
        {"Bucket": "Turnover savings", "Annual Value ($)": turnover_savings},
        {"Bucket": "Training savings", "Annual Value ($)": training_savings},
        {"Bucket": "Productivity value", "Annual Value ($)": annual_productivity_value},
        {"Bucket": "Risk reduction value", "Annual Value ($)": annual_risk_reduction_value},
    ]
)

m1, m2, m3, m4 = st.columns(4)
m1.metric("Annual Gross Savings", money(total_gross_savings))
m2.metric("Annual Investment", money(annual_investment))
m3.metric("Net Annual Benefit", money(net_benefit))
m4.metric("ROI", pct_from_ratio(roi_ratio))

if math.isfinite(payback_months):
    st.info(f"Estimated payback period: {payback_months:,.1f} months")
else:
    st.warning("Payback period cannot be computed because gross savings is $0.")

left, right = st.columns([1.1, 0.9])

with left:
    st.subheader("Savings Breakdown")
    display_df = breakdown_df.copy()
    display_df["Annual Value ($)"] = display_df["Annual Value ($)"].map(lambda x: round(float(x), 0))
    st.dataframe(display_df, use_container_width=True, hide_index=True)

with right:
    st.subheader("Context")
    st.markdown(
        f"""
- **Manual QA labor baseline:** {money(manual_qa_labor_cost)}
- **Manual supervisor QA baseline:** {money(manual_supervisor_qa_cost)}
- **Manual total QA baseline:** {money(manual_total_cost)}
- **Coverage uplift vs manual review:** {coverage_uplift_pct:,.1f}%
- **Annual turnovers at baseline:** {annual_turnovers:,.1f}
- **Turnover cost baseline:** {money(turnover_cost_baseline)}
- **Baseline training cost:** {money(baseline_training_cost)}
"""
    )

st.subheader("Export")
csv_bytes = pd.DataFrame(list(results.items()), columns=["Metric", "Value"]).to_csv(index=False).encode("utf-8")
st.download_button(
    "Download results as CSV",
    data=csv_bytes,
    file_name="commscoach_roi_results.csv",
    mime="text/csv",
)

excel_bytes = build_excel_export(inputs, results, breakdown_df)
st.download_button(
    "Download full export as Excel",
    data=excel_bytes,
    file_name="commscoach_roi_export.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.caption(f"Last calculated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
