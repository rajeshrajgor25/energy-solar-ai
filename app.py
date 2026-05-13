import streamlit as st
from pathlib import Path
import sys
import math

sys.path.insert(0, str(Path(__file__).parent))

from extractor import extract_bill_data
from excel_writer import fill_excel
from config import GROQ_API_KEY

st.set_page_config(
    page_title="EnergyBae Solar Calculator",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #FF6B00, #FF8C00);
        color: white;
        padding: 1.5rem 2rem;
        border-radius: 12px;
        margin-bottom: 1.5rem;
        text-align: center;
    }
    .main-header h1 { margin: 0; font-size: 2rem; font-weight: 700; }
    .main-header p  { margin: 0.3rem 0 0; opacity: 0.9; font-size: 0.95rem; }
    .card {
        background: #fff;
        border: 1px solid #e5e7eb;
        border-radius: 10px;
        padding: 1.2rem 1.5rem;
        margin-bottom: 1rem;
        box-shadow: 0 1px 4px rgba(0,0,0,0.06);
    }
    .card-title {
        font-size: 0.78rem;
        color: #6b7280;
        text-transform: uppercase;
        letter-spacing: 0.06em;
        margin-bottom: 0.3rem;
    }
    .card-value { font-size: 1.3rem; font-weight: 700; color: #111827; }
    .highlight { color: #FF6B00; }
    .section-header {
        background: #fff7ed;
        border-left: 4px solid #FF6B00;
        padding: 0.6rem 1rem;
        border-radius: 0 8px 8px 0;
        margin: 1rem 0 0.5rem;
        font-weight: 600;
        color: #7c3500;
        font-size: 0.9rem;
    }
    .stDownloadButton > button {
        background: linear-gradient(135deg, #FF6B00, #FF8C00) !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        font-size: 1rem !important;
        padding: 0.7rem 2rem !important;
        width: 100%;
    }
    .stDownloadButton > button:hover { opacity: 0.9 !important; }
    div[data-testid="stAlert"] { border-radius: 8px; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
    <h1>EnergyBae - Solar Load Calculator</h1>
    <p>Upload your electricity bill — AI extracts data — Download filled solar report</p>
</div>
""", unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.markdown("""
    <div style='background:linear-gradient(135deg,#FF6B00,#FF8C00);
                border-radius:10px;padding:14px 16px;margin-bottom:16px;text-align:center'>
        <div style='color:white;font-weight:700;font-size:1.2rem;letter-spacing:0.03em'>EnergyBae</div>
        <div style='color:rgba(255,255,255,0.85);font-size:0.72rem;margin-top:4px'>
            Renewable Energy Solutions
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("#### Solar Assumptions")
    st.caption("Adjust default values for your region")

    peak_sun = st.slider("Peak Sun Hours (hrs/day)", min_value=3.5, max_value=6.0, value=4.5, step=0.1, format="%.1f hrs")

    efficiency_pct = st.slider("System Efficiency (%)", min_value=70, max_value=90, value=80, step=1, format="%d%%")
    efficiency = efficiency_pct / 100

    panel_wp = st.selectbox("Panel Wattage (Wp)", [400, 450, 500, 540, 550, 600], index=3)

    cost_kwp = st.number_input("Cost per kWp (Rs.)", min_value=30000, max_value=100000, value=55000, step=1000)

    subsidy = st.slider("Govt. Subsidy (%)", min_value=0, max_value=40, value=30, step=5, format="%d%%")

    st.markdown("---")
    st.markdown("**Current Settings**")
    st.markdown(f"""
    <div style='background:#fff7ed;border-radius:8px;padding:10px 12px;
                font-size:0.78rem;color:#7c3500;line-height:2.0'>
        Sun Hours &nbsp;&nbsp;&nbsp; <b>{peak_sun:.1f} hrs/day</b><br>
        Efficiency &nbsp;&nbsp;&nbsp; <b>{efficiency_pct}%</b><br>
        Panel Size &nbsp;&nbsp;&nbsp; <b>{panel_wp} Wp</b><br>
        Cost/kWp &nbsp;&nbsp;&nbsp;&nbsp; <b>Rs. {cost_kwp:,}</b><br>
        Subsidy &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <b>{subsidy}%</b>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("""
    <div style='font-size:0.74rem;color:#6b7280;line-height:2.0'>
        energy@gmail.com<br>
        +91 91145 67890<br>
        Energy, Pune
    </div>
    """, unsafe_allow_html=True)

# Main layout
col_upload, col_results = st.columns([1, 1.2], gap="large")

with col_upload:
    st.markdown('<div class="section-header">Upload Electricity Bill</div>', unsafe_allow_html=True)

    uploaded_file = st.file_uploader(
        "Drag and drop or click to upload",
        type=["pdf", "png", "jpg", "jpeg"],
        help="Supports MSEDCL and most Indian electricity bills"
    )

    if uploaded_file:
        st.success(f"Uploaded: {uploaded_file.name} ({uploaded_file.size/1024:.1f} KB)")
        if uploaded_file.name.lower().endswith((".png", ".jpg", ".jpeg")):
            st.image(uploaded_file, caption="Uploaded Bill", use_column_width=True)

    st.markdown("---")

    already_processed = "bill_data" in st.session_state

    with st.expander("Manual Override (optional)", expanded=False):
        st.caption("Correct any field if AI extraction was inaccurate")
        prev    = st.session_state.get("bill_data", {})
        m_units = st.number_input("Units Consumed (kWh)", 0.0, 100000.0, float(prev.get("units_consumed") or 0.0), key="m_units")
        m_load  = st.number_input("Sanctioned Load (kW)", 0.0, 1000.0,   float(prev.get("sanctioned_load") or 0.0), key="m_load")
        m_rate  = st.number_input("Per Unit Rate (Rs./kWh)", 0.0, 50.0,  float(prev.get("per_unit_rate") or 0.0), key="m_rate")
        m_bill  = st.number_input("Total Bill Amount (Rs.)", 0.0, 10000000.0, float(prev.get("total_monthly_bill") or 0.0), key="m_bill")

        if already_processed:
            if st.button("Recalculate with These Values", use_container_width=True):
                bd = dict(st.session_state["bill_data"])
                if m_units > 0: bd["units_consumed"]     = m_units
                if m_load  > 0: bd["sanctioned_load"]    = m_load
                if m_rate  > 0: bd["per_unit_rate"]      = m_rate
                if m_bill  > 0: bd["total_monthly_bill"] = m_bill
                st.session_state["bill_data"] = bd
                st.rerun()

    run_btn = st.button(
        "Extract and Calculate Solar System",
        type="primary",
        disabled=(not uploaded_file),
        use_container_width=True
    )

    if not uploaded_file and not already_processed:
        st.info("Upload a bill to get started.")

# Processing
if run_btn and uploaded_file:
    with col_results:
        progress = st.progress(0, text="Starting...")
        status   = st.empty()
        try:
            progress.progress(10, text="Reading file...")
            file_bytes = uploaded_file.read()

            progress.progress(30, text="AI extracting bill data...")
            status.info("Sending to Groq AI for extraction...")
            bill_data = extract_bill_data(file_bytes, uploaded_file.name, GROQ_API_KEY)

            if st.session_state.get("m_units", 0) > 0: bill_data["units_consumed"]     = st.session_state.m_units
            if st.session_state.get("m_load",  0) > 0: bill_data["sanctioned_load"]    = st.session_state.m_load
            if st.session_state.get("m_rate",  0) > 0: bill_data["per_unit_rate"]      = st.session_state.m_rate
            if st.session_state.get("m_bill",  0) > 0: bill_data["total_monthly_bill"] = st.session_state.m_bill

            progress.progress(60, text="Filling Excel template...")
            status.info("Mapping extracted data to Excel...")

            solar_assumptions = {
                "peak_sun_hours":    peak_sun,
                "system_efficiency": efficiency,
                "panel_wattage":     panel_wp,
                "cost_per_kwp":      cost_kwp,
                "subsidy_pct":       subsidy / 100,
                "degradation_rate":  0.005,
                "lifespan_years":    25,
            }

            output_path = fill_excel(bill_data, solar_assumptions)

            progress.progress(90, text="Finalizing report...")
            status.empty()
            progress.progress(100, text="Done!")

            st.session_state["bill_data"]         = bill_data
            st.session_state["output_path"]       = str(output_path)
            st.session_state["solar_assumptions"] = solar_assumptions

        except Exception as e:
            progress.empty()
            st.error(f"Error: {str(e)}")
            st.exception(e)

# Results
with col_results:
    if "bill_data" in st.session_state:
        bill_data = st.session_state["bill_data"]

        solar_assump = {
            "peak_sun_hours":    peak_sun,
            "system_efficiency": efficiency,
            "panel_wattage":     panel_wp,
            "cost_per_kwp":      cost_kwp,
            "subsidy_pct":       subsidy / 100,
            "degradation_rate":  0.005,
            "lifespan_years":    25,
        }

        output_path = fill_excel(bill_data, solar_assump)

        st.markdown('<div class="section-header">Extracted Bill Information</div>', unsafe_allow_html=True)

        c1, c2 = st.columns(2)
        with c1:
            st.markdown(f'<div class="card"><div class="card-title">Consumer Name</div><div class="card-value">{bill_data.get("consumer_name","—")}</div></div>', unsafe_allow_html=True)
            st.markdown(f'<div class="card"><div class="card-title">Billing Month</div><div class="card-value">{bill_data.get("billing_month","—")}</div></div>', unsafe_allow_html=True)
            st.markdown(f'<div class="card"><div class="card-title">Units Consumed</div><div class="card-value highlight">{bill_data.get("units_consumed","—")} kWh</div></div>', unsafe_allow_html=True)
        with c2:
            st.markdown(f'<div class="card"><div class="card-title">Consumer Number</div><div class="card-value" style="font-size:1rem">{bill_data.get("consumer_number","—")}</div></div>', unsafe_allow_html=True)
            st.markdown(f'<div class="card"><div class="card-title">Total Bill Amount</div><div class="card-value highlight">Rs. {bill_data.get("total_monthly_bill","—")}</div></div>', unsafe_allow_html=True)
            st.markdown(f'<div class="card"><div class="card-title">Per Unit Rate</div><div class="card-value">Rs. {bill_data.get("per_unit_rate","—")}/kWh</div></div>', unsafe_allow_html=True)

        st.markdown('<div class="section-header">Solar System Recommendation</div>', unsafe_allow_html=True)

        try:
            units      = float(bill_data.get("units_consumed") or 0)
            daily      = units / 30 if units else 0
            eff        = solar_assump["system_efficiency"]
            psh        = solar_assump["peak_sun_hours"]
            req_kWp    = (daily / (psh * eff)) if (psh and eff and daily) else 0
            rec_kWp    = math.ceil(req_kWp * 2) / 2 if req_kWp else 0
            panels     = math.ceil((rec_kWp * 1000) / solar_assump["panel_wattage"]) if rec_kWp else 0
            cost_gross = rec_kWp * solar_assump["cost_per_kwp"]
            net_cost   = cost_gross - (cost_gross * solar_assump["subsidy_pct"])
            rate       = float(bill_data.get("per_unit_rate") or 0)
            annual_gen = rec_kWp * psh * 365 * eff if rec_kWp else 0
            savings    = annual_gen * rate if rate else 0
            payback    = (net_cost / savings) if savings else 0
            rooftop    = rec_kWp * 100
        except Exception:
            rec_kWp = panels = net_cost = savings = payback = rooftop = 0

        ca, cb, cc = st.columns(3)
        with ca:
            st.metric("Recommended Size", f"{rec_kWp:.1f} kWp")
            st.metric("Solar Panels", f"{panels} nos.")
        with cb:
            st.metric("Net Cost (after subsidy)", f"Rs. {net_cost:,.0f}")
            st.metric("Rooftop Area", f"{rooftop:.0f} sq.ft.")
        with cc:
            st.metric("Annual Savings", f"Rs. {savings:,.0f}")
            st.metric("Payback Period", f"{payback:.1f} yrs")

        st.markdown("---")
        st.markdown("### Download Report")

        if output_path.exists():
            with open(output_path, "rb") as f:
                excel_bytes = f.read()
            st.download_button(
                label="Download Filled Excel Report",
                data=excel_bytes,
                file_name=output_path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            st.caption("Open in Excel or LibreOffice to see all calculated values")

        st.success("Solar report generated successfully. Download your Excel file above.")

st.markdown("---")
st.markdown("""
<div style='text-align:center;color:#9ca3af;font-size:0.8rem;padding:0.5rem'>
    Energy — Empowering People with Renewable Energy Solutions | Pune<br>
</div>
""", unsafe_allow_html=True)