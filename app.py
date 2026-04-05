
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import google.generativeai as genai
import math

# ── Page config
st.set_page_config(
    page_title="Inventory Optimization Simulator",
    page_icon="📦",
    layout="wide"
)

# ── Custom CSS
st.markdown("""
<style>
  .main-title { font-size: 2rem; font-weight: 700; color: #1a1a2e; margin-bottom: 0; }
  .sub-title  { font-size: 1rem; color: #666; margin-bottom: 2rem; }
  .metric-card {
    background: #f8f9ff;
    border: 1px solid #e0e4ff;
    border-radius: 12px;
    padding: 1rem 1.2rem;
    margin-bottom: 0.5rem;
  }
  .metric-label { font-size: 0.75rem; color: #888; text-transform: uppercase; letter-spacing: 0.05em; }
  .metric-value { font-size: 1.5rem; font-weight: 700; color: #1a1a2e; }
  .metric-unit  { font-size: 0.8rem; color: #888; }
  .winner-card {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    border-radius: 12px;
    padding: 1.5rem;
    color: white;
    margin-bottom: 1rem;
  }
  .section-header {
    font-size: 1.1rem;
    font-weight: 600;
    color: #1a1a2e;
    border-left: 4px solid #667eea;
    padding-left: 0.75rem;
    margin: 1.5rem 0 1rem;
  }
  .warning-box {
    background: #fff8e1;
    border: 1px solid #ffcc02;
    border-radius: 8px;
    padding: 0.75rem 1rem;
    font-size: 0.875rem;
    color: #7a5c00;
    margin-bottom: 0.5rem;
  }
  .info-box {
    background: #e8f4fd;
    border: 1px solid #90caf9;
    border-radius: 8px;
    padding: 0.75rem 1rem;
    font-size: 0.875rem;
    color: #0d47a1;
    margin-bottom: 0.5rem;
  }
  .stButton > button {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    color: white;
    border: none;
    border-radius: 8px;
    padding: 0.6rem 2rem;
    font-size: 1rem;
    font-weight: 600;
    width: 100%;
  }
</style>
""", unsafe_allow_html=True)

# ════════════════════════════════════════════
# CORE CALCULATION FUNCTIONS
# ════════════════════════════════════════════

def calc_eoq(annual_demand, order_cost, holding_cost_per_unit):
    if holding_cost_per_unit <= 0 or annual_demand <= 0:
        return 0
    return math.sqrt((2 * annual_demand * order_cost) / holding_cost_per_unit)

def calc_safety_stock(z_score, demand_std, lead_time_avg, lead_time_std, daily_demand):
    demand_variability   = (z_score ** 2) * (lead_time_avg) * (demand_std ** 2)
    lead_time_variability = (z_score ** 2) * (daily_demand ** 2) * (lead_time_std ** 2)
    return math.sqrt(demand_variability + lead_time_variability)

def calc_reorder_point(daily_demand, lead_time_avg, safety_stock):
    return (daily_demand * lead_time_avg) + safety_stock

def calc_total_cost(annual_demand, eoq, order_cost, holding_cost_per_unit, safety_stock):
    if eoq <= 0:
        return 0
    order_cost_annual   = (annual_demand / eoq) * order_cost
    holding_cost_annual = ((eoq / 2) + safety_stock) * holding_cost_per_unit
    return order_cost_annual + holding_cost_annual

def apply_seasonal_adjustment(base_demand, is_peak, peak_multiplier, offpeak_multiplier):
    return base_demand * (peak_multiplier if is_peak else offpeak_multiplier)

def apply_supplier_risk_buffer(safety_stock, num_suppliers, reliability_pct):
    if num_suppliers == 1:
        risk_factor = 1 + ((100 - reliability_pct) / 100) * 1.5
    elif num_suppliers == 2:
        risk_factor = 1 + ((100 - reliability_pct) / 100) * 0.8
    else:
        risk_factor = 1 + ((100 - reliability_pct) / 100) * 0.4
    return safety_stock * risk_factor

def apply_price_volatility_adjustment(eoq, price_trend, price_change_pct):
    if price_trend == "Rising":
        return eoq * (1 + price_change_pct / 200)
    elif price_trend == "Falling":
        return eoq * (1 - price_change_pct / 300)
    return eoq

def calc_obsolescence_cap(shelf_life_days, daily_demand):
    return shelf_life_days * daily_demand * 0.8

def calc_working_capital_cost(dead_stock_units, unit_cost, holding_cost_pct):
    return dead_stock_units * unit_cost * (holding_cost_pct / 100)

# Z-score lookup
Z_SCORES = {90: 1.28, 95: 1.645, 97: 1.88, 99: 2.326}

def run_scenario(service_level, annual_demand, order_cost, unit_cost, holding_cost_pct,
                 lead_time_avg, lead_time_std, demand_std, working_days,
                 num_suppliers, reliability_pct,
                 price_trend, price_change_pct,
                 shelf_life_days, apply_obsolescence,
                 dead_stock_units, apply_dead_stock,
                 is_peak_season, peak_multiplier, offpeak_multiplier,
                 moq, warehouse_cap, budget_limit):

    z = Z_SCORES.get(service_level, 1.645)
    holding_cost_per_unit = unit_cost * (holding_cost_pct / 100)
    daily_demand = annual_demand / working_days

    # Apply seasonal adjustment
    adj_demand = apply_seasonal_adjustment(
        annual_demand, is_peak_season, peak_multiplier, offpeak_multiplier)
    adj_daily  = adj_demand / working_days

    # EOQ
    eoq = calc_eoq(adj_demand, order_cost, holding_cost_per_unit)

    # Price volatility adjustment
    eoq = apply_price_volatility_adjustment(eoq, price_trend, price_change_pct)

    # Enforce MOQ
    eoq = max(eoq, moq)

    # Obsolescence cap
    if apply_obsolescence and shelf_life_days > 0:
        obs_cap = calc_obsolescence_cap(shelf_life_days, adj_daily)
        eoq = min(eoq, obs_cap)

    # Safety stock
    ss = calc_safety_stock(z, demand_std, lead_time_avg, lead_time_std, adj_daily)

    # Supplier risk buffer
    ss = apply_supplier_risk_buffer(ss, num_suppliers, reliability_pct)

    # Reorder point
    rop = calc_reorder_point(adj_daily, lead_time_avg, ss)

    # Enforce warehouse cap
    eoq = min(eoq, warehouse_cap - ss) if warehouse_cap > 0 else eoq

    # Budget check
    budget_ok = (eoq * unit_cost) <= budget_limit if budget_limit > 0 else True

    # Total cost
    total_cost = calc_total_cost(adj_demand, eoq, order_cost, holding_cost_per_unit, ss)

    # Dead stock working capital cost
    wc_cost = 0
    if apply_dead_stock and dead_stock_units > 0:
        wc_cost = calc_working_capital_cost(dead_stock_units, unit_cost, holding_cost_pct)
        total_cost += wc_cost

    # Orders per year
    orders_per_year = adj_demand / eoq if eoq > 0 else 0

    # Stockout risk
    stockout_risk = (100 - service_level)

    return {
        "eoq":             round(eoq),
        "safety_stock":    round(ss),
        "reorder_point":   round(rop),
        "total_cost":      round(total_cost, 2),
        "orders_per_year": round(orders_per_year, 1),
        "holding_cost":    round(((eoq / 2) + ss) * holding_cost_per_unit, 2),
        "order_cost_ann":  round((adj_demand / eoq) * order_cost if eoq > 0 else 0, 2),
        "wc_cost":         round(wc_cost, 2),
        "stockout_risk":   stockout_risk,
        "budget_ok":       budget_ok,
        "adj_demand":      round(adj_demand),
        "obs_warning":     apply_obsolescence and (eoq >= calc_obsolescence_cap(shelf_life_days, adj_demand/working_days) * 0.95) if apply_obsolescence else False,
    }

# ════════════════════════════════════════════
# AI RECOMMENDATION FUNCTION
# ════════════════════════════════════════════

def get_ai_recommendation(api_key, product_name, scenarios_data, constraints_summary):
    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel("gemini-1.5-flash")

        prompt = f"""
You are a senior supply chain analyst. A user has run an Inventory Optimization Simulator for their product.

PRODUCT: {product_name}

BUSINESS CONSTRAINTS:
{constraints_summary}

SCENARIO RESULTS:
{scenarios_data}

Your task:
1. Compare all scenarios across: total annual cost, safety stock levels, stockout risk, and feasibility given constraints.
2. Clearly declare ONE winning scenario and explain WHY it is best for this specific business situation.
3. Highlight the top 3 risks the user should be aware of based on their inputs.
4. Give 3 specific, actionable recommendations the user should implement immediately.
5. Flag any red flags in their current inventory setup.

Write in clear, plain English. Be specific with numbers. Format your response with these exact headers:
**RECOMMENDED SCENARIO**
**WHY THIS SCENARIO WINS**
**TOP 3 RISKS**
**3 IMMEDIATE ACTIONS**
**RED FLAGS**
"""
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        return f"AI Error: {str(e)}. Please check your API key."

# ════════════════════════════════════════════
# VISUALIZATION FUNCTIONS
# ════════════════════════════════════════════

def plot_cost_comparison(results, scenario_names):
    fig = go.Figure()
    colors = ["#667eea", "#f093fb", "#4facfe", "#43e97b"]

    for i, (name, r) in enumerate(zip(scenario_names, results)):
        fig.add_trace(go.Bar(
            name=name,
            x=["Holding Cost", "Order Cost", "Working Capital Cost"],
            y=[r["holding_cost"], r["order_cost_ann"], r["wc_cost"]],
            marker_color=colors[i],
            text=[f"${r['holding_cost']:,.0f}", f"${r['order_cost_ann']:,.0f}", f"${r['wc_cost']:,.0f}"],
            textposition="auto",
        ))

    fig.update_layout(
        barmode="group", title="Cost Breakdown by Scenario",
        yaxis_title="Annual Cost (USD)", xaxis_title="Cost Type",
        plot_bgcolor="white", paper_bgcolor="white",
        font=dict(family="Arial", size=12),
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        height=400
    )
    return fig

def plot_safety_stock_rop(results, scenario_names):
    colors = ["#667eea", "#f093fb", "#4facfe", "#43e97b"]
    fig = make_subplots(rows=1, cols=2,
                        subplot_titles=("Safety Stock (units)", "Reorder Point (units)"))

    for i, (name, r) in enumerate(zip(scenario_names, results)):
        fig.add_trace(go.Bar(
            name=name, x=[name], y=[r["safety_stock"]],
            marker_color=colors[i], showlegend=False,
            text=[f"{r['safety_stock']}"], textposition="auto"
        ), row=1, col=1)
        fig.add_trace(go.Bar(
            name=name, x=[name], y=[r["reorder_point"]],
            marker_color=colors[i], showlegend=False,
            text=[f"{r['reorder_point']}"], textposition="auto"
        ), row=1, col=2)

    fig.update_layout(
        plot_bgcolor="white", paper_bgcolor="white",
        height=380, title="Safety Stock & Reorder Point Comparison"
    )
    return fig

def plot_total_cost_risk(results, scenario_names):
    colors = ["#667eea", "#f093fb", "#4facfe", "#43e97b"]
    fig = go.Figure()

    for i, (name, r) in enumerate(zip(scenario_names, results)):
        fig.add_trace(go.Scatter(
            x=[r["stockout_risk"]], y=[r["total_cost"]],
            mode="markers+text",
            marker=dict(size=20, color=colors[i]),
            text=[name], textposition="top center",
            name=name
        ))

    fig.update_layout(
        title="Cost vs Stockout Risk Tradeoff",
        xaxis_title="Stockout Risk (%)",
        yaxis_title="Total Annual Cost (USD)",
        plot_bgcolor="white", paper_bgcolor="white",
        height=400,
        xaxis=dict(autorange="reversed")
    )
    return fig

def plot_eoq_comparison(results, scenario_names):
    colors = ["#667eea", "#f093fb", "#4facfe", "#43e97b"]
    fig = go.Figure(go.Bar(
        x=scenario_names,
        y=[r["eoq"] for r in results],
        marker_color=colors[:len(results)],
        text=[f"{r['eoq']} units" for r in results],
        textposition="auto"
    ))
    fig.update_layout(
        title="Economic Order Quantity (EOQ) per Scenario",
        yaxis_title="Order Quantity (units)",
        plot_bgcolor="white", paper_bgcolor="white",
        height=360
    )
    return fig

# ════════════════════════════════════════════
# MAIN APP UI
# ════════════════════════════════════════════

st.markdown('<p class="main-title">📦 Inventory Optimization Simulator</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">Enter your product data, define scenarios, and let AI recommend the best inventory strategy.</p>', unsafe_allow_html=True)

# ── SIDEBAR — API Key
with st.sidebar:
    st.markdown("### ⚙️ Settings")
    api_key = st.text_input("Gemini API Key", type="password",
                            placeholder="AIzaSy...",
                            help="Get your free key at aistudio.google.com")
    st.markdown("---")
    st.markdown("### 📖 How to use")
    st.markdown("""
1. Enter your product details
2. Set business constraints
3. Choose scenarios to compare
4. Click **Run Simulation**
5. Get AI recommendation
""")

# ════════════════════════════════════════
# TAB LAYOUT
# ════════════════════════════════════════
tab1, tab2, tab3 = st.tabs(["📋 Product & Constraints", "🎯 Scenarios", "📊 Results & AI"])

# ════════════════════════════════════════
# TAB 1 — PRODUCT & CONSTRAINTS
# ════════════════════════════════════════
with tab1:
    st.markdown('<p class="section-header">Basic Product Information</p>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)

    with col1:
        product_name     = st.text_input("Product Name", value="Product A")
        annual_demand    = st.number_input("Annual Demand (units)", min_value=100, value=10000, step=100)
        demand_std       = st.number_input("Demand Std Dev (units/day)", min_value=0.0, value=15.0, step=1.0,
                                           help="How much daily demand varies. Higher = more unpredictable.")
    with col2:
        unit_cost        = st.number_input("Unit Cost (USD)", min_value=0.1, value=25.0, step=0.5)
        order_cost       = st.number_input("Order Cost per Order (USD)", min_value=1.0, value=150.0, step=10.0,
                                           help="Cost of placing one order — admin, shipping, processing.")
        holding_cost_pct = st.number_input("Annual Holding Cost (%)", min_value=1.0, max_value=50.0, value=20.0, step=1.0,
                                           help="Usually 20–30% of unit cost. Covers storage, insurance, capital.")
    with col3:
        working_days     = st.number_input("Working Days per Year", min_value=200, max_value=365, value=250)
        moq              = st.number_input("Min Order Quantity (MOQ)", min_value=0, value=50, step=10,
                                           help="Minimum units your supplier will ship per order.")
        warehouse_cap    = st.number_input("Warehouse Capacity (units, 0 = no limit)", min_value=0, value=0, step=100)
        budget_limit     = st.number_input("Budget per Order (USD, 0 = no limit)", min_value=0.0, value=0.0, step=100.0)

    st.markdown('<p class="section-header">Supply & Lead Time</p>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    with col1:
        lead_time_avg    = st.number_input("Avg Lead Time (days)", min_value=1, value=14, step=1)
        lead_time_std    = st.number_input("Lead Time Std Dev (days)", min_value=0.0, value=3.0, step=0.5,
                                           help="How much lead time varies. High value = unreliable supplier.")
    with col2:
        num_suppliers    = st.selectbox("Number of Suppliers", [1, 2, 3, 4, 5], index=0)
        reliability_pct  = st.slider("Supplier Reliability (%)", min_value=50, max_value=100, value=85,
                                     help="% of orders delivered on time.")
    with col3:
        backup_supplier  = st.radio("Backup Supplier Available?", ["No", "Yes"], horizontal=True)

    st.markdown('<p class="section-header">Business Constraints</p>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**🌡️ Seasonal Demand**")
        apply_seasonal   = st.checkbox("Apply seasonal demand adjustment", value=False)
        if apply_seasonal:
            is_peak      = st.radio("Current season:", ["Peak Season", "Off-Peak Season"], horizontal=True)
            peak_mult    = st.slider("Peak demand multiplier", 1.0, 3.0, 1.8, 0.1,
                                     help="e.g. 1.8 means demand is 80% higher in peak season")
            offpeak_mult = st.slider("Off-peak demand multiplier", 0.3, 1.0, 0.7, 0.1)
        else:
            is_peak      = False
            peak_mult    = 1.0
            offpeak_mult = 1.0

        st.markdown("**💰 Raw Material Price Volatility**")
        apply_price      = st.checkbox("Apply price volatility adjustment", value=False)
        if apply_price:
            price_trend  = st.selectbox("Price trend next 3 months", ["Rising", "Stable", "Falling"])
            price_chg    = st.slider("Expected price change (%)", 0, 50, 10)
        else:
            price_trend  = "Stable"
            price_chg    = 0

    with col2:
        st.markdown("**⏰ Obsolescence Risk**")
        apply_obs        = st.checkbox("Apply obsolescence / shelf life cap", value=False)
        if apply_obs:
            shelf_life   = st.number_input("Product shelf life (days)", min_value=30, value=180, step=30)
            obs_risk     = st.selectbox("Obsolescence risk level", ["Low", "Medium", "High"])
        else:
            shelf_life   = 9999
            obs_risk     = "Low"

        st.markdown("**🏚️ Dead Stock**")
        apply_dead       = st.checkbox("Account for existing dead stock", value=False)
        if apply_dead:
            dead_units   = st.number_input("Dead stock units on hand", min_value=0, value=500, step=50)
            dead_months  = st.number_input("Months since last movement", min_value=1, value=6)
        else:
            dead_units   = 0
            dead_months  = 0

# Store in session state for use in other tabs
st.session_state["inputs"] = dict(
    product_name=product_name, annual_demand=annual_demand, demand_std=demand_std,
    unit_cost=unit_cost, order_cost=order_cost, holding_cost_pct=holding_cost_pct,
    working_days=working_days, moq=moq, warehouse_cap=warehouse_cap,
    budget_limit=budget_limit, lead_time_avg=lead_time_avg, lead_time_std=lead_time_std,
    num_suppliers=num_suppliers, reliability_pct=reliability_pct, backup_supplier=backup_supplier,
    apply_seasonal=apply_seasonal, is_peak=(is_peak=="Peak Season"),
    peak_mult=peak_mult, offpeak_mult=offpeak_mult,
    apply_price=apply_price, price_trend=price_trend, price_chg=price_chg,
    apply_obs=apply_obs, shelf_life=shelf_life, obs_risk=obs_risk,
    apply_dead=apply_dead, dead_units=dead_units,
)

# ════════════════════════════════════════
# TAB 2 — SCENARIOS
# ════════════════════════════════════════
with tab2:
    st.markdown('<p class="section-header">Define Your Scenarios</p>', unsafe_allow_html=True)
    st.markdown("Each scenario uses a different **service level** — the probability of never running out of stock. Higher service level = more safety stock = higher cost but lower risk.")

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.markdown("#### 🛡️ Scenario A — Conservative")
        use_a  = st.checkbox("Include Scenario A", value=True)
        sl_a   = st.selectbox("Service Level A", [90, 95, 97, 99], index=3, key="sla")
        name_a = st.text_input("Name", value="Conservative", key="na")

    with col2:
        st.markdown("#### ⚖️ Scenario B — Balanced")
        use_b  = st.checkbox("Include Scenario B", value=True)
        sl_b   = st.selectbox("Service Level B", [90, 95, 97, 99], index=1, key="slb")
        name_b = st.text_input("Name", value="Balanced", key="nb")

    with col3:
        st.markdown("#### 🏃 Scenario C — Lean")
        use_c  = st.checkbox("Include Scenario C", value=True)
        sl_c   = st.selectbox("Service Level C", [90, 95, 97, 99], index=0, key="slc")
        name_c = st.text_input("Name", value="Lean", key="nc")

    with col4:
        st.markdown("#### ✏️ Scenario D — Custom")
        use_d  = st.checkbox("Include Scenario D", value=False)
        sl_d   = st.selectbox("Service Level D", [90, 95, 97, 99], index=1, key="sld")
        name_d = st.text_input("Name", value="Custom", key="nd")

    st.session_state["scenarios"] = [
        (use_a, sl_a, name_a), (use_b, sl_b, name_b),
        (use_c, sl_c, name_c), (use_d, sl_d, name_d)
    ]

# ════════════════════════════════════════
# TAB 3 — RESULTS & AI
# ════════════════════════════════════════
with tab3:
    if st.button("🚀 Run Simulation & Get AI Recommendation"):
        inp = st.session_state.get("inputs", {})
        scenarios = st.session_state.get("scenarios", [])

        if not inp:
            st.error("Please fill in your product details in Tab 1 first.")
        else:
            active = [(sl, name) for (use, sl, name) in scenarios if use]
            if len(active) < 2:
                st.warning("Please enable at least 2 scenarios in Tab 2 to compare.")
            else:
                results      = []
                scenario_names = []

                with st.spinner("Calculating scenarios..."):
                    for sl, name in active:
                        r = run_scenario(
                            service_level    = sl,
                            annual_demand    = inp["annual_demand"],
                            order_cost       = inp["order_cost"],
                            unit_cost        = inp["unit_cost"],
                            holding_cost_pct = inp["holding_cost_pct"],
                            lead_time_avg    = inp["lead_time_avg"],
                            lead_time_std    = inp["lead_time_std"],
                            demand_std       = inp["demand_std"],
                            working_days     = inp["working_days"],
                            num_suppliers    = inp["num_suppliers"],
                            reliability_pct  = inp["reliability_pct"],
                            price_trend      = inp["price_trend"],
                            price_change_pct = inp["price_chg"],
                            shelf_life_days  = inp["shelf_life"],
                            apply_obsolescence = inp["apply_obs"],
                            dead_stock_units = inp["dead_units"],
                            apply_dead_stock = inp["apply_dead"],
                            is_peak_season   = inp["is_peak"],
                            peak_multiplier  = inp["peak_mult"],
                            offpeak_multiplier = inp["offpeak_mult"],
                            moq              = inp["moq"],
                            warehouse_cap    = inp["warehouse_cap"] if inp["warehouse_cap"] > 0 else 999999,
                            budget_limit     = inp["budget_limit"] if inp["budget_limit"] > 0 else 999999999,
                        )
                        results.append(r)
                        scenario_names.append(name)

                # ── Results metrics
                st.markdown('<p class="section-header">📊 Scenario Results</p>', unsafe_allow_html=True)
                cols = st.columns(len(results))
                for i, (col, name, r) in enumerate(zip(cols, scenario_names, results)):
                    with col:
                        st.markdown(f"**{name} (SL: {active[i][0]}%)**")
                        st.markdown(f"""
<div class="metric-card">
  <div class="metric-label">EOQ</div>
  <div class="metric-value">{r['eoq']:,}</div>
  <div class="metric-unit">units per order</div>
</div>
<div class="metric-card">
  <div class="metric-label">Safety Stock</div>
  <div class="metric-value">{r['safety_stock']:,}</div>
  <div class="metric-unit">units</div>
</div>
<div class="metric-card">
  <div class="metric-label">Reorder Point</div>
  <div class="metric-value">{r['reorder_point']:,}</div>
  <div class="metric-unit">units</div>
</div>
<div class="metric-card">
  <div class="metric-label">Total Annual Cost</div>
  <div class="metric-value">${r['total_cost']:,.0f}</div>
  <div class="metric-unit">USD/year</div>
</div>
<div class="metric-card">
  <div class="metric-label">Orders Per Year</div>
  <div class="metric-value">{r['orders_per_year']}</div>
  <div class="metric-unit">orders</div>
</div>
""", unsafe_allow_html=True)
                        if not r["budget_ok"]:
                            st.markdown('<div class="warning-box">⚠️ Exceeds budget limit</div>', unsafe_allow_html=True)
                        if r["obs_warning"]:
                            st.markdown('<div class="warning-box">⚠️ EOQ near shelf life cap</div>', unsafe_allow_html=True)

                # ── Charts
                st.markdown('<p class="section-header">📈 Visual Comparison</p>', unsafe_allow_html=True)
                c1, c2 = st.columns(2)
                with c1:
                    st.plotly_chart(plot_cost_comparison(results, scenario_names), use_container_width=True)
                with c2:
                    st.plotly_chart(plot_safety_stock_rop(results, scenario_names), use_container_width=True)
                c3, c4 = st.columns(2)
                with c3:
                    st.plotly_chart(plot_total_cost_risk(results, scenario_names), use_container_width=True)
                with c4:
                    st.plotly_chart(plot_eoq_comparison(results, scenario_names), use_container_width=True)

                # ── AI Recommendation
                st.markdown('<p class="section-header">🤖 AI Recommendation</p>', unsafe_allow_html=True)

                if not api_key:
                    st.markdown('<div class="info-box">ℹ️ Enter your Gemini API key in the sidebar to get AI recommendations.</div>', unsafe_allow_html=True)
                else:
                    with st.spinner("AI is analysing your scenarios..."):
                        scenarios_text = ""
                        for name, r in zip(scenario_names, results):
                            scenarios_text += f"""
Scenario: {name}
  - EOQ: {r['eoq']} units
  - Safety Stock: {r['safety_stock']} units
  - Reorder Point: {r['reorder_point']} units
  - Total Annual Cost: ${r['total_cost']:,.2f}
  - Orders Per Year: {r['orders_per_year']}
  - Holding Cost: ${r['holding_cost']:,.2f}
  - Order Cost: ${r['order_cost_ann']:,.2f}
  - Stockout Risk: {r['stockout_risk']}%
  - Budget OK: {r['budget_ok']}
"""
                        constraints_text = f"""
- Suppliers: {inp['num_suppliers']} (Reliability: {inp['reliability_pct']}%)
- Backup supplier: {inp['backup_supplier']}
- Lead time: {inp['lead_time_avg']} days avg (±{inp['lead_time_std']} days)
- Seasonal adjustment: {'Yes — ' + ('Peak' if inp['is_peak'] else 'Off-peak') if inp['apply_seasonal'] else 'No'}
- Price trend: {inp['price_trend']} ({inp['price_chg']}% change expected)
- Obsolescence risk: {inp['obs_risk'] if inp['apply_obs'] else 'Not applied'}
- Dead stock: {inp['dead_units']} units ({inp['dead_months']} months idle) if inp['apply_dead'] else 'Not applied'
- Warehouse capacity: {inp['warehouse_cap'] if inp['warehouse_cap'] > 0 else 'No limit'}
- Budget per order: ${inp['budget_limit'] if inp['budget_limit'] > 0 else 'No limit'}
"""
                        ai_output = get_ai_recommendation(
                            api_key, inp["product_name"], scenarios_text, constraints_text)
                        st.markdown(ai_output)

                # ── Download CSV
                st.markdown('<p class="section-header">💾 Export Results</p>', unsafe_allow_html=True)
                df_export = pd.DataFrame(results, index=scenario_names)
                csv = df_export.to_csv().encode("utf-8")
                st.download_button(
                    label="⬇️ Download Results as CSV",
                    data=csv,
                    file_name=f"{inp['product_name']}_inventory_optimization.csv",
                    mime="text/csv"
                )
