
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import math
import io
from groq import Groq
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

GROQ_API_KEY = st.secrets["GROQ_API_KEY"]

st.set_page_config(
    page_title="StockSense — Inventory Optimizer",
    page_icon="📦",
    layout="wide"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Plus Jakarta Sans', sans-serif !important;
}

/* Hide default streamlit elements */
#MainMenu { visibility: hidden; }
footer { visibility: hidden; }
header { visibility: hidden; }
.block-container { padding: 0 !important; max-width: 100% !important; }
.stApp { background: #111827 !important; }

/* Top header bar */
.top-header {
    background: #111827;
    border-bottom: 1px solid #374151;
    padding: 0 32px;
    height: 56px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    position: sticky;
    top: 0;
    z-index: 999;
}
.logo-wrap { display: flex; align-items: center; gap: 10px; }
.logo-icon {
    width: 32px; height: 32px;
    background: linear-gradient(135deg, #4f46e5, #7c3aed);
    border-radius: 8px;
    display: flex; align-items: center; justify-content: center;
    font-size: 16px;
}
.logo-text { font-size: 16px; font-weight: 700; color: #f9fafb; letter-spacing: -0.02em; }
.logo-badge {
    font-size: 9px; background: #312e8122; color: #818cf8;
    padding: 2px 8px; border-radius: 20px; font-weight: 600;
    border: 1px solid #312e8144;
}
.status-wrap { display: flex; align-items: center; gap: 6px; font-size: 11px; color: #6b7280; }
.status-dot { width: 6px; height: 6px; border-radius: 50%; background: #22c55e; box-shadow: 0 0 6px #22c55e88; display: inline-block; }

/* Hero section */
.hero-section {
    background: linear-gradient(135deg, #1a1040 0%, #111827 60%);
    border-bottom: 1px solid #374151;
    padding: 32px 32px 28px;
    position: relative;
    overflow: hidden;
}
.hero-section::before {
    content: '';
    position: absolute; top: -80px; right: -60px;
    width: 300px; height: 300px;
    background: radial-gradient(circle, #4f46e522 0%, transparent 70%);
    border-radius: 50%;
}
.hero-tag {
    font-size: 10px; font-weight: 600; letter-spacing: 0.12em;
    text-transform: uppercase; color: #818cf8; margin-bottom: 10px;
    display: flex; align-items: center; gap: 8px;
}
.hero-tag::before { content: ''; width: 20px; height: 1px; background: #6366f1; }
.hero-title {
    font-size: 28px; font-weight: 700; color: #f9fafb;
    letter-spacing: -0.03em; line-height: 1.2; margin-bottom: 8px;
}
.hero-title .accent {
    background: linear-gradient(90deg, #818cf8, #a78bfa);
    -webkit-background-clip: text; -webkit-text-fill-color: transparent;
}
.hero-sub { font-size: 13px; color: #6b7280; line-height: 1.7; max-width: 560px; margin-bottom: 18px; }
.industry-pills { display: flex; flex-wrap: wrap; gap: 6px; }
.industry-pill {
    font-size: 10px; padding: 4px 11px; border-radius: 20px;
    border: 1px solid #374151; color: #6b7280; background: #1f2937;
}

/* Main content area */
.main-content { padding: 28px 32px; }

/* Section header */
.section-hdr {
    font-size: 11px; font-weight: 600; letter-spacing: 0.08em;
    text-transform: uppercase; color: #6b7280;
    margin-bottom: 14px; display: flex; align-items: center; gap: 8px;
}
.section-hdr::after { content: ''; flex: 1; height: 1px; background: #1f2937; }

/* Cards */
.metric-card {
    background: #1f2937;
    border: 1px solid #374151;
    border-radius: 12px;
    padding: 16px 18px;
    margin-bottom: 8px;
}
.metric-label { font-size: 10px; color: #6b7280; text-transform: uppercase; letter-spacing: 0.06em; margin-bottom: 4px; }
.metric-value { font-size: 22px; font-weight: 700; color: #f9fafb; }
.metric-unit { font-size: 11px; color: #6b7280; margin-top: 2px; }
.metric-value.indigo { color: #818cf8; }
.metric-value.red { color: #f87171; }
.metric-value.amber { color: #fbbf24; }
.metric-value.green { color: #34d399; }

/* Step cards */
.step-card {
    background: #1f2937;
    border: 1px solid #374151;
    border-radius: 14px;
    padding: 22px;
    height: 100%;
    transition: border-color 0.2s;
}
.step-card:hover { border-color: #4f46e5; }
.step-num-badge {
    width: 30px; height: 30px; border-radius: 8px;
    background: linear-gradient(135deg, #4f46e5, #7c3aed);
    display: inline-flex; align-items: center; justify-content: center;
    font-size: 13px; font-weight: 700; color: white; margin-bottom: 12px;
}
.step-title { font-size: 14px; font-weight: 600; color: #f9fafb; margin-bottom: 6px; }
.step-desc { font-size: 12px; color: #6b7280; line-height: 1.65; margin-bottom: 14px; }
.hint-bar {
    display: flex; align-items: center; gap: 7px;
    background: #111827; padding: 9px 12px; border-radius: 8px;
    border: 1px solid #374151; font-size: 11px; margin-bottom: 14px;
}
.hint-required { color: #818cf8; font-weight: 600; }
.hint-optional { color: #6b7280; }

/* Download button */
.dl-button {
    display: inline-flex; align-items: center; justify-content: center; gap: 7px;
    width: 100%; padding: 11px; border-radius: 9px;
    background: linear-gradient(135deg, #4f46e5, #7c3aed);
    color: white; font-size: 12px; font-weight: 600;
    border: none; cursor: pointer; transition: opacity 0.15s;
}
.dl-button:hover { opacity: 0.9; }

/* Upload zone */
.upload-zone {
    border: 2px dashed #374151;
    border-radius: 12px; padding: 30px 20px;
    text-align: center; background: #111827; cursor: pointer;
    transition: border-color 0.2s, background 0.2s;
    margin-bottom: 10px;
}
.upload-zone:hover { border-color: #4f46e5; background: #1a1040; }
.upload-zone-icon {
    width: 46px; height: 46px; border-radius: 12px;
    background: #374151; display: flex; align-items: center;
    justify-content: center; margin: 0 auto 12px; font-size: 20px;
}
.upload-zone-title { font-size: 14px; font-weight: 600; color: #f9fafb; margin-bottom: 4px; }
.upload-zone-sub { font-size: 11px; color: #6b7280; }

/* Info cards bottom */
.info-cards { display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; margin-top: 20px; }
.info-card {
    background: #1f2937; border: 1px solid #374151;
    border-radius: 10px; padding: 14px;
    display: flex; gap: 10px; align-items: flex-start;
}
.info-card-icon { font-size: 16px; flex-shrink: 0; margin-top: 1px; }
.info-card-title { font-size: 11px; font-weight: 600; color: #9ca3af; margin-bottom: 3px; }
.info-card-text { font-size: 10px; color: #6b7280; line-height: 1.55; }

/* Scenario winner */
.winner-banner {
    background: linear-gradient(135deg, #312e81, #4c1d95);
    border: 1px solid #4f46e5;
    border-radius: 10px; padding: 12px 16px;
    display: flex; align-items: center; justify-content: space-between;
    margin-bottom: 16px;
}
.winner-label { font-size: 11px; color: #a5b4fc; font-weight: 500; }
.winner-name { font-size: 15px; font-weight: 700; color: #e0e7ff; }
.winner-cost { font-size: 13px; color: #818cf8; font-weight: 600; }

/* Risk badges */
.risk-high { background: #450a0a; border: 1px solid #7f1d1d; color: #fca5a5; border-radius: 8px; padding: 8px 12px; font-size: 12px; margin-bottom: 10px; }
.risk-med  { background: #451a03; border: 1px solid #78350f; color: #fcd34d; border-radius: 8px; padding: 8px 12px; font-size: 12px; margin-bottom: 10px; }
.risk-low  { background: #052e16; border: 1px solid #14532d; color: #6ee7b7; border-radius: 8px; padding: 8px 12px; font-size: 12px; margin-bottom: 10px; }

/* Privacy notice */
.privacy-notice {
    background: #1f2937; border: 1px solid #374151;
    border-radius: 8px; padding: 10px 14px;
    font-size: 11px; color: #6b7280; line-height: 1.6;
    margin-bottom: 16px;
}
.privacy-notice strong { color: #9ca3af; }

/* Scenario cards */
.scenario-card {
    background: #1f2937; border: 1px solid #374151;
    border-radius: 12px; padding: 16px;
    margin-bottom: 8px;
}
.scenario-card.winner-card { border-color: #4f46e5; background: #1e1b4b; }
.scenario-name { font-size: 12px; font-weight: 600; color: #e5e7eb; margin-bottom: 2px; }
.scenario-sl { font-size: 10px; color: #6b7280; margin-bottom: 10px; }

/* Streamlit overrides */
div[data-testid="stButton"] button {
    background: linear-gradient(135deg, #4f46e5, #7c3aed) !important;
    color: white !important; border: none !important;
    border-radius: 9px !important; padding: 0.6rem 1.5rem !important;
    font-family: 'Plus Jakarta Sans', sans-serif !important;
    font-weight: 600 !important; font-size: 13px !important;
    width: 100% !important;
}
div[data-testid="stSelectbox"] { color: #f9fafb !important; }
div[data-testid="stTabs"] [data-baseweb="tab-list"] {
    background: #111827 !important;
    border-bottom: 1px solid #374151 !important;
    padding: 0 32px !important;
    gap: 0 !important;
}
div[data-testid="stTabs"] [data-baseweb="tab"] {
    background: transparent !important;
    color: #6b7280 !important;
    font-family: 'Plus Jakarta Sans', sans-serif !important;
    font-size: 12px !important; font-weight: 500 !important;
    padding: 12px 20px !important;
    border-bottom: 2px solid transparent !important;
}
div[data-testid="stTabs"] [aria-selected="true"] {
    color: #818cf8 !important;
    border-bottom: 2px solid #6366f1 !important;
    background: transparent !important;
}
div[data-testid="stTabs"] [data-baseweb="tab-panel"] {
    background: #111827 !important;
    padding: 0 !important;
}
div[data-testid="stFileUploader"] {
    background: #111827 !important;
}
div[data-testid="stFileUploader"] > div {
    background: #1f2937 !important;
    border: 2px dashed #374151 !important;
    border-radius: 12px !important;
}
div[data-testid="stDataFrame"] { background: #1f2937 !important; }
div[data-testid="stChatMessage"] {
    background: #1f2937 !important;
    border: 1px solid #374151 !important;
    border-radius: 10px !important;
    margin-bottom: 8px !important;
}
div[data-testid="stChatInputContainer"] {
    background: #1f2937 !important;
    border: 1px solid #4f46e5 !important;
    border-radius: 10px !important;
}
.stMarkdown { color: #d1d5db !important; }
h1,h2,h3 { color: #f9fafb !important; }
p { color: #9ca3af !important; }
label { color: #9ca3af !important; }
.stSelectbox label { color: #9ca3af !important; }
div[data-testid="stDownloadButton"] button {
    background: #1f2937 !important;
    color: #818cf8 !important;
    border: 1px solid #4f46e5 !important;
    border-radius: 9px !important;
    font-weight: 600 !important;
    width: 100% !important;
}
</style>
""", unsafe_allow_html=True)

# ════════════════════════════════════════════
# CORE MATH
# ════════════════════════════════════════════
Z_SCORES = {90: 1.28, 95: 1.645, 97: 1.88, 99: 2.326}

def calc_eoq(annual_demand, order_cost, holding_cost_per_unit):
    if holding_cost_per_unit <= 0 or annual_demand <= 0: return 0
    return math.sqrt((2 * annual_demand * order_cost) / holding_cost_per_unit)

def calc_safety_stock(z, demand_std_daily, lead_time_avg, lead_time_std, daily_demand):
    part1 = (z**2) * lead_time_avg * (demand_std_daily**2)
    part2 = (z**2) * (daily_demand**2) * (lead_time_std**2)
    return math.sqrt(part1 + part2)

def calc_supplier_risk_factor(num_suppliers, reliability_pct):
    if num_suppliers == 1:   return 1 + ((100 - reliability_pct) / 100) * 1.5
    elif num_suppliers == 2: return 1 + ((100 - reliability_pct) / 100) * 0.8
    else:                    return 1 + ((100 - reliability_pct) / 100) * 0.4

def run_sku(row):
    results = {}
    monthly_demand  = float(row["Monthly_Demand"])
    unit_cost       = float(row["Unit_Cost_USD"])
    order_cost      = float(row["Order_Cost_USD"])
    lead_time_avg   = float(row["Lead_Time_Days"])
    num_suppliers   = int(row["Num_Suppliers"])
    reliability     = float(row["Supplier_Reliability_Pct"])
    service_level   = float(row["Service_Level_Pct"])

    working_days_month  = float(row.get("Working_Days_Month", 22)) if pd.notna(row.get("Working_Days_Month", np.nan)) else 22
    monthly_hold_pct    = float(row.get("Monthly_Holding_Cost_Pct", 2.0)) if pd.notna(row.get("Monthly_Holding_Cost_Pct", np.nan)) else 2.0
    daily_demand        = monthly_demand / working_days_month
    demand_std          = float(row.get("Demand_Std_Dev", daily_demand * 0.2)) if pd.notna(row.get("Demand_Std_Dev", np.nan)) else daily_demand * 0.2
    lead_time_std       = float(row.get("Lead_Time_Std_Dev", lead_time_avg * 0.2)) if pd.notna(row.get("Lead_Time_Std_Dev", np.nan)) else lead_time_avg * 0.2
    moq                 = float(row.get("MOQ", 0)) if pd.notna(row.get("MOQ", np.nan)) else 0
    shelf_life          = float(row.get("Shelf_Life_Days", 9999)) if pd.notna(row.get("Shelf_Life_Days", np.nan)) else 9999
    dead_stock          = float(row.get("Dead_Stock_Units", 0)) if pd.notna(row.get("Dead_Stock_Units", np.nan)) else 0
    price_trend         = str(row.get("Price_Trend", "Stable")) if pd.notna(row.get("Price_Trend", np.nan)) else "Stable"
    peak_mult           = float(row.get("Peak_Season_Multiplier", 1.0)) if pd.notna(row.get("Peak_Season_Multiplier", np.nan)) else 1.0

    annual_demand         = monthly_demand * 12 * peak_mult
    annual_hold_pct       = monthly_hold_pct * 12
    holding_cost_per_unit = unit_cost * (annual_hold_pct / 100)
    risk_factor           = calc_supplier_risk_factor(num_suppliers, reliability)
    price_adj             = 1.075 if price_trend == "Rising" else (0.93 if price_trend == "Falling" else 1.0)

    for sl_name, sl_val in [("Conservative", 99), ("Balanced", 95), ("Lean", 90)]:
        z   = Z_SCORES[sl_val]
        eoq = calc_eoq(annual_demand, order_cost, holding_cost_per_unit) * price_adj
        eoq = max(eoq, moq)
        if shelf_life < 9999:
            eoq = min(eoq, shelf_life * daily_demand * 0.8)
        ss  = calc_safety_stock(z, demand_std, lead_time_avg, lead_time_std, daily_demand) * risk_factor
        rop = (daily_demand * lead_time_avg) + ss
        hold_cost      = ((eoq / 2) + ss) * holding_cost_per_unit
        order_cost_ann = (annual_demand / eoq) * order_cost if eoq > 0 else 0
        wc_cost        = dead_stock * unit_cost * (annual_hold_pct / 100)
        total_cost     = hold_cost + order_cost_ann + wc_cost
        results[sl_name] = {
            "eoq": round(eoq), "safety_stock": round(ss),
            "reorder_point": round(rop), "holding_cost": round(hold_cost, 2),
            "order_cost_ann": round(order_cost_ann, 2), "wc_cost": round(wc_cost, 2),
            "total_cost": round(total_cost, 2), "orders_per_year": round(annual_demand / eoq if eoq > 0 else 0, 1),
            "stockout_risk": 100 - sl_val, "service_level": sl_val,
        }

    best = "Balanced"
    if service_level >= 97: best = "Conservative"
    elif service_level < 93: best = "Lean"
    if num_suppliers == 1 and reliability < 80: best = "Conservative"
    if dead_stock > monthly_demand * 2: best = "Lean"

    results["recommended"] = best
    results["daily_demand"] = round(daily_demand, 1)
    results["annual_demand"] = round(annual_demand)
    results["risk_score"] = round(
        ((100 - reliability) * 0.4) + ((1 if num_suppliers == 1 else 0) * 30) +
        ((dead_stock / (monthly_demand + 1)) * 10), 1)
    return results

# ════════════════════════════════════════════
# CHARTS — Theme D colors
# ════════════════════════════════════════════
CHART_COLORS  = {"Conservative": "#4f46e5", "Balanced": "#7c3aed", "Lean": "#a78bfa"}
CHART_BG      = "#1f2937"
CHART_PAPER   = "#1f2937"
CHART_GRID    = "#374151"
CHART_TEXT    = "#9ca3af"
CHART_TITLE   = "#f9fafb"

def base_layout(title):
    return dict(
        title=dict(text=title, font=dict(color=CHART_TITLE, size=13, family="Plus Jakarta Sans"), x=0),
        plot_bgcolor=CHART_BG, paper_bgcolor=CHART_PAPER,
        font=dict(color=CHART_TEXT, family="Plus Jakarta Sans", size=11),
        xaxis=dict(gridcolor=CHART_GRID, zerolinecolor=CHART_GRID, showgrid=False),
        yaxis=dict(gridcolor=CHART_GRID, zerolinecolor=CHART_GRID, gridwidth=0.5),
        margin=dict(l=12, r=12, t=44, b=12),
        height=300,
        legend=dict(orientation="h", y=1.12, font=dict(size=10), bgcolor="rgba(0,0,0,0)"),
        hoverlabel=dict(bgcolor="#374151", font_color="#f9fafb", font_family="Plus Jakarta Sans")
    )

def chart_costs(r, name):
    scenarios = ["Conservative", "Balanced", "Lean"]
    fig = go.Figure()
    categories = ["Holding Cost", "Order Cost", "Dead Stock Cost"]
    for s in scenarios:
        fig.add_trace(go.Bar(
            name=s, x=categories,
            y=[r[s]["holding_cost"], r[s]["order_cost_ann"], r[s]["wc_cost"]],
            marker=dict(color=CHART_COLORS[s], line=dict(width=0)),
            text=[f"${v:,.0f}" for v in [r[s]["holding_cost"], r[s]["order_cost_ann"], r[s]["wc_cost"]]],
            textposition="auto", textfont=dict(size=10, color="white")
        ))
    fig.update_layout(**base_layout(f"Cost Breakdown"), barmode="group")
    return fig

def chart_ss_rop(r, name):
    scenarios = ["Conservative", "Balanced", "Lean"]
    fig = make_subplots(rows=1, cols=2,
                        subplot_titles=["Safety Stock (units)", "Reorder Point (units)"],
                        horizontal_spacing=0.12)
    for s in scenarios:
        fig.add_trace(go.Bar(
            name=s, x=[s], y=[r[s]["safety_stock"]],
            marker=dict(color=CHART_COLORS[s], line=dict(width=0)),
            text=[f"{r[s]['safety_stock']:,}"], textposition="auto",
            textfont=dict(size=10, color="white"), showlegend=True
        ), row=1, col=1)
        fig.add_trace(go.Bar(
            name=s, x=[s], y=[r[s]["reorder_point"]],
            marker=dict(color=CHART_COLORS[s], line=dict(width=0)),
            text=[f"{r[s]['reorder_point']:,}"], textposition="auto",
            textfont=dict(size=10, color="white"), showlegend=False
        ), row=1, col=2)

    layout = base_layout("Safety Stock & Reorder Point")
    layout["showlegend"] = False
    for ann in layout.get("annotations", []):
        ann["font"] = dict(color=CHART_TEXT, size=11)
    fig.update_layout(**layout)
    fig.update_xaxes(showgrid=False, tickfont=dict(color=CHART_TEXT, size=10))
    fig.update_yaxes(gridcolor=CHART_GRID, gridwidth=0.5, tickfont=dict(color=CHART_TEXT, size=10))

    # Fix subplot title colors
    for annotation in fig.layout.annotations:
        annotation.font.color = CHART_TEXT
        annotation.font.size  = 11
    return fig

def chart_total_cost(r, name):
    scenarios = ["Conservative", "Balanced", "Lean"]
    fig = go.Figure(go.Bar(
        x=scenarios,
        y=[r[s]["total_cost"] for s in scenarios],
        marker=dict(
            color=[CHART_COLORS[s] for s in scenarios],
            line=dict(width=0)
        ),
        text=[f"${r[s]['total_cost']:,.0f}" for s in scenarios],
        textposition="auto",
        textfont=dict(size=11, color="white")
    ))
    fig.update_layout(**base_layout("Total Annual Cost (USD)"))
    return fig

def chart_risk_cost(r, name):
    scenarios = ["Conservative", "Balanced", "Lean"]
    fig = go.Figure()
    for s in scenarios:
        fig.add_trace(go.Scatter(
            x=[r[s]["stockout_risk"]], y=[r[s]["total_cost"]],
            mode="markers+text",
            marker=dict(size=22, color=CHART_COLORS[s],
                       line=dict(width=2, color="white")),
            text=[s], textposition="top center",
            textfont=dict(size=10, color=CHART_TEXT),
            name=s
        ))
    layout = base_layout("Risk vs Cost Tradeoff")
    layout["xaxis"]["title"] = dict(text="Stockout Risk (%)", font=dict(color=CHART_TEXT, size=10))
    layout["yaxis"]["title"] = dict(text="Total Annual Cost ($)", font=dict(color=CHART_TEXT, size=10))
    layout["xaxis"]["autorange"] = "reversed"
    layout["xaxis"]["showgrid"] = True
    layout["xaxis"]["gridcolor"] = CHART_GRID
    fig.update_layout(**layout)
    return fig

# ════════════════════════════════════════════
# AI FUNCTIONS
# ════════════════════════════════════════════
def get_ai_analysis(all_results, df):
    try:
        client = Groq(api_key=GROQ_API_KEY)
        top_risk = sorted(all_results.items(), key=lambda x: x[1]["risk_score"], reverse=True)[:10]
        risk_summary = ""
        for sku_id, r in top_risk:
            row  = df[df["SKU_ID"] == sku_id].iloc[0]
            rec  = r["recommended"]
            risk_summary += f"\nSKU: {sku_id} — {row['Product_Name']} | Risk: {r['risk_score']} | Recommended: {rec} | EOQ: {r[rec]['eoq']} | SS: {r[rec]['safety_stock']} | Cost: ${r[rec]['total_cost']:,.0f} | Suppliers: {row['Num_Suppliers']} | Reliability: {row['Supplier_Reliability_Pct']}%"

        total_cost    = sum(r[r["recommended"]]["total_cost"] for r in all_results.values())
        high_risk_cnt = sum(1 for r in all_results.values() if r["risk_score"] > 40)
        single_src    = sum(1 for _, row in df.iterrows() if int(row["Num_Suppliers"]) == 1)

        prompt = f"""You are a senior supply chain analyst.
Portfolio: {len(all_results)} SKUs | Total annual cost: ${total_cost:,.0f} | High risk: {high_risk_cnt} | Single source: {single_src}
Top risk SKUs: {risk_summary}

Provide analysis with these exact headers:
**PORTFOLIO HEALTH SUMMARY**
**TOP 5 CRITICAL SKUs**
**3 BIGGEST RISKS**
**5 IMMEDIATE ACTIONS**
**COST SAVINGS OPPORTUNITIES**

Be specific with numbers. Write for a supply chain director."""

        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3, max_tokens=1200
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"AI Error: {str(e)}"

def get_chat_response(client, question, history, all_results, df):
    try:
        context = f"Portfolio: {len(all_results)} SKUs\n"
        for sku_id, r in list(all_results.items())[:20]:
            row = df[df["SKU_ID"] == sku_id].iloc[0]
            rec = r["recommended"]
            context += f"- {sku_id} ({row['Product_Name']}): EOQ={r[rec]['eoq']}, SS={r[rec]['safety_stock']}, ROP={r[rec]['reorder_point']}, Cost=${r[rec]['total_cost']:,.0f}, Risk={r['risk_score']}\n"

        messages = [{"role": "system", "content": f"You are a senior inventory analyst. Answer using this data:\n{context}"}]
        for msg in history:
            messages.append(msg)
        messages.append({"role": "user", "content": question})

        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=messages, temperature=0.3, max_tokens=600
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"Error: {str(e)}"

# ════════════════════════════════════════════
# EXCEL FUNCTIONS
# ════════════════════════════════════════════
def generate_template():
    wb  = openpyxl.Workbook()
    ws  = wb.active
    ws.title = "Inventory Data"

    req_fill = PatternFill("solid", fgColor="1e1b4b")
    opt_fill = PatternFill("solid", fgColor="1c1917")
    req_font = Font(bold=True, color="818cf8", size=10)
    opt_font = Font(bold=True, color="d97706", size=10)
    center   = Alignment(horizontal="center", vertical="center")

    required_cols = ["SKU_ID","Product_Name","Monthly_Demand","Unit_Cost_USD","Order_Cost_USD","Lead_Time_Days","Num_Suppliers","Supplier_Reliability_Pct","Service_Level_Pct"]
    optional_cols = ["Working_Days_Month","Monthly_Holding_Cost_Pct","Demand_Std_Dev","Lead_Time_Std_Dev","MOQ","Shelf_Life_Days","Dead_Stock_Units","Price_Trend","Peak_Season_Multiplier"]
    all_cols = required_cols + optional_cols
    ws.append(all_cols)

    for i, col in enumerate(all_cols, 1):
        cell = ws.cell(row=1, column=i)
        cell.fill      = req_fill if col in required_cols else opt_fill
        cell.font      = req_font if col in required_cols else opt_font
        cell.alignment = center
        cell.value     = col

    samples = [
        ["SKU-001","Product A",4200,5.00,200,14,1,78,95,22,2.0,15,3,100,1095,200,"Stable",1.0],
        ["SKU-002","Product B",3100,8.50,180,21,2,88,95,22,2.0,12,4,50,730,0,"Rising",1.3],
        ["SKU-003","Product C",890,45.00,350,30,1,72,99,22,2.5,5,5,20,180,50,"Stable",1.0],
        ["SKU-004","Product D",5600,3.20,150,18,3
