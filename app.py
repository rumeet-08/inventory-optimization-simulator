
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

# ── Your Groq API key — hardcoded, users never see it
GROQ_API_KEY = st.secrets["GROQ_API_KEY"]

# ════════════════════════════════════════════
# PAGE CONFIG
# ════════════════════════════════════════════
st.set_page_config(
    page_title="Pharma Inventory Optimizer",
    page_icon="💊",
    layout="wide"
)

st.markdown("""
<style>
  .main-title { font-size: 2rem; font-weight: 700; color: #1a1a2e; }
  .sub-title { font-size: 0.95rem; color: #666; margin-bottom: 1.5rem; }
  .metric-card {
    background: #f8f9ff;
    border: 1px solid #e0e4ff;
    border-radius: 10px;
    padding: 0.9rem 1.1rem;
    margin-bottom: 0.5rem;
  }
  .metric-label { font-size: 0.72rem; color: #888; text-transform: uppercase; letter-spacing: 0.05em; }
  .metric-value { font-size: 1.4rem; font-weight: 700; color: #1a1a2e; }
  .metric-unit { font-size: 0.75rem; color: #888; }
  .section-header {
    font-size: 1rem;
    font-weight: 600;
    color: #1a1a2e;
    border-left: 4px solid #667eea;
    padding-left: 0.75rem;
    margin: 1.2rem 0 0.8rem;
  }
  .risk-high { background: #fff0f0; border: 1px solid #ffcccc; border-radius: 8px; padding: 0.6rem 1rem; font-size: 0.85rem; color: #cc0000; margin-bottom: 0.4rem; }
  .risk-med  { background: #fffbf0; border: 1px solid #ffe0a0; border-radius: 8px; padding: 0.6rem 1rem; font-size: 0.85rem; color: #996600; margin-bottom: 0.4rem; }
  .risk-low  { background: #f0fff4; border: 1px solid #b0f0c0; border-radius: 8px; padding: 0.6rem 1rem; font-size: 0.85rem; color: #006620; margin-bottom: 0.4rem; }
  .stButton > button {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    color: white; border: none; border-radius: 8px;
    padding: 0.6rem 2rem; font-size: 1rem; font-weight: 600; width: 100%;
  }
  .upload-box {
    background: #f8f9ff; border: 2px dashed #667eea;
    border-radius: 12px; padding: 2rem; text-align: center;
    margin: 1rem 0;
  }
  .scenario-winner {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    border-radius: 10px; padding: 0.8rem 1.2rem;
    color: white; font-weight: 600; font-size: 0.9rem;
    margin-bottom: 0.5rem;
  }
</style>
""", unsafe_allow_html=True)

# ════════════════════════════════════════════
# CORE MATH FUNCTIONS
# ════════════════════════════════════════════
Z_SCORES = {90: 1.28, 95: 1.645, 97: 1.88, 99: 2.326}

def to_annual(monthly_val):
    return monthly_val * 12

def calc_eoq(annual_demand, order_cost, holding_cost_per_unit):
    if holding_cost_per_unit <= 0 or annual_demand <= 0:
        return 0
    return math.sqrt((2 * annual_demand * order_cost) / holding_cost_per_unit)

def calc_safety_stock(z, demand_std_daily, lead_time_avg, lead_time_std, daily_demand):
    part1 = (z ** 2) * lead_time_avg * (demand_std_daily ** 2)
    part2 = (z ** 2) * (daily_demand ** 2) * (lead_time_std ** 2)
    return math.sqrt(part1 + part2)

def calc_supplier_risk_factor(num_suppliers, reliability_pct):
    if num_suppliers == 1:
        return 1 + ((100 - reliability_pct) / 100) * 1.5
    elif num_suppliers == 2:
        return 1 + ((100 - reliability_pct) / 100) * 0.8
    else:
        return 1 + ((100 - reliability_pct) / 100) * 0.4

def run_sku(row):
    results = {}

    # ── Parse required fields
    monthly_demand   = float(row["Monthly_Demand"])
    unit_cost        = float(row["Unit_Cost_USD"])
    order_cost       = float(row["Order_Cost_USD"])
    lead_time_avg    = float(row["Lead_Time_Days"])
    num_suppliers    = int(row["Num_Suppliers"])
    reliability      = float(row["Supplier_Reliability_Pct"])
    service_level    = float(row["Service_Level_Pct"])

    # ── Parse optional fields with smart defaults
    working_days_month = float(row.get("Working_Days_Month", 22)) if pd.notna(row.get("Working_Days_Month", np.nan)) else 22
    monthly_hold_pct   = float(row.get("Monthly_Holding_Cost_Pct", 2.0)) if pd.notna(row.get("Monthly_Holding_Cost_Pct", np.nan)) else 2.0
    daily_demand       = monthly_demand / working_days_month
    demand_std         = float(row.get("Demand_Std_Dev", daily_demand * 0.2)) if pd.notna(row.get("Demand_Std_Dev", np.nan)) else daily_demand * 0.2
    lead_time_std      = float(row.get("Lead_Time_Std_Dev", lead_time_avg * 0.2)) if pd.notna(row.get("Lead_Time_Std_Dev", np.nan)) else lead_time_avg * 0.2
    moq                = float(row.get("MOQ", 0)) if pd.notna(row.get("MOQ", np.nan)) else 0
    shelf_life         = float(row.get("Shelf_Life_Days", 9999)) if pd.notna(row.get("Shelf_Life_Days", np.nan)) else 9999
    dead_stock         = float(row.get("Dead_Stock_Units", 0)) if pd.notna(row.get("Dead_Stock_Units", np.nan)) else 0
    price_trend        = str(row.get("Price_Trend", "Stable")) if pd.notna(row.get("Price_Trend", np.nan)) else "Stable"
    peak_mult          = float(row.get("Peak_Season_Multiplier", 1.0)) if pd.notna(row.get("Peak_Season_Multiplier", np.nan)) else 1.0

    # ── Conversions
    annual_demand        = to_annual(monthly_demand) * peak_mult
    annual_hold_pct      = monthly_hold_pct * 12
    holding_cost_per_unit = unit_cost * (annual_hold_pct / 100)
    risk_factor          = calc_supplier_risk_factor(num_suppliers, reliability)

    # ── Price trend EOQ adjustment
    price_adj = 1.0
    if price_trend == "Rising":   price_adj = 1.075
    elif price_trend == "Falling": price_adj = 0.93

    # ── Run 3 scenarios
    for sl_name, sl_val in [("Conservative", 99), ("Balanced", 95), ("Lean", 90)]:
        z   = Z_SCORES[sl_val]
        eoq = calc_eoq(annual_demand, order_cost, holding_cost_per_unit) * price_adj
        eoq = max(eoq, moq)

        # Shelf life cap
        if shelf_life < 9999:
            max_qty = shelf_life * daily_demand * 0.8
            eoq     = min(eoq, max_qty)

        ss  = calc_safety_stock(z, demand_std, lead_time_avg, lead_time_std, daily_demand)
        ss  = ss * risk_factor

        rop = (daily_demand * lead_time_avg) + ss

        hold_cost  = ((eoq / 2) + ss) * holding_cost_per_unit
        order_cost_ann = (annual_demand / eoq) * order_cost if eoq > 0 else 0
        wc_cost    = dead_stock * unit_cost * (annual_hold_pct / 100)
        total_cost = hold_cost + order_cost_ann + wc_cost
        orders_yr  = annual_demand / eoq if eoq > 0 else 0

        results[sl_name] = {
            "eoq":            round(eoq),
            "safety_stock":   round(ss),
            "reorder_point":  round(rop),
            "holding_cost":   round(hold_cost, 2),
            "order_cost_ann": round(order_cost_ann, 2),
            "wc_cost":        round(wc_cost, 2),
            "total_cost":     round(total_cost, 2),
            "orders_per_year": round(orders_yr, 1),
            "stockout_risk":  100 - sl_val,
            "service_level":  sl_val,
        }

    # ── Pick best scenario
    user_sl   = service_level
    best_name = "Balanced"
    if user_sl >= 97:   best_name = "Conservative"
    elif user_sl >= 93: best_name = "Balanced"
    else:               best_name = "Lean"

    # Override if single supplier unreliable
    if num_suppliers == 1 and reliability < 80:
        best_name = "Conservative"

    # Override if lots of dead stock
    if dead_stock > monthly_demand * 2:
        best_name = "Lean"

    results["recommended"] = best_name
    results["daily_demand"] = round(daily_demand, 1)
    results["annual_demand"] = round(annual_demand)
    results["risk_score"] = round(
        ((100 - reliability) * 0.4) +
        ((1 if num_suppliers == 1 else 0) * 30) +
        ((dead_stock / (monthly_demand + 1)) * 10), 1
    )

    return results

# ════════════════════════════════════════════
# CHART FUNCTIONS
# ════════════════════════════════════════════
COLORS = {"Conservative": "#667eea", "Balanced": "#f093fb", "Lean": "#4facfe"}

def chart_cost_breakdown(sku_results, sku_name):
    scenarios = ["Conservative", "Balanced", "Lean"]
    fig = go.Figure()
    for s in scenarios:
        r = sku_results[s]
        fig.add_trace(go.Bar(
            name=s,
            x=["Holding Cost", "Order Cost", "Dead Stock Cost"],
            y=[r["holding_cost"], r["order_cost_ann"], r["wc_cost"]],
            marker_color=COLORS[s],
            text=[f"${r['holding_cost']:,.0f}", f"${r['order_cost_ann']:,.0f}", f"${r['wc_cost']:,.0f}"],
            textposition="auto",
        ))
    fig.update_layout(
        barmode="group", title=f"Cost Breakdown — {sku_name}",
        yaxis_title="USD/year", plot_bgcolor="white",
        paper_bgcolor="white", height=360,
        legend=dict(orientation="h", y=1.1)
    )
    return fig

def chart_safety_rop(sku_results, sku_name):
    scenarios = ["Conservative", "Balanced", "Lean"]
    fig = make_subplots(rows=1, cols=2,
                        subplot_titles=("Safety Stock (units)", "Reorder Point (units)"))
    for i, s in enumerate(scenarios):
        r = sku_results[s]
        fig.add_trace(go.Bar(
            name=s, x=[s], y=[r["safety_stock"]],
            marker_color=COLORS[s], showlegend=False,
            text=[f"{r['safety_stock']:,}"], textposition="auto"
        ), row=1, col=1)
        fig.add_trace(go.Bar(
            name=s, x=[s], y=[r["reorder_point"]],
            marker_color=COLORS[s], showlegend=False,
            text=[f"{r['reorder_point']:,}"], textposition="auto"
        ), row=1, col=2)
    fig.update_layout(
        plot_bgcolor="white", paper_bgcolor="white",
        height=340, title=f"Safety Stock & Reorder Point — {sku_name}"
    )
    return fig

def chart_total_cost(sku_results, sku_name):
    scenarios = ["Conservative", "Balanced", "Lean"]
    fig = go.Figure(go.Bar(
        x=scenarios,
        y=[sku_results[s]["total_cost"] for s in scenarios],
        marker_color=[COLORS[s] for s in scenarios],
        text=[f"${sku_results[s]['total_cost']:,.0f}" for s in scenarios],
        textposition="auto"
    ))
    fig.update_layout(
        title=f"Total Annual Cost — {sku_name}",
        yaxis_title="USD/year",
        plot_bgcolor="white", paper_bgcolor="white", height=340
    )
    return fig

def chart_risk_vs_cost(sku_results, sku_name):
    scenarios = ["Conservative", "Balanced", "Lean"]
    fig = go.Figure()
    for s in scenarios:
        r = sku_results[s]
        fig.add_trace(go.Scatter(
            x=[r["stockout_risk"]], y=[r["total_cost"]],
            mode="markers+text",
            marker=dict(size=18, color=COLORS[s]),
            text=[s], textposition="top center", name=s
        ))
    fig.update_layout(
        title=f"Cost vs Stockout Risk — {sku_name}",
        xaxis_title="Stockout Risk (%)",
        yaxis_title="Total Annual Cost (USD)",
        xaxis=dict(autorange="reversed"),
        plot_bgcolor="white", paper_bgcolor="white", height=340
    )
    return fig

# ════════════════════════════════════════════
# AI FUNCTION — GROQ
# ════════════════════════════════════════════
def get_ai_analysis(all_results, df):
    try:
        client = Groq(api_key=GROQ_API_KEY)

        # Build a compact summary of top risk SKUs
        risk_summary = ""
        top_risk = sorted(all_results.items(),
                         key=lambda x: x[1]["risk_score"], reverse=True)[:10]

        for sku_id, r in top_risk:
            row  = df[df["SKU_ID"] == sku_id].iloc[0]
            name = row["Product_Name"]
            rec  = r["recommended"]
            risk_summary += f"""
SKU: {sku_id} — {name}
  Risk Score: {r['risk_score']}
  Recommended Scenario: {rec}
  EOQ: {r[rec]['eoq']} units | Safety Stock: {r[rec]['safety_stock']} units
  Reorder Point: {r[rec]['reorder_point']} units
  Total Annual Cost: ${r[rec]['total_cost']:,.0f}
  Suppliers: {row['Num_Suppliers']} | Reliability: {row['Supplier_Reliability_Pct']}%
"""

        total_skus    = len(all_results)
        total_cost    = sum(r[r["recommended"]]["total_cost"] for r in all_results.values())
        high_risk_cnt = sum(1 for r in all_results.values() if r["risk_score"] > 40)
        single_src    = sum(1 for _, row in df.iterrows() if int(row["Num_Suppliers"]) == 1)

        prompt = f"""
You are a senior pharmaceutical supply chain analyst.
You have just run an inventory optimization analysis across {total_skus} SKUs.

PORTFOLIO SUMMARY:
- Total SKUs analysed: {total_skus}
- Total annual inventory cost (recommended scenarios): ${total_cost:,.0f}
- High risk SKUs (risk score > 40): {high_risk_cnt}
- Single-source supplier SKUs: {single_src}

TOP 10 HIGHEST RISK SKUs:
{risk_summary}

Provide a concise but insightful portfolio-level analysis with these exact sections:

**PORTFOLIO HEALTH SUMMARY**
2-3 sentences on the overall state of this pharma company's inventory.

**TOP 5 CRITICAL SKUs NEEDING IMMEDIATE ATTENTION**
List the 5 highest risk SKUs with one specific action for each.

**3 BIGGEST SUPPLY CHAIN RISKS**
Based on the data — what are the top 3 systemic risks in this portfolio?

**5 IMMEDIATE ACTIONS**
Specific, numbered, actionable steps this company should take this month.

**COST OPTIMISATION OPPORTUNITIES**
Where can they reduce inventory costs without increasing risk?

Be specific with numbers. Write for a supply chain director who will act on this today.
"""

        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3,
            max_tokens=1500
        )
        return response.choices[0].message.content

    except Exception as e:
        return f"AI Error: {str(e)}"

# ════════════════════════════════════════════
# EXCEL REPORT GENERATOR
# ════════════════════════════════════════════
def generate_excel_report(all_results, df):
    wb = openpyxl.Workbook()

    # ── Styles
    header_fill   = PatternFill("solid", fgColor="667EEA")
    header_font   = Font(bold=True, color="FFFFFF", size=11)
    risk_high_fill = PatternFill("solid", fgColor="FFE0E0")
    risk_med_fill  = PatternFill("solid", fgColor="FFF8E0")
    risk_low_fill  = PatternFill("solid", fgColor="E8FFE8")
    center        = Alignment(horizontal="center", vertical="center")
    bold          = Font(bold=True)

    def style_header_row(ws, row_num, num_cols):
        for c in range(1, num_cols + 1):
            cell            = ws.cell(row=row_num, column=c)
            cell.fill       = header_fill
            cell.font       = header_font
            cell.alignment  = center

    def auto_width(ws):
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                try:
                    if cell.value:
                        max_len = max(max_len, len(str(cell.value)))
                except:
                    pass
            ws.column_dimensions[col_letter].width = min(max_len + 4, 30)

    # ════ SHEET 1 — Summary Dashboard ════
    ws1 = wb.active
    ws1.title = "Summary Dashboard"
    headers = ["SKU ID", "Product Name", "Risk Score", "Recommended",
               "EOQ (units)", "Safety Stock", "Reorder Point",
               "Total Cost/Year", "Orders/Year", "Stockout Risk %"]
    ws1.append(headers)
    style_header_row(ws1, 1, len(headers))

    for sku_id, r in all_results.items():
        row  = df[df["SKU_ID"] == sku_id].iloc[0]
        rec  = r["recommended"]
        data = [
            sku_id, row["Product_Name"], r["risk_score"], rec,
            r[rec]["eoq"], r[rec]["safety_stock"], r[rec]["reorder_point"],
            f"${r[rec]['total_cost']:,.0f}", r[rec]["orders_per_year"],
            f"{r[rec]['stockout_risk']}%"
        ]
        ws1.append(data)
        row_idx = ws1.max_row
        fill = risk_high_fill if r["risk_score"] > 40 else (
               risk_med_fill  if r["risk_score"] > 20 else risk_low_fill)
        for c in range(1, len(headers) + 1):
            ws1.cell(row=row_idx, column=c).fill = fill

    auto_width(ws1)

    # ════ SHEET 2 — High Risk SKUs ════
    ws2 = wb.create_sheet("High Risk SKUs")
    ws2.append(["SKU ID", "Product Name", "Risk Score", "Risk Reason",
                "Suppliers", "Reliability %", "Recommended Action"])
    style_header_row(ws2, 1, 7)

    high_risk = [(k, v) for k, v in all_results.items() if v["risk_score"] > 40]
    high_risk.sort(key=lambda x: x[1]["risk_score"], reverse=True)

    for sku_id, r in high_risk:
        row = df[df["SKU_ID"] == sku_id].iloc[0]
        reasons = []
        if int(row["Num_Suppliers"]) == 1:
            reasons.append("Single source supplier")
        if float(row["Supplier_Reliability_Pct"]) < 80:
            reasons.append("Low supplier reliability")
        if float(row.get("Dead_Stock_Units", 0) or 0) > float(row["Monthly_Demand"]):
            reasons.append("High dead stock")

        ws2.append([
            sku_id, row["Product_Name"], r["risk_score"],
            " | ".join(reasons) if reasons else "Multiple factors",
            row["Num_Suppliers"], row["Supplier_Reliability_Pct"],
            f"Increase safety stock to {r['Conservative']['safety_stock']} units"
        ])
        for c in range(1, 8):
            ws2.cell(row=ws2.max_row, column=c).fill = risk_high_fill

    auto_width(ws2)

    # ════ SHEET 3 — Full Cost Analysis ════
    ws3 = wb.create_sheet("Cost Analysis")
    ws3.append(["SKU ID", "Product Name", "Scenario",
                "Holding Cost", "Order Cost", "Dead Stock Cost", "Total Cost"])
    style_header_row(ws3, 1, 7)

    for sku_id, r in all_results.items():
        row = df[df["SKU_ID"] == sku_id].iloc[0]
        for scenario in ["Conservative", "Balanced", "Lean"]:
            s = r[scenario]
            ws3.append([
                sku_id, row["Product_Name"], scenario,
                f"${s['holding_cost']:,.0f}",
                f"${s['order_cost_ann']:,.0f}",
                f"${s['wc_cost']:,.0f}",
                f"${s['total_cost']:,.0f}"
            ])

    auto_width(ws3)

    # ════ SHEET 4 — Reorder Schedule ════
    ws4 = wb.create_sheet("Reorder Schedule")
    ws4.append(["SKU ID", "Product Name", "Daily Demand",
                "Reorder Point", "EOQ", "Orders Per Year",
                "Lead Time (days)", "Safety Stock"])
    style_header_row(ws4, 1, 8)

    for sku_id, r in all_results.items():
        row = df[df["SKU_ID"] == sku_id].iloc[0]
        rec = r["recommended"]
        ws4.append([
            sku_id, row["Product_Name"], r["daily_demand"],
            r[rec]["reorder_point"], r[rec]["eoq"],
            r[rec]["orders_per_year"], row["Lead_Time_Days"],
            r[rec]["safety_stock"]
        ])

    auto_width(ws4)

    # ════ SHEET 5 — Scenario Comparison ════
    ws5 = wb.create_sheet("Scenario Comparison")
    ws5.append(["SKU ID", "Product Name",
                "Conservative EOQ", "Conservative SS", "Conservative Cost",
                "Balanced EOQ",     "Balanced SS",     "Balanced Cost",
                "Lean EOQ",         "Lean SS",          "Lean Cost",
                "Recommended"])
    style_header_row(ws5, 1, 12)

    for sku_id, r in all_results.items():
        row = df[df["SKU_ID"] == sku_id].iloc[0]
        ws5.append([
            sku_id, row["Product_Name"],
            r["Conservative"]["eoq"], r["Conservative"]["safety_stock"], f"${r['Conservative']['total_cost']:,.0f}",
            r["Balanced"]["eoq"],     r["Balanced"]["safety_stock"],     f"${r['Balanced']['total_cost']:,.0f}",
            r["Lean"]["eoq"],         r["Lean"]["safety_stock"],          f"${r['Lean']['total_cost']:,.0f}",
            r["recommended"]
        ])

    auto_width(ws5)

    # ── Save to bytes
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ════════════════════════════════════════════
# EXCEL TEMPLATE GENERATOR
# ════════════════════════════════════════════
def generate_template():
    wb   = openpyxl.Workbook()
    ws   = wb.active
    ws.title = "Inventory Data"

    req_fill  = PatternFill("solid", fgColor="E8F4FD")
    opt_fill  = PatternFill("solid", fgColor="FFF8E1")
    req_font  = Font(bold=True, color="0D47A1")
    opt_font  = Font(bold=True, color="E65100")
    center    = Alignment(horizontal="center")

    required_cols = [
        "SKU_ID", "Product_Name", "Monthly_Demand",
        "Unit_Cost_USD", "Order_Cost_USD", "Lead_Time_Days",
        "Num_Suppliers", "Supplier_Reliability_Pct", "Service_Level_Pct"
    ]
    optional_cols = [
        "Working_Days_Month", "Monthly_Holding_Cost_Pct",
        "Demand_Std_Dev", "Lead_Time_Std_Dev", "MOQ",
        "Shelf_Life_Days", "Dead_Stock_Units", "Price_Trend",
        "Peak_Season_Multiplier"
    ]

    all_cols = required_cols + optional_cols
    ws.append(all_cols)

    for i, col in enumerate(all_cols, 1):
        cell = ws.cell(row=1, column=i)
        if col in required_cols:
            cell.fill      = req_fill
            cell.font      = req_font
        else:
            cell.fill      = opt_fill
            cell.font      = opt_font
        cell.alignment = center

    # ── Sample rows
    samples = [
        ["PAR-500", "Paracetamol 500mg", 4200, 5.00, 200, 14, 1, 78, 95,
         22, 2.0, 15, 3, 100, 1095, 200, "Stable", 1.0],
        ["AMO-250", "Amoxicillin 250mg", 3100, 8.50, 180, 21, 2, 88, 95,
         22, 2.0, "", "", 50, 730, 0, "Rising", 1.3],
        ["INS-100", "Insulin 100IU",      890, 45.00, 350, 30, 1, 72, 99,
         22, 2.5, 5, 5, 20, 180, 50, "Stable", 1.0],
        ["MET-500", "Metformin 500mg",   5600, 3.20, 150, 18, 3, 92, 90,
         22, 1.8, 20, 4, 200, 1825, 0, "Falling", 1.0],
        ["CIP-500", "Ciprofloxacin 500mg",2400, 12.00, 220, 25, 1, 65, 99,
         22, 2.0, 10, 6, 50, 730, 300, "Rising", 1.5],
    ]
    for s in samples:
        ws.append(s)

    # ── Instructions sheet
    ws2         = wb.create_sheet("Instructions")
    ws2["A1"]   = "HOW TO FILL THE TEMPLATE"
    ws2["A1"].font = Font(bold=True, size=14)
    instructions = [
        ("", ""),
        ("BLUE columns", "REQUIRED — must be filled for every product"),
        ("ORANGE columns", "OPTIONAL — leave blank if unknown, app will use smart defaults"),
        ("", ""),
        ("SKU_ID", "Unique code for your product e.g. PAR-500"),
        ("Product_Name", "Full product name"),
        ("Monthly_Demand", "How many units you sell per month"),
        ("Unit_Cost_USD", "Cost per unit in USD"),
        ("Order_Cost_USD", "Cost to place one order (admin + shipping)"),
        ("Lead_Time_Days", "Average days your supplier takes to deliver"),
        ("Num_Suppliers", "Number of suppliers for this product (1, 2, 3...)"),
        ("Supplier_Reliability_Pct", "% of orders delivered on time (0-100)"),
        ("Service_Level_Pct", "Target service level: 90, 95, 97, or 99"),
        ("", ""),
        ("Working_Days_Month", "Working days per month (default: 22)"),
        ("Monthly_Holding_Cost_Pct", "Monthly storage cost as % of unit cost (default: 2%)"),
        ("Demand_Std_Dev", "Daily demand standard deviation (default: 20% of daily demand)"),
        ("Lead_Time_Std_Dev", "Lead time standard deviation in days (default: 20% of lead time)"),
        ("MOQ", "Minimum order quantity from supplier (default: none)"),
        ("Shelf_Life_Days", "Product shelf life in days (default: no limit)"),
        ("Dead_Stock_Units", "Units currently sitting idle (default: 0)"),
        ("Price_Trend", "Rising / Stable / Falling (default: Stable)"),
        ("Peak_Season_Multiplier", "Demand multiplier in peak season e.g. 1.8 (default: 1.0)"),
        ("", ""),
        ("MAX SKUs", "Up to 500 products per upload"),
        ("FILE FORMAT", "Save as .xlsx before uploading"),
    ]
    for row in instructions:
        ws2.append(list(row))

    ws2.column_dimensions["A"].width = 28
    ws2.column_dimensions["B"].width = 55

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ════════════════════════════════════════════
# CHAT FUNCTION
# ════════════════════════════════════════════
def get_chat_response(client, question, conversation_history, all_results, df):
    try:
        # Build compact data context
        top_risk = sorted(all_results.items(),
                         key=lambda x: x[1]["risk_score"], reverse=True)[:5]
        context = "INVENTORY PORTFOLIO CONTEXT:\n"
        context += f"Total SKUs: {len(all_results)}\n"
        context += f"High risk SKUs: {sum(1 for r in all_results.values() if r['risk_score'] > 40)}\n\n"
        context += "TOP 5 HIGH RISK SKUs:\n"
        for sku_id, r in top_risk:
            row  = df[df["SKU_ID"] == sku_id].iloc[0]
            rec  = r["recommended"]
            context += f"- {sku_id} ({row['Product_Name']}): Risk={r['risk_score']}, EOQ={r[rec]['eoq']}, Safety Stock={r[rec]['safety_stock']}, Reorder Point={r[rec]['reorder_point']}, Cost=${r[rec]['total_cost']:,.0f}\n"

        context += "\nALL SKUs SUMMARY:\n"
        for sku_id, r in all_results.items():
            row = df[df["SKU_ID"] == sku_id].iloc[0]
            rec = r["recommended"]
            context += f"- {sku_id} ({row['Product_Name']}): Recommended={rec}, EOQ={r[rec]['eoq']}, SS={r[rec]['safety_stock']}, ROP={r[rec]['reorder_point']}, Cost=${r[rec]['total_cost']:,.0f}, Risk={r['risk_score']}\n"

        messages = [
            {
                "role": "system",
                "content": f"""You are a senior pharmaceutical supply chain analyst assistant.
You have access to the user's complete inventory data. Answer questions specifically using their data.
Be concise, specific, and actionable. Always reference actual SKU IDs and numbers from the data.

{context}"""
            }
        ]

        # Add conversation history
        for msg in conversation_history:
            messages.append(msg)

        # Add current question
        messages.append({"role": "user", "content": question})

        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=messages,
            temperature=0.3,
            max_tokens=800
        )
        return response.choices[0].message.content

    except Exception as e:
        return f"Error: {str(e)}"

# ════════════════════════════════════════════
# MAIN APP
# ════════════════════════════════════════════
st.markdown('<p class="main-title">💊 Pharma Inventory Optimizer</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">Upload your product data and get instant inventory optimization across all SKUs — EOQ, safety stock, reorder points, and AI recommendations.</p>', unsafe_allow_html=True)

tab1, tab2, tab3 = st.tabs(["📂 Upload", "📊 Results", "🤖 AI & Export"])

# ════════════════════════════════════════════
# TAB 1 — UPLOAD
# ════════════════════════════════════════════
with tab1:
    st.markdown('<p class="section-header">Step 1 — Download the Template</p>', unsafe_allow_html=True)
    st.markdown("Download the Excel template, fill in your product data, then upload it back. Only 9 columns are required — the rest are optional.")

    template_file = generate_template()
    st.download_button(
        label="⬇️ Download Excel Template",
        data=template_file,
        file_name="pharma_inventory_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.markdown("**Template has 2 sheets:**")
    st.markdown("- **Inventory Data** — fill your products here (blue = required, orange = optional)")
    st.markdown("- **Instructions** — explains every column")

    st.markdown('<p class="section-header">Step 2 — Upload Your Filled Template</p>', unsafe_allow_html=True)

    uploaded = st.file_uploader(
        "Upload your filled Excel file",
        type=["xlsx"],
        help="Max 500 SKUs. Make sure all required columns are filled."
    )

    if uploaded:
        try:
            df = pd.read_excel(uploaded, sheet_name="Inventory Data")

            required_cols = [
                "SKU_ID", "Product_Name", "Monthly_Demand",
                "Unit_Cost_USD", "Order_Cost_USD", "Lead_Time_Days",
                "Num_Suppliers", "Supplier_Reliability_Pct", "Service_Level_Pct"
            ]

            missing = [c for c in required_cols if c not in df.columns]

            if missing:
                st.error(f"❌ Missing required columns: {', '.join(missing)}")
            else:
                df = df.dropna(subset=required_cols)

                if len(df) == 0:
                    st.error("❌ No valid rows found. Please check your required columns.")
                elif len(df) > 500:
                    st.warning(f"⚠️ {len(df)} SKUs found. Trimming to first 500.")
                    df = df.head(500)
                else:
                    st.success(f"✅ {len(df)} SKUs validated successfully!")
                    st.dataframe(df.head(10), use_container_width=True)
                    if len(df) > 10:
                        st.caption(f"Showing first 10 of {len(df)} rows.")
                    st.session_state["df"] = df
                    st.info("✅ Data ready! Go to the **Results** tab to see your analysis.")

        except Exception as e:
            st.error(f"❌ Error reading file: {str(e)}")

# ════════════════════════════════════════════
# TAB 2 — RESULTS
# ════════════════════════════════════════════
with tab2:
    if "df" not in st.session_state:
        st.info("📂 Please upload your data in the Upload tab first.")
    else:
        df = st.session_state["df"]

        # ── Run all calculations
        if "all_results" not in st.session_state:
            with st.spinner("Calculating all SKUs..."):
                all_results = {}
                for _, row in df.iterrows():
                    try:
                        all_results[str(row["SKU_ID"])] = run_sku(row)
                    except Exception as e:
                        st.warning(f"Skipped {row['SKU_ID']}: {e}")
                st.session_state["all_results"] = all_results

        all_results = st.session_state["all_results"]

        # ── Portfolio summary cards
        st.markdown('<p class="section-header">Portfolio Overview</p>', unsafe_allow_html=True)

        total_cost    = sum(r[r["recommended"]]["total_cost"] for r in all_results.values())
        total_ss      = sum(r[r["recommended"]]["safety_stock"] for r in all_results.values())
        high_risk_cnt = sum(1 for r in all_results.values() if r["risk_score"] > 40)
        single_src    = sum(1 for _, row in df.iterrows() if int(row["Num_Suppliers"]) == 1)

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown(f"""<div class="metric-card">
              <div class="metric-label">Total SKUs</div>
              <div class="metric-value">{len(all_results)}</div>
              <div class="metric-unit">products analysed</div>
            </div>""", unsafe_allow_html=True)
        with c2:
            st.markdown(f"""<div class="metric-card">
              <div class="metric-label">Total Annual Cost</div>
              <div class="metric-value">${total_cost:,.0f}</div>
              <div class="metric-unit">across all SKUs</div>
            </div>""", unsafe_allow_html=True)
        with c3:
            st.markdown(f"""<div class="metric-card">
              <div class="metric-label">High Risk SKUs</div>
              <div class="metric-value" style="color:#cc0000">{high_risk_cnt}</div>
              <div class="metric-unit">need immediate attention</div>
            </div>""", unsafe_allow_html=True)
        with c4:
            st.markdown(f"""<div class="metric-card">
              <div class="metric-label">Single Source SKUs</div>
              <div class="metric-value" style="color:#996600">{single_src}</div>
              <div class="metric-unit">supplier dependency risk</div>
            </div>""", unsafe_allow_html=True)

        # ── SKU navigator
        st.markdown('<p class="section-header">SKU Deep Dive</p>', unsafe_allow_html=True)

        sku_options = [f"{row['SKU_ID']} — {row['Product_Name']}"
                       for _, row in df.iterrows()]
        selected    = st.selectbox("Select a product to analyse:", sku_options)
        sku_id      = selected.split(" — ")[0]
        sku_name    = selected.split(" — ")[1]
        r           = all_results[str(sku_id)]
        rec         = r["recommended"]

        # ── Recommended scenario banner
        st.markdown(f'<div class="scenario-winner">⭐ Recommended Scenario: {rec} — Total Annual Cost: ${r[rec]["total_cost"]:,.0f}</div>',
                    unsafe_allow_html=True)

        # ── Risk indicator
        rs = r["risk_score"]
        if rs > 40:
            st.markdown(f'<div class="risk-high">🔴 High Risk SKU (score: {rs}) — Immediate attention required</div>', unsafe_allow_html=True)
        elif rs > 20:
            st.markdown(f'<div class="risk-med">🟡 Medium Risk SKU (score: {rs}) — Monitor closely</div>', unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="risk-low">🟢 Low Risk SKU (score: {rs}) — Well managed</div>', unsafe_allow_html=True)

        # ── Metrics for all 3 scenarios
        cols = st.columns(3)
        for col, scenario in zip(cols, ["Conservative", "Balanced", "Lean"]):
            s = r[scenario]
            star = " ⭐" if scenario == rec else ""
            with col:
                st.markdown(f"**{scenario} (SL: {s['service_level']}%){star}**")
                st.markdown(f"""<div class="metric-card">
                  <div class="metric-label">EOQ</div>
                  <div class="metric-value">{s['eoq']:,}</div>
                  <div class="metric-unit">units per order</div>
                </div>
                <div class="metric-card">
                  <div class="metric-label">Safety Stock</div>
                  <div class="metric-value">{s['safety_stock']:,}</div>
                  <div class="metric-unit">units</div>
                </div>
                <div class="metric-card">
                  <div class="metric-label">Reorder Point</div>
                  <div class="metric-value">{s['reorder_point']:,}</div>
                  <div class="metric-unit">units</div>
                </div>
                <div class="metric-card">
                  <div class="metric-label">Total Annual Cost</div>
                  <div class="metric-value">${s['total_cost']:,.0f}</div>
                  <div class="metric-unit">USD/year</div>
                </div>
                <div class="metric-card">
                  <div class="metric-label">Orders Per Year</div>
                  <div class="metric-value">{s['orders_per_year']}</div>
                  <div class="metric-unit">orders</div>
                </div>""", unsafe_allow_html=True)

        # ── Charts
        st.markdown('<p class="section-header">Visual Analysis</p>', unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            st.plotly_chart(chart_cost_breakdown(r, sku_name), use_container_width=True)
        with c2:
            st.plotly_chart(chart_safety_rop(r, sku_name), use_container_width=True)
        c3, c4 = st.columns(2)
        with c3:
            st.plotly_chart(chart_total_cost(r, sku_name), use_container_width=True)
        with c4:
            st.plotly_chart(chart_risk_vs_cost(r, sku_name), use_container_width=True)

# ════════════════════════════════════════════
# TAB 3 — AI & EXPORT
# ════════════════════════════════════════════
with tab3:
    if "all_results" not in st.session_state:
        st.info("📂 Please upload your data in the Upload tab first.")
    else:
        all_results = st.session_state["all_results"]
        df          = st.session_state["df"]

        # ── Privacy notice
        st.markdown("""
<div style="background:#1a1a2e;border:1px solid #333;border-radius:8px;padding:0.8rem 1rem;margin-bottom:1rem;font-size:0.82rem;color:#aaa;">
🔒 <strong style="color:#ccc;">Privacy Notice:</strong> Your data is processed locally and never stored.
AI analysis sends summarised metrics (not raw data) to Groq for processing.
No data is retained after your session ends. Closing this tab permanently deletes all data.
</div>
""", unsafe_allow_html=True)

        # ── AI Portfolio Analysis
        st.markdown('<p class="section-header">AI Portfolio Analysis</p>', unsafe_allow_html=True)
        st.markdown("Click below to get an AI-powered analysis of your entire inventory portfolio — risks, opportunities, and immediate actions.")

        if st.button("🤖 Generate AI Analysis"):
            with st.spinner("AI is analysing your portfolio..."):
                ai_output = get_ai_analysis(all_results, df)
                st.session_state["ai_output"]       = ai_output
                st.session_state["chat_history"]    = []
                st.session_state["chat_messages"]   = []

        if "ai_output" in st.session_state:
            st.markdown(st.session_state["ai_output"])

            # ── Chat section
            st.markdown('<p class="section-header">💬 Ask Follow-up Questions</p>', unsafe_allow_html=True)
            st.markdown("Ask anything about your inventory data — specific SKUs, reorder decisions, risk explanations, cost savings, and more.")

            # Show chat history
            if "chat_messages" not in st.session_state:
                st.session_state["chat_messages"] = []
            if "chat_history" not in st.session_state:
                st.session_state["chat_history"]  = []

            for msg in st.session_state["chat_messages"]:
                with st.chat_message(msg["role"]):
                    st.markdown(msg["content"])

            # Chat input
            if prompt := st.chat_input("Ask a question about your inventory... e.g. Which SKU should I reorder first?"):
                # Show user message
                st.session_state["chat_messages"].append({"role": "user", "content": prompt})
                with st.chat_message("user"):
                    st.markdown(prompt)

                # Get AI response
                with st.chat_message("assistant"):
                    with st.spinner("Thinking..."):
                        client   = Groq(api_key=GROQ_API_KEY)
                        response = get_chat_response(
                            client, prompt,
                            st.session_state["chat_history"],
                            all_results, df
                        )
                        st.markdown(response)

                # Save to history
                st.session_state["chat_messages"].append({"role": "assistant", "content": response})
                st.session_state["chat_history"].append({"role": "user",      "content": prompt})
                st.session_state["chat_history"].append({"role": "assistant", "content": response})

            # Clear chat button
            if st.session_state.get("chat_messages"):
                if st.button("🗑️ Clear Chat History"):
                    st.session_state["chat_messages"] = []
                    st.session_state["chat_history"]  = []
                    st.rerun()

        # ── Download
        st.markdown('<p class="section-header">Download Full Report</p>', unsafe_allow_html=True)
        st.markdown("Download a complete Excel report with 5 organized sheets — summary dashboard, high risk SKUs, cost analysis, reorder schedule, and scenario comparison.")

        report = generate_excel_report(all_results, df)
        st.download_button(
            label="⬇️ Download Full Excel Report",
            data=report,
            file_name="pharma_inventory_optimization_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
