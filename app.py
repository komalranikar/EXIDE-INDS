import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px

st.set_page_config(
    page_title="Exide Industries - Financial Analysis",
    page_icon="📊",
    layout="wide"
)

st.markdown("""
<style>
.main-title {font-size:34px;font-weight:700;margin-bottom:0}
.sub-title {font-size:16px;color:#666;margin-bottom:25px}
.section-title {font-size:24px;font-weight:650;margin-top:10px;margin-bottom:15px}
div[data-testid="stMetric"] {
    border:1px solid #dddddd;
    border-radius:10px;
    padding:15px;
    background-color:#fafafa;
}
</style>
""", unsafe_allow_html=True)

FILE_NAME = "Exide Inds.xlsx"
SHEET_NAME = "Data Sheet"

def clean_number(value):
    try:
        if pd.isna(value):
            return np.nan
        if isinstance(value, str):
            value = value.replace(",", "").replace("%", "").strip()
        return float(value)
    except (ValueError, TypeError):
        return np.nan

def safe_divide(numerator, denominator):
    numerator = np.asarray(numerator, dtype=float)
    denominator = np.asarray(denominator, dtype=float)
    with np.errstate(divide="ignore", invalid="ignore"):
        result = np.divide(numerator, denominator)
    result[~np.isfinite(result)] = np.nan
    return result

def format_value(value, suffix=""):
    if pd.isna(value):
        return "N/A"
    return f"{value:,.2f}{suffix}"

try:
    raw = pd.read_excel(FILE_NAME, sheet_name=SHEET_NAME, header=None)
except FileNotFoundError:
    st.error(f"'{FILE_NAME}' was not found. Keep the Excel file in the same folder as app.py.")
    st.stop()
except Exception as e:
    st.error(f"Could not read the Excel file: {e}")
    st.stop()

# Source workbook rows. Excel row numbers are used here.
source_rows = {
    "Sales":17, "Raw Materials":18, "Change in Inventory":19,
    "Power":20, "Other Manufacturing":21, "Employee Cost":22,
    "Selling/Admin":23, "Other Expenses":24, "Other Income":25,
    "Depreciation":26, "Interest":27, "PBT":28, "Tax":29, "Net Profit":30,
    "Share Capital":57, "Reserves":58, "Borrowings":59,
    "Other Liabilities":60, "Net Block":62, "CWIP":63,
    "Investments":64, "Other Assets":65, "Total Assets":66,
    "Receivables":67, "Inventory":68
}

years = [f"FY{y}" for y in range(2017, 2027)]

def get_row(excel_row):
    values = raw.iloc[excel_row - 1, 1:11].tolist()
    return np.array([clean_number(x) for x in values], dtype=float)

data = {name: get_row(row) for name, row in source_rows.items()}

sales = data["Sales"]

expenses = (
    data["Raw Materials"] + data["Change in Inventory"] +
    data["Power"] + data["Other Manufacturing"] +
    data["Employee Cost"] + data["Selling/Admin"] +
    data["Other Expenses"]
)

operating_profit = sales - expenses
ebit = data["PBT"] + data["Interest"]
shareholders_equity = data["Share Capital"] + data["Reserves"]
capital_employed = shareholders_equity + data["Borrowings"]
working_capital = data["Other Assets"] - data["Other Liabilities"]

# Ratio calculations
net_profit_margin = safe_divide(data["Net Profit"] * 100, sales)
operating_profit_margin = safe_divide(operating_profit * 100, sales)
debt_equity = safe_divide(data["Borrowings"], shareholders_equity)
debt_ratio = safe_divide(data["Borrowings"] * 100, data["Total Assets"])
interest_coverage = safe_divide(ebit, data["Interest"])
asset_turnover = safe_divide(sales, data["Total Assets"])
inventory_turnover = safe_divide(sales, data["Inventory"])
debtor_days = safe_divide(data["Receivables"] * 365, sales)
roe = safe_divide(data["Net Profit"] * 100, shareholders_equity)

average_capital_employed = np.full(len(years), np.nan)
roce = np.full(len(years), np.nan)
for i in range(1, len(years)):
    average_capital_employed[i] = (capital_employed[i] + capital_employed[i-1]) / 2
    if average_capital_employed[i] != 0:
        roce[i] = ebit[i] * 100 / average_capital_employed[i]

# Du-Pont analysis
average_assets = np.full(len(years), np.nan)
average_equity = np.full(len(years), np.nan)
for i in range(1, len(years)):
    average_assets[i] = (data["Total Assets"][i] + data["Total Assets"][i-1]) / 2
    average_equity[i] = (shareholders_equity[i] + shareholders_equity[i-1]) / 2

dupont_npm = net_profit_margin.copy()
dupont_asset_turnover = safe_divide(sales, average_assets)
dupont_equity_multiplier = safe_divide(average_assets, average_equity)
dupont_roe = (
    (dupont_npm / 100) *
    dupont_asset_turnover *
    dupont_equity_multiplier * 100
)

analysis = pd.DataFrame({
    "Year": years,
    "Sales": sales,
    "Operating Profit": operating_profit,
    "Net Profit": data["Net Profit"],
    "EBIT": ebit,
    "Shareholders' Equity": shareholders_equity,
    "Capital Employed": capital_employed,
    "Working Capital": working_capital,
    "Net Profit Margin (%)": net_profit_margin,
    "Operating Profit Margin (%)": operating_profit_margin,
    "Debt-Equity (x)": debt_equity,
    "Debt Ratio (%)": debt_ratio,
    "Interest Coverage (x)": interest_coverage,
    "Asset Turnover (x)": asset_turnover,
    "Inventory Turnover (x)": inventory_turnover,
    "Debtor Days": debtor_days,
    "ROE (%)": roe,
    "ROCE (%)": roce,
    "DuPont NPM (%)": dupont_npm,
    "DuPont Asset Turnover (x)": dupont_asset_turnover,
    "DuPont Equity Multiplier (x)": dupont_equity_multiplier,
    "DuPont ROE (%)": dupont_roe
})

# Trend analysis: FY2017 = 100
trend = pd.DataFrame({"Year": years})
for item in ["Sales", "Operating Profit", "Net Profit"]:
    base = analysis[item].iloc[0]
    trend[item] = np.nan if pd.isna(base) or base == 0 else analysis[item] / base * 100

# Horizontal analysis
horizontal = pd.DataFrame({"Year": years})
for item in ["Sales", "Operating Profit", "Net Profit"]:
    values = analysis[item].values
    abs_change = np.full(len(values), np.nan)
    pct_change = np.full(len(values), np.nan)
    for i in range(1, len(values)):
        abs_change[i] = values[i] - values[i-1]
        if values[i-1] != 0:
            pct_change[i] = (values[i] - values[i-1]) / abs(values[i-1]) * 100
    horizontal[f"{item} Change"] = abs_change
    horizontal[f"{item} Change (%)"] = pct_change

# Vertical/common-size analysis
pnl_items = [
    "Sales","Raw Materials","Change in Inventory","Power",
    "Other Manufacturing","Employee Cost","Selling/Admin",
    "Other Expenses","Other Income","Depreciation","Interest",
    "PBT","Tax","Net Profit"
]
bs_items = [
    "Share Capital","Reserves","Borrowings","Other Liabilities",
    "Net Block","CWIP","Investments","Other Assets",
    "Receivables","Inventory","Total Assets"
]

vertical_pnl = pd.DataFrame({"Year": years})
for item in pnl_items:
    vertical_pnl[item] = safe_divide(data[item] * 100, sales)

vertical_bs = pd.DataFrame({"Year": years})
for item in bs_items:
    vertical_bs[item] = safe_divide(data[item] * 100, data["Total Assets"])

# Sidebar
st.sidebar.title("📊 Exide Industries")
page = st.sidebar.radio(
    "Select Analysis",
    ["Dashboard","Ratio Analysis","Du-Pont Analysis",
     "Trend Analysis","Horizontal Analysis","Vertical Analysis"]
)
st.sidebar.markdown("---")
st.sidebar.write("**Company:** Exide Industries Ltd.")
st.sidebar.write("**Period:** FY2017 – FY2026")
st.sidebar.write("**Source:** P&L + Balance Sheet")

st.markdown('<div class="main-title">Exide Industries Ltd. – Financial Analysis Dashboard</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Financial Ratio Analysis based on Profit & Loss Statement and Balance Sheet</div>', unsafe_allow_html=True)

# Dashboard
if page == "Dashboard":
    st.markdown('<div class="section-title">📌 Financial Snapshot</div>', unsafe_allow_html=True)
    selected_year = st.selectbox("Select Year", years, index=len(years)-1)
    i = years.index(selected_year)

    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric("Revenue / Sales", format_value(sales[i]))
    c2.metric("Net Profit", format_value(data["Net Profit"][i]))
    c3.metric("ROE", format_value(roe[i], "%"))
    c4.metric("ROCE", format_value(roce[i], "%"))
    c5.metric("Debt-Equity", format_value(debt_equity[i], "x"))

    st.markdown("---")
    st.markdown('<div class="section-title">📈 Revenue & Net Profit Trend</div>', unsafe_allow_html=True)
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=years, y=sales, mode="lines+markers", name="Sales"))
    fig.add_trace(go.Scatter(x=years, y=data["Net Profit"], mode="lines+markers", name="Net Profit"))
    fig.update_layout(xaxis_title="Financial Year", yaxis_title="Value", hovermode="x unified", height=420)
    st.plotly_chart(fig, use_container_width=True)

    st.markdown('<div class="section-title">💰 Profitability</div>', unsafe_allow_html=True)
    st.line_chart(analysis.set_index("Year")[["Net Profit Margin (%)","ROE (%)","ROCE (%)"]])

    col1,col2 = st.columns(2)
    with col1:
        st.markdown('<div class="section-title">🏦 Solvency</div>', unsafe_allow_html=True)
        st.line_chart(analysis.set_index("Year")[["Debt-Equity (x)","Interest Coverage (x)"]])
    with col2:
        st.markdown('<div class="section-title">⚙️ Efficiency</div>', unsafe_allow_html=True)
        st.line_chart(analysis.set_index("Year")[["Inventory Turnover (x)","Debtor Days"]])

    st.markdown("---")
    latest = analysis.iloc[-1]
    k1,k2,k3,k4 = st.columns(4)
    k1.metric("Operating Profit Margin", f"{latest['Operating Profit Margin (%)']:.2f}%")
    k2.metric("Interest Coverage", f"{latest['Interest Coverage (x)']:.2f}x")
    k3.metric("Inventory Turnover", f"{latest['Inventory Turnover (x)']:.2f}x")
    k4.metric("Debtor Days", f"{latest['Debtor Days']:.2f}")
    st.info("Calculations use only the Profit & Loss Statement and Balance Sheet data in the source workbook.")

# Ratio Analysis
elif page == "Ratio Analysis":
    st.markdown('<div class="section-title">📊 Financial Ratio Analysis</div>', unsafe_allow_html=True)
    ratio_columns = [
        "Year","Net Profit Margin (%)","Operating Profit Margin (%)",
        "Debt-Equity (x)","Debt Ratio (%)","Interest Coverage (x)",
        "Asset Turnover (x)","Inventory Turnover (x)","Debtor Days",
        "ROE (%)","ROCE (%)","Working Capital"
    ]
    st.dataframe(analysis[ratio_columns], use_container_width=True, hide_index=True)

    ratio_choice = st.selectbox("Select Ratio to Visualize", ratio_columns[1:-1])
    fig = px.line(analysis, x="Year", y=ratio_choice, markers=True, title=f"{ratio_choice} Trend")
    fig.update_layout(height=450)
    st.plotly_chart(fig, use_container_width=True)

    meanings = {
        "Net Profit Margin (%)":"Percentage of sales converted into net profit.",
        "Operating Profit Margin (%)":"Operating profit earned from sales.",
        "Debt-Equity (x)":"Relationship between borrowings and shareholders' equity.",
        "Debt Ratio (%)":"Proportion of total assets financed by borrowings.",
        "Interest Coverage (x)":"How many times operating earnings cover interest expense.",
        "Asset Turnover (x)":"How efficiently total assets generate sales.",
        "Inventory Turnover (x)":"How many times inventory is converted through sales.",
        "Debtor Days":"Approximate days taken to collect receivables.",
        "ROE (%)":"Return generated on shareholders' equity.",
        "ROCE (%)":"Return generated on capital employed."
    }
    st.info(meanings[ratio_choice])

# Du-Pont
elif page == "Du-Pont Analysis":
    st.markdown('<div class="section-title">🔎 Du-Pont Analysis</div>', unsafe_allow_html=True)
    st.write("Du-Pont breaks ROE into Net Profit Margin, Asset Turnover and Equity Multiplier.")
    dupont_table = analysis[[
        "Year","DuPont NPM (%)","DuPont Asset Turnover (x)",
        "DuPont Equity Multiplier (x)","DuPont ROE (%)"
    ]]
    st.dataframe(dupont_table, use_container_width=True, hide_index=True)

    fig = go.Figure()
    fig.add_trace(go.Scatter(x=years, y=dupont_npm, mode="lines+markers", name="Net Profit Margin"))
    fig.add_trace(go.Scatter(x=years, y=dupont_asset_turnover, mode="lines+markers", name="Asset Turnover"))
    fig.add_trace(go.Scatter(x=years, y=dupont_equity_multiplier, mode="lines+markers", name="Equity Multiplier"))
    fig.update_layout(title="Du-Pont Components", xaxis_title="Financial Year", yaxis_title="Value", hovermode="x unified", height=450)
    st.plotly_chart(fig, use_container_width=True)
    st.line_chart(analysis.set_index("Year")[["ROE (%)","DuPont ROE (%)"]])

# Trend
elif page == "Trend Analysis":
    st.markdown('<div class="section-title">📈 Trend Analysis</div>', unsafe_allow_html=True)
    st.write("FY2017 is the base year and is assigned an index value of 100.")
    st.dataframe(trend, use_container_width=True, hide_index=True)
    selected_items = st.multiselect(
        "Select items",
        ["Sales","Operating Profit","Net Profit"],
        default=["Sales","Operating Profit","Net Profit"]
    )
    if selected_items:
        fig = px.line(trend, x="Year", y=selected_items, markers=True, title="Trend Index (FY2017 = 100)")
        fig.update_layout(yaxis_title="Index", height=450)
        st.plotly_chart(fig, use_container_width=True)

# Horizontal
elif page == "Horizontal Analysis":
    st.markdown('<div class="section-title">↔️ Horizontal Analysis</div>', unsafe_allow_html=True)
    st.write("Each year is compared with the immediately previous year.")
    st.dataframe(horizontal, use_container_width=True, hide_index=True)
    selected_item = st.selectbox("Select Item", ["Sales","Operating Profit","Net Profit"])
    fig = px.bar(horizontal, x="Year", y=f"{selected_item} Change (%)", title=f"{selected_item} Year-on-Year Percentage Change")
    fig.update_layout(height=450)
    st.plotly_chart(fig, use_container_width=True)

# Vertical
elif page == "Vertical Analysis":
    st.markdown('<div class="section-title">📐 Vertical / Common-Size Analysis</div>', unsafe_allow_html=True)
    analysis_type = st.radio("Select Statement", ["Profit & Loss","Balance Sheet"], horizontal=True)

    if analysis_type == "Profit & Loss":
        st.write("Each Profit & Loss item is shown as a percentage of Sales.")
        st.dataframe(vertical_pnl, use_container_width=True, hide_index=True)
        selected_item = st.selectbox("Select P&L Item", pnl_items)
        fig = px.line(vertical_pnl, x="Year", y=selected_item, markers=True, title=f"{selected_item} as % of Sales")
    else:
        st.write("Each Balance Sheet item is shown as a percentage of Total Assets.")
        st.dataframe(vertical_bs, use_container_width=True, hide_index=True)
        selected_item = st.selectbox("Select Balance Sheet Item", bs_items)
        fig = px.line(vertical_bs, x="Year", y=selected_item, markers=True, title=f"{selected_item} as % of Total Assets")

    fig.update_layout(yaxis_title="Percentage", height=450)
    st.plotly_chart(fig, use_container_width=True)

st.markdown("---")
st.caption("Exide Industries Ltd. | Financial Ratio Analysis | FY2017–FY2026 | Source: P&L Statement and Balance Sheet")
