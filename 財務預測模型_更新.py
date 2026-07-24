import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as pgo

# 頁面設置
st.set_page_config(
    page_title="財務預測與敏銳度模擬系統 (支援 Excel 上傳)",
    page_icon="📈",
    layout="wide"
)

st.title("📈 財務預測與 What-if 敏銳度模擬 Web App")
st.caption("支援直接上傳 Excel 模型檔，自動讀取工程進度、園區數據並進行動態模擬")

# -----------------------------------------------------------------------------
# 側邊欄：1. Excel 檔案上傳區  2. What-if 控制面板
# -----------------------------------------------------------------------------
st.sidebar.header("📁 1. 數據源設定")
uploaded_file = st.sidebar.file_uploader("上傳你的 Excel 財務模型檔 (.xlsx)", type=["xlsx"])

st.sidebar.markdown("---")
st.sidebar.header("🎛️ 2. What-if 敏銳度模擬")

scenario = st.sidebar.selectbox(
    "選擇預設情境 (Preset Scenario)",
    ["自訂情境 (Custom)", "樂觀情境 (Optimistic)", "基準情境 (Base Case)", "悲觀情境 (Pessimistic)"]
)

if scenario == "樂觀情境 (Optimistic)":
    default_esc, default_vac, default_prog, default_ifrs = 3.5, 2.0, 1.2, 2.0
elif scenario == "悲觀情境 (Pessimistic)":
    default_esc, default_vac, default_prog, default_ifrs = 0.5, 8.0, 0.8, 3.5
else:
    default_esc, default_vac, default_prog, default_ifrs = 2.0, 4.0, 1.0, 2.5

esc_rate = st.sidebar.slider("每年租金調漲率 (%)", 0.0, 5.0, float(default_esc), 0.5) / 100
vac_rate = st.sidebar.slider("預期招租空置率 (%)", 0.0, 15.0, float(default_vac), 0.5) / 100
prj_progress_mult = st.sidebar.slider("工程完工進度倍率", 0.5, 1.5, float(default_prog), 0.1)
ifrs_rate = st.sidebar.slider("IFRS 16 年折現率 (%)", 1.0, 5.0, float(default_ifrs), 0.25) / 100

# -----------------------------------------------------------------------------
# 數據讀取與處理引擎
# -----------------------------------------------------------------------------
months_36 = [f"{m}-{str(y)[2:]}" for y in [2027, 2028, 2029] for m in ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']]

if uploaded_file is not None:
    st.sidebar.success("✅ 已成功讀取上傳的 Excel 檔案！")
    try:
        # 讀取 Excel 中的各工作表
        xls = pd.ExcelFile(uploaded_file)
        
        # 讀取 Construction 工程數據
        df_con_raw = pd.read_excel(xls, "Construction")
        total_con_rev = df_con_raw.iloc[3, 2] if df_con_raw.shape[0] > 3 else 1000000000
        total_con_cogs = df_con_raw.iloc[3, 3] if df_con_raw.shape[0] > 3 else 900000000
        
        # 讀取完工進度
        progress_row = df_con_raw.iloc[9, 2:38].values if df_con_raw.shape[0] > 9 else [0.03] + [0.02]*35
        progress_list = [float(p) if pd.notnull(p) else 0.0 for p in progress_row]
        
        # 讀取 Leasing 數據
        df_lea_raw = pd.read_excel(xls, "Leasing")
        base_rev_sum = df_lea_raw.iloc[11, 11] if df_lea_raw.shape[0] > 11 else 1060
        base_cogs_sum = df_lea_raw.iloc[19, 11] if df_lea_raw.shape[0] > 19 else 796
        
        monthly_base_rev = base_rev_sum * 1000 * (1 - vac_rate)
        monthly_base_cogs = base_cogs_sum * 1000
        
    except Exception as e:
        st.error(f"讀取 Excel 內容時出現格式差異，系統已切換至預設模型計算。錯誤訊息：{e}")
        # 預設範例數據
        total_con_rev, total_con_cogs = 1700000000, 1480000000
        progress_list = [0.03, 0.04, 0.05] + [0.02]*33
        monthly_base_rev = 1060000 * (1 - vac_rate)
        monthly_base_cogs = 796000
else:
    st.info("💡 提示：目前正使用範例數據運作。您可以在左側面板「上傳你的 Excel 檔」匯入真實工程與園區資料！")
    total_con_rev, total_con_cogs = 1700000000, 1480000000
    progress_list = [0.03, 0.04, 0.05] + [0.02]*33
    monthly_base_rev = 1060000 * (1 - vac_rate)
    monthly_base_cogs = 796000

# -----------------------------------------------------------------------------
# 36 個月財務模型動態計算
# -----------------------------------------------------------------------------
con_rev_36 = [total_con_rev * p * prj_progress_mult for p in progress_list]
con_cogs_36 = [total_con_cogs * p * prj_progress_mult for p in progress_list]

leasing_rev_36 = [monthly_base_rev * ((1 + esc_rate) ** (idx // 12)) for idx in range(36)]
leasing_cogs_36 = [monthly_base_cogs] * 36

sga_36 = [5200000] * 36
capex_dep_36 = [0]*2 + [1000000]*2 + [1333333]*2 + [1633333]*6 + [2133333]*24

ifrs_rou = 1200000000
ifrs_dep_36 = [ifrs_rou / 36] * 36
ifrs_payment_36 = [34500000] * 36

ifrs_interest_36 = []
bop = ifrs_rou
for i in range(36):
    interest = bop * (ifrs_rate / 12)
    ifrs_interest_36.append(interest)
    bop = bop + interest - ifrs_payment_36[i]

# 彙總全公司 P&L 表格
df_pnl = pd.DataFrame(index=months_36[:len(con_rev_36)])
df_pnl["Construction 營收"] = con_rev_36
df_pnl["Leasing 營收"] = leasing_rev_36
df_pnl["總營收 (Total Revenue)"] = df_pnl["Construction 營收"] + df_pnl["Leasing 營收"]

df_pnl["Construction 成本"] = con_cogs_36
df_pnl["Leasing 成本"] = leasing_cogs_36
df_pnl["總營業成本 (COGS)"] = df_pnl["Construction 成本"] + df_pnl["Leasing 成本"]

df_pnl["營業毛利 (Gross Profit)"] = df_pnl["總營收 (Total Revenue)"] - df_pnl["總營業成本 (COGS)"]
df_pnl["SG&A 費用"] = sga_36[:len(con_rev_36)]
df_pnl["CAPEX 折舊"] = capex_dep_36[:len(con_rev_36)]
df_pnl["營業利益 (Operating Profit)"] = df_pnl["營業毛利 (Gross Profit)"] - df_pnl["SG&A 費用"] - df_pnl["CAPEX 折舊"]

df_pnl["IFRS 16 利息費用"] = ifrs_interest_36[:len(con_rev_36)]
df_pnl["稅前淨利 (EBT)"] = df_pnl["營業利益 (Operating Profit)"] - df_pnl["IFRS 16 利息費用"]
df_pnl["所得稅"] = df_pnl["稅前淨利 (EBT)"].apply(lambda x: x * 0.2 if x > 0 else 0)
df_pnl["稅後淨利 (Net Income)"] = df_pnl["稅前淨利 (EBT)"] - df_pnl["所得稅"]

# EBITDA 計算
df_pnl["- Monthly Lease Payment"] = ifrs_payment_36[:len(con_rev_36)]
df_pnl["+ IFRS 16 ROU 折舊"] = ifrs_dep_36[:len(con_rev_36)]
df_pnl["+ CAPEX 折舊"] = capex_dep_36[:len(con_rev_36)]
df_pnl["EBITDA"] = df_pnl["營業利益 (Operating Profit)"] - df_pnl["- Monthly Lease Payment"] + df_pnl["+ IFRS 16 ROU 折舊"] + df_pnl["+ CAPEX 折舊"]

# -----------------------------------------------------------------------------
# 視覺化與數據呈現
# -----------------------------------------------------------------------------
kpi1, kpi2, kpi3, kpi4 = st.columns(4)
total_3y_rev = df_pnl["總營收 (Total Revenue)"].sum()
total_3y_ebitda = df_pnl["EBITDA"].sum()
avg_ebitda_margin = (df_pnl["EBITDA"].sum() / total_3y_rev) * 100 if total_3y_rev > 0 else 0
total_3y_net_profit = df_pnl["稅後淨利 (Net Income)"].sum()

kpi1.metric("3 年累積總營收", f"NT$ {total_3y_rev/1e8:.2f} 億")
kpi2.metric("3 年累積 EBITDA", f"NT$ {total_3y_ebitda/1e8:.2f} 億")
kpi3.metric("平均 EBITDA Margin", f"{avg_ebitda_margin:.1f}%")
kpi4.metric("3 年累積稅後淨利", f"NT$ {total_3y_net_profit/1e8:.2f} 億")

st.markdown("---")

tab1, tab2 = st.tabs(["📊 36個月 營收與 EBITDA 走勢圖", "📑 36個月 P&L 數據總表"])

with tab1:
    fig = pgo.Figure()
    fig.add_trace(pgo.Bar(x=df_pnl.index, y=df_pnl["總營收 (Total Revenue)"]/1e6, name="總營收 (百萬元)", marker_color="#1F4E79"))
    fig.add_trace(pgo.Scatter(x=df_pnl.index, y=df_pnl["EBITDA"]/1e6, name="EBITDA (百萬元)", line=dict(color="#ED7D31", width=3)))
    fig.update_layout(title="每月營收與 EBITDA 趨勢變化 (含 What-if 模擬受惠/影響)", xaxis_title="月份", yaxis_title="新台幣 (百萬元)", barmode="group", height=480)
    st.plotly_chart(fig, use_container_width=True)

with tab2:
    st.subheader("📑 36 個月綜合損益與 EBITDA 明細表 (NTD)")
    st.dataframe(df_pnl.T.style.format("{:,.0f}"), use_container_width=True)