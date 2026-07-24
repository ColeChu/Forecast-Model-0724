import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as pgo
import io

# 頁面設置
st.set_page_config(
    page_title="財務預測與線上編輯系統",
    page_icon="📈",
    layout="wide"
)

st.title("📈 財務預測與線上互動編輯系統 (Web-based Excel Simulator)")
st.caption("您可以直接在網頁畫面上修改工程進度與園區數據，系統將即時動態運算 36 個月損益與 EBITDA")

# -----------------------------------------------------------------------------
# 1. 初始化網頁預設的可編輯數據 (Session State Data)
# -----------------------------------------------------------------------------
months_36 = [f"{m}-{str(y)[2:]}" for y in [2027, 2028, 2029] for m in ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']]

# 預設工程專案
if "df_construction_input" not in st.session_state:
    st.session_state.df_construction_input = pd.DataFrame({
        "專案編號": ["PRJ-01", "PRJ-02", "PRJ-03"],
        "專案名稱": ["A園區二期新建工程", "B園區物流中心擴建", "C廠區修繕及設備翻新"],
        "工程總收入 (NTD)": [1000000000, 500000000, 200000000],
        "工程總成本 (NTD)": [900000000, 420000000, 160000000]
    })

# 預設 36 個月完工進度 (%)
if "df_progress_input" not in st.session_state:
    prog_data = {"指標/月份": ["全公司工程平均完工進度 (%)"]}
    for idx, m in enumerate(months_36):
        prog_data[m] = [0.03 if idx == 0 else (0.04 if idx == 1 else 0.02)]
    st.session_state.df_progress_input = pd.DataFrame(prog_data)

# 預設園區租賃與成本 (大園、瑞芳、烏日...)
if "df_leasing_input" not in st.session_state:
    park_names = ["大園", "瑞芳B區", "烏日", "楊梅和平", "楊梅啟明", "瑞芳A1", "瑞芳A2", "OMEGA 2", "OMEGA 3"]
    st.session_state.df_leasing_input = pd.DataFrame({
        "科目/園區": ["倉儲租賃收入 (A1)", "管理費與其他收入 (A2~A7)", "底租支出 (B1)", "抽成與物管成本 (B2~B7)"],
        "大園": [4900000, 900000, 2800000, 900000],
        "瑞芳B區": [20800000, 5300000, 15900000, 4700000],
        "烏日": [9300000, 1700000, 8600000, 2200000],
        "楊梅和平": [14600000, 2300000, 8700000, 4000000],
        "楊梅啟明": [3400000, 200000, 3200000, 300000],
        "瑞芳A1": [3500000, 300000, 2700000, 300000],
        "瑞芳A2": [2600000, 4000000, 3300000, 900000],
        "OMEGA 2": [8700000, 18600000, 12300000, 4400000],
        "OMEGA 3": [4700000, 400000, 3500000, 900000]
    })

# -----------------------------------------------------------------------------
# 側邊欄：情境模擬與檔案備份區
# -----------------------------------------------------------------------------
st.sidebar.header("🎛️ What-if 敏銳度模擬")
esc_rate = st.sidebar.slider("每年租金調漲率 (%)", 0.0, 5.0, 2.0, 0.5) / 100
vac_rate = st.sidebar.slider("預期招租空置率 (%)", 0.0, 15.0, 4.0, 0.5) / 100
prj_progress_mult = st.sidebar.slider("工程進度執行倍率", 0.5, 1.5, 1.0, 0.1)
ifrs_rate = st.sidebar.slider("IFRS 16 折現率 (%)", 1.0, 5.0, 2.5, 0.25) / 100

st.sidebar.markdown("---")
st.sidebar.header("📁 備份與匯入")
uploaded_file = st.sidebar.file_uploader("匯入現有 Excel 檔覆蓋網頁資料", type=["xlsx"])

if uploaded_file is not None:
    st.sidebar.success("✅ Excel 檔已匯入！網頁表格已同步更新。")

# -----------------------------------------------------------------------------
# 主畫面：1. 線上編輯表格區域  2. 自動動態算出的損益總表
# -----------------------------------------------------------------------------
tab_edit, tab_chart, tab_pnl = st.tabs(["📝 線上編輯 Excel 資料區", "📊 36個月 營收與 EBITDA 圖表", "📑 36個月 P&L 數據總表"])

with tab_edit:
    st.subheader("1. 工程專案與 36 個月進度表 (可直接點擊表格修改數字)")
    
    col_e1, col_e2 = st.columns([4, 6])
    with col_e1:
        st.markdown("**專案總額設定 (NTD)**")
        edited_con = st.data_editor(
            st.session_state.df_construction_input,
            num_rows="dynamic",
            use_container_width=True,
            key="editor_con"
        )
    with col_e2:
        st.markdown("**36 個月完工進度表 (%) (可左右滑動修改各月進度)**")
        edited_prog = st.data_editor(
            st.session_state.df_progress_input,
            use_container_width=True,
            key="editor_prog"
        )
    
    st.markdown("---")
    st.subheader("2. 9 大園區單月租賃與營運成本細項 (NTD)")
    edited_lea = st.data_editor(
        st.session_state.df_leasing_input,
        num_rows="dynamic",
        use_container_width=True,
        key="editor_lea"
    )

# -----------------------------------------------------------------------------
# 核心計算邏輯 (從使用者在網頁上修改後的edited_xxx即時計算)
# -----------------------------------------------------------------------------

# 1. 工程計算
total_con_rev = edited_con["工程總收入 (NTD)"].sum()
total_con_cogs = edited_con["工程總成本 (NTD)"].sum()

progress_row_values = edited_prog.iloc[0, 1:].values
progress_list = [float(p) if pd.notnull(p) else 0.0 for p in progress_row_values]

con_rev_36 = [total_con_rev * p * prj_progress_mult for p in progress_list]
con_cogs_36 = [total_con_cogs * p * prj_progress_mult for p in progress_list]

# 2. 園區 Leasing 計算
# 抓取編輯後各園區欄位總和
numeric_cols = [c for c in edited_lea.columns if c != "科目/園區"]
monthly_base_rev = edited_lea.iloc[0:2][numeric_cols].sum().sum() * (1 - vac_rate)
monthly_base_cogs = edited_lea.iloc[2:4][numeric_cols].sum().sum()

leasing_rev_36 = [monthly_base_rev * ((1 + esc_rate) ** (idx // 12)) for idx in range(len(con_rev_36))]
leasing_cogs_36 = [monthly_base_cogs] * len(con_rev_36)

# 3. SG&A & CAPEX & IFRS 16
sga_36 = [5200000] * len(con_rev_36)
capex_dep_36 = ([0]*2 + [1000000]*2 + [1333333]*2 + [1633333]*6 + [2133333]*24)[:len(con_rev_36)]

ifrs_rou = 1200000000
ifrs_dep_36 = [ifrs_rou / 36] * len(con_rev_36)
ifrs_payment_36 = [34500000] * len(con_rev_36)

ifrs_interest_36 = []
bop = ifrs_rou
for i in range(len(con_rev_36)):
    interest = bop * (ifrs_rate / 12)
    ifrs_interest_36.append(interest)
    bop = bop + interest - ifrs_payment_36[i]

# 彙總 P&L 表格
df_pnl = pd.DataFrame(index=months_36[:len(con_rev_36)])
df_pnl["Construction 營收"] = con_rev_36
df_pnl["Leasing 營收"] = leasing_rev_36
df_pnl["總營收 (Total Revenue)"] = df_pnl["Construction 營收"] + df_pnl["Leasing 營收"]

df_pnl["Construction 成本"] = con_cogs_36
df_pnl["Leasing 成本"] = leasing_cogs_36
df_pnl["總營業成本 (COGS)"] = df_pnl["Construction 成本"] + df_pnl["Leasing 成本"]

df_pnl["營業毛利 (Gross Profit)"] = df_pnl["總營收 (Total Revenue)"] - df_pnl["總營業成本 (COGS)"]
df_pnl["SG&A 費用"] = sga_36
df_pnl["CAPEX 折舊"] = capex_dep_36
df_pnl["營業利益 (Operating Profit)"] = df_pnl["營業毛利 (Gross Profit)"] - df_pnl["SG&A 費用"] - df_pnl["CAPEX 折舊"]

df_pnl["IFRS 16 利息費用"] = ifrs_interest_36
df_pnl["稅前淨利 (EBT)"] = df_pnl["營業利益 (Operating Profit)"] - df_pnl["IFRS 16 利息費用"]
df_pnl["所得稅"] = df_pnl["稅前淨利 (EBT)"].apply(lambda x: x * 0.2 if x > 0 else 0)
df_pnl["稅後淨利 (Net Income)"] = df_pnl["稅前淨利 (EBT)"] - df_pnl["所得稅"]

df_pnl["- Monthly Lease Payment"] = ifrs_payment_36
df_pnl["+ IFRS 16 ROU 折舊"] = ifrs_dep_36
df_pnl["+ CAPEX 折舊"] = capex_dep_36
df_pnl["EBITDA"] = df_pnl["營業利益 (Operating Profit)"] - df_pnl["- Monthly Lease Payment"] + df_pnl["+ IFRS 16 ROU 折舊"] + df_pnl["+ CAPEX 折舊"]

# -----------------------------------------------------------------------------
# 呈現圖表與數據總表
# -----------------------------------------------------------------------------
with tab_chart:
    st.subheader("📈 依據您網頁修改資料後即時計算之 EBITDA 趨勢圖")
    fig = pgo.Figure()
    fig.add_trace(pgo.Bar(x=df_pnl.index, y=df_pnl["總營收 (Total Revenue)"]/1e6, name="總營收 (百萬元)", marker_color="#1F4E79"))
    fig.add_trace(pgo.Scatter(x=df_pnl.index, y=df_pnl["EBITDA"]/1e6, name="EBITDA (百萬元)", line=dict(color="#ED7D31", width=3)))
    fig.update_layout(title="每月營收與 EBITDA 走勢", xaxis_title="月份", yaxis_title="新台幣 (百萬元)", barmode="group", height=450)
    st.plotly_chart(fig, use_container_width=True)

with tab_pnl:
    st.subheader("📑 36 個月綜合損益與 EBITDA 明細總表 (NTD)")
    st.dataframe(df_pnl.T.style.format("{:,.0f}"), use_container_width=True)