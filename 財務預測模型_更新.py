import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as pgo

# 頁面配置
st.set_page_config(
    page_title="三年間財務預測與分園區損益模型 Web App",
    page_icon="📊",
    layout="wide"
)

st.title("📊 三年間財務預測與分園區損益模型 (Web 版)")
st.caption("原汁原味還原 Excel 多工作表架構，支援各頁面線上填寫與跨頁自動動態勾稽")

# 時間軸 (2027 Jan - 2029 Dec)
months_36 = [f"{m}-{str(y)[2:]}" for y in [2027, 2028, 2029] for m in ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']]

# -----------------------------------------------------------------------------
# 1. 全局數據狀態初始化 (Session State for 8 Sheets)
# -----------------------------------------------------------------------------

# Sheet 1: Construction 專案主表 & 每月進度 (%)
if "df_con_projects" not in st.session_state:
    st.session_state.df_con_projects = pd.DataFrame({
        "專案編號": ["PRJ-01", "PRJ-02", "PRJ-03"],
        "專案名稱": ["A園區二期新建工程", "B園區物流中心擴建", "C廠區修繕及設備翻新"],
        "工程總收入 (NTD)": [1000000000, 500000000, 200000000],
        "工程總成本 (NTD)": [900000000, 420000000, 160000000]
    })

if "df_con_progress" not in st.session_state:
    p1 = [0.03, 0.05, 0.05, 0.05, 0.05, 0.05, 0.05, 0.05, 0.05] + [0.03]*27
    p2 = [0.02]*36
    p3 = [0.02]*36
    data_p = {"專案名稱": ["A園區二期新建工程", "B園區物流中心擴建", "C廠區修繕及設備翻新"]}
    for idx, m in enumerate(months_36):
        data_p[m] = [p1[idx], p2[idx], p3[idx]]
    st.session_state.df_con_progress = pd.DataFrame(data_p)

# Sheet 2: Leasing 9大園區細項 (每月基準 NTD)
if "df_leasing" not in st.session_state:
    st.session_state.df_leasing = pd.DataFrame({
        "科目/園區": ["倉儲租賃收入 (A1)", "倉儲管理費收入 (A2)", "儲位服務收入 (A3)", "設備租賃收入 (A4)", "ASRS服務收入 (A5)", "水電收入 (A6)", "其他收入 (A7)", "底租支出 (B1)", "抽成 (B2)", "物管-修繕與維護 (B3)", "物管-清潔與管理 (B4)", "折舊費用-PPE (B5)", "水電成本 (B6)", "其他成本 (B7)"],
        "大園": [49000, 2000, 0, 3000, 0, 3000, 1000, 28000, 2000, 0, 1000, 1000, 3000, 1000],
        "瑞芳B區": [208000, 18000, 0, 14000, 0, 20000, 1000, 159000, 10000, 3000, 3000, 11000, 18000, 2000],
        "烏日": [93000, 4000, 0, 0, 0, 13000, 0, 86000, 2000, 1000, 5000, 0, 13000, 1000],
        "楊梅和平": [146000, 6000, 0, 0, 0, 16000, 1000, 87000, 17000, 2000, 3000, 1000, 16000, 1000],
        "楊梅啟明": [34000, 2000, 0, 0, 0, 0, 0, 32000, 0, 1000, 2000, 0, 0, 0],
        "瑞芳A1": [35000, 1000, 0, 1000, 0, 0, 0, 27000, 0, 1000, 1000, 1000, 0, 0],
        "瑞芳A2": [26000, 1000, 0, 38000, 0, 0, 0, 33000, 0, 2000, 1000, 6000, 0, 0],
        "OMEGA 2": [87000, 5000, 134000, 23000, 19000, 3000, 1000, 123000, 0, 5000, 4000, 20000, 6000, 8000],
        "OMEGA 3": [47000, 2000, 0, 0, 0, 2000, 0, 35000, 0, 1000, 2000, 1000, 3000, 1000]
    })

# Sheet 3: CAPEX 投資清單
if "df_capex" not in st.session_state:
    st.session_state.df_capex = pd.DataFrame({
        "專案編號": ["CAPEX-01", "CAPEX-02", "CAPEX-03", "CAPEX-04"],
        "所屬園區/項目名稱": ["桃園 Park A 自動化分揀設備升級", "新竹 Omega Park 太陽能屋頂建置", "台中 Central Park 溫控倉儲設備擴建", "全園區智慧安防與 IT 骨幹更新"],
        "投資總金額 (NTD)": [60000000, 36000000, 48000000, 12000000],
        "預計發生月份": ["Mar-27", "Jun-27", "Jan-28", "May-27"],
        "耐用年限 (年)": [5, 10, 8, 3]
    })

# Sheet 4: SG&A 人力與管銷
if "df_sga_hr" not in st.session_state:
    st.session_state.df_sga_hr = pd.DataFrame({
        "部門名稱": ["管理總部 (HQ)", "營運物流部 (Operations)", "工程管理部 (Construction)", "業務與行銷部 (Sales)"],
        "員工人數 (Headcount)": [15, 30, 12, 8],
        "平均月薪/人 (NTD)": [80000, 55000, 70000, 65000],
        "勞健保與福利加成 (%)": [0.20, 0.18, 0.18, 0.18]
    })

if "df_sga_other" not in st.session_state:
    st.session_state.df_sga_other = pd.DataFrame({
        "費用項目": ["租金與辦公室雜費", "行銷推廣費用", "資訊系統與軟體授權", "專業服務費 (會計/法律)", "其他日常管銷費用"],
        "每月預估金額 (NTD)": [500000, 300000, 250000, 200000, 350000]
    })

# Sheet 5: IFRS 16 合約
if "df_ifrs16" not in st.session_state:
    st.session_state.df_ifrs16 = pd.DataFrame({
        "合約編號": ["LEASE-01", "LEASE-02", "LEASE-03"],
        "園區名稱": ["桃園 Logistics Park A 土地租約", "新竹 Omega Park 廠房租約", "台中 Central Park 倉儲租約"],
        "原始ROU資產 (NTD)": [500000000, 400000000, 300000000],
        "年折現率 (%)": [0.025, 0.025, 0.025],
        "租賃期限 (月)": [36, 36, 36],
        "每月實際支付 (-MLA pmt)": [14500000, 11500000, 8500000]
    })

# -----------------------------------------------------------------------------
# 2. 建立多頁面 (Multi-Tab Container) 完全對應 Excel 8 個 Sheet
# -----------------------------------------------------------------------------
tabs = st.tabs([
    "📖 SOP 說明", 
    "🏗️ Construction", 
    "🏢 Leasing", 
    "🏬 Park P&L", 
    "⚙️ CAPEX", 
    "👥 SG&A", 
    "📜 IFRS16_Lease", 
    "📈 P&L 總表"
])

# -----------------------------------------------------------------------------
# 分頁 1: SOP
# -----------------------------------------------------------------------------
with tabs[0]:
    st.header("三年間財務預測與分園區損益模型 — 標準作業程序 (SOP)")
    st.markdown("""
    ### 一、 模型架構與工作表總覽
    * **`SOP`**：本說明頁，提供模型填寫流程與規範。
    * **`Construction`**：工程專案參數設定與 36 個月完工進度 (%) 填寫區。
    * **`Leasing`**：現有 9 大園區之單月租賃收入與營運成本細項填寫區。
    * **`Park P&L`**：分園區年度損益分析表，支援動態切換檢視年度。
    * **`CAPEX`**：資本支出投資清單及 36 個月月折舊費用試算。
    * **`SG&A`**：公司總部與各部門人力編制及日常管銷費用預估。
    * **`IFRS16_Lease`**：多園區 IFRS 16 租賃合約清單及 36 個月攤銷表。
    * **`P&L`**：全公司 36 個月綜合損益表與 EBITDA 分析總表（自動勾稽串接所有子頁面）。
    """)

# -----------------------------------------------------------------------------
# 分頁 2: Construction (動態連動進度列)
# -----------------------------------------------------------------------------
with tabs[1]:
    st.header("工程專案參數設定與收入成本拆解表 (Construction)")
    
    st.subheader("1. 專案總額設定（可在下方新增/刪除專案列）")
    edited_con_projects = st.data_editor(
        st.session_state.df_con_projects,
        num_rows="dynamic",
        use_container_width=True,
        key="edit_con_p"
    )
    
    # 同步更新進度列
    current_p_names = edited_con_projects["專案名稱"].dropna().tolist()
    existing_p_names = st.session_state.df_con_progress["專案名稱"].tolist()
    
    prog_rows = []
    for p_name in current_p_names:
        if p_name in existing_p_names:
            row = st.session_state.df_con_progress[st.session_state.df_con_progress["專案名稱"] == p_name].iloc[0].to_dict()
        else:
            row = {"專案名稱": p_name}
            for m in months_36:
                row[m] = 0.0
        prog_rows.append(row)
    st.session_state.df_con_progress = pd.DataFrame(prog_rows)

    st.subheader("2. 專案 36 個月完工進度設定 (%)")
    edited_con_progress = st.data_editor(
        st.session_state.df_con_progress,
        use_container_width=True,
        key="edit_con_prog"
    )

# -----------------------------------------------------------------------------
# 分頁 3: Leasing
# -----------------------------------------------------------------------------
with tabs[2]:
    st.header("各園區租賃收入與營運成本細項設定表 (Leasing - 每月基準 NTD)")
    st.caption("可以直接點擊儲存格修改 9 大園區的單月基準數值：")
    edited_leasing = st.data_editor(
        st.session_state.df_leasing,
        num_rows="dynamic",
        use_container_width=True,
        key="edit_leasing"
    )

# -----------------------------------------------------------------------------
# 分頁 4: Park P&L
# -----------------------------------------------------------------------------
with tabs[3]:
    st.header("分園區損益分析表 (Park P&L)")
    selected_year = st.selectbox("選擇試算年度：", [2027, 2028, 2029], key="park_year")
    y_mult = (1 + 0.02) ** (selected_year - 2027)
    
    park_cols = [c for c in edited_leasing.columns if c != "科目/園區"]
    df_park_calc = edited_leasing.copy()
    for c in park_cols:
        df_park_calc[c] = df_park_calc[c] * 12 * y_mult
    
    st.dataframe(df_park_calc.style.format({c: "{:,.0f}" for c in park_cols}), use_container_width=True)

# -----------------------------------------------------------------------------
# 分頁 5: CAPEX
# -----------------------------------------------------------------------------
with tabs[4]:
    st.header("資本支出 (CAPEX) 與固定資產折舊規劃表")
    edited_capex = st.data_editor(
        st.session_state.df_capex,
        num_rows="dynamic",
        use_container_width=True,
        key="edit_capex"
    )

# -----------------------------------------------------------------------------
# 分頁 6: SG&A
# -----------------------------------------------------------------------------
with tabs[5]:
    st.header("管銷費用 (SG&A) 與人力成本規劃表")
    st.subheader("1. 各部門人力編制")
    edited_sga_hr = st.data_editor(st.session_state.df_sga_hr, num_rows="dynamic", use_container_width=True, key="edit_sga_hr")
    
    st.subheader("2. 其他固定營運費用")
    edited_sga_other = st.data_editor(st.session_state.df_sga_other, num_rows="dynamic", use_container_width=True, key="edit_sga_other")

# -----------------------------------------------------------------------------
# 分頁 7: IFRS16_Lease
# -----------------------------------------------------------------------------
with tabs[6]:
    st.header("多園區 IFRS 16 租賃合約總表")
    edited_ifrs16 = st.data_editor(
        st.session_state.df_ifrs16,
        num_rows="dynamic",
        use_container_width=True,
        key="edit_ifrs16"
    )

# -----------------------------------------------------------------------------
# 分頁 8: P&L 綜合損益總表 (跨分頁自動勾稽運算)
# -----------------------------------------------------------------------------
with tabs[7]:
    st.header("全公司三年財務預測綜合損益表與 EBITDA 分析 (P&L)")
    
    # 1. 計算 Construction
    con_rev_36 = np.zeros(36)
    con_cogs_36 = np.zeros(36)
    for idx, row_p in edited_con_projects.iterrows():
        p_name = row_p["專案名稱"]
        rev_t = row_p["工程總收入 (NTD)"]
        cogs_t = row_p["工程總成本 (NTD)"]
        if pd.notnull(p_name) and p_name in edited_con_progress["專案名稱"].values:
            prog_row = edited_con_progress[edited_con_progress["專案名稱"] == p_name].iloc[0, 1:].values
            prog_arr = np.array([float(p) if pd.notnull(p) else 0.0 for p in prog_row])
            con_rev_36 += rev_t * prog_arr
            con_cogs_36 += cogs_t * prog_arr

    # 2. 計算 Leasing
    leasing_num_cols = [c for c in edited_leasing.columns if c != "科目/園區"]
    leasing_rev_m = edited_leasing.iloc[0:7][leasing_num_cols].sum().sum()
    leasing_cogs_m = edited_leasing.iloc[7:14][leasing_num_cols].sum().sum()
    
    leasing_rev_36 = [leasing_rev_m * ((1 + 0.02) ** (i // 12)) for i in range(36)]
    leasing_cogs_36 = [leasing_cogs_m] * 36

    # 3. 計算 SG&A
    hr_total_m = (edited_sga_hr["員工人數 (Headcount)"] * edited_sga_hr["平均月薪/人 (NTD)"] * (1 + edited_sga_hr["勞健保與福利加成 (%)"])).sum()
    other_sga_m = edited_sga_other["每月預估金額 (NTD)"].sum()
    sga_36 = [hr_total_m + other_sga_m] * 36

    # 4. 計算 CAPEX 折舊
    capex_dep_36 = np.zeros(36)
    for idx, row_c in edited_capex.iterrows():
        total_inv = row_c["投資總金額 (NTD)"]
        m_start = row_c["預計發生月份"] if "預計發生月份" in row_c else row_c.iloc[3]
        years = row_c["耐用年限 (年)"]
        if pd.notnull(total_inv) and years > 0 and m_start in months_36:
            start_idx = months_36.index(m_start)
            monthly_dep = total_inv / (years * 12)
            for i in range(start_idx, 36):
                capex_dep_36[i] += monthly_dep

    # 5. 計算 IFRS 16
    ifrs_rou_total = edited_ifrs16["原始ROU資產 (NTD)"].sum()
    ifrs_pmt_total = edited_ifrs16["每月實際支付 (-MLA pmt)"].sum()
    ifrs_dep_36 = [ifrs_rou_total / 36] * 36
    ifrs_pmt_36 = [ifrs_pmt_total] * 36
    
    ifrs_interest_36 = []
    bop = ifrs_rou_total
    avg_rate = edited_ifrs16["年折現率 (%)"].mean() if len(edited_ifrs16) > 0 else 0.025
    for i in range(36):
        interest = bop * (avg_rate / 12)
        ifrs_interest_36.append(interest)
        bop = bop + interest - ifrs_pmt_total

    # 組裝 36 個月 P&L
    df_pnl = pd.DataFrame(index=months_36)
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
    df_pnl["- Monthly Lease Payment"] = ifrs_pmt_36
    df_pnl["+ IFRS 16 ROU 折舊"] = ifrs_dep_36
    df_pnl["+ CAPEX 折舊"] = capex_dep_36
    df_pnl["EBITDA"] = df_pnl["營業利益 (Operating Profit)"] - df_pnl["- Monthly Lease Payment"] + df_pnl["+ IFRS 16 ROU 折舊"] + df_pnl["+ CAPEX 折舊"]

    # 繪製走勢圖與顯示表格
    fig = pgo.Figure()
    fig.add_trace(pgo.Bar(x=df_pnl.index, y=df_pnl["總營收 (Total Revenue)"]/1e6, name="總營收 (百萬元)", marker_color="#1F4E79"))
    fig.add_trace(pgo.Scatter(x=df_pnl.index, y=df_pnl["EBITDA"]/1e6, name="EBITDA (百萬元)", line=dict(color="#ED7D31", width=3)))
    fig.update_layout(title="跨分頁數據連動 — 36 個月營收與 EBITDA 走勢", xaxis_title="月份", yaxis_title="新台幣 (百萬元)", barmode="group", height=420)
    st.plotly_chart(fig, use_container_width=True)

    st.subheader("📑 36 個月綜合損益總表 (跨頁勾稽算結果)")
    st.dataframe(df_pnl.T.style.format("{:,.0f}"), use_container_width=True)
