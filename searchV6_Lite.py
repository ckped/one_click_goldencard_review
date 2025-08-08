import streamlit as st
import pandas as pd
import sqlite3
import io
from datetime import datetime

# === 資料庫連線 ===
sqlite3.connect('mydb.sqlite', check_same_thread=False)

st.title("🔎 企業履歷綜合查詢工具")

# --- 查詢條件輸入 ---
company_id_input = st.text_input("公司統編", "")
company_name_input = st.text_input("公司名稱", "")

years_sql = """
    SELECT DISTINCT apply_year FROM rd_project
    UNION
    SELECT DISTINCT apply_year FROM smart_project
"""
years = pd.read_sql(years_sql, conn)
year_options = sorted(years['apply_year'].dropna().unique(), reverse=True)
selected_year = st.selectbox("申請年度", options=[None] + list(year_options))
group_options = ['新興跨域組', '平台經濟組', '數位服務組', '通訊傳播組']
selected_group = st.selectbox("組別篩選", options=[None] + group_options)

# --- session_state 初始化 ---
if 'dfs' not in st.session_state:
    st.session_state['dfs'] = []

# --- 開始查詢按鈕觸發 ---
if st.button("開始查詢"):
    st.session_state['dfs'] = []  # 清空舊結果

    # 顯示查詢條件
    st.write("🎯 查詢條件")
    st.write(f"公司統編：{company_id_input or '（未填）'}")
    st.write(f"公司名稱：{company_name_input or '（未填）'}")
    st.write(f"申請年度：{selected_year or '（未選）'}")
    st.write(f"收案組別：{selected_group or '（未填）'}")

    # --- 篩選 helper 函式 ---
    def make_filter(table_alias, field, use_group=False, name_field="project_name"):
        clause = ""
        if company_name_input:
             clause += f" AND {table_alias}.company_name LIKE '%{company_name_input}%'"
        elif company_id_input:
            clause += f" AND {table_alias}.{field} = '{company_id_input}'"
        if selected_year:
            clause += f" AND {table_alias}.apply_year = {selected_year}"
        if use_group and selected_group:
            clause += f" AND trim({table_alias}.\"group\") = '{selected_group}'"
        return clause
    # --- 研發資料 ---
    rd_query = """
        SELECT a.*, b.apply_amount, b.approved
        FROM rd_project AS a
        LEFT JOIN rd_item AS b
        ON a.apply_year = b.apply_year
           AND a.company_id = b.company_id
           AND a.project_name = b.project_name
        WHERE 1=1
    """
    rd_query += make_filter('a', 'company_id', use_group=True,name_field='project_name')

    rd_df = pd.read_sql(rd_query, conn)

    # 改中文欄位
    rd_df = rd_df.rename(columns={
        "company_id": "公司統編",
        "company_name": "公司名稱",
        "apply_year": "申請年度",
        "project_name": "計畫名稱",
        "group": "收案組別",
        "type_innovation_sme": "產創/中小企",
        "industry_category": "產業類別",
        "apply_amount": "申請金額",
        "approved": "是否通過"
    })
    #儀表板功能
    if not rd_df.empty and "申請金額" in rd_df.columns and "是否通過" in rd_df.columns:
        rd_df["申請金額"] = pd.to_numeric(rd_df["申請金額"], errors="coerce")

        total_cases_rd = len(rd_df)
        total_amount_rd = rd_df["申請金額"].sum()

        approved_df_rd = rd_df[rd_df["是否通過"] == "通過"]
        approved_cases_rd = len(approved_df_rd)
        approved_amount_rd = approved_df_rd["申請金額"].sum()

        case_rate_rd = f"{(approved_cases_rd / total_cases_rd * 100):.1f}%" if total_cases_rd > 0 else "0%"
        amount_rate_rd = f"{(approved_amount_rd / total_amount_rd * 100):.1f}%" if total_amount_rd > 0 else "0%"

        st.markdown("### 研發投抵 統計摘要")
        # 第一排
        col1, col2 = st.columns(2)
        col1.metric("總件數", total_cases_rd)
        col2.metric("申請金額總額", f"{total_amount_rd:,.0f}元")

        # 第二排
        col3, col4 = st.columns(2)
        col3.metric("通過件數", approved_cases_rd)
        col4.metric("通過金額", f"{approved_amount_rd:,.0f}元")

        # 第三排
        col5, col6 = st.columns(2)
        col5.metric("案件通過率", case_rate_rd)
        col6.metric("金額通過率", amount_rate_rd)
    #展示研發資料
    st.subheader("🧪 研發資料")
    st.dataframe(rd_df)
    st.session_state['dfs'].append(('研發資料', rd_df))

    # --- 設備資料 ---
    smart_query = """
        SELECT a.*, b.item_no, b.item_name, b.item_type, b.total_amount, b.subsidy, b.apply_amount, b.first_review, b.final_review
        FROM smart_project AS a
        LEFT JOIN smart_item AS b
        ON a.apply_year = b.apply_year
           AND a.company_id = b.company_id
           AND a.plan_name = b.plan_name
        WHERE 1=1
    """
    smart_query += make_filter('a', 'company_id',name_field='plan_name')

    raw_smart = pd.read_sql(smart_query, conn)

    mapping = {
        '資安產業': '新興跨域組',
        **dict.fromkeys([
            '其他無店面零售業-電子購物','其他無店面零售業-第三方支付',
            '軟體出版業-線上遊戲','資訊服務業-電簽服務','軟體出版業-數位內容'
        ], '平台經濟組'),
        **dict.fromkeys([
            '軟體出版業-套裝軟體','電腦程式設計、諮詢相關服務業','資訊服務業','資訊服務產業'
        ], '數位服務組'),
        **dict.fromkeys(['電信產業','傳播事業'], '通訊傳播組')
    }
    raw_smart['mapped_group'] = raw_smart['industry_category'].map(mapping)

    if selected_group:
        smart_df = raw_smart[raw_smart['mapped_group'] == selected_group]
    else:
        smart_df = raw_smart.copy()
    
    smart_df = smart_df.rename(columns={
        "company_id": "公司統編",
        "company_name": "公司名稱",
        "apply_year": "申請年度",
        "plan_name": "計畫名稱",
        "mapped_group": "收案組別",
        "industry_category": "產業類別",
        "item_no": "項目編號",
        "item_name": "項目名稱",
        "item_type": "項目類型",
        "total_amount": "購買總金額",
        "subsidy": "補助款",
        "apply_amount": "申請金額",
        "first_review": "初審結果",
        "final_review": "複審結果"
    })
    if not smart_df.empty and "申請金額" in smart_df.columns and "複審結果" in smart_df.columns:
        smart_df["申請金額"] = pd.to_numeric(smart_df["申請金額"], errors="coerce")

        total_cases_s = len(smart_df)
        total_amount_s = smart_df["申請金額"].sum()

        approved_df_s = smart_df[smart_df["複審結果"] == "複審項目核定"]
        approved_cases_s = len(approved_df_s)
        approved_amount_s = approved_df_s["申請金額"].sum()

        case_rate_s = f"{(approved_cases_s / total_cases_s * 100):.1f}%" if total_cases_s > 0 else "0%"
        amount_rate_s = f"{(approved_amount_s / total_amount_s * 100):.1f}%" if total_amount_s > 0 else "0%"
        
        with st.container():
            # 設備投抵 統計摘要
            st.markdown("### 設備投抵 統計摘要")
            col1, col2 = st.columns(2)
            col1.metric("項目數", total_cases_s)
            col2.metric("申請金額總額", f"{total_amount_s:,.0f}元")

            col3, col4 = st.columns(2)
            col3.metric("通過項目數", approved_cases_s)
            col4.metric("通過金額", f"{approved_amount_s:,.0f}元")

            col5, col6 = st.columns(2)
            col5.metric("項目通過率", case_rate_s)
            col6.metric("金額通過率", amount_rate_s)


    st.subheader("🤖 設備投抵資料")
    st.dataframe(smart_df)
    st.session_state['dfs'].append(('設備資料', smart_df))

    # --- 上市櫃資料 ---
    ipo_query = "SELECT * FROM ipo_info WHERE 1=1"
    if company_name_input:
        ipo_query += f" AND company_name LIKE '%{company_name_input}%'"
    elif company_id_input:
        ipo_query += f" AND company_id = '{company_id_input}'"
    if selected_group:
        ipo_query += f" AND trim(\"group\") = '{selected_group}'"

    ipo_df = pd.read_sql(ipo_query, conn)
    # 日期轉民國年 #轉中文欄位
    ipo_df = ipo_df.rename(columns={
        "company_id": "公司統編",
        "company_name": "公司名稱",
        "capital": "實收資本額",
        "apply_date": "申請日期",
        "visit_date": "拜會日期",
        "meeting_date": "評估會議日期",
        "ipo_type": "上市櫃類型",
        "group": "業務組別",
        "country": "國內/國外",
        "broker": "主辦券商",
        "status": "拜會時狀態",
        "result": "申請結果",
        "remark": "結果備註"
    })
    # --- 日期轉民國年 ---
    def to_roc_str(dt):
        if pd.isnull(dt):
            return ""
        roc_year = dt.year - 1911
        return f"{roc_year}/{dt.month}/{dt.day}"
    if "申請日期" in ipo_df.columns:
        ipo_df["申請日期"] = pd.to_datetime(ipo_df["申請日期"], errors='coerce').apply(to_roc_str)
    if "拜會日期" in ipo_df.columns:
        ipo_df["拜會日期"] = pd.to_datetime(ipo_df["拜會日期"], errors='coerce').apply(to_roc_str)
    if "評估會議日期" in ipo_df.columns:
        ipo_df["評估會議日期"] = pd.to_datetime(ipo_df["評估會議日期"], errors='coerce').apply(to_roc_str)

    if not ipo_df.empty:
        st.subheader("📈 上市櫃資料")
        st.dataframe(ipo_df)
        st.session_state['dfs'].append(('上市櫃資料', ipo_df))
    else:
        st.info("ℹ️ 無符合條件的上市櫃資料")

# --- 下載按鈕 ---
if st.session_state['dfs']:
    company_label = company_id_input or (company_name_input or 'ALL')
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"查詢結果_{company_label}_{selected_year or 'ALL'}_{timestamp}.xlsx"

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        for name, df in st.session_state['dfs']:
            df.to_excel(writer, sheet_name=name[:31], index=False)
    buffer.seek(0)

    st.download_button(
        label="📥 下載查詢結果 Excel",
        data=buffer,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# 關閉
conn.close()
