import streamlit as st
import pandas as pd
import io

# --- 页面专业配置 ---
st.set_page_config(layout="wide", page_title="财务智能凭证中心Pro")

st.markdown("""
    <style>
    .main { background-color: #f5f7f9; }
    .stTabs [data-baseweb="tab-list"] { gap: 24px; }
    .stTabs [data-baseweb="tab"] { height: 50px; white-space: pre-wrap; background-color: #fff; border-radius: 4px 4px 0 0; gap: 1px; }
    .stTabs [aria-selected="true"] { background-color: #e6f7ff !important; border-bottom: 2px solid #1890ff !important; }
    </style>
""", unsafe_allow_html=True)

st.title("🛡️ 财务智能凭证转换中台 (旗舰版)")

# --- 1. 核心存储逻辑 ---
if 'standard_rules' not in st.session_state:
    st.session_state.standard_rules = pd.DataFrame([
        {"业务场景": "日常报销", "摘要": "报销-{单位}-{备注}", "方向": "借", "科目编码": "660201", "科目名称": "管理费用", "金额公式": "100%"},
        {"业务场景": "日常报销", "摘要": "付现金", "方向": "贷", "科目编码": "1001", "科目名称": "库存现金", "金额公式": "100%"},
        {"业务场景": "销售回款", "摘要": "收到-{单位}-货款", "方向": "借", "科目编码": "1002", "科目名称": "银行存款", "金额公式": "100%"},
        {"业务场景": "销售回款", "摘要": "核销-{单位}-应收", "方向": "贷", "科目编码": "1122", "科目名称": "应收账款", "金额公式": "100%"}
    ])

# --- 2. 交互界面 ---
t1, t2, t3 = st.tabs(["📋 1. 预设标准规则库", "📥 2. 流水导入与校验", "📊 3. 结果预览与导出"])

# --- TAB 1: 规则定义 (DIY 的灵魂) ---
with t1:
    st.subheader("⚙️ 业务场景分录模板")
    st.info("💡 这里定义您的『财务逻辑』。您可以随意增加行、删除行、改列名。支持 {单位}、{备注} 等变量。")
    # 完全自由的行列编辑，模仿 Excel
    st.session_state.standard_rules = st.data_editor(
        st.session_state.standard_rules,
        num_rows="dynamic",
        use_container_width=True,
        key="rule_editor"
    )

# --- TAB 2: 数据处理 (最牛的校验逻辑) ---
with t2:
    st.subheader("📑 数据导入")
    c1, c2 = st.columns([1, 3])
    with c1:
        st.write("**导入规范：**")
        st.caption("1. 必须包含业务场景列\n2. 必须包含金额列\n3. 建议包含单位/备注列")
        file = st.file_uploader("选择您的流水账 Excel", type=['xlsx'])
    
    if file:
        df_raw = pd.read_excel(file).fillna("")
        cols = df_raw.columns.tolist()
        
        with c2:
            st.write("**字段映射：**")
            m1, m2, m3, m4 = st.columns(4)
            with m1: map_date = st.selectbox("日期列", cols)
            with m2: map_biz = st.selectbox("业务场景列", cols)
            with m3: map_amt = st.selectbox("金额列", cols)
            with m4: map_unit = st.selectbox("单位/摘要补充列", cols)

        if st.button("🚀 运行智能校验并生成", type="primary"):
            results = []
            diag_errors = []
            voucher_no = 1
            
            rules = st.session_state.standard_rules
            
            for i, row in df_raw.iterrows():
                biz_key = str(row[map_biz]).strip()
                # 核心匹配逻辑
                matches = rules[rules["业务场景"] == biz_key]
                
                if matches.empty:
                    diag_errors.append({"原始行号": i+2, "业务关键词": biz_key, "失败原因": "规则库中未定义此业务", "解决方案": "请在Tab 1补充规则"})
                    continue
                
                # 严格按照一笔业务一个凭证号
                v_code = f"记-{voucher_no:03d}"
                for _, r in matches.iterrows():
                    # 变量动态替换
                    final_memo = str(r["摘要"]).replace("{单位}", str(row[map_unit])).replace("{备注}", str(row[map_unit]))
                    
                    results.append({
                        "凭证号": v_code,
                        "日期": str(row[map_date]).split(" ")[0],
                        "摘要": final_memo,
                        "科目编码": r["科目编码"],
                        "科目名称": r["科目名称"],
                        "借方金额": row[map_amt] if r["方向"] == "借" else 0,
                        "贷方金额": row[map_amt] if r["方向"] == "贷" else 0,
                    })
                voucher_no += 1
            
            if diag_errors:
                st.error("❌ 校验失败：发现未定义的业务，请修正后重试。")
                st.table(pd.DataFrame(diag_errors))
            
            if results:
                st.session_state.final_vouchers = pd.DataFrame(results)
                st.success(f"🎉 转换成功！共生成 {voucher_no-1} 张凭证。请前往 Tab 3 查看。")

# --- TAB 3: 预览与导出 ---
with t3:
    if 'final_vouchers' in st.session_state:
        st.subheader("👁️ 凭证预览")
        st.write("这是根据您的规则生成的标准分录：")
        st.data_editor(st.session_state.final_vouchers, use_container_width=True)
        
        # 导出 Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            st.session_state.final_vouchers.to_excel(writer, index=False)
        
        st.download_button(
            "📥 导出为财务标准 Excel",
            data=output.getvalue(),
            file_name="凭证结果.xlsx",
            mime="application/vnd.ms-excel",
            type="primary"
        )
    else:
        st.warning("☕ 请先完成前两个步骤的操作。")
