import streamlit as st
import pandas as pd
import io

st.set_page_config(layout="wide", page_title="全能财务凭证工厂")

st.title("🏗️ 全能财务凭证工厂 (多业务+自动单位版)")

# --- 1. 初始化规则库 (支持多业务类型) ---
if 'rules_df' not in st.session_state:
    st.session_state.rules_df = pd.DataFrame([
        {"业务类型": "收汇", "方向": "借", "科目编码": "1002", "摘要": "收到{单位}货款"},
        {"业务类型": "收汇", "方向": "贷", "科目编码": "1122", "摘要": "核销{单位}往来"},
        {"业务类型": "发货", "方向": "借", "科目编码": "6401", "摘要": "结转{单位}成本"},
        {"业务类型": "发货", "方向": "贷", "科目编码": "1405", "摘要": "库存商品出库"}
    ])

# --- 2. 标签页 ---
tab1, tab2, tab3 = st.tabs(["⚙️ 1. DIY 规则定义", "📥 2. 数据导入与诊断", "👁️ 3. 结果预览与导出"])

# --- TAB 1: 规则定义 ---
with tab1:
    st.subheader("🛠️ 定义您的业务逻辑库")
    st.markdown("""
    **操作指南：**
    1. **多业务处理**：在【业务类型】列填入不同的名字（如：收汇、发货、计提、调账）。
    2. **自动单位抓取**：在【摘要】中写 `{单位}`，生成时会自动替换为流水里的公司名。
    3. **增减自由**：点击表格下方 `+` 增加分录行。
    """)
    st.session_state.rules_df = st.data_editor(st.session_state.rules_df, num_rows="dynamic", use_container_width=True)

# --- TAB 2: 导入与诊断 ---
with tab2:
    col_l, col_r = st.columns([1, 2])
    
    with col_l:
        st.info("📋 **上传格式要求**")
        st.markdown("""
        您的 Excel 必须包含以下列：
        - **日期** (2024-01-01)
        - **往来单位** (公司全称)
        - **金额** (数字)
        - **业务名称** (需对应规则表)
        """)
        
    with col_r:
        f = st.file_uploader("点击或拖拽上传流水 Excel", type=['xlsx'])
        if f:
            df_raw = pd.read_excel(f).fillna("")
            raw_cols = df_raw.columns.tolist()
            
            st.write("🔍 **请关联流水表字段：**")
            c1, c2, c3, c4 = st.columns(4)
            with c1: d_date = st.selectbox("日期列", raw_cols)
            with c2: d_unit = st.selectbox("单位列", raw_cols)
            with c3: d_amt = st.selectbox("金额列", raw_cols)
            with c4: d_type = st.selectbox("业务名称列", raw_cols)
            
            if st.button("🚀 运行诊断并生成"):
                results = []
                errors = []
                
                for i, row in df_raw.iterrows():
                    biz_name = str(row[d_type]).strip()
                    # 匹配规则
                    matches = st.session_state.rules_df[st.session_state.rules_df['业务类型'] == biz_name]
                    
                    if matches.empty:
                        errors.append({"行号": i+2, "业务名": biz_name, "问题": "未定义分录规则", "建议": f"在Tab 1增加'{biz_name}'的规则"})
                        continue
                    
                    v_no = str(i + 1).zfill(4)
                    for _, r in matches.iterrows():
                        # 自动抓取单位并替换摘要
                        memo = str(r['摘要']).replace("{单位}", str(row[d_unit]))
                        
                        results.append({
                            "凭证号": v_no,
                            "日期": str(row[d_date]).split(" ")[0],
                            "摘要": memo,
                            "科目编码": r['科目编码'],
                            "借方": row[d_amt] if r['方向'] == "借" else 0,
                            "贷方": row[d_amt] if r['方向'] == "贷" else 0,
                            "辅助单位": row[d_unit]
                        })
                
                if errors:
                    st.error("❌ 发现错误，以下数据无法生成：")
                    st.table(pd.DataFrame(errors))
                
                if results:
                    st.session_state.final_res = pd.DataFrame(results)
                    st.success(f"✅ 成功生成 {len(df_raw)-len(errors)} 条业务的凭证分录！")

# --- TAB 3: 结果与导出 ---
with tab3:
    if 'final_res' in st.session_state:
        st.subheader("👁️ 结果预览与微调")
        st.info("提示：您可以直接在下方表格修改任何数字或文字，修改后导出的就是最终版本。")
        # 结果也支持编辑
        st.session_state.final_res = st.data_editor(st.session_state.final_res, num_rows="dynamic", use_container_width=True)
        
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
            st.session_state.final_res.to_excel(writer, index=False)
        st.download_button("📥 导出好会计/易代账格式文件", data=out.getvalue(), file_name="凭证结果.xlsx")
    else:
        st.warning("暂无数据，请先在 Tab 2 上传并生成。")
