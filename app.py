import streamlit as st
import pandas as pd
import io

# ==========================================
# 0. 全局配置与样式
# ==========================================
st.set_page_config(layout="wide", page_title="财务凭证智能转换引擎 Pro")

st.markdown("""
<style>
    .big-font { font-size:20px !important; font-weight: bold; color: #2c3e50; }
    .success-box { padding: 10px; background-color: #d4edda; color: #155724; border-radius: 5px; }
    .error-box { padding: 10px; background-color: #f8d7da; color: #721c24; border-radius: 5px; }
</style>
""", unsafe_allow_html=True)

st.title("🚀 财务凭证智能转换引擎 (商业旗舰版)")

# ==========================================
# 1. 核心逻辑：规则引擎初始化
# ==========================================
def get_default_templates():
    """定义标准商业软件的预设模板"""
    return pd.DataFrame([
        # 场景1：一借一贷
        {"业务场景": "收汇", "摘要模板": "收{往来单位}货款", "科目编码": "1002", "方向": "借", "取值公式": "金额", "辅助核算": ""},
        {"业务场景": "收汇", "摘要模板": "核销应收账款", "科目编码": "1122", "方向": "贷", "取值公式": "金额", "辅助核算": "{往来单位}"},
        
        # 场景2：一借多贷 (含税)
        {"业务场景": "销售开票", "摘要模板": "销售给{往来单位}", "科目编码": "1122", "方向": "借", "取值公式": "金额", "辅助核算": "{往来单位}"},
        {"业务场景": "销售开票", "摘要模板": "确认收入", "科目编码": "6001", "方向": "贷", "取值公式": "金额/1.13", "辅助核算": ""}, # 模拟除税
        {"业务场景": "销售开票", "摘要模板": "计提销项税", "科目编码": "2221", "方向": "贷", "取值公式": "金额-金额/1.13", "辅助核算": ""},
    ])

if 'template_df' not in st.session_state:
    st.session_state.template_df = get_default_templates()

# ==========================================
# 2. 界面布局：分步式工作流
# ==========================================
tab1, tab2, tab3 = st.tabs(["🏗️ 1. 策略配置 (规则库)", "📥 2. 智能映射与诊断", "📊 3. 结果生成"])

# --- Tab 1: 规则配置 (最灵活的部分) ---
with tab1:
    st.markdown('<p class="big-font">核心策略库 (Mapping Rules)</p>', unsafe_allow_html=True)
    st.info("""
    **💡 高级用法说明：**
    1. **业务场景**：这是匹配 Excel 的唯一钥匙。
    2. **变量注入**：使用 `{列名}` 代表 Excel 中的那一列数据。例如 `{往来单位}`。
    3. **多行分录**：同一个“业务场景”可以写多行，生成时会自动组合在一起。
    4. **取值公式**：支持简单计算，如 `金额` 或 `金额*0.06`。
    """)
    
    # 允许用户像 Excel 一样操作规则库
    st.session_state.template_df = st.data_editor(
        st.session_state.template_df,
        num_rows="dynamic",
        use_container_width=True,
        height=400
    )

# --- Tab 2: 导入与映射 (商业软件的 Mapping 逻辑) ---
with tab2:
    st.markdown('<p class="big-font">数据源导入</p>', unsafe_allow_html=True)
    
    uploaded_file = st.file_uploader("上传您的业务流水 Excel", type=['xlsx', 'xls'])
    
    if uploaded_file:
        df_raw = pd.read_excel(uploaded_file).fillna("")
        st.write("✅ **数据预览 (前3行)：**")
        st.dataframe(df_raw.head(3), use_container_width=True)
        
        raw_cols = df_raw.columns.tolist()
        
        st.markdown("---")
        st.markdown("#### 🔗 字段映射 (Mapping)")
        st.caption("请告诉系统，您的 Excel 列分别代表什么意义？")
        
        c1, c2, c3, c4 = st.columns(4)
        with c1: map_biz = st.selectbox("【业务场景】对应哪一列？", raw_cols, index=0, help="用于去规则库匹配")
        with c2: map_date = st.selectbox("【日期】对应哪一列？", raw_cols, index=min(1, len(raw_cols)-1))
        with c3: map_amt = st.selectbox("【金额】对应哪一列？", raw_cols, index=min(2, len(raw_cols)-1))
        with c4: map_unit = st.selectbox("【往来/辅助】对应哪一列？", raw_cols, index=min(3, len(raw_cols)-1))

        # 存入 session 供下一步使用
        st.session_state.mapping = {
            "biz": map_biz, "date": map_date, "amt": map_amt, "unit": map_unit, "raw_data": df_raw
        }
        st.success("映射完成！请前往【步骤3】生成凭证。")

# --- Tab 3: 生成引擎 (核心计算) ---
with tab3:
    st.markdown('<p class="big-font">凭证生成控制台</p>', unsafe_allow_html=True)
    
    if 'mapping' not in st.session_state or 'raw_data' not in st.session_state.mapping:
        st.warning("⚠️ 请先在【步骤2】上传数据并完成映射。")
    else:
        if st.button("🚀 启动转换引擎", type="primary"):
            df = st.session_state.mapping['raw_data']
            mapping = st.session_state.mapping
            rules = st.session_state.template_df
            
            results = []
            errors = []
            voucher_id = 1 # 凭证号计数器
            
            # --- 核心循环逻辑 ---
            for idx, row in df.iterrows():
                # 1. 提取当前行的业务关键词
                biz_key = str(row[mapping['biz']]).strip()
                
                # 2. 在规则库中查找所有匹配的行 (Filter)
                # 这一步实现了“一个业务 -> 多行分录”
                matched_rules = rules[rules["业务场景"] == biz_key]
                
                # 3. 失败拦截 (Fail-Fast)
                if matched_rules.empty:
                    errors.append({
                        "原始行号": idx + 2,
                        "业务内容": biz_key,
                        "错误类型": "❌ 未定义规则",
                        "解决办法": "请在 Tab 1 添加该业务场景的规则"
                    })
                    continue # 跳过此行
                
                # 4. 成功匹配，开始生成分录
                current_voucher_no = f"{voucher_id:04d}" # 0001
                
                for _, r in matched_rules.iterrows():
                    # --- A. 智能变量替换 (Abstract Parsing) ---
                    # 摘要处理：将 {往来单位} 替换为 Excel 里的实际值
                    memo = str(r["摘要模板"])
                    # 支持替换用户指定的任何列
                    for col_name in df.columns:
                        if f"{{{col_name}}}" in memo:
                            memo = memo.replace(f"{{{col_name}}}", str(row[col_name]))
                    # 默认替换映射的列
                    memo = memo.replace("{往来单位}", str(row[mapping['unit']]))

                    # --- B. 辅助核算处理 ---
                    aux = str(r["辅助核算"])
                    if "{往来单位}" in aux:
                        aux = str(row[mapping['unit']])
                    
                    # --- C. 金额计算 (Formula Eval) ---
                    # 这是一个高级功能：支持简单的加减乘除
                    try:
                        base_amt = float(row[mapping['amt']])
                        # 简单的 eval 安全处理，允许使用 '金额' 这个词
                        calc_formula = str(r["取值公式"]).replace("金额", str(base_amt))
                        final_amt = eval(calc_formula)
                    except:
                        final_amt = 0
                        
                    # --- D. 组装结果 ---
                    results.append({
                        "凭证号": current_voucher_no,
                        "日期": str(row[mapping['date']]).split(" ")[0],
                        "摘要": memo,
                        "科目编码": r["科目编码"],
                        "借方": round(final_amt, 2) if r["方向"] == "借" else 0,
                        "贷方": round(final_amt, 2) if r["方向"] == "贷" else 0,
                        "辅助核算": aux
                    })
                
                # 处理完一行 Excel，凭证号 +1
                voucher_id += 1
            
            # --- 结果展示 ---
            if errors:
                st.markdown("### 🚫 诊断报告 (生成失败)")
                st.dataframe(pd.DataFrame(errors), use_container_width=True)
                st.error(f"警告：有 {len(errors)} 笔业务无法生成，请修正规则库后重试。")
            
            if results:
                st.markdown("### ✅ 生成结果预览")
                final_df = pd.DataFrame(results)
                
                # 提供最终编辑能力
                final_df = st.data_editor(final_df, height=500, use_container_width=True)
                
                st.success(f"成功生成 {len(df)-len(errors)} 张凭证，共 {len(results)} 行分录。")
                
                # 导出
                out = io.BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                    final_df.to_excel(writer, index=False)
                
                st.download_button(
                    label="📥 导出最终凭证文件 (Excel)",
                    data=out.getvalue(),
                    file_name="Vouchers_Final.xlsx",
                    mime="application/vnd.ms-excel",
                    type="primary"
                )
