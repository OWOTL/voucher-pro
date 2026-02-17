import streamlit as st
import pandas as pd
import io

st.set_page_config(layout="wide", page_title="财务凭证终极工厂")

st.title("🏗️ 财务凭证终极工厂 (列/行全自由-防崩溃版)")

# --- 1. 动态列名初始化 ---
if 'columns_list' not in st.session_state:
    st.session_state.columns_list = ["业务关键词", "借贷方向", "科目编码", "摘要模板", "辅助核算1"]

# --- 2. 侧边栏：映射逻辑 (加入安全保护) ---
with st.sidebar:
    st.header("⚙️ 1. 功能映射")
    cols = st.session_state.columns_list
    
    # 使用try-except防止因索引不存在而崩溃
    def safe_index(target, column_list, default=0):
        try: return column_list.index(target)
        except: return default

    m_biz = st.selectbox("哪列是【业务关键词】？", cols, index=safe_index("业务关键词", cols))
    m_dir = st.selectbox("哪列是【借贷方向】？", cols, index=safe_index("借贷方向", cols))
    m_code = st.selectbox("哪列是【科目编码】？", cols, index=safe_index("科目编码", cols))
    m_memo = st.selectbox("哪列是【摘要模板】？", cols, index=safe_index("摘要模板", cols))

    st.divider()
    st.header("📂 2. 导入与诊断")
    st.caption("摘要模板支持抓取流水列：使用 {列名}")

# --- 3. 初始化/同步数据 ---
if 'rules_df' not in st.session_state:
    st.session_state.rules_df = pd.DataFrame([{"业务关键词": "发货", "借贷方向": "借", "科目编码": "1122", "摘要模板": "发货-{客户}", "辅助核算1": "{客户}"}])

# 自动处理列增减
if list(st.session_state.rules_df.columns) != cols:
    st.session_state.rules_df = st.session_state.rules_df.reindex(columns=cols).fillna("")

# --- 4. 界面布局 ---
tab1, tab2, tab3 = st.tabs(["🎯 规则全自定义 (增删列/行)", "📥 数据导入与诊断", "👁️ 预览与导出"])

with tab1:
    st.subheader("🛠️ 定义你的业务逻辑库")
    # 增加一个明显的列管理器
    col_input = st.text_input("🔧 在此修改表头列名（用中文或英文逗号隔开，改完按回车）：", value=",".join(cols))
    if st.button("🔄 应用新表头结构"):
        new_cols = [c.strip() for c in col_input.replace("，", ",").split(",") if c.strip()]
        if new_cols:
            st.session_state.columns_list = new_cols
            st.rerun()

    st.info("💡 操作指南：下方表格直接双击改内容，点击下方 + 增加行。")
    st.session_state.rules_df = st.data_editor(st.session_state.rules_df, num_rows="dynamic", use_container_width=True)

with tab2:
    f = st.file_uploader("上传您的原始流水 Excel", type=['xlsx'])
    if f:
        df_raw = pd.read_excel(f).fillna("")
        raw_cols = df_raw.columns.tolist()
        
        st.write("🔍 **请选择流水表对应的关键列：**")
        c1, c2, c3 = st.columns(3)
        with c1: d_date = st.selectbox("流水-日期列", raw_cols)
        with c2: d_biz = st.selectbox("流水-业务关键词列", raw_cols)
        with c3: d_amt = st.selectbox("流水-金额列", raw_cols)
            
        if st.button("🚀 开始生成凭证"):
            results = []
            errors = []
            
            for i, row in df_raw.iterrows():
                biz_key = str(row[d_biz]).strip()
                matches = st.session_state.rules_df[st.session_state.rules_df[m_biz] == biz_key]
                
                if matches.empty:
                    errors.append({"行号": i+2, "原始业务名": biz_key, "原因": "规则库没定义"})
                    continue
                
                v_no = str(i + 1).zfill(4)
                for _, r in matches.iterrows():
                    # 智能解析函数
                    def parse_val(text):
                        text = str(text)
                        for rc in raw_cols:
                            if f"{{{rc}}}" in text: text = text.replace(f"{{{rc}}}", str(row[rc]))
                        return text

                    # 自动搜集所有辅助列
                    aux_parts = []
                    for c in cols:
                        if c not in [m_biz, m_dir, m_code, m_memo] and r[c]:
                            aux_parts.append(f"{c}:{parse_val(r[c])}")

                    results.append({
                        "凭证号": v_no,
                        "日期": str(row[d_date]).split(" ")[0],
                        "摘要": parse_val(r[m_memo]),
                        "科目编码": r[m_code],
                        "借方": row[d_amt] if r[m_dir] in ["借", "J"] else 0,
                        "贷方": row[d_amt] if r[m_dir] in ["贷", "D"] else 0,
                        "辅助核算汇总": " / ".join(aux_parts)
                    })
            
            if errors:
                st.error("❌ 诊断到以下业务未配置规则，请在 Tab 1 补充：")
                st.table(pd.DataFrame(errors))
            
            if results:
                st.session_state.final_res = pd.DataFrame(results)
                st.success(f"✅ 生成完毕，共处理 {len(df_raw)-len(errors)} 笔业务。")

with tab3:
    if 'final_res' in st.session_state:
        st.subheader("👀 结果预览 (满意后导出)")
        st.data_editor(st.session_state.final_res, num_rows="dynamic", use_container_width=True)
        
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
            st.session_state.final_res.to_excel(writer, index=False)
        st.download_button("📥 导出好会计导入文件", data=out.getvalue(), file_name="凭证结果.xlsx")
