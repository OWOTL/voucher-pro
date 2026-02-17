import streamlit as st
import pandas as pd
import io

st.set_page_config(layout="wide", page_title="骨灰级财务DIY生成器")

st.title("🏗️ 骨灰级财务凭证生成器 (完全自定义列版)")

# --- 1. 这里是你的“司令部”：想加减列，直接在这里改文字 ---
# 修改这里的列表，网页上的表格列会立刻同步增减
if 'columns_list' not in st.session_state:
    st.session_state.columns_list = ["业务关键词", "方向", "科目编码", "科目名称", "摘要模板", "辅助单位", "项目信息"]

# --- 2. 侧边栏：让系统知道哪一列是干嘛的 ---
with st.sidebar:
    st.header("🛠️ 1. 指定功能列")
    st.info("如果修改了右侧的表头，请在这里对齐关系。")
    
    # 动态获取当前所有列
    cols = st.session_state.columns_list
    
    m_biz = st.selectbox("哪列是【业务关键词】？", cols, index=0)
    m_dir = st.selectbox("哪列是【方向(借/贷)】？", cols, index=1)
    m_code = st.selectbox("哪列是【科目编码】？", cols, index=2)
    m_memo = st.selectbox("哪列是【摘要模板】？", cols, index=4)
    
    st.divider()
    st.header("📋 2. 导入与诊断说明")
    st.write("- 流水表必须有：日期、业务名、金额。")
    st.write("- 摘要模板支持 `{列名}` 自动抓取。")

# --- 3. 初始化规则数据 ---
if 'rules_data' not in st.session_state:
    # 根据上面的初始列创建一个默认行
    init_row = {c: "" for c in cols}
    init_row[cols[0]] = "发货"
    init_row[cols[1]] = "借"
    st.session_state.rules_data = pd.DataFrame([init_row])

# 检查列是否同步
if list(st.session_state.rules_data.columns) != cols:
    # 如果列变了，重新调整 DataFrame 结构但不丢失数据
    st.session_state.rules_data = st.session_state.rules_data.reindex(columns=cols).fillna("")

# --- 4. 标签页 ---
tab1, tab2, tab3 = st.tabs(["⚙️ 规则设置 (行/列管理)", "📥 数据导入与智能诊断", "👁️ 预览、修改与导出"])

with tab1:
    st.subheader("定义分录逻辑")
    # 这里我们用文本框让用户“增加/删除列名”
    new_cols_str = st.text_area("🔧 修改列名（用中文逗号或英文逗号隔开）：", value=",".join(cols))
    if st.button("💾 更新列结构"):
        st.session_state.columns_list = [c.strip() for c in new_cols_str.replace("，", ",").split(",")]
        st.rerun()

    st.markdown("---")
    # 动态表格编辑
    st.session_state.rules_data = st.data_editor(
        st.session_state.rules_data, 
        num_rows="dynamic", 
        use_container_width=True
    )

with tab2:
    f = st.file_uploader("上传业务 Excel", type=['xlsx'])
    if f:
        df_raw = pd.read_excel(f).fillna("")
        raw_cols = df_raw.columns.tolist()
        
        # 映射流水列
        c1, c2, c3 = st.columns(3)
        with c1: d_date = st.selectbox("流水：日期列", raw_cols)
        with c2: d_biz = st.selectbox("流水：业务类型列", raw_cols)
        with c3: d_amt = st.selectbox("流水：金额列", raw_cols)
            
        if st.button("🚀 运行诊断并生成结果"):
            results = []
            errors = []
            
            for i, row in df_raw.iterrows():
                biz_key = str(row[d_biz]).strip()
                matches = st.session_state.rules_data[st.session_state.rules_data[m_biz] == biz_key]
                
                if matches.empty:
                    errors.append({"行号": i+2, "业务名": biz_key, "原因": "规则库没定义这个业务"})
                    continue
                
                v_no = str(i + 1).zfill(4)
                
                for _, r in matches.iterrows():
                    # 自动抓取辅助核算：把规则表里除了核心功能列之外的所有列，都作为辅助项拼起来
                    # 摘要处理
                    memo = str(r[m_memo])
                    for rc in raw_cols:
                        if f"{{{rc}}}" in memo: memo = memo.replace(f"{{{rc}}}", str(row[rc]))
                    
                    # 动态抓取所有非核心列作为辅助信息
                    aux_info = []
                    for c in cols:
                        if c not in [m_biz, m_dir, m_code, m_memo]:
                            val = str(r[c])
                            # 如果规则里写了 {单位}，就去流水里抓
                            for rc in raw_cols:
                                if f"{{{rc}}}" in val: val = val.replace(f"{{{rc}}}", str(row[rc]))
                            if val: aux_info.append(f"{c}:{val}")

                    results.append({
                        "凭证号": v_no,
                        "日期": str(row[d_date]).split(" ")[0],
                        "摘要": memo,
                        "科目编码": r[m_code],
                        "借方": row[d_amt] if r[m_dir] == "借" else 0,
                        "贷方": row[d_amt] if r[m_dir] == "贷" else 0,
                        "辅助核算": " | ".join(aux_info)
                    })
            
            if errors:
                st.error("❌ 诊断报告：请修复以下业务规则")
                st.table(pd.DataFrame(errors))
            
            if results:
                st.session_state.final_res = pd.DataFrame(results)
                st.success("✅ 生成预览成功！")

with tab3:
    if 'final_res' in st.session_state:
        st.subheader("凭证修改与导出")
        final_edited = st.data_editor(st.session_state.final_res, num_rows="dynamic", use_container_width=True)
        
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
            final_edited.to_excel(writer, index=False)
        st.download_button("📥 点击下载导入文件", data=out.getvalue(), file_name="好会计凭证.xlsx")
