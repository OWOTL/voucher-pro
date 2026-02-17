import streamlit as st
import pandas as pd
import io

st.set_page_config(layout="wide", page_title="通用凭证DIY工厂")

st.title("🏗️ 通用凭证DIY工厂 (支持发货/流水/多种场景)")

# --- 1. 初始化规则表 (如果没有则建立默认结构) ---
if 'rule_df' not in st.session_state:
    # 默认列名，用户可以在界面上改
    st.session_state.rule_df = pd.DataFrame([
        {"业务关键词": "发货", "借贷": "借", "科目编码": "1122", "摘要": "发货给{客户}", "关联列": "客户"},
        {"业务关键词": "发货", "借贷": "贷", "科目编码": "6001", "摘要": "销售商品", "关联列": ""},
        {"业务关键词": "收汇", "借贷": "借", "科目编码": "1002", "摘要": "收到{单位}货款", "关联列": "单位"},
        {"业务关键词": "收汇", "借贷": "贷", "科目编码": "1122", "摘要": "核销", "关联列": ""}
    ])

# --- 侧边栏：映射中心 (防止DIY标题后崩溃的关键) ---
with st.sidebar:
    st.header("⚙️ 规则列映射")
    st.info("如果你修改了【规则设置】里的标题，请在这里重新对应。")
    r_cols = st.session_state.rule_df.columns.tolist()
    m_biz = st.selectbox("哪一列代表【业务类型/关键词】?", r_cols, index=0)
    m_dir = st.selectbox("哪一列代表【借/贷方向】?", r_cols, index=1)
    m_code = st.selectbox("哪一列代表【科目编码】?", r_cols, index=2)
    m_memo = st.selectbox("哪一列代表【摘要内容】?", r_cols, index=3)
    m_aux = st.selectbox("哪一列代表【辅助项关联列名】?", r_cols, index=4)

    st.divider()
    st.header("📊 流水/发货表列映射")
    st.caption("上传文件后，请在这里选择流水表对应的列")

# --- 主界面标签页 ---
tab1, tab2, tab3 = st.tabs(["🎯 1. 设置分录规则 (DIY标题)", "文件上传与生成", "👀 预览、微调与导出"])

# --- TAB 1: DIY 规则 ---
with tab1:
    st.subheader("定义你的业务逻辑")
    st.markdown("""
    - **DIY标题**：双击表头即可修改标题（如把'业务关键词'改成'场景'）。
    - **增删内容**：右键点击表格可以增加或删除行。
    - **注意**：改完标题后，请确保**左侧侧边栏**的映射关系依然正确。
    """)
    st.session_state.rule_df = st.data_editor(st.session_state.rule_df, num_rows="dynamic", use_container_width=True)

# --- TAB 2: 上传与生成 ---
with tab2:
    f = st.file_uploader("上传您的原始数据 (Excel)", type=['xlsx'])
    if f:
        df_src = pd.read_excel(f).fillna("")
        src_cols = df_src.columns.tolist()
        
        with st.sidebar:
            d_biz = st.selectbox("原始表中：哪一列是业务类型?", src_cols)
            d_amt = st.selectbox("原始表中：哪一列是金额?", src_cols)
            d_date = st.selectbox("原始表中：哪一列是日期?", src_cols)

        if st.button("🚀 开始生成凭证号及分录"):
            final_list = []
            for i, row in df_src.iterrows():
                # 匹配规则
                biz_key = str(row[d_biz])
                matches = st.session_state.rule_df[st.session_state.rule_df[m_biz] == biz_key]
                
                # 凭证号逻辑：每行原始数据生成一个独立编号 (如 001, 002)
                v_no = str(i + 1).zfill(3)
                
                for _, r in matches.iterrows():
                    # 摘要处理
                    memo = str(r[m_memo])
                    for c in src_cols:
                        if f"{{{c}}}" in memo: memo = memo.replace(f"{{{c}}}", str(row[c]))
                    
                    # 辅助项处理
                    aux_col = r[m_aux]
                    aux_val = row[aux_col] if aux_col in src_cols else ""

                    final_list.append({
                        "凭证类别": "记",
                        "凭证号": v_no,
                        "凭证日期": str(row[d_date]).split(" ")[0],
                        "摘要": memo,
                        "科目编码": r[m_code],
                        "借方金额": row[d_amt] if r[m_dir] == "借" else 0,
                        "贷方金额": row[d_amt] if r[m_dir] == "贷" else 0,
                        "辅助核算": aux_val
                    })
            
            if final_list:
                st.session_state.preview_df = pd.DataFrame(final_list)
                st.success("✅ 生成成功！请点击上方‘预览、微调与导出’标签页。")
            else:
                st.error("❌ 未匹配到任何规则，请检查业务关键词是否一致。")

# --- TAB 3: 预览与导出 ---
with tab3:
    if 'preview_df' in st.session_state:
        st.subheader("最终凭证预览")
        st.info("可以在下方直接修改凭证号、金额、摘要。满意后点下载。")
        
        # 用户在此处进行最后的DIY微调
        final_edited = st.data_editor(st.session_state.preview_df, num_rows="dynamic", use_container_width=True, height=500)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            final_edited.to_excel(writer, index=False)
        
        st.divider()
        st.download_button("📥 确认无误，导出好会计文件", data=output.getvalue(), file_name="好会计导入包.xlsx")