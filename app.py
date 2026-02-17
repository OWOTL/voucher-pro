import streamlit as st
import pandas as pd
import io

st.set_page_config(layout="wide", page_title="全自由财务凭证生成器")

st.title("🏗️ 全自由财务凭证生成器 (列/行全自定义版)")

# --- 1. 初始化列名和数据 ---
if 'columns_list' not in st.session_state:
    # 默认初始列
    st.session_state.columns_list = ["业务关键词", "借贷方向", "科目编码", "摘要模板", "辅助单位"]

if 'rules_df' not in st.session_state:
    # 默认初始数据
    st.session_state.rules_df = pd.DataFrame([
        {"业务关键词": "收汇", "借贷方向": "借", "科目编码": "1002", "摘要模板": "收到{单位}货款", "辅助单位": "{单位}"},
        {"业务关键词": "收汇", "借贷方向": "贷", "科目编码": "1122", "摘要模板": "核销回款", "辅助单位": "{单位}"}
    ])

# --- 2. 侧边栏：核心功能映射 (防止改列名后崩溃) ---
with st.sidebar:
    st.header("⚙️ 1. 功能映射配置")
    st.info("如果你修改了右侧的表头名，请在这里重新指定对应关系：")
    
    current_cols = st.session_state.columns_list
    
    # 安全索引函数
    def get_idx(name, col_list):
        try: return col_list.index(name)
        except: return 0

    map_biz = st.selectbox("哪一列是【业务关键词】(匹配开关)？", current_cols, index=get_idx("业务关键词", current_cols))
    map_dir = st.selectbox("哪一列是【借/贷方向】？", current_cols, index=get_idx("借贷方向", current_cols))
    map_code = st.selectbox("哪一列是【科目编码】？", current_cols, index=get_idx("科目编码", current_cols))
    map_memo = st.selectbox("哪一列是【摘要内容】？", current_cols, index=get_idx("摘要模板", current_cols))

# --- 3. 页面布局 ---
tab1, tab2, tab3 = st.tabs(["🎯 1. DIY 规则与列管理", "📥 2. 上传流水并生成", "👁️ 3. 结果预览与导出"])

# --- TAB 1: 真正意义上的列/行自由修改 ---
with tab1:
    st.subheader("🛠️ 定义你的规则表结构")
    
    # 列名管理器
    new_cols_str = st.text_input("📝 在这里修改/增减表头（用逗号隔开）：", value=",".join(st.session_state.columns_list))
    if st.button("🔄 应用新列名结构"):
        updated_cols = [c.strip() for c in new_cols_str.replace("，", ",").split(",") if c.strip()]
