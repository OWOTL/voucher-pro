import streamlit as st
import pandas as pd
import io

# 设置网页宽屏
st.set_page_config(layout="wide", page_title="极简财务凭证工具")

st.title("📑 极简财务凭证工具 (像Excel一样自由)")

# --- 1. 核心逻辑：自动寻找功能列 ---
def find_col(df, keywords):
    """在用户乱序/改名后的表格里，凭直觉找到对应的功能列"""
    for col in df.columns:
        if any(key in str(col) for key in keywords):
            return col
    return None

# --- 2. 初始化规则表 (Tab 1) ---
if 'rules_data' not in st.session_state:
    # 初始给一个最标准的结构，后续用户随便怎么增删列都行
    st.session_state.rules_data = pd.DataFrame([
        {"业务关键词": "收汇", "借/贷": "借", "科目编码": "1002", "摘要": "收到{单位}货款"},
        {"业务关键词": "收汇", "借/贷": "贷", "科目编码": "1122", "摘要": "回款"},
    ])

# --- 3. 页面布局 ---
t1, t2 = st.tabs(["⚙️ 1. 定义你的规则 (直接操作下方表格)", "🚀 2. 上传流水并生成凭证"])

with t1:
    st.markdown("""
    ### 🛠️ 规则实验室
    - **增加/删除列**：在表头点击**鼠标右键**，选择 `Insert column` 或 `Delete column`。
    - **修改表头**：直接**双击表头**改名。
    - **自动抓取单位**：在摘要里写 `{单位}`。
    """)
    # 这一行开启了 Streamlit 最强大的自由编辑模式
    st.session_state.rules_data = st.data_editor(
        st.session_state.rules_data,
        num_rows="dynamic",
        column_config={}, # 留空以允许右键操作
        use_container_width=True,
    )

with t2:
    f = st.file_uploader("第一步：上传您的流水 Excel", type=['xlsx'])
    if f:
        df_raw = pd.read_excel(f).fillna("")
        
        # 自动探测流水的列
        raw_cols = df_raw.columns.tolist()
        st.write("🔍 **识别流水信息：**")
        c1, c2, c3, c4 = st.columns(4)
        with c1: d_date = st.selectbox("哪列是【日期】？", raw_cols, index=0)
        with c2: d_unit = st.selectbox("哪列是【单位/往来】？", raw_cols, index=min(1, len(raw_cols)-1))
        with c3: d_amt = st.selectbox("哪列是【金额】？", raw_cols, index=min(2, len(raw_cols)-1))
        with c4: d_biz = st.selectbox("哪列是【匹配业务的列】？", raw_cols, index=min(1, len(raw_cols)-1))

        if st.button("✨ 按照规则生成凭证"):
            rules = st.session_state.rules_data
            
            # 自动定位规则表里的关键列，不管用户怎么改名都能找
            col_biz = find_col(rules, ["业务", "关键词", "匹配"])
            col_dir = find_col(rules, ["方向", "借", "贷"])
            col_code = find_col(rules, ["科目", "编码", "代码"])
            col_memo = find_col(rules, ["摘要", "内容", "描述"])

            if not all([col_biz, col_dir, col_code, col_memo]):
                st.error(f"❌ 规则表里缺少必要列！请确保包含：业务关键词、方向、科目编码、摘要。当前识别到：{list(rules.columns)}")
            else:
                results = []
                for i, row in df_raw.iterrows():
                    biz_val = str(row[d_biz]).strip()
                    # 在规则表里找匹配项
                    matches = rules[rules[col_biz].apply(lambda x: str(x) in biz_val and str(x) != "")]
                    
                    if matches.empty:
                        st.warning(f"⚠️ 行 {i+2}: '{biz_val}' 没找到规则，已跳过。")
                        continue
                    
                    for _, r in matches.iterrows():
                        results.append({
                            "凭证号": str(i + 1).zfill(4),
                            "日期": str(row[d_date]).split(" ")[0],
                            "摘要": str(r[col_memo]).replace("{单位}", str(row[d_unit])),
                            "科目编码": r[col_code],
                            "借方": row[d_amt] if "借" in str(r[col_dir]) else 0,
                            "贷方": row[d_amt] if "贷" in str(r[col_dir]) else 0,
                            "辅助单位": row[d_unit]
                        })
                
                if results:
                    final_df = pd.DataFrame(results)
                    st.success("✅ 生成成功！")
                    st.data_editor(final_df, use_container_width=True)
                    
                    # 导出
                    out = io.BytesIO()
                    with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                        final_df.to_excel(writer, index=False)
                    st.download_button("📥 导出结果", data=out.getvalue(), file_name="凭证.xlsx")
