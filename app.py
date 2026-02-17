import streamlit as st
import pandas as pd
import io

st.set_page_config(layout="wide", page_title="全维度财务凭证工厂")

st.title("🏗️ 全维度财务凭证工厂 (列/行全自由DIY版)")

# --- 1. 初始化逻辑：规则表现在完全是空的，由你来定义列 ---
if 'rule_df' not in st.session_state:
    # 初始只给最基础的 5 列作为模板
    st.session_state.rule_df = pd.DataFrame([
        {"业务关键词": "发货", "方向": "借", "科目编码": "1122", "摘要": "发货给{客户}", "辅助项": "{客户}"},
        {"业务关键词": "发货", "方向": "贷", "科目编码": "6001", "摘要": "确认收入", "辅助项": ""},
    ])

# --- 2. 侧边栏：映射中心 (无论你把列改成什么名字，在这里指定一下) ---
with st.sidebar:
    st.header("⚙️ 1. 定义核心功能列")
    st.info("请在右侧表格修改/增加列，完成后在此处指定对应关系：")
    
    current_rule_cols = st.session_state.rule_df.columns.tolist()
    
    # 动态映射：让系统知道哪一列负责什么
    map_biz = st.selectbox("哪一列代表【业务关键词】？", current_rule_cols, index=0)
    map_dir = st.selectbox("哪一列代表【借/贷方向】？", current_rule_cols, index=1)
    map_code = st.selectbox("哪一列代表【科目编码】？", current_rule_cols, index=2)
    map_memo = st.selectbox("哪一列代表【摘要内容】？", current_rule_cols, index=3)
    map_aux = st.selectbox("哪一列代表【辅助核算】？", current_rule_cols, index=4)

    st.divider()
    st.header("📂 2. 原始数据映射")
    st.caption("上传流水后，匹配流水表的关键信息")

# --- 3. 标签页设计 ---
tab1, tab2, tab3 = st.tabs(["📋 1. 自由定义规则 (增删列/行)", "📥 2. 诊断导入与生成", "👁️ 3. 预览、微调与导出"])

# --- TAB 1: 真正的自由 DIY ---
with tab1:
    st.subheader("🛠️ 规则工厂")
    st.markdown("""
    **在这里，你可以像操作 Excel 一样：**
    - **增加/删除行**：点击表格下方的 `+` 或选中行按 `Delete`。
    - **增加/删除列**：右键点击表头，选择 **`Insert column`** 或 **`Delete column`**。
    - **改列名**：双击表头即可直接修改标题。
    - **重要**：改完后，请务必在**左侧侧边栏**检查映射是否正确。
    """)
    
    # 开启全功能编辑
    st.session_state.rule_df = st.data_editor(
        st.session_state.rule_df, 
        num_rows="dynamic", # 动态行
        use_container_width=True,
        column_config={col: st.column_config.TextColumn(col) for col in st.session_state.rule_df.columns} 
    )

# --- TAB 2: 数据导入与深度诊断 ---
with tab2:
    st.subheader("上传业务数据")
    f = st.file_uploader("请上传 Excel (.xlsx)", type=['xlsx'])
    
    if f:
        df_raw = pd.read_excel(f).fillna("")
        raw_cols = df_raw.columns.tolist()
        
        # 让用户映射流水表的关键列
        c1, c2, c3 = st.columns(3)
        with c1: d_date = st.selectbox("流水：日期列", raw_cols)
        with c2: d_biz = st.selectbox("流水：业务关键词列", raw_cols)
        with c3: d_amt = st.selectbox("流水：金额列", raw_cols)
            
        if st.button("🚀 开始生成并诊断"):
            diags = []
            results = []
            
            for i, row in df_raw.iterrows():
                biz_key = str(row[d_biz]).strip()
                # 在规则表里找匹配的所有分录
                matched = st.session_state.rule_df[st.session_state.rule_df[map_biz] == biz_key]
                
                if matched.empty:
                    diags.append({"行号": i+2, "内容": biz_key, "状态": "❌ 缺失规则", "对策": "请在规则页新增此业务"})
                    continue
                
                # 凭证号逻辑：流水表的一行对应一个号
                v_no = str(i + 1).zfill(4)
                
                for _, r in matched.iterrows():
                    # 摘要与辅助项支持多字段动态替换
                    def smart_replace(text):
                        text = str(text)
                        for col in raw_cols:
                            if f"{{{col}}}" in text:
                                text = text.replace(f"{{{col}}}", str(row[col]))
                        return text

                    results.append({
                        "凭证号": v_no,
                        "日期": str(row[d_date]).split(" ")[0],
                        "摘要": smart_replace(r[map_memo]),
                        "科目编码": r[map_code],
                        "借方": row[d_amt] if r[map_dir] == "借" else 0,
                        "贷方": row[d_amt] if r[map_dir] == "贷" else 0,
                        "辅助核算": smart_replace(r[map_aux])
                    })
            
            if diags:
                st.error("诊断报告：发现无法生成的行")
                st.table(pd.DataFrame(diags))
            
            if results:
                st.session_state.final_vouchers = pd.DataFrame(results)
                st.success(f"✅ 生成完毕！成功转换 {len(df_raw)-len(diags)} 笔业务。")

# --- TAB 3: 预览与微调 ---
with tab3:
    if 'final_vouchers' in st.session_state:
        st.subheader("结果预览 (所见即所得)")
        st.info("💡 提示：你可以直接在这里修改凭证号实现『合单』或『分单』，或者修改任意摘要和金额。")
        
        # 结果编辑器同样支持动态修改
        edited_df = st.data_editor(st.session_state.final_vouchers, num_rows="dynamic", use_container_width=True, height=500)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            edited_df.to_excel(writer, index=False)
        
        st.download_button("📥 导出修改后的文件", data=output.getvalue(), file_name="好会计导入文件.xlsx")
