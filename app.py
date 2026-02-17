import streamlit as st
import pandas as pd
import io

st.set_page_config(layout="wide", page_title="智能财务凭证工厂")

st.title("🧮 智能财务凭证工厂 (全动态DIY版)")

# --- 1. 侧边栏：标题 DIY 控制区 ---
with st.sidebar:
    st.header("🎨 1. 自定义规则标题")
    st.info("在这里改名，右侧【规则设置】的表头会实时变化。")
    t_biz = st.text_input("【业务类型】列名", "业务关键词")
    t_dir = st.text_input("【借贷方向】列名", "借/贷")
    t_code = st.text_input("【科目编码】列名", "科目号")
    t_memo = st.text_input("【摘要模板】列名", "摘要内容")
    t_aux = st.text_input("【辅助核算关联列】列名", "关联流水列")

    st.divider()
    st.header("📖 2. 导入规范")
    st.warning("上传的 Excel 必须包含：\n- **日期** (如: 2023/01/01)\n- **金额** (数字)\n- **业务关键词** (需匹配规则)")

# --- 2. 初始化规则表 (全动态行数) ---
if 'rules' not in st.session_state:
    st.session_state.rules = pd.DataFrame([
        {t_biz: "发货", t_dir: "借", t_code: "1122", t_memo: "发货-{客户}", t_aux: "客户"},
        {t_biz: "发货", t_dir: "贷", t_code: "6001", t_memo: "销售收入", t_aux: ""},
        {t_biz: "银行收汇", t_dir: "借", t_code: "1002", t_memo: "收到货款", t_aux: ""},
        {t_biz: "银行收汇", t_dir: "贷", t_code: "1122", t_memo: "核销-{客户}", t_aux: "客户"},
    ])

# 同步表头
if list(st.session_state.rules.columns) != [t_biz, t_dir, t_code, t_memo, t_aux]:
    st.session_state.rules.columns = [t_biz, t_dir, t_code, t_memo, t_aux]

# --- 3. 页面布局 ---
tab1, tab2, tab3 = st.tabs(["⚙️ 规则设置 (DIY)", "📥 数据导入与诊断", "👁️ 预览、修改与导出"])

# --- TAB 1: 规则设置 (行数可增删) ---
with tab1:
    st.subheader("🛠️ 自定义分录逻辑库")
    st.markdown("""
    **操作指南：**
    - **修改标题**：在左侧侧边栏输入新名字。
    - **增加规则行**：点击表格下方的 **`+` (Add row)** 按钮。
    - **删除规则行**：选中行前面的复选框，按键盘 **Delete** 键。
    - **修改内容**：直接双击单元格。
    """)
    # 开启 num_rows="dynamic" 允许用户无限增加科目种类
    st.session_state.rules = st.data_editor(
        st.session_state.rules, 
        num_rows="dynamic", 
        use_container_width=True,
        key="rule_editor"
    )

# --- TAB 2: 数据导入与智能诊断 ---
with tab2:
    st.subheader("数据导入与合法性检查")
    f = st.file_uploader("请上传您的业务流水 Excel", type=['xlsx'])
    
    if f:
        df_raw = pd.read_excel(f).fillna("")
        raw_cols = df_raw.columns.tolist()
        
        # 字段映射
        st.write("🔍 **请告诉系统数据在哪一列：**")
        c1, c2, c3 = st.columns(3)
        with c1: m_date = st.selectbox("哪列是【日期】？", raw_cols)
        with c2: m_biz = st.selectbox("哪列是【业务类型/关键词】？", raw_cols)
        with c3: m_amt = st.selectbox("哪列是【金额】？", raw_cols)
            
        if st.button("🚀 运行诊断并生成预览"):
            error_log = []
            vouchers = []
            
            for i, row in df_raw.iterrows():
                biz_val = str(row[m_biz]).strip()
                # 寻找匹配
                matches = st.session_state.rules[st.session_state.rules[t_biz] == biz_val]
                
                if matches.empty:
                    error_log.append({
                        "原始行号": i + 2,
                        "业务内容": biz_val,
                        "错误原因": "在规则表中未找到对应的分录逻辑",
                        "解决办法": f"请在『规则设置』页面的【{t_biz}】列中增加一行“{biz_val}”"
                    })
                    continue
                
                # 自动生成凭证号
                v_no = str(i + 1).zfill(4)
                
                for _, r in matches.iterrows():
                    # 动态摘要替换
                    memo = str(r[t_memo])
                    for c in raw_cols:
                        if f"{{{c}}}" in memo: memo = memo.replace(f"{{{c}}}", str(row[c]))
                    
                    # 辅助项处理
                    aux_col = r[t_aux]
                    aux_val = row[aux_col] if aux_col in raw_cols else ""

                    vouchers.append({
                        "凭证号": v_no,
                        "日期": str(row[m_date]).split(" ")[0],
                        "摘要": memo,
                        "科目编码": r[t_code],
                        "借方金额": row[m_amt] if r[t_dir] == '借' else 0,
                        "贷方金额": row[m_amt] if r[t_dir] == '贷' else 0,
                        "辅助核算": aux_val
                    })
            
            # 显示诊断报告
            if error_log:
                st.error("❌ 诊断发现错误：部分数据无法匹配规则")
                st.table(pd.DataFrame(error_log))
                st.warning("⚠️ **请按照上表的解决办法操作，补全规则后再重新生成。**")
            
            if vouchers:
                st.session_state.final_vouchers = pd.DataFrame(vouchers)
                st.success(f"✅ 处理完成！共生成 {len(df_raw) - len(error_log)} 笔凭证。")

# --- TAB 3: 预览与导出 ---
with tab3:
    if 'final_vouchers' in st.session_state:
        st.subheader("👀 结果预览 (支持直接修改)")
        st.info("您可以直接双击下方任意单元格进行最后微调，改完后点击下载。")
        
        final_edited = st.data_editor(
            st.session_state.final_vouchers, 
            num_rows="dynamic", 
            use_container_width=True,
            height=600
        )
        
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
            final_edited.to_excel(writer, index=False)
        
        st.divider()
        st.download_button("📥 确认无误，导出好会计文件", data=out.getvalue(), file_name="好会计凭证包.xlsx")
    else:
        st.info("暂无生成结果，请先在『数据导入』页面上传并运行。")
