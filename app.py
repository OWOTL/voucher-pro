import streamlit as st
import pandas as pd
import io

st.set_page_config(layout="wide", page_title="财务凭证自动化系统")

st.title("🚀 财务凭证自动化生成系统 (标准校验版)")

# --- 1. 侧边栏：定义规则表的标题 (完全 DIY) ---
with st.sidebar:
    st.header("🎨 1. 定义规则表标题")
    t_biz = st.text_input("【业务类型】叫什么", "业务场景")
    t_dir = st.text_input("【借贷方向】叫什么", "方向")
    t_code = st.text_input("【科目编码】叫什么", "科目编号")
    t_memo = st.text_input("【摘要模板】叫什么", "摘要内容")
    t_aux = st.text_input("【辅助核算关联列】叫什么", "关联流水列")

    st.divider()
    st.header("📋 2. 原始数据导入要求")
    st.info("您的 Excel 必须包含以下三类信息：\n1. 日期\n2. 业务类型关键词\n3. 金额")

# --- 2. 初始化规则数据 ---
if 'rules' not in st.session_state:
    st.session_state.rules = pd.DataFrame([
        {t_biz: "发货", t_dir: "借", t_code: "1122", t_memo: "发货给{客户}", t_aux: "客户"},
        {t_biz: "发货", t_dir: "贷", t_code: "6001", t_memo: "销售收入", t_aux: ""}
    ])

# 同步标题修改
if list(st.session_state.rules.columns) != [t_biz, t_dir, t_code, t_memo, t_aux]:
    st.session_state.rules.columns = [t_biz, t_dir, t_code, t_memo, t_aux]

# --- 3. 页面布局 ---
tab1, tab2, tab3 = st.tabs(["⚙️ 规则 DIY 设置", "📥 数据导入与校验", "👁️ 结果预览与导出"])

# --- TAB 1: 规则设置 ---
with tab1:
    st.subheader("设置您的分录规则")
    st.markdown("在这里定义不同业务对应的会计分录。标题已根据左侧设置更新。")
    st.session_state.rules = st.data_editor(st.session_state.rules, num_rows="dynamic", use_container_width=True)

# --- TAB 2: 导入与校验 ---
with tab2:
    st.subheader("上传原始业务数据")
    uploaded_file = st.file_uploader("支持 Excel (.xlsx) 格式", type=['xlsx'])
    
    if uploaded_file:
        df_raw = pd.read_excel(uploaded_file).fillna("")
        raw_cols = df_raw.columns.tolist()
        
        st.write("已成功读取文件，请匹配数据列：")
        col1, col2, col3 = st.columns(3)
        with col1:
            map_date = st.selectbox("哪一列是【日期】？", raw_cols)
        with col2:
            map_biz = st.selectbox("哪一列是【业务类型】？", raw_cols)
        with col3:
            map_amt = st.selectbox("哪一列是【金额】？", raw_cols)
            
        if st.button("开始校验并生成凭证"):
            errors = []
            vouchers = []
            
            for i, row in df_raw.iterrows():
                biz_val = str(row[map_biz]).strip()
                # 匹配规则
                matches = st.session_state.rules[st.session_state.rules[t_biz] == biz_val]
                
                if matches.empty:
                    errors.append(f"第 {i+2} 行：业务类型 '{biz_val}' 尚未定义分录规则。")
                    continue
                
                # 凭证号（一行原始数据一个号）
                v_no = str(i + 1).zfill(4)
                
                for _, r in matches.iterrows():
                    # 摘要替换
                    memo = str(r[t_memo])
                    for c in raw_cols:
                        if f"{{{c}}}" in memo: memo = memo.replace(f"{{{c}}}", str(row[c]))
                    
                    # 辅助核算
                    aux_col = r[t_aux]
                    aux_val = row[aux_col] if aux_col in raw_cols else ""

                    vouchers.append({
                        "凭证号": v_no,
                        "日期": str(row[map_date]).split(" ")[0],
                        "摘要": memo,
                        "科目编码": r[t_code],
                        "借方": row[map_amt] if r[t_dir] == '借' else 0,
                        "贷方": row[map_amt] if r[t_dir] == '贷' else 0,
                        "辅助核算": aux_val
                    })
            
            if errors:
                st.error("⚠️ 导入存在问题：")
                for err in errors: st.write(err)
                st.warning("请在 TAB 1 补全缺失规则后重新点击生成。")
            
            if vouchers:
                st.session_state.final_df = pd.DataFrame(vouchers)
                st.success(f"✅ 生成成功！共转换 {len(df_raw)} 笔业务。请前往预览。")

# --- TAB 3: 预览与导出 ---
with tab3:
    if 'final_df' in st.session_state:
        st.subheader("结果预览 (支持直接修改)")
        st.info("您可以直接双击下方单元格修改任何内容（如科目、金额、摘要）。")
        
        # 结果编辑器
        edited_df = st.data_editor(st.session_state.final_df, num_rows="dynamic", use_container_width=True, height=500)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            edited_df.to_excel(writer, index=False)
        
        st.divider()
        st.download_button(
            label="📥 导出为好会计导入文件",
            data=output.getvalue(),
            file_name="好会计凭证导入表.xlsx",
            mime="application/vnd.ms-excel"
        )
    else:
        st.info("暂无预览数据，请先完成第二步导入。")
