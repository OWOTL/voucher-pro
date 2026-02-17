import streamlit as st
import pandas as pd
import io

# --- 0. 全局配置 ---
st.set_page_config(layout="wide", page_title="专业财务凭证DIY系统")

# 标题与理念
st.title("🧮 专业财务凭证 DIY 生成系统")
st.markdown("### 💡 设计理念：规则即模板，变量即数据，严格的一单一号。")

# --- 1. 核心功能函数 ---

def load_default_rules():
    """初始化一个标准的、空的、但是包含必要表头的规则模板"""
    return pd.DataFrame([
        # 示例数据，帮助用户理解
        {"匹配关键词": "收汇", "摘要模板": "收{单位}货款", "借/贷": "借", "科目编码": "1002", "辅助核算": ""},
        {"匹配关键词": "收汇", "摘要模板": "核销应收账款", "借/贷": "贷", "科目编码": "1122", "辅助核算": "{单位}"},
        {"匹配关键词": "费用", "摘要模板": "报销{单位}费用", "借/贷": "借", "科目编码": "6602", "辅助核算": "{部门}"},
        {"匹配关键词": "费用", "摘要模板": "付现金", "借/贷": "贷", "科目编码": "1001", "辅助核算": ""},
    ])

def validate_upload(df):
    """验证上传格式，返回 (是否通过, 消息/列名列表)"""
    if df.empty:
        return False, "文件为空"
    return True, df.columns.tolist()

# --- 2. 侧边栏：状态控制 ---
if 'rules_df' not in st.session_state:
    st.session_state.rules_df = load_default_rules()

# --- 3. 主界面分步引导 ---

# ==========================================================
# 步骤 1：定义标准分录规则 (DIY 核心)
# ==========================================================
st.markdown("---")
st.subheader("1️⃣ 设定分录规则 (模板库)")
st.info("""
**操作说明：**
1. **匹配关键词**：这是系统识别业务的唯一标识。例如流水表里有“收汇”，这里就必须写“收汇”。
2. **智能变量**：支持使用 `{列名}`。例如摘要写 `收{单位}款`，系统会自动把流水里的单位填进去。
3. **可空项**：除了关键词，其他都可以根据需要留空。
4. **增删改**：像操作 Excel 一样，直接在表格里修改、添加行、删除行。
""")

# 使用 data_editor 实现完全自由的行列编辑
st.session_state.rules_df = st.data_editor(
    st.session_state.rules_df,
    num_rows="dynamic",
    use_container_width=True,
    key="editor_rules"
)

# ==========================================================
# 步骤 2：导入前说明与文件上传
# ==========================================================
st.markdown("---")
st.subheader("2️⃣ 导入流水数据")

c1, c2 = st.columns([1, 2])
with c1:
    st.warning("📋 **导入格式要求**")
    st.markdown("""
    您的 Excel 文件建议包含以下列（列名可自定义，但逻辑要清晰）：
    * **日期列**：业务发生时间
    * **摘要/业务列**：用于匹配规则（如“收汇”、“提现”）
    * **单位/辅助列**：用于填充变量
    * **金额列**：借贷金额
    """)

with c2:
    uploaded_file = st.file_uploader("拖拽或点击上传 Excel 流水单", type=['xlsx', 'xls'])

# ==========================================================
# 步骤 3：自动提取与映射 (智能化)
# ==========================================================
if uploaded_file:
    df_raw = pd.read_excel(uploaded_file).fillna("")
    
    st.success(f"✅ 读取成功！共 {len(df_raw)} 条数据。请告诉系统每一列的含义：")
    
    # 智能映射区域
    cols = df_raw.columns.tolist()
    
    # 布局映射选择器
    m1, m2, m3, m4 = st.columns(4)
    with m1: map_date = st.selectbox("哪一列是【日期】？", cols, index=0)
    with m2: map_biz = st.selectbox("哪一列是【业务类型/匹配词】？", cols, index=1 if len(cols)>1 else 0)
    with m3: map_unit = st.selectbox("哪一列是【单位/辅助信息】？", cols, index=2 if len(cols)>2 else 0)
    with m4: map_amt = st.selectbox("哪一列是【金额】？", cols, index=3 if len(cols)>3 else 0)

    # ==========================================================
    # 步骤 4：生成与诊断 (严谨逻辑)
    # ==========================================================
    st.markdown("---")
    if st.button("🚀 开始生成凭证 (自动编号)", type="primary"):
        results = []
        errors = []
        rules = st.session_state.rules_df
        
        # 预检查规则表必要列
        required_rule_cols = ["匹配关键词", "摘要模板", "借/贷", "科目编码", "辅助核算"]
        if not all(col in rules.columns for col in required_rule_cols):
             st.error(f"❌ 规则表表头被破坏，请确保包含：{required_rule_cols}")
             st.stop()

        # 开始循环处理每一行流水
        voucher_counter = 1 # 凭证号计数器，从 001 开始
        
        progress_bar = st.progress(0)
        
        for idx, row in df_raw.iterrows():
            # 更新进度条
            progress_bar.progress((idx + 1) / len(df_raw))
            
            # 1. 获取当前流水的业务类型
            biz_key = str(row[map_biz]).strip()
            
            # 2. 去规则库里找匹配的行
            # 逻辑：查找规则表中“匹配关键词”列等于当前业务类型的所有行
            matched_rules = rules[rules["匹配关键词"] == biz_key]
            
            # 3. 诊断：如果没找到规则
            if matched_rules.empty:
                errors.append({
                    "行号": idx + 2, # Excel从第2行开始
                    "日期": row[map_date],
                    "业务内容": biz_key,
                    "失败原因": "❌ 未在规则表中找到对应的关键词",
                    "建议操作": f"请在步骤1的规则表中添加一行，关键词填入 '{biz_key}'"
                })
                continue # 跳过这一行
            
            # 4. 生成凭证 (严格一单一号)
            v_num = f"{voucher_counter:03d}" # 格式化为 001, 002...
            
            for _, rule in matched_rules.iterrows():
                # 智能变量替换 (DIY的核心)
                # 摘要处理
                summary_text = str(rule["摘要模板"])
                summary_text = summary_text.replace("{单位}", str(row[map_unit]))
                summary_text = summary_text.replace("{金额}", str(row[map_amt]))
                
                # 辅助核算处理
                aux_text = str(rule["辅助核算"])
                aux_text = aux_text.replace("{单位}", str(row[map_unit]))
                aux_text = aux_text.replace("{部门}", str(row[map_unit])) # 假设部门也在单位列，或者用户自己扩展

                # 构建最终行
                results.append({
                    "凭证号": v_num,
                    "日期": str(row[map_date]).split(" ")[0], # 只要日期部分
                    "摘要": summary_text,
                    "科目编码": rule["科目编码"],
                    "借方金额": row[map_amt] if str(rule["借/贷"]) == "借" else 0,
                    "贷方金额": row[map_amt] if str(rule["借/贷"]) == "贷" else 0,
                    "辅助核算": aux_text
                })
            
            # 处理完一笔业务，凭证号+1
            voucher_counter += 1

        # ==========================================================
        # 结果反馈区
        # ==========================================================
        if errors:
            st.error(f"⚠️ 生成完成，但有 {len(errors)} 条数据失败！请查看下方诊断报告：")
            st.dataframe(pd.DataFrame(errors), use_container_width=True)
            st.markdown("### 🛠️ 怎么修复？")
            st.markdown("1. 查看上方表格的**业务内容**。")
            st.markdown("2. 回到**步骤 1**，在规则表里补上缺失的关键词规则。")
            st.markdown("3. 再次点击生成按钮。")
        else:
            st.success(f"🎉 完美！成功生成 {len(df_raw)} 笔凭证，无任何错误。")

        if results:
            st.subheader("3️⃣ 结果预览与导出")
            final_df = pd.DataFrame(results)
            
            # 预览
            st.dataframe(final_df, use_container_width=True)
            
            # 导出
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                final_df.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 下载标准凭证 Excel (可以直接导入财务软件)",
                data=output.getvalue(),
                file_name="generated_vouchers_001.xlsx",
                mime="application/vnd.ms-excel",
                type="primary"
            )

else:
    st.info("👋 请先上传 Excel 文件以开始工作。")
