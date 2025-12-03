import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from openai import OpenAI
import io
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openpyxl import load_workbook

# --- 页面配置 ---
st.set_page_config(page_title="CTR 视觉重构系统 (V37)", layout="wide")
st.title("🎯 首页卡片 CTR 视觉重构系统 (V37.0 Leader看板)")

# ==========================================
# 🛠️ 绘图函数集 (V37 重构)
# ==========================================
def plot_paired_bar(df, category_col, val_a, val_b, title):
    """绘制 A/B 时期对比柱状图 (分组)"""
    # 转换数据格式为长表，方便 Plotly 分组
    df_melt = df.melt(id_vars=[category_col], value_vars=[val_a, val_b], 
                      var_name='时期', value_name='CTR')
    
    # 映射友好的名字
    df_melt['时期'] = df_melt['时期'].map({val_a: '时期 A (基准)', val_b: '时期 B (当前)'})
    
    fig = px.bar(df_melt, y=category_col, x='CTR', color='时期', barmode='group',
                 title=title, orientation='h', text_auto='.2%',
                 color_discrete_map={'时期 A (基准)': '#95A5A6', '时期 B (当前)': '#3498DB'})
    
    fig.update_layout(
        yaxis={'categoryorder':'total ascending', 'type': 'category'}, # 强制分类轴
        xaxis_tickformat=".2%",
        legend=dict(orientation="h", y=1.1),
        height=500,
        margin=dict(l=20, r=20, t=50, b=20)
    )
    return fig

def plot_impact_diverging(df, category_col, impact_col, title):
    """绘制贡献度/涨跌幅 红色/绿色图"""
    # 根据正负值上色
    df['Color'] = df[impact_col].apply(lambda x: '#E74C3C' if x >= 0 else '#2ECC71')
    
    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=df[category_col],
        x=df[impact_col],
        orientation='h',
        marker=dict(color=df['Color']),
        text=df[impact_col],
        texttemplate='%{text:+.2%}',
        textposition='outside'
    ))
    
    fig.update_layout(
        title=title,
        yaxis={'categoryorder':'total ascending', 'type': 'category'},
        xaxis_tickformat=".2%",
        height=500,
        showlegend=False
    )
    return fig

# ... (保留原有的导出函数) ...
def generate_excel(dfs_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in dfs_dict.items():
            safe_sheet_name = sheet_name[:30]
            df.to_excel(writer, sheet_name=safe_sheet_name, index=False)
            worksheet = writer.sheets[safe_sheet_name]
            worksheet.set_column(0, len(df.columns) - 1, 15)
    return output.getvalue()

def generate_word_report(title, metrics, summary_text, tables_data):
    doc = Document()
    head = doc.add_heading(title, 0)
    head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f"生成时间: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M')}")
    
    doc.add_heading('一、核心大盘战报', level=1)
    p = doc.add_paragraph()
    for k, v in metrics.items():
        run = p.add_run(f"{k}: {v}\n")
        run.font.size = Pt(12)
        run.bold = True
    
    doc.add_heading('二、深度归因与洞察', level=1)
    doc.add_paragraph(summary_text)
    
    for table_title, df in tables_data.items():
        if df.empty: continue
        doc.add_heading(f"三、{table_title}", level=1)
        t = doc.add_table(rows=1, cols=len(df.columns))
        t.style = 'Table Grid'
        hdr_cells = t.rows[0].cells
        for i, col_name in enumerate(df.columns): hdr_cells[i].text = str(col_name)
        for _, row in df.iterrows():
            row_cells = t.add_row().cells
            for i, val in enumerate(row):
                if isinstance(val, float): row_cells[i].text = f"{val:.2%}" if abs(val) < 1 else f"{val:,.0f}"
                else: row_cells[i].text = str(val)
    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()

# ==========================================
# 🤖 AI 助手
# ==========================================
def init_ai_sidebar(context_data):
    st.sidebar.markdown("---")
    st.sidebar.header("🤖 AI 智能分析助手")
    with st.sidebar.expander("⚙️ 模型配置", expanded=False):
        api_key = st.text_input("API Key", type="password")
        base_url = st.text_input("Base URL", value="https://api.deepseek.com")
        model_name = st.text_input("Model Name", value="deepseek-chat")
    
    if "messages" not in st.session_state: st.session_state.messages = []
    for message in st.session_state.messages:
        with st.sidebar.chat_message(message["role"]): st.markdown(message["content"])
    
    if prompt := st.sidebar.chat_input("问我..."):
        if not api_key: st.sidebar.error("请填入 API Key")
        else:
            st.session_state.messages.append({"role": "user", "content": prompt})
            with st.sidebar.chat_message("user"): st.markdown(prompt)
            with st.sidebar.chat_message("assistant"):
                msg_ph = st.empty()
                full_res = ""
                try:
                    client = OpenAI(api_key=api_key, base_url=base_url)
                    sys_prompt = f"你是一个资深数据分析师。基于以下数据回答：\n{context_data}"
                    stream = client.chat.completions.create(
                        model=model_name,
                        messages=[{"role": "system", "content": sys_prompt}] + [{"role": m["role"], "content": m["content"]} for m in st.session_state.messages],
                        stream=True,
                    )
                    for chunk in stream:
                        if chunk.choices[0].delta.content:
                            full_res += chunk.choices[0].delta.content
                            msg_ph.markdown(full_res + "▌")
                    msg_ph.markdown(full_res)
                    st.session_state.messages.append({"role": "assistant", "content": full_res})
                except Exception as e: st.error(f"Error: {e}")

GLOBAL_DATA_CONTEXT = "暂无数据。"

# ==========================================
# 1. 数据接入
# ==========================================
st.sidebar.header("1. 数据接入")
manual_country = st.sidebar.text_input("✍️ 所属国家", value="US").upper()
read_visible_only = st.sidebar.checkbox("👁️ 只读取显示行 (剔除筛选隐藏)", value=False)

file_a = st.sidebar.file_uploader("上传主表格 (A)", type=["xlsx", "xls"], key="file_a")
sheet_name_a = 0
if file_a:
    try:
        xls = pd.ExcelFile(file_a)
        if len(xls.sheet_names) > 1: sheet_name_a = st.sidebar.selectbox(f"表A工作表:", xls.sheet_names, key="s_a")
    except: pass

file_b = st.sidebar.file_uploader("上传对比表格 (B)", type=["xlsx", "xls"], key="file_b")
sheet_name_b = 0
if file_b:
    try:
        xls = pd.ExcelFile(file_b)
        if len(xls.sheet_names) > 1: sheet_name_b = st.sidebar.selectbox(f"表B工作表:", xls.sheet_names, key="s_b")
    except: pass

st.sidebar.markdown("---")
min_exp_noise = st.sidebar.number_input("最小曝光阈值", value=50, step=10)

def extract_start_date(header_str):
    s = str(header_str).strip()
    if "~" in s: return s.split("~")[0].strip()
    if "～" in s: return s.split("～")[0].strip()
    return s

@st.cache_data
def process_data(file, sheet_name=0, visible_only=False):
    try:
        if visible_only:
            wb = load_workbook(file, data_only=True, read_only=False)
            ws = wb.active if sheet_name == 0 else wb[sheet_name]
            data = []
            rows_iter = ws.iter_rows(values_only=False)
            headers = None
            for row in rows_iter:
                if ws.row_dimensions[row[0].row].hidden: continue
                row_values = [cell.value for cell in row]
                if headers is None: headers = row_values
                else: data.append(row_values)
            raw_df = pd.DataFrame(data, columns=headers)
        else:
            raw_df = pd.read_excel(file, sheet_name=sheet_name)
            
        rename_map = {}
        for col in raw_df.columns:
            if "卡片" in col or "Card" in col: rename_map[col] = 'card_id'
            elif "坑位" in col or "Slot" in col: rename_map[col] = 'slot_id'
            elif "指标" in col: rename_map[col] = 'metric_name'
        df = raw_df.rename(columns=rename_map)
        
        if 'slot_id' not in df.columns: df['slot_id'] = 'Default'
        df['card_id'] = df['card_id'].astype(str)
        df['slot_id'] = df['slot_id'].astype(str)
        
        fixed_cols = ['card_id', 'slot_id', 'metric_name', '合计', '均值', '总计', 'Total']
        potential_date_cols = [c for c in df.columns if c not in fixed_cols and "Unnamed" not in str(c)]
        melted = df.melt(id_vars=['card_id', 'slot_id', 'metric_name'], value_vars=potential_date_cols, var_name='original_header', value_name='count')
        
        melted['date_str'] = melted['original_header'].apply(extract_start_date)
        melted['date'] = pd.to_datetime(melted['date_str'], errors='coerce').dt.date
        melted = melted.dropna(subset=['date'])
        melted['count'] = pd.to_numeric(melted['count'], errors='coerce').fillna(0)
        
        def get_type(t):
            if "曝光" in str(t): return "exposure_uv"
            if "点击" in str(t): return "click_uv"
            return None
        melted['type'] = melted['metric_name'].apply(get_type)
        melted = melted.dropna(subset=['type'])
        
        final_df = melted.pivot_table(index=['date', 'card_id', 'slot_id'], columns='type', values='count', aggfunc='sum').reset_index()
        for c in ['exposure_uv', 'click_uv']:
            if c not in final_df.columns: final_df[c] = 0
        final_df = final_df.fillna(0)
        final_df = final_df[final_df['exposure_uv'] >= min_exp_noise]
        final_df = final_df[final_df['click_uv'] <= final_df['exposure_uv']]
        return final_df
    except Exception as e: return None

# --- 3. 单文件分析 ---
def render_analysis_view(data, group_cols, view_name, unique_key_prefix):
    period_stats = data.groupby(group_cols).agg({'exposure_uv': 'sum', 'click_uv': 'sum'}).reset_index()
    period_stats['加权均值CTR'] = period_stats['click_uv'] / period_stats['exposure_uv']
    
    daily_agg = data.groupby(group_cols + ['date']).agg({'exposure_uv': 'sum', 'click_uv': 'sum'}).reset_index()
    daily_agg['daily_ctr'] = daily_agg['click_uv'] / daily_agg['exposure_uv']
    arithmetic_stats = daily_agg.groupby(group_cols)['daily_ctr'].mean().reset_index().rename(columns={'daily_ctr': '算术均值CTR'})
    
    daily_pivot = daily_agg.pivot_table(index=group_cols, columns='date', values='daily_ctr', aggfunc='mean')
    daily_pivot.columns = [d.strftime('%m-%d') for d in daily_pivot.columns]
    
    merged = pd.merge(period_stats, arithmetic_stats, on=group_cols, how='left')
    merged = pd.merge(merged, daily_pivot, on=group_cols, how='left')
    merged = merged.sort_values('exposure_uv', ascending=False)
    
    display_df = merged.copy()
    if 'slot_id' in group_cols:
        display_df['label'] = display_df['card_id'] + " (坑位 " + display_df['slot_id'] + ")"
    else:
        display_df['label'] = display_df['card_id']
    
    date_cols = [c for c in display_df.columns if '-' in c]
    final_cols = ['card_id', 'slot_id', '加权均值CTR', '算术均值CTR', 'exposure_uv', 'click_uv'] + date_cols
    # 确保列存在
    final_cols = [c for c in final_cols if c in display_df.columns]
    
    show_df = display_df[final_cols].rename(columns={'card_id': '卡片ID', 'slot_id': '坑位ID', 'exposure_uv': '总曝光', 'click_uv': '总点击'})

    st.markdown(f"#### 📋 {view_name} - 详细数据")
    format_dict = {'加权均值CTR': '{:.2%}', '算术均值CTR': '{:.2%}', '总曝光': '{:,.0f}', '总点击': '{:,.0f}'}
    for d in date_cols: format_dict[d] = '{:.2%}'
    styled_df = show_df.style.format(format_dict).background_gradient(subset=['加权均值CTR'], cmap='RdYlGn', axis=0)
    st.dataframe(styled_df, use_container_width=True, height=400)
    
    st.markdown(f"#### 📈 {view_name} - 趋势图")
    unique_key = f"ms_{view_name}_{unique_key_prefix}"
    top_labels = display_df['label'].head(5).tolist()
    sel = st.multiselect(f"选择要对比的{view_name}", display_df['label'].unique(), default=top_labels, key=unique_key)
    if sel:
        plot_df = daily_agg.copy()
        if 'slot_id' in group_cols: plot_df['label'] = plot_df['card_id'] + " (坑位 " + plot_df['slot_id'] + ")"
        else: plot_df['label'] = plot_df['card_id']
        plot_df = plot_df[plot_df['label'].isin(sel)]
        fig = px.line(plot_df, x='date', y='daily_ctr', color='label', markers=True)
        fig.update_yaxes(tickformat=".2%")
        st.plotly_chart(fig, use_container_width=True)

def show_single_analysis(df, label="表格 A"):
    st.markdown(f"## 🔎 {label} - 深度分析")
    
    enable_internal = st.checkbox("⚔️ 开启表内时段对比", key=f"ec_{label}")
    if enable_internal:
        show_comparison_logic(df, df, f"{label}-A", f"{label}-B")
        return

    min_d, max_d = df['date'].min(), df['date'].max()
    date_range = st.date_input("选择时间段", [min_d, max_d], key=f"d_{label}")
    if len(date_range) != 2: return
    sub_df = df[(df['date'] >= date_range[0]) & (df['date'] <= date_range[1])].copy()
    
    total_exp = sub_df['exposure_uv'].sum()
    total_clk = sub_df['click_uv'].sum()
    weighted_ctr = total_clk / total_exp if total_exp > 0 else 0
    daily_agg = sub_df.groupby('date').agg({'exposure_uv':'sum', 'click_uv':'sum'})
    daily_agg['day_ctr'] = daily_agg['click_uv'] / daily_agg['exposure_uv']
    arithmetic_ctr = daily_agg['day_ctr'].mean()
    
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("总曝光", f"{total_exp:,.0f}")
    c2.metric("总点击", f"{total_clk:,.0f}")
    c3.metric("加权均值 CTR", f"{weighted_ctr:.2%}")
    c4.metric("算术均值 CTR", f"{arithmetic_ctr:.2%}")
    
    global GLOBAL_DATA_CONTEXT
    GLOBAL_DATA_CONTEXT = f"单表分析: {label}, CTR: {weighted_ctr:.2%}, 点击: {total_clk}"
    
    export_df = sub_df.groupby(['card_id', 'slot_id']).agg({'exposure_uv':'sum', 'click_uv':'sum'}).reset_index()
    export_df['weighted_ctr'] = export_df['click_uv'] / export_df['exposure_uv']
    export_df = export_df.sort_values('exposure_uv', ascending=False)
    top_5 = export_df.head(5).rename(columns={'card_id':'卡片ID', 'weighted_ctr':'CTR', 'exposure_uv':'曝光'})
    
    st.divider()
    t1, t2 = st.tabs(["💳 视图一：只看卡片", "📍 视图二：细分卡片+坑位"])
    with t1: render_analysis_view(sub_df, ['card_id'], "卡片维度", label)
    with t2: render_analysis_view(sub_df, ['card_id', 'slot_id'], "卡片+坑位细分", label)
    
    st.divider()
    st.header("📥 导出中心")
    c_e1, c_e2 = st.columns(2)
    word_file = generate_word_report(f"报告-{manual_country}", {"周期": str(date_range), "曝光": f"{total_exp:,.0f}"}, "数据附表", {"Top5": top_5})
    excel_file = generate_excel({"聚合": export_df, "明细": sub_df})
    with c_e1: st.download_button("📄 下载 Word", word_file, f"Report_{label}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=f"bw_{label}")
    with c_e2: st.download_button("📊 下载 Excel", excel_file, f"Data_{label}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"be_{label}")

# --- 5. 双表对比 (V37: Dashboard + Charts) ---
def show_comparison_logic(d1_raw, d2_raw, label_a_name="表格A", label_b_name="表格B"):
    st.markdown("### ⚙️ 对比配置")
    compare_mode = st.radio("👉 维度：", ["💳 仅对比卡片", "📍 对比 卡片+坑位"], horizontal=True, key=f"rad_{label_a_name}")
    group_cols = ['card_id'] if "仅对比卡片" in compare_mode else ['card_id', 'slot_id']
    
    all_cards = sorted(list(set(d1_raw['card_id'].unique()) | set(d2_raw['card_id'].unique())))
    exclude_list = st.multiselect("🚫 剔除指定卡片", all_cards, key=f"exc_{label_a_name}")
    
    if exclude_list:
        d1 = d1_raw[~d1_raw['card_id'].isin(exclude_list)].copy()
        d2 = d2_raw[~d2_raw['card_id'].isin(exclude_list)].copy()
    else:
        d1, d2 = d1_raw.copy(), d2_raw.copy()
    
    c1, c2 = st.columns(2)
    with c1: d1_range = st.date_input(f"{label_a_name} 时间段", [d1['date'].min(), d1['date'].max()], key=f"dr1_{label_a_name}")
    with c2: d2_range = st.date_input(f"{label_b_name} 时间段", [d2['date'].min(), d2['date'].max()], key=f"dr2_{label_a_name}")
        
    if len(d1_range)==2 and len(d2_range)==2:
        d1_final = d1[(d1['date'] >= d1_range[0]) & (d1['date'] <= d1_range[1])]
        d2_final = d2[(d2['date'] >= d2_range[0]) & (d2['date'] <= d2_range[1])]
        
        def calc_global(d):
            e = d['exposure_uv'].sum()
            c = d['click_uv'].sum()
            return e, c, (c/e if e>0 else 0)
        
        ea, ca, ctra = calc_global(d1_final)
        eb, cb, ctrb = calc_global(d2_final)
        ctr_multiple = (ctrb / ctra) if ctra > 0 else 0
        exp_diff_pct = (eb - ea) / ea if ea > 0 else 0
        
        # 归因
        top_row = d2_final.groupby('card_id')['click_uv'].sum().sort_values(ascending=False).head(1)
        summary_text_report, top_info = "", "无明显头部"
        
        if not top_row.empty:
            top_id = top_row.index[0]
            top_contrib = top_row.values[0] / cb if cb > 0 else 0
            d1_no = d1_final[d1_final['card_id'] != top_id]
            d2_no = d2_final[d2_final['card_id'] != top_id]
            _, _, ctra_no = calc_global(d1_no)
            _, _, ctrb_no = calc_global(d2_no)
            ctr_mult_no = (ctrb_no / ctra_no) if ctra_no > 0 else 0
            conclusion = "✅ 普涨型" if ctr_mult_no > 1.05 else "⚠️ 头部依赖型"
            
            st.markdown(f"""
            <div style="background-color: #F0F2F6; padding: 20px; border-radius: 10px; border-left: 6px solid #FF9800; color: #111;">
                <h3 style="margin:0; color: #000;">📝 深度归因总结</h3>
                1. <b>整体表现：</b> CTR 是上周期的 <b>{ctr_multiple:.2f} 倍</b> ({ctrb:.2%} vs {ctra:.2%})。<br>
                2. <b>剔除验证：</b> 剔除头部【{top_id}】后，CTR 倍数为 <b>{ctr_mult_no:.2f} 倍</b>。<br>
                <div style="margin-top:5px;">{conclusion}</div>
            </div>
            """, unsafe_allow_html=True)
            summary_text_report = f"剔除头部 {top_id} 后，倍数为 {ctr_mult_no:.2f}。结论: {conclusion}"
            top_info = f"Top1: {top_id} (贡献{top_contrib:.1%})"
        
        st.subheader("📊 全盘战报")
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("全盘 CTR", f"{ctrb:.2%}", f"{ctrb-ctra:+.2%}", delta_color="normal")
        k2.metric("总曝光", f"{eb:,.0f}", f"{exp_diff_pct:+.1%}", delta_color="normal")
        k3.metric("总点击", f"{cb:,.0f}", f"{(cb-ca)/ca if ca>0 else 0:+.1%}", delta_color="normal")
        
        diag = "⚪ 常规"
        if exp_diff_pct < -0.2 and (ctrb-ctra) > 0 and (cb-ca) < 0: diag = "⚠️ 虚假提效 (萎缩)"
        elif exp_diff_pct > 0.2 and (ctrb-ctra) < 0: diag = "🟠 流量稀释"
        elif ctr_multiple > 1.05 and (cb-ca) > 0: diag = "🟢 有效增长"
        k4.info(diag)
        
        global GLOBAL_DATA_CONTEXT
        GLOBAL_DATA_CONTEXT = f"对比战报\nA表CTR: {ctra:.2%} B表CTR: {ctrb:.2%}\n诊断: {diag}\n归因: {top_info}"

        # === V37 新增：双表对比 Dashboard ===
        st.divider()
        st.subheader("📊 双表对比驾驶舱 (Dashboard)")
        
        # 准备数据
        stat1 = d1_final.groupby(group_cols).agg({'exposure_uv':'sum', 'click_uv':'sum'}).reset_index()
        stat2 = d2_final.groupby(group_cols).agg({'exposure_uv':'sum', 'click_uv':'sum'}).reset_index()
        stat1 = stat1.rename(columns={'exposure_uv':'Exp_A', 'click_uv':'Clk_A'})
        stat2 = stat2.rename(columns={'exposure_uv':'Exp_B', 'click_uv':'Clk_B'})
        stat1['CTR_A'] = stat1['Clk_A'] / stat1['Exp_A']
        stat2['CTR_B'] = stat2['Clk_B'] / stat2['Exp_B']
        
        comp = pd.merge(stat1, stat2, on=group_cols, how='outer', indicator=True)
        comp['_merge'] = comp['_merge'].astype(str)
        comp = comp.fillna(0)
        def label_status(row):
            if row['_merge'] == 'both': return '🔵 延续'
            if row['_merge'] == 'right_only': return '🟢 新上架'
            if row['_merge'] == 'left_only': return '🔴 已下架'
        comp['状态'] = comp.apply(label_status, axis=1)
        comp['CTR差值'] = comp['CTR_B'] - comp['CTR_A']
        
        # Label 处理
        if 'slot_id' in group_cols:
            comp['label'] = comp['card_id'] + " (" + comp['slot_id'] + ")"
        else:
            comp['label'] = comp['card_id']
            
        # 1. 核心卡片对比图 (Top 10 High Traffic)
        top_traffic = comp.sort_values('Exp_B', ascending=False).head(10)
        if not top_traffic.empty:
            fig_bar = plot_paired_bar(top_traffic, 'label', 'CTR_A', 'CTR_B', "🔥 流量 Top 10 卡片 CTR 对比 (A vs B)")
            st.plotly_chart(fig_bar, use_container_width=True)
            
        # 2. 贡献度图 (Impact)
        # 贡献度 = (CTR_B - CTR_A) * 权重(这里简单用平均曝光占比近似)
        comp['Impact'] = comp['CTR差值'] * ((comp['Exp_A'] + comp['Exp_B'])/2)
        top_impact = comp.sort_values('Impact', ascending=False).head(5) # 拉升 Top 5
        bot_impact = comp.sort_values('Impact', ascending=True).head(5) # 拖累 Top 5
        impact_df = pd.concat([top_impact, bot_impact])
        
        fig_impact = plot_impact_diverging(impact_df, 'label', 'CTR差值', "🏆 涨跌幅 Top 榜 (红涨绿跌)")
        st.plotly_chart(fig_impact, use_container_width=True)

        st.divider()
        st.subheader("📋 详细数据表")
        comp = comp.sort_values(['状态', 'CTR差值'])
        show_cols = ['状态'] + group_cols + ['CTR_A', 'CTR_B', 'CTR差值', 'Exp_A', 'Exp_B']
        fmt = {'CTR_A':'{:.2%}', 'CTR_B':'{:.2%}', 'CTR差值':'{:+.2%}', 'Exp_A':'{:,.0f}', 'Exp_B':'{:,.0f}'}
        def highlight_status(val):
            if '新' in str(val): return 'color: green; font-weight: bold'
            if '下架' in str(val): return 'color: red; font-weight: bold'
            return 'color: blue'
        st.dataframe(comp[show_cols].style.format(fmt).applymap(highlight_status, subset=['状态']).background_gradient(subset=['CTR差值'], cmap='RdYlGn', vmin=-0.02, vmax=0.02), use_container_width=True)
        
        st.divider()
        st.header("📥 导出中心")
        c_ex1, c_ex2 = st.columns(2)
        word_file = generate_word_report(f"对比战报-{manual_country}", {"CTR变化": f"{ctra:.2%}->{ctrb:.2%}"}, summary_text_report, {"红榜": top_impact, "黑榜": bot_impact})
        excel_file = generate_excel({"全盘": comp, "红榜": top_impact})
        with c_ex1: st.download_button("📄 下载 Word", word_file, f"Report_Compare_{label_a_name}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=f"bw_{label_a_name}")
        with c_ex2: st.download_button("📊 下载 Excel", excel_file, f"Data_Compare_{label_a_name}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"be_{label_a_name}")

def show_comparison(df1, df2):
    show_comparison_logic(df1, df2, "表格A", "表格B")

# --- 主逻辑 ---
df_a = None
if file_a: df_a = process_data(file_a, sheet_name_a, visible_only=read_visible_only)
df_b = None
if file_b: df_b = process_data(file_b, sheet_name_b, visible_only=read_visible_only)

if df_a is not None:
    if df_b is not None:
        mode = st.radio("👇 模式", ["📄 单文件分析", "⚔️ 双表对比"], horizontal=True)
        st.divider()
        if mode == "📄 单文件分析":
            t1, t2 = st.tabs(["表格 A", "表格 B"])
            with t1: show_single_analysis(df_a, "表格 A")
            with t2: show_single_analysis(df_b, "表格 B")
        else:
            show_comparison(df_a, df_b)
    else:
        show_single_analysis(df_a, "表格 A")
else:
    st.info("👈 请在左侧上传 Excel 文件。")

if GLOBAL_DATA_CONTEXT != "暂无数据。":
    init_ai_sidebar(GLOBAL_DATA_CONTEXT)
