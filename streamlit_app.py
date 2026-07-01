"""
2026年海外客户投诉台账 - Streamlit 云端数据看板
部署于 Streamlit Community Cloud, 通过 GitHub 自动更新
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime
import os
import io
import zipfile

# ============================================================
# Page Config
# ============================================================
st.set_page_config(
    page_title="2026海外客诉看板",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# Data Loading (cached)
# ============================================================
@st.cache_data(ttl=3600)
def load_data():
    """从 Excel 读取并清洗数据"""
    file_path = os.path.join(os.path.dirname(__file__), "2026年海外客户投诉台账.xlsx")

    df = pd.read_excel(file_path, sheet_name='所有客诉', header=1)
    df = df.dropna(how='all')
    df.columns = [
        '编号','分公司','国家或地区','是否大客户','投诉日期','应结案日期','实际完成日期',
        '完成周期（天）','机型属性','故障比例','问题描述','客户诉求','跟进人','处理类型',
        '结案状态','应急措施','原因分析','长期整改措施','责任单位','故障大类','品质负责人',
        '8D报告','8D措施点检','备注'
    ]

    # Strict filtering
    df['编号_num'] = pd.to_numeric(df['编号'], errors='coerce')
    df = df[df['编号_num'].notna() & df['分公司'].notna() & (df['分公司'].astype(str).str.strip() != '')].copy()
    df = df.drop_duplicates(subset='编号_num', keep='first').sort_values('编号_num').reset_index(drop=True)
    df['编号_int'] = df['编号_num'].astype(int)

    # Date parsing
    for col in ['投诉日期','应结案日期','实际完成日期']:
        df[col] = pd.to_datetime(df[col], errors='coerce')

    df['完成周期（天）'] = pd.to_numeric(df['完成周期（天）'], errors='coerce')
    df['故障比例'] = pd.to_numeric(df['故障比例'], errors='coerce')
    df['投诉月份'] = df['投诉日期'].dt.to_period('M').astype(str)

    return df


def apply_filters(df):
    """Apply sidebar filters to dataframe"""
    filtered = df.copy()

    # Branch filter
    branches = st.sidebar.multiselect(
        "分公司", options=sorted(df['分公司'].dropna().unique()),
        default=[]
    )
    if branches:
        filtered = filtered[filtered['分公司'].isin(branches)]

    # Status filter
    statuses = st.sidebar.multiselect(
        "结案状态", options=sorted(df['结案状态'].dropna().unique()),
        default=[]
    )
    if statuses:
        filtered = filtered[filtered['结案状态'].isin(statuses)]

    # VIP filter
    vip_options = st.sidebar.multiselect(
        "客户类型", options=['是','否'],
        default=[]
    )
    if vip_options:
        filtered = filtered[filtered['是否大客户'].isin(vip_options)]

    # Date range
    min_date = df['投诉日期'].min()
    max_date = df['投诉日期'].max()
    if pd.notna(min_date) and pd.notna(max_date):
        date_range = st.sidebar.date_input(
            "投诉日期范围",
            value=(min_date.date(), max_date.date()),
            min_value=min_date.date(),
            max_value=max_date.date(),
        )
        if len(date_range) == 2:
            filtered = filtered[
                (filtered['投诉日期'] >= pd.Timestamp(date_range[0])) &
                (filtered['投诉日期'] <= pd.Timestamp(date_range[1]))
            ]

    # Fault category filter
    faults = st.sidebar.multiselect(
        "故障大类", options=sorted(df['故障大类'].dropna().unique()),
        default=[]
    )
    if faults:
        filtered = filtered[filtered['故障大类'].isin(faults)]

    return filtered


# ============================================================
# Color Palette
# ============================================================
COLORS = ['#1890FF','#27AE60','#F39C12','#E74C3C','#9B59B6','#1ABC9C','#E67E22','#3498DB',
          '#2ECC71','#E91E63','#00BCD4','#FF9800','#795548','#607D8B','#CDDC39']

STATUS_COLORS = {'结案': '#27AE60', '关闭': '#95A5A6', '未结案': '#F39C12', '暂停': '#E74C3C'}

# ============================================================
# Chart Functions
# ============================================================
def make_branch_chart(df):
    """分公司投诉分布 - 横向柱状图"""
    counts = df['分公司'].value_counts().reset_index()
    counts.columns = ['分公司','数量']
    counts = counts.sort_values('数量')

    colors = ['#E74C3C' if v >= 8 else '#F39C12' if v >= 5 else '#1890FF' for v in counts['数量']]

    fig = go.Figure(go.Bar(
        x=counts['数量'], y=counts['分公司'],
        orientation='h',
        marker_color=colors,
        text=counts['数量'],
        textposition='outside',
        hovertemplate='%{y}: %{x} 条<extra></extra>',
    ))
    fig.update_layout(
        title='分公司投诉分布',
        xaxis_title='投诉数量',
        height=360,
        margin=dict(l=10, r=30, t=40, b=10),
        bargap=0.3,
    )
    return fig


def make_country_chart(df):
    """国家/地区 TOP15"""
    counts = df['国家或地区'].value_counts().nlargest(15).reset_index()
    counts.columns = ['国家','数量']
    counts = counts.sort_values('数量')

    colors = ['#E74C3C' if i >= len(counts)-3 else '#1890FF' for i in range(len(counts))]

    fig = go.Figure(go.Bar(
        x=counts['数量'], y=counts['国家'],
        orientation='h',
        marker_color=colors,
        text=counts['数量'],
        textposition='outside',
        hovertemplate='%{y}: %{x} 条<extra></extra>',
    ))
    fig.update_layout(
        title='国家/地区投诉量 TOP15',
        xaxis_title='投诉数量',
        height=360,
        margin=dict(l=10, r=30, t=40, b=10),
        bargap=0.25,
    )
    return fig


def make_monthly_trend(df):
    """月度投诉趋势 + 结案率"""
    monthly = df.groupby('投诉月份').agg(
        投诉量=('编号_int','count'),
        结案量=('结案状态', lambda x: x.isin(['结案','关闭']).sum())
    ).reset_index()
    monthly['结案率'] = (monthly['结案量'] / monthly['投诉量'] * 100).round(1)

    fig = make_subplots(specs=[[{"secondary_y": True}]])

    fig.add_trace(go.Bar(
        name='投诉量', x=monthly['投诉月份'], y=monthly['投诉量'],
        marker_color='#1890FF', text=monthly['投诉量'], textposition='outside',
    ), secondary_y=False)

    fig.add_trace(go.Bar(
        name='已结案', x=monthly['投诉月份'], y=monthly['结案量'],
        marker_color='#27AE60',
    ), secondary_y=False)

    fig.add_trace(go.Scatter(
        name='结案率', x=monthly['投诉月份'], y=monthly['结案率'],
        mode='lines+markers+text', text=monthly['结案率'].apply(lambda x: f'{x}%'),
        textposition='top center', line=dict(color='#F39C12', width=3),
        marker=dict(size=10, symbol='diamond'),
    ), secondary_y=True)

    fig.add_hline(y=monthly['结案率'].mean(), line_dash="dash", line_color="gray",
                  annotation_text=f"平均结案率 {monthly['结案率'].mean():.0f}%", secondary_y=True)

    fig.update_layout(
        title='月度投诉趋势 & 结案率',
        hovermode='x unified',
        height=380,
        margin=dict(l=10, r=10, t=40, b=10),
        legend=dict(orientation='h', y=-0.15),
    )
    fig.update_yaxes(title_text="条数", secondary_y=False)
    fig.update_yaxes(title_text="结案率 (%)", range=[0, 105], secondary_y=True)
    return fig


def make_status_pie(df):
    """结案状态环形图"""
    counts = df['结案状态'].value_counts().reset_index()
    counts.columns = ['状态','数量']

    fig = go.Figure(go.Pie(
        labels=counts['状态'], values=counts['数量'],
        hole=0.55,
        marker_colors=[STATUS_COLORS.get(s, '#999') for s in counts['状态']],
        textinfo='label+percent',
        textfont_size=13,
        hovertemplate='%{label}: %{value} 条 (%{percent})<extra></extra>',
    ))
    fig.update_layout(
        title='结案状态分布',
        height=380,
        margin=dict(l=10, r=10, t=40, b=10),
    )
    return fig


def make_fault_pareto(df):
    """故障大类帕累托图（标准垂直柱状图+累计折线）"""
    counts = df['故障大类'].value_counts().reset_index()
    counts.columns = ['故障大类','数量']
    total = counts['数量'].sum()
    counts = counts.sort_values('数量', ascending=False)
    counts['累计占比'] = counts['数量'].cumsum() / total * 100

    colors = ['#E74C3C' if i < 3 else '#F39C12' if i < 6 else '#1890FF'
              for i in range(len(counts))]

    fig = make_subplots(specs=[[{"secondary_y": True}]])

    fig.add_trace(go.Bar(
        name='投诉数量', x=counts['故障大类'], y=counts['数量'],
        marker_color=colors,
        text=counts['数量'], textposition='outside',
    ), secondary_y=False)

    fig.add_trace(go.Scatter(
        name='累计占比', x=counts['故障大类'], y=counts['累计占比'],
        mode='lines+markers', line=dict(color='#9B59B6', width=2.5),
        marker=dict(size=8, symbol='diamond'),
    ), secondary_y=True)

    fig.add_hline(y=80, line_dash="dash", line_color="#E74C3C",
                  annotation_text="80% 线", secondary_y=True)

    fig.update_layout(
        title='故障大类帕累托分析',
        height=400,
        margin=dict(l=10, r=10, t=40, b=40),
        legend=dict(orientation='h', y=-0.15),
        xaxis=dict(tickangle=-35, tickfont=dict(size=11)),
    )
    fig.update_yaxes(title_text="投诉数量", secondary_y=False)
    fig.update_yaxes(title_text="累计占比 (%)", range=[0, 105], secondary_y=True)
    return fig


def make_model_vip_chart(df):
    """机型属性 + 大客户占比"""
    model_counts = df['机型属性'].value_counts().reset_index()
    model_counts.columns = ['机型','数量']

    fig = make_subplots(
        rows=1, cols=2,
        specs=[[{'type': 'pie'}, {'type': 'pie'}]],
        subplot_titles=('机型属性', '大客户 vs 非大客户'),
    )

    fig.add_trace(go.Pie(
        labels=model_counts['机型'], values=model_counts['数量'],
        hole=0.4, textinfo='label+percent',
    ), row=1, col=1)

    vip_counts = df['是否大客户'].value_counts()
    fig.add_trace(go.Pie(
        labels=['非大客户','大客户'],
        values=[vip_counts.get('否',0), vip_counts.get('是',0)],
        hole=0.4, textinfo='label+percent',
        marker_colors=['#1890FF','#E74C3C'],
    ), row=1, col=2)

    fig.update_layout(
        height=380,
        margin=dict(l=10, r=10, t=40, b=10),
    )
    return fig


def make_follower_chart(df):
    """跟进人工作量 TOP8"""
    counts = df['跟进人'].value_counts().nlargest(8).reset_index()
    counts.columns = ['跟进人','数量']
    counts = counts.sort_values('数量')

    colors = ['#E74C3C' if v >= 15 else '#F39C12' if v >= 10 else '#1890FF' for v in counts['数量']]

    fig = go.Figure(go.Bar(
        x=counts['数量'], y=counts['跟进人'],
        orientation='h',
        marker_color=colors,
        text=counts['数量'],
        textposition='outside',
    ))
    fig.update_layout(
        title='跟进人工作量 TOP8',
        xaxis_title='处理数量',
        height=300,
        margin=dict(l=10, r=30, t=40, b=10),
        bargap=0.3,
    )
    return fig


def make_d8_qa_chart(df):
    """8D报告覆盖率 + 品质负责人"""
    d8_total = len(df)
    d8_yes = len(df[df['8D报告'] == '有'])
    d8_rate = round(d8_yes / d8_total * 100, 1) if d8_total > 0 else 0

    fig = make_subplots(
        rows=1, cols=2,
        specs=[[{'type': 'indicator'}, {'type': 'pie'}]],
        column_widths=[0.4, 0.6],
        subplot_titles=('8D报告覆盖率', '品质负责人分布'),
    )

    fig.add_trace(go.Indicator(
        mode='gauge+number+delta',
        value=d8_rate,
        number={'suffix': '%', 'font': {'size': 36}},
        delta={'reference': 80, 'relative': False},
        gauge={
            'axis': {'range': [0, 100], 'tickwidth': 1},
            'bar': {'color': '#1890FF'},
            'steps': [
                {'range': [0, 60], 'color': 'rgba(231,76,60,0.3)'},
                {'range': [60, 85], 'color': 'rgba(243,156,18,0.3)'},
                {'range': [85, 100], 'color': 'rgba(39,174,96,0.3)'},
            ],
            'threshold': {
                'line': {'color': '#E74C3C', 'width': 2},
                'thickness': 0.8, 'value': 80
            }
        },
        title={'text': f'已出具: {d8_yes}/{d8_total}'},
    ), row=1, col=1)

    qa_counts = df['品质负责人'].value_counts().reset_index()
    qa_counts.columns = ['负责人','数量']

    fig.add_trace(go.Pie(
        labels=qa_counts['负责人'], values=qa_counts['数量'],
        hole=0.5, textinfo='label+percent',
    ), row=1, col=2)

    fig.update_layout(
        height=340,
        margin=dict(l=10, r=10, t=40, b=10),
    )
    return fig


def make_cycle_chart(df):
    """完成周期分布"""
    cycle = df['完成周期（天）'].dropna()
    labels = ['1-3天','4-6天','7-9天','10天以上']
    values = [
        int((cycle <= 3).sum()),
        int(((cycle > 3) & (cycle <= 6)).sum()),
        int(((cycle > 6) & (cycle <= 9)).sum()),
        int((cycle > 9).sum()),
    ]

    fig = go.Figure(go.Bar(
        x=labels, y=values,
        marker_color=['#27AE60','#1890FF','#F39C12','#E74C3C'],
        text=values, textposition='outside',
    ))
    fig.update_layout(
        title=f'完成周期分布（平均 {cycle.mean():.1f} 天，中位数 {cycle.median():.0f} 天）',
        xaxis_title='周期',
        yaxis_title='条数',
        height=300,
        margin=dict(l=10, r=10, t=40, b=10),
    )
    return fig


# ============================================================
# Main App
# ============================================================
def generate_export_zip(df, now):
    """生成包含全部分析维度 + 明细数据的 ZIP 文件"""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:

        # --- 01 客诉明细数据 ---
        detail_cols = {
            '编号_int': '编号', '分公司': '分公司', '国家或地区': '国家', '是否大客户': '大客户',
            '投诉日期': '投诉日期', '应结案日期': '应结案日期', '实际完成日期': '实际完成日期',
            '完成周期（天）': '完成周期(天)', '机型属性': '机型', '故障大类': '故障大类',
            '结案状态': '结案状态', '处理类型': '处理类型', '跟进人': '跟进人',
            '品质负责人': '品质负责人', '8D报告': '8D报告', '问题描述': '问题描述',
        }
        detail = df[list(detail_cols.keys())].copy()
        detail.columns = list(detail_cols.values())
        for dc in ['投诉日期','应结案日期','实际完成日期']:
            if dc in detail.columns:
                detail[dc] = detail[dc].apply(lambda x: x.strftime('%Y-%m-%d') if pd.notna(x) and hasattr(x,'strftime') else '')
        zf.writestr('01_客诉明细数据.csv', detail.to_csv(index=False, encoding='utf-8-sig'))

        # --- 02 分公司分布 ---
        branch = df['分公司'].value_counts().reset_index()
        branch.columns = ['分公司','投诉数量']
        branch['占比(%)'] = (branch['投诉数量'] / len(df) * 100).round(1)
        zf.writestr('02_分公司分布.csv', branch.to_csv(index=False, encoding='utf-8-sig'))

        # --- 03 国家地区TOP ---
        country = df['国家或地区'].value_counts().reset_index()
        country.columns = ['国家或地区','投诉数量']
        country['占比(%)'] = (country['投诉数量'] / len(df) * 100).round(1)
        zf.writestr('03_国家地区统计.csv', country.to_csv(index=False, encoding='utf-8-sig'))

        # --- 04 月度趋势 ---
        monthly = df.groupby(df['投诉日期'].dt.to_period('M').astype(str)).agg(
            投诉量=('编号_int','count'),
            已结案量=('结案状态', lambda x: x.isin(['结案','关闭']).sum()),
            未结案量=('结案状态', lambda x: (~x.isin(['结案','关闭'])).sum()),
        ).reset_index()
        monthly.columns = ['月份','投诉量','已结案量','未结案量']
        monthly['结案率(%)'] = (monthly['已结案量'] / monthly['投诉量'] * 100).round(1)
        zf.writestr('04_月度趋势.csv', monthly.to_csv(index=False, encoding='utf-8-sig'))

        # --- 05 结案状态 ---
        status = df['结案状态'].value_counts().reset_index()
        status.columns = ['结案状态','数量']
        status['占比(%)'] = (status['数量'] / len(df) * 100).round(1)
        zf.writestr('05_结案状态分布.csv', status.to_csv(index=False, encoding='utf-8-sig'))

        # --- 06 故障大类帕累托 ---
        fault = df['故障大类'].value_counts().reset_index()
        fault.columns = ['故障大类','数量']
        fault = fault.sort_values('数量', ascending=False)
        fault['占比(%)'] = (fault['数量'] / fault['数量'].sum() * 100).round(1)
        fault['累计占比(%)'] = fault['占比(%)'].cumsum().round(1)
        zf.writestr('06_故障大类帕累托.csv', fault.to_csv(index=False, encoding='utf-8-sig'))

        # --- 07 机型属性 ---
        model = df['机型属性'].value_counts().reset_index()
        model.columns = ['机型属性','数量']
        model['占比(%)'] = (model['数量'] / len(df) * 100).round(1)
        zf.writestr('07_机型属性分布.csv', model.to_csv(index=False, encoding='utf-8-sig'))

        # --- 08 跟进人工作量 ---
        follower = df['跟进人'].value_counts().reset_index()
        follower.columns = ['跟进人','处理数量']
        follower['占比(%)'] = (follower['处理数量'] / len(df) * 100).round(1)
        zf.writestr('08_跟进人工作量.csv', follower.to_csv(index=False, encoding='utf-8-sig'))

        # --- 09 品质负责人 ---
        qa = df['品质负责人'].value_counts().reset_index()
        qa.columns = ['品质负责人','负责数量']
        qa['占比(%)'] = (qa['负责数量'] / len(df) * 100).round(1)
        zf.writestr('09_品质负责人分布.csv', qa.to_csv(index=False, encoding='utf-8-sig'))

        # --- 10 超期未结案 ---
        overdue_mask = ~df['结案状态'].isin(['结案','关闭']) & df['应结案日期'].notna() & (df['应结案日期'] < now)
        overdue = df[overdue_mask][['编号_int','国家或地区','投诉日期','应结案日期','结案状态','跟进人','问题描述']].copy()
        if len(overdue) > 0:
            overdue['超期天数'] = (now - overdue['应结案日期']).dt.days
            overdue.columns = ['编号','国家','投诉日期','应结案日期','结案状态','跟进人','问题描述','超期天数']
            for dc in ['投诉日期','应结案日期']:
                overdue[dc] = overdue[dc].apply(lambda x: x.strftime('%Y-%m-%d') if pd.notna(x) and hasattr(x,'strftime') else '')
        else:
            overdue = pd.DataFrame({'提示': ['无超期未结案记录']})
        zf.writestr('10_超期未结案预警.csv', overdue.to_csv(index=False, encoding='utf-8-sig'))

        # --- 11 分公司×结案状态交叉表 ---
        cross = pd.crosstab(df['分公司'], df['结案状态'])
        cross['合计'] = cross.sum(axis=1)
        cross = cross.sort_values('合计', ascending=False)
        zf.writestr('11_分公司×结案状态交叉表.csv', cross.to_csv(encoding='utf-8-sig'))

        # --- 12 故障大类×机型交叉表 ---
        cross2 = pd.crosstab(df['故障大类'], df['机型属性'])
        cross2['合计'] = cross2.sum(axis=1)
        cross2 = cross2.sort_values('合计', ascending=False)
        zf.writestr('12_故障大类×机型交叉表.csv', cross2.to_csv(encoding='utf-8-sig'))

        # --- 13 8D报告统计 ---
        d8 = df['8D报告'].value_counts().reset_index()
        d8.columns = ['8D报告状态','数量']
        d8['占比(%)'] = (d8['数量'] / len(df) * 100).round(1)
        zf.writestr('13_8D报告覆盖率.csv', d8.to_csv(index=False, encoding='utf-8-sig'))

        # --- 14 完成周期统计 ---
        cycle = df['完成周期（天）'].dropna()
        cycle_stats = pd.DataFrame({
            '指标': ['记录数','平均值','中位数','最小值','最大值','标准差'],
            '数值': [
                len(cycle),
                round(cycle.mean(), 1),
                round(cycle.median(), 1),
                int(cycle.min()),
                int(cycle.max()),
                round(cycle.std(), 1),
            ]
        })
        zf.writestr('14_完成周期统计.csv', cycle_stats.to_csv(index=False, encoding='utf-8-sig'))

    buf.seek(0)
    return buf


def main():
    # Load data
    df = load_data()
    now = datetime.now()

    # ---- Sidebar ----
    with st.sidebar:
        st.image("https://img.icons8.com/color/96/bar-chart.png", width=48)
        st.markdown("## 📊 筛选器")

        # Data info
        st.caption(f"数据文件更新时间: {datetime.fromtimestamp(os.path.getmtime(os.path.join(os.path.dirname(__file__), '2026年海外客户投诉台账.xlsx'))).strftime('%Y-%m-%d %H:%M')}")

        filtered = apply_filters(df)

        st.divider()
        st.caption(f"筛选后记录数: **{len(filtered)}** / {len(df)}")
        st.caption(f"数据生成时间: {now.strftime('%Y-%m-%d %H:%M:%S')}")

        st.divider()
        st.markdown("### 📥 导出数据分析")

        export_zip = generate_export_zip(filtered, now)
        st.download_button(
            label="⬇️ 一键导出全部分析数据 (ZIP)",
            data=export_zip,
            file_name=f"海外客诉分析数据_{now.strftime('%Y%m%d_%H%M')}.zip",
            mime="application/zip",
            use_container_width=True,
        )
        st.caption("包含 14 个CSV文件：明细数据 + 各维度统计 + 交叉表 + 超期预警")

    # ---- Main Area ----
    # Header
    st.markdown("""
    <div style="background:linear-gradient(135deg,#1a1a2e,#16213e,#0f3460);padding:20px 30px;border-radius:12px;margin-bottom:16px;">
        <h1 style="color:#fff;margin:0;font-size:26px;">📊 2026年海外客户投诉数据看板</h1>
        <p style="color:#aaa;margin:4px 0 0;font-size:13px;">Overseas Customer Complaint Dashboard · 基于 Streamlit + Plotly 构建</p>
    </div>
    """, unsafe_allow_html=True)

    # ---- KPI Row ----
    total = len(filtered)
    done = len(filtered[filtered['结案状态'].isin(['结案','关闭'])])
    rate = round(done / total * 100, 1) if total > 0 else 0
    pending = len(filtered[~filtered['结案状态'].isin(['结案','关闭'])])
    cycle_avg = round(filtered['完成周期（天）'].dropna().mean(), 1)
    overdue = len(filtered[~filtered['结案状态'].isin(['结案','关闭']) &
                           filtered['应结案日期'].notna() &
                           (filtered['应结案日期'] < now)])

    k1, k2, k3, k4, k5 = st.columns(5)
    with k1:
        st.metric("📋 总投诉量", f"{total} 条")
    with k2:
        st.metric("✅ 结案率", f"{rate}%", delta=f"结案{len(filtered[filtered['结案状态']=='结案'])}条" if total > 0 else None)
    with k3:
        st.metric("⏱️ 平均处理周期", f"{cycle_avg} 天")
    with k4:
        st.metric("⚠️ 未结案", f"{pending} 条", delta_color="inverse")
    with k5:
        st.metric("🔴 超期未结案", f"{overdue} 条", delta_color="inverse")

    st.divider()

    # ---- Tabs ----
    tab1, tab2, tab3, tab4 = st.tabs(["📊 概览仪表盘", "🔍 故障分析", "👤 质量管理", "📋 数据明细"])

    # ============ TAB 1: 概览仪表盘 ============
    with tab1:
        col1, col2 = st.columns(2)
        with col1:
            st.plotly_chart(make_branch_chart(filtered), use_container_width=True, key='tab1_branch')
        with col2:
            st.plotly_chart(make_country_chart(filtered), use_container_width=True, key='tab1_country')

        col3, col4 = st.columns(2)
        with col3:
            st.plotly_chart(make_monthly_trend(filtered), use_container_width=True, key='tab1_monthly')
        with col4:
            st.plotly_chart(make_status_pie(filtered), use_container_width=True, key='tab1_status')

    # ============ TAB 2: 故障分析 ============
    with tab2:
        col1, col2 = st.columns(2)
        with col1:
            st.plotly_chart(make_fault_pareto(filtered), use_container_width=True, key='tab2_pareto')
        with col2:
            st.plotly_chart(make_model_vip_chart(filtered), use_container_width=True, key='tab2_model')

        # Fault × Model cross table
        st.markdown("#### 故障大类 × 机型属性 交叉矩阵")
        cross = pd.crosstab(filtered['故障大类'], filtered['机型属性'])
        cross['合计'] = cross.sum(axis=1)
        cross_sorted = cross.sort_values('合计', ascending=False)
        st.dataframe(cross_sorted, use_container_width=True, height=250)

    # ============ TAB 3: 质量管理 ============
    with tab3:
        col1, col2 = st.columns(2)
        with col1:
            st.plotly_chart(make_follower_chart(filtered), use_container_width=True, key='tab3_follower')
        with col2:
            st.plotly_chart(make_cycle_chart(filtered), use_container_width=True, key='tab3_cycle')

        col3, col4 = st.columns(2)
        with col3:
            st.plotly_chart(make_d8_qa_chart(filtered), use_container_width=True, key='tab3_d8')

        with col4:
            st.markdown("#### 🚨 超期未结案预警")
            overdue_df = filtered[~filtered['结案状态'].isin(['结案','关闭']) &
                                  filtered['应结案日期'].notna() &
                                  (filtered['应结案日期'] < now)]

            if len(overdue_df) > 0:
                for _, r in overdue_df.iterrows():
                    days = (now - r['应结案日期']).days
                    severity = 'error' if days > 30 else 'warning'
                    with st.container():
                        getattr(st, severity)(
                            f"**编号{r['编号_int']}** | {r['国家或地区']} | "
                            f"应结案: {r['应结案日期'].strftime('%Y-%m-%d')} | "
                            f"超期 **{days} 天**\n\n"
                            f"{str(r['问题描述'])[:120]}..."
                        )
            else:
                st.success("✅ 当前无超期未结案记录")

        # Branch × Status cross table
        st.markdown("#### 分公司 × 结案状态 明细表")
        cross2 = pd.crosstab(filtered['分公司'], filtered['结案状态'])
        cross2['合计'] = cross2.sum(axis=1)
        st.dataframe(cross2, use_container_width=True, height=280)

    # ============ TAB 4: 数据明细 ============
    with tab4:
        # Search & filter row
        search_col1, search_col2, search_col3 = st.columns([3, 1, 1])
        with search_col1:
            search = st.text_input("🔍 全局搜索", placeholder="输入编号/国家/故障/跟进人...")
        with search_col2:
            fault_filter = st.selectbox("故障大类筛选", options=['全部'] + sorted(filtered['故障大类'].dropna().unique().tolist()))
        with search_col3:
            status_filter = st.selectbox("结案状态筛选", options=['全部'] + sorted(filtered['结案状态'].dropna().unique().tolist()))

        # Apply filters
        display = filtered.copy()
        if search:
            mask = (
                display['编号_int'].astype(str).str.contains(search, case=False) |
                display['国家或地区'].str.contains(search, case=False, na=False) |
                display['故障大类'].str.contains(search, case=False, na=False) |
                display['跟进人'].str.contains(search, case=False, na=False) |
                display['分公司'].str.contains(search, case=False, na=False) |
                display['问题描述'].str.contains(search, case=False, na=False)
            )
            display = display[mask]
        if fault_filter != '全部':
            display = display[display['故障大类'] == fault_filter]
        if status_filter != '全部':
            display = display[display['结案状态'] == status_filter]

        st.caption(f"显示 {len(display)} / {len(filtered)} 条记录")

        # Build display table
        table_cols = {
            '编号_int': '编号',
            '分公司': '分公司',
            '国家或地区': '国家',
            '是否大客户': '大客户',
            '投诉日期': '投诉日期',
            '应结案日期': '应结案',
            '实际完成日期': '实际完成',
            '完成周期（天）': '周期',
            '机型属性': '机型',
            '故障大类': '故障大类',
            '结案状态': '状态',
            '处理类型': '类型',
            '跟进人': '跟进人',
            '品质负责人': '品质负责人',
            '8D报告': '8D',
            '问题描述': '问题描述',
        }

        show = display[list(table_cols.keys())].copy()
        show.columns = list(table_cols.values())
        # Format dates
        for dc in ['投诉日期','应结案','实际完成']:
            if dc in show.columns:
                show[dc] = show[dc].apply(lambda x: x.strftime('%Y-%m-%d') if pd.notna(x) and hasattr(x, 'strftime') else str(x)[:10])

        st.dataframe(
            show,
            use_container_width=True,
            height=600,
            hide_index=True,
            column_config={
                '编号': st.column_config.NumberColumn('编号', width='small'),
                '大客户': st.column_config.TextColumn('大客户', width='small'),
                '周期': st.column_config.NumberColumn('周期(天)', width='small'),
                '问题描述': st.column_config.TextColumn('问题描述', width='large'),
            }
        )

        # Export
        csv = show.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="📥 导出为 CSV",
            data=csv,
            file_name=f"海外客诉数据_{now.strftime('%Y%m%d')}.csv",
            mime='text/csv',
        )

    # ---- Footer ----
    st.divider()
    st.caption(f"© 2026 海外客户投诉数据看板 · 数据更新: {now.strftime('%Y-%m-%d %H:%M')} · Powered by Streamlit")


if __name__ == '__main__':
    main()
