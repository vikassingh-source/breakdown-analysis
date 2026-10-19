import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
import io
import re
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.drawing.image import Image as XLImage
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

st.set_page_config(page_title="Breakdown Analysis", layout="wide")
st.markdown("<h1 style='text-align: center;'>Dyeing Department</h1>",unsafe_allow_html=True)

st.title("🔧 Breakdown & Production Loss Analysis Dashboard")
st.markdown(
    """
    <div style="background-color:#1F4E78; padding:10px 15px; border-radius:8px; color:white; margin-bottom:15px;">
        <b>👤 Team Leader:</b> JC Roy(Mechanical) &nbsp;|&nbsp
        <b>👤 Team Member:</b> Vikas Singh(Instrumentation) &nbsp;|&nbsp
        <b>👤 Team Member:</b> Ajit Basha(Electrical) &nbsp;|&nbsp
        <b>👤 Team Member:</b> Bhaskar Ghormare(Dyeing) &nbsp;|&nbsp
        <b>📅 Version:</b> 2.0
    </div>
    """,
    unsafe_allow_html=True
)

# ============================================================
# HELPERS
# ============================================================
MONTH_MAP = {
    'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
    'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12,
    'january': 1, 'february': 2, 'march': 3, 'april': 4, 'june': 6,
    'july': 7, 'august': 8, 'september': 9, 'october': 10,
    'november': 11, 'december': 12
}

MONTH_NAMES = {
    '01': 'Jan', '02': 'Feb', '03': 'Mar', '04': 'Apr',
    '05': 'May', '06': 'Jun', '07': 'Jul', '08': 'Aug',
    '09': 'Sep', '10': 'Oct', '11': 'Nov', '12': 'Dec'
}


def parse_period(filename):
    name = filename.lower().replace('.xlsx', '').replace('.xls', '')
    pattern = r'(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|january|february|march|april|june|july|august|september|october|november|december)[\s_\-]*(\d{2,4})'
    m = re.search(pattern, name)
    if m:
        month = MONTH_MAP.get(m.group(1)[:3], 0)
        year = int(m.group(2))
        if year < 100:
            year += 2000
        label = MONTH_NAMES.get(f"{month:02d}", m.group(1).title()) + "-" + str(year)[-2:]
        return f"{year}-{month:02d}", label
    m = re.search(r'(\d{4})[\s_\-]?(\d{2})', name)
    if m:
        year, month = int(m.group(1)), int(m.group(2))
        label = MONTH_NAMES.get(f"{month:02d}", str(month)) + "-" + str(year)
        return f"{year}-{month:02d}", label
    return name, name


def map_type(value):
    v = str(value).lower().strip()
    if 'inst' in v:
        return 'Instrumentation'
    if 'elec' in v or 'electric' in v:
        return 'Electrical'
    if 'mech' in v:
        return 'Mechanical'
    return str(value).title()


def read_breakdown_file(uploaded):
    file_ext = uploaded.name.split('.')[-1].lower()
    engine = 'xlrd' if file_ext == 'xls' else 'openpyxl'
    sheets = pd.read_excel(uploaded, sheet_name=None, engine=engine)

    frames = []
    for sheet_name, df in sheets.items():
        df.columns = [str(c).replace('.', '').strip() for c in df.columns]
        if 'Worked By' in df.columns:
            df["Breakdown Type"] = df["Worked By"].apply(map_type)
        else:
            df["Breakdown Type"] = map_type(sheet_name)
        frames.append(df)

    data = pd.concat(frames, ignore_index=True)
    if 'M/C Name' in data.columns:
        data['M/C Name'] = data['M/C Name'].ffill()
    if 'Date' in data.columns:
        data['Date'] = pd.to_datetime(data['Date'], errors='coerce').ffill()
        # NEW:  Calculate Fiscal Year (April to March)
        data['FY_Year'] = data['Date'].dt.year + (data['Date'].dt.month >=4).astype(int)
        data['FY_label'] = data['FY_Year'].apply(lambda y: f"{y-1}-{str(y)[-2:]}")
    data['Total Hrs'] = pd.to_numeric(data['Total Hrs'], errors='coerce').fillna(0)
    data['Production Loss'] = pd.to_numeric(data['Production Loss'], errors='coerce').fillna(0)
    data = data.dropna(subset=['M/C Name', 'Reason of Breakdown'])

    period_key, period_label = parse_period(uploaded.name)
    data['Period_Key'] = period_key
    data['Period'] = period_label
    data['Source File'] = uploaded.name
    return data


def fig_to_png(fig, width=900, height=500):
    try:
        return pio.to_image(fig, format='png', width=width, height=height, scale=1.5)
    except Exception:
        return None


def insert_chart(ws, fig, cell, width=650, height=350):
    png_bytes = fig_to_png(fig)
    if png_bytes is None:
        return False
    img_stream = io.BytesIO(png_bytes)
    img = XLImage(img_stream)
    img.width = width
    img.height = height
    ws.add_image(img, cell)
    return True


# ============================================================
# SIDEBAR
# ============================================================
st.sidebar.header("📤 Upload Files")

breakdown_files = st.sidebar.file_uploader(
    "Breakdown Excel file(s) — one per month",
    type=["xlsx", "xls"], accept_multiple_files=True, key="breakdown"
)

machines_file = st.sidebar.file_uploader(
    "Master Machine List (machines.xlsx)",
    type=["xlsx", "xls"], key="machines"
)

calendar_files = st.sidebar.file_uploader(
    "Production Calendar(s) — one per month",
    type=["xlsx", "xls"], accept_multiple_files=True, key="calendars"
)

st.sidebar.markdown("---")
if st.sidebar.button("Clear All Uploaded Files", use_container_width=True):
    for key in ["breakdown", "machines", "calendars"]:
        if key in st.session_state:
            del st.session_state[key]
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.markdown("### 👤 Created by")
st.sidebar.markdown("**Vikas Singh**")
st.sidebar.caption("Instrumentation Dept.")

# ============================================================
# MAIN
# ============================================================
if breakdown_files:
    all_data = []
    for f in breakdown_files:
        try:
            all_data.append(read_breakdown_file(f))
        except Exception as e:
            st.error(f"❌ Could not read {f.name}: {e}")

    if not all_data:
        st.stop()

    data = pd.concat(all_data, ignore_index=True).sort_values('Period_Key')

    master_machines = []
    if machines_file:
        try:
            mdf = pd.read_excel(machines_file)
            mdf.columns = [str(c).strip() for c in mdf.columns]
            col = 'Machine Name' if 'Machine Name' in mdf.columns else mdf.columns[0]
            master_machines = sorted(mdf[col].dropna().astype(str).unique().tolist())
            st.sidebar.success(f"✅ {len(master_machines)} machines in master list")
        except Exception as e:
            st.sidebar.error(f"❌ Machine list error: {e}")

    calendar_data = {}
    if calendar_files:
        for cf in calendar_files:
            try:
                cdf = pd.read_excel(cf)
                cdf.columns = [str(c).strip() for c in cdf.columns]
                date_col = None
                hr_col = None
                for c in cdf.columns:
                    if 'date' in c.lower() and date_col is None:
                        date_col = c
                    if ('hour' in c.lower() or 'avail' in c.lower()) and hr_col is None:
                        hr_col = c
                if date_col is None or hr_col is None:
                    continue
                cdf[date_col] = pd.to_datetime(cdf[date_col], errors='coerce')
                cdf = cdf.dropna(subset=[date_col])
                cdf['Period_Key'] = cdf[date_col].dt.strftime('%Y-%m')
                calendar_data.update(cdf.groupby('Period_Key')[hr_col].sum().to_dict())
            except Exception as e:
                st.sidebar.warning(f"⚠️ Calendar {cf.name}: {e}")

    st.success(f"✅ Loaded {len(data)} records from {len(breakdown_files)} file(s): "
               f"{', '.join(sorted(data['Period'].unique()))}")

    # ---- SIDEBAR FILTERS ----
    st.sidebar.header("🔍 Filters")
    periods = sorted(data['Period'].unique())
    sel_periods = st.sidebar.multiselect("Period", periods, default=periods)
    machines = sorted(data['M/C Name'].astype(str).unique())
    sel_machines = st.sidebar.multiselect("Machine", machines, default=machines)
    types = sorted(data['Breakdown Type'].unique())
    sel_types = st.sidebar.multiselect("Breakdown Type", types, default=types)
    shifts = sorted(data['Shift'].dropna().astype(str).unique())
    sel_shifts = st.sidebar.multiselect("Shift", shifts, default=shifts)

    df = data[
        data['Period'].isin(sel_periods)
        & data['M/C Name'].astype(str).isin(sel_machines)
        & data['Breakdown Type'].isin(sel_types)
        & data['Shift'].astype(str).isin(sel_shifts)
    ]

    df['Year'] = df['Period_Key'].str[:4]
    df['Month_Num'] = df['Period_Key'].str[5:7]
    df['Month_Name'] = df['Month_Num'].map(MONTH_NAMES)

    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
        "📊 Overview", "📅 Month Comparison", "📈 Year Comparison",
        "🏆 Yearly Analysis", "🏭 Capacity Analysis",
        "🔧 Machine-wise", "🔁 Causes","📅 YTD (Fiscal Year)"
    ])

    # ---------- TAB 1: OVERVIEW ----------
    with tab1:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Breakdown Hours", f"{df['Total Hrs'].sum():,.1f}")
        c2.metric("Total Production Loss", f"{df['Production Loss'].sum():,.0f}")
        c3.metric("Total Breakdown Events", len(df))
        c4.metric("Avg Hours/Event", f"{df['Total Hrs'].mean():.2f}" if len(df) else "0")
        st.divider()

        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Breakdown Hours by Type")
            pie = df.groupby('Breakdown Type')['Total Hrs'].sum().reset_index()
            if len(pie) > 0:
                fig = px.pie(pie, names='Breakdown Type', values='Total Hrs', hole=0.4,
                             color_discrete_sequence=px.colors.qualitative.Set2)
                st.plotly_chart(fig, use_container_width=True)
        with col2:
            st.subheader("Production Loss by Type")
            pie2 = df.groupby('Breakdown Type')['Production Loss'].sum().reset_index()
            if len(pie2) > 0:
                fig = px.pie(pie2, names='Breakdown Type', values='Production Loss', hole=0.4,
                             color_discrete_sequence=px.colors.qualitative.Pastel)
                st.plotly_chart(fig, use_container_width=True)

        st.subheader("📅 Monthly Trend — Hours & Loss")
        trend = df.groupby('Period_Key').agg(
            Period=('Period', 'first'), Hours=('Total Hrs', 'sum'),
            Loss=('Production Loss', 'sum'), Events=('Total Hrs', 'count')
        ).reset_index().sort_values('Period_Key')
        if len(trend) > 0:
            fig = go.Figure()
            fig.add_trace(go.Bar(x=trend['Period'], y=trend['Hours'],
                                 name='Breakdown Hours', marker_color='#1F4E78'))
            fig.add_trace(go.Scatter(x=trend['Period'], y=trend['Loss'],
                                     name='Production Loss', mode='lines+markers',
                                     line=dict(color='#E74C3C', width=3), yaxis='y2'))
            fig.update_layout(yaxis=dict(title='Hours'),
                              yaxis2=dict(title='Loss', overlaying='y', side='right'),
                              hovermode='x unified', height=450)
            st.plotly_chart(fig, use_container_width=True)

        st.subheader("📋 Detailed Records (Top 200)")
        st.dataframe(df.head(200), use_container_width=True, hide_index=True)

    # ---------- TAB 2: MONTH COMPARISON ----------
    with tab2:
        if len(sel_periods) < 2:
            st.info("👆 Upload at least 2 files to compare months.")
        else:
            st.subheader("📅 Month-over-Month Comparison")
            p1, p2 = sel_periods[0], sel_periods[-1]
            df1, df2 = df[df['Period'] == p1], df[df['Period'] == p2]
            hours1, hours2 = df1['Total Hrs'].sum(), df2['Total Hrs'].sum()
            loss1, loss2 = df1['Production Loss'].sum(), df2['Production Loss'].sum()
            ev1, ev2 = len(df1), len(df2)
            dh = ((hours2 - hours1) / hours1 * 100) if hours1 else 0
            dl = ((loss2 - loss1) / loss1 * 100) if loss1 else 0
            de = ((ev2 - ev1) / ev1 * 100) if ev1 else 0
            c1, c2, c3 = st.columns(3)
            c1.metric(f"Hours ({p2} vs {p1})", f"{hours2:,.1f}", f"{dh:+.1f}%")
            c2.metric(f"Loss ({p2} vs {p1})", f"{loss2:,.0f}", f"{dl:+.1f}%")
            c3.metric(f"Events ({p2} vs {p1})", f"{ev2}", f"{de:+.1f}%")
            st.divider()
            cmp_type = df.groupby(['Period_Key', 'Period', 'Breakdown Type'])['Total Hrs'].sum().reset_index()
            if len(cmp_type) > 0:
                fig = px.bar(cmp_type, x='Breakdown Type', y='Total Hrs', color='Period',
                             barmode='group', text_auto='.1f')
                st.plotly_chart(fig, use_container_width=True)
            cmp_mach = df.groupby(['Period_Key', 'Period', 'M/C Name'])['Total Hrs'].sum().reset_index()
            if len(cmp_mach) > 0:
                fig = px.bar(cmp_mach, x='M/C Name', y='Total Hrs', color='Period',
                             barmode='group', text_auto='.1f')
                st.plotly_chart(fig, use_container_width=True)
            st.dataframe(df.pivot_table(index='M/C Name', columns='Period',
                                        values='Total Hrs', aggfunc='sum', fill_value=0),
                         use_container_width=True)
            st.dataframe(df.pivot_table(index='Breakdown Type', columns='Period',
                                        values='Total Hrs', aggfunc='sum', fill_value=0),
                         use_container_width=True)

    # ---------- TAB 3: YEAR COMPARISON ----------
    with tab3:
        years = sorted(df['Year'].unique())
        if len(years) < 2:
            st.info("👆 Upload files from 2 different years to compare.")
        else:
            st.subheader("📈 Year-over-Year Comparison")
            yr_totals = df.groupby('Year').agg(
                Hours=('Total Hrs', 'sum'), Loss=('Production Loss', 'sum'),
                Events=('Total Hrs', 'count')).reset_index()
            fig = px.bar(yr_totals, x='Year', y=['Hours', 'Loss'], barmode='group', text_auto='.1f')
            st.plotly_chart(fig, use_container_width=True)
            monthly = df.groupby(['Year', 'Month_Name'])['Total Hrs'].sum().reset_index()
            fig = px.bar(monthly, x='Month_Name', y='Total Hrs', color='Year',
                         barmode='group', text_auto='.1f',
                         category_orders={'Month_Name': list(MONTH_NAMES.values())})
            st.plotly_chart(fig, use_container_width=True)
            st.dataframe(df.pivot_table(index='Breakdown Type', columns='Year',
                                        values='Total Hrs', aggfunc='sum', fill_value=0),
                         use_container_width=True)
            st.dataframe(df.pivot_table(index='M/C Name', columns='Year',
                                        values='Total Hrs', aggfunc='sum', fill_value=0),
                         use_container_width=True)

    # ---------- TAB 4: YEARLY ANALYSIS ----------
    with tab4:
        years_available = sorted(df['Year'].unique())
        if len(years_available) == 0:
            st.info("Upload at least one file with a year in the filename.")
        else:
            selected_year = st.selectbox("Select Year", years_available, index=len(years_available) - 1)
            year_df = df[df['Year'] == selected_year]
            st.subheader(f"🏆 Total Year Summary — {selected_year}")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Total Hours", f"{year_df['Total Hrs'].sum():,.1f}")
            c2.metric("Total Loss", f"{year_df['Production Loss'].sum():,.0f}")
            c3.metric("Total Events", len(year_df))
            c4.metric("Avg Hours/Event", f"{year_df['Total Hrs'].mean():.2f}" if len(year_df) else "0")
            c5, c6, c7, c8 = st.columns(4)
            c5.metric("Machines", year_df['M/C Name'].nunique())
            c6.metric("Causes", year_df['Reason of Breakdown'].nunique())
            c7.metric("Worst Machine",
                      year_df.groupby('M/C Name')['Total Hrs'].sum().idxmax() if len(year_df) else "-")
            c8.metric("Top Type",
                      year_df.groupby('Breakdown Type')['Total Hrs'].sum().idxmax() if len(year_df) else "-")
            st.divider()
            month_table = year_df.groupby(['Month_Num', 'Month_Name']).agg(
                Breakdown_Hours=('Total Hrs', 'sum'),
                Production_Loss=('Production Loss', 'sum'),
                Events=('Total Hrs', 'count')).reset_index().sort_values('Month_Num')
            month_table['Cumulative_Hours'] = month_table['Breakdown_Hours'].cumsum()
            st.dataframe(month_table.drop(columns=['Month_Num']), use_container_width=True, hide_index=True)
            if len(month_table) > 0:
                fig = go.Figure()
                fig.add_trace(go.Bar(x=month_table['Month_Name'], y=month_table['Breakdown_Hours'],
                                     name='Monthly Hours', marker_color='#1F4E78'))
                fig.add_trace(go.Scatter(x=month_table['Month_Name'], y=month_table['Cumulative_Hours'],
                                         name='Cumulative', mode='lines+markers',
                                         line=dict(color='#E74C3C', width=3), yaxis='y2'))
                fig.update_layout(yaxis=dict(title='Monthly Hours'),
                                  yaxis2=dict(title='Cumulative', overlaying='y', side='right'),
                                  hovermode='x unified', height=450)
                st.plotly_chart(fig, use_container_width=True)
            month_order = [m for m in MONTH_NAMES.values()]
            st.dataframe(year_df.pivot_table(index='Month_Name', columns='Breakdown Type',
                                              values='Total Hrs', aggfunc='sum', fill_value=0).reindex(
                [m for m in month_order if m in year_df['Month_Name'].unique()]), use_container_width=True)
            year_totals = df.groupby('Year').agg(
                Breakdown_Hours=('Total Hrs', 'sum'), Production_Loss=('Production Loss', 'sum'),
                Events=('Total Hrs', 'count'), Machines=('M/C Name', 'nunique'),
                Causes=('Reason of Breakdown', 'nunique')).reset_index()
            st.dataframe(year_totals, use_container_width=True, hide_index=True)

    # ---------- TAB 5: CAPACITY ANALYSIS ----------
    with tab5:
        st.subheader("🏭 Plant Capacity & Availability Analysis")
        if not master_machines:
            st.warning("⚠️ Upload **machines.xlsx** in the sidebar for full capacity analysis.")
        period_options = sorted(df['Period_Key'].unique())
        sel_period_cap = st.selectbox("Select Period", period_options, index=len(period_options) - 1,
                                       format_func=lambda x: df[df['Period_Key'] == x]['Period'].iloc[0]
                                       if len(df[df['Period_Key'] == x]) else x)
        period_df = df[df['Period_Key'] == sel_period_cap]

        if sel_period_cap in calendar_data:
            available_hours_per_machine = calendar_data[sel_period_cap]
            st.success(f"✅ Calendar loaded: {available_hours_per_machine} hrs/machine")
        else:
            days_in_month = pd.Period(sel_period_cap).days_in_month
            available_hours_per_machine = 20 * days_in_month
            st.warning(f"⚠️ No calendar. Using default {available_hours_per_machine} hrs/machine")

        all_machines = master_machines if master_machines else sorted(period_df['M/C Name'].astype(str).unique())
        machines_with_bd = set(period_df['M/C Name'].astype(str).unique())
        total_machines = len(all_machines)
        with_bd_count = len([m for m in all_machines if m in machines_with_bd])
        without_bd_count = total_machines - with_bd_count

        plant_capacity = total_machines * available_hours_per_machine
        breakdown_hours = period_df['Total Hrs'].sum()
        production_loss = period_df['Production Loss'].sum()
        availability_pct = ((plant_capacity - breakdown_hours) / plant_capacity * 100) if plant_capacity else 0
        downtime_pct = (breakdown_hours / plant_capacity * 100) if plant_capacity else 0
        events = len(period_df)
        mttr = (breakdown_hours / events) if events else 0
        mtbf = ((plant_capacity - breakdown_hours) / events) if events else 0

        st.markdown("### 📊 Plant-wise KPIs")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Machines", total_machines)
        c2.metric("With Breakdown", f"{with_bd_count} ({with_bd_count/total_machines*100:.1f}%)" if total_machines else "0")
        c3.metric("Without Breakdown", f"{without_bd_count} ({without_bd_count/total_machines*100:.1f}%)" if total_machines else "0")
        c4.metric("Available Hrs/Machine", f"{available_hours_per_machine:,.0f}")
        c5, c6, c7, c8 = st.columns(4)
        c5.metric("Plant Capacity (hrs)", f"{plant_capacity:,.0f}")
        c6.metric("Breakdown Hours", f"{breakdown_hours:,.1f}")
        c7.metric("Availability %", f"{availability_pct:.2f}%")
        c8.metric("Downtime %", f"{downtime_pct:.2f}%")
        c9, c10, c11, c12 = st.columns(4)
        c9.metric("MTBF (hrs)", f"{mtbf:,.1f}" if events else "—")
        c10.metric("MTTR (hrs)", f"{mttr:.2f}" if events else "—")
        c11.metric("Production Loss", f"{production_loss:,.0f}")
        c12.metric("Loss per Breakdown Hr", f"{production_loss/breakdown_hours:.1f}" if breakdown_hours else "—")

        st.divider()
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Machines — Breakdown Status")
            pie_df = pd.DataFrame({"Status": ["With Breakdown", "Without Breakdown"],
                                   "Count": [with_bd_count, without_bd_count]})
            fig = px.pie(pie_df, names="Status", values="Count", hole=0.4, color="Status",
                         color_discrete_map={"With Breakdown": "#E74C3C", "Without Breakdown": "#27AE60"})
            st.plotly_chart(fig, use_container_width=True)
        with col2:
            pie2 = period_df.groupby('Breakdown Type')['Total Hrs'].sum().reset_index()
            if len(pie2) > 0:
                fig = px.pie(pie2, names='Breakdown Type', values='Total Hrs', hole=0.4)
                st.plotly_chart(fig, use_container_width=True)

        machine_stats = []
        for m in all_machines:
            m_df = period_df[period_df['M/C Name'].astype(str) == m]
            m_hours = m_df['Total Hrs'].sum()
            m_events = len(m_df)
            m_loss = m_df['Production Loss'].sum()
            m_avail = ((available_hours_per_machine - m_hours) / available_hours_per_machine * 100) if available_hours_per_machine else 100
            m_mtbf = ((available_hours_per_machine - m_hours) / m_events) if m_events else None
            m_mttr = (m_hours / m_events) if m_events else 0
            machine_stats.append({"Machine": m, "Breakdown Hours": round(m_hours, 2),
                                   "Events": m_events, "Production Loss": round(m_loss, 1),
                                   "Availability %": round(m_avail, 2),
                                   "MTBF (hrs)": round(m_mtbf, 1) if m_mtbf else "—",
                                   "MTTR (hrs)": round(m_mttr, 2) if m_events else "—"})
        machine_stats_df = pd.DataFrame(machine_stats).sort_values('Availability %')

        st.subheader("📋 Per-Machine Availability (Overall)")
        st.dataframe(machine_stats_df, use_container_width=True, hide_index=True)

        st.divider()
        st.subheader("🔬 MTBF & MTTR — Split by Breakdown Type")
        type_stats = []
        for m in all_machines:
            m_df = period_df[period_df['M/C Name'].astype(str) == m]
            for btype in ['Instrumentation', 'Electrical', 'Mechanical']:
                t_df = m_df[m_df['Breakdown Type'] == btype]
                t_hours = t_df['Total Hrs'].sum()
                t_events = len(t_df)
                t_loss = t_df['Production Loss'].sum()
                t_mttr = (t_hours / t_events) if t_events else None
                t_mtbf = ((available_hours_per_machine - t_hours) / t_events) if t_events else None
                type_stats.append({"Machine": m, "Breakdown Type": btype, "Events": t_events,
                                    "Breakdown Hours": round(t_hours, 2), "Production Loss": round(t_loss, 1),
                                    "MTTR (hrs)": round(t_mttr, 2) if t_mttr else "—",
                                    "MTBF (hrs)": round(t_mtbf, 1) if t_mtbf else "—"})
        type_stats_df = pd.DataFrame(type_stats)
        type_stats_display = type_stats_df[type_stats_df['Events'] > 0].copy().sort_values(
            ['Machine', 'Breakdown Type']).reset_index(drop=True)
        st.dataframe(type_stats_display, use_container_width=True, hide_index=True)

        if len(type_stats_display) > 0:
            col1, col2 = st.columns(2)
            with col1:
                fig = px.bar(type_stats_display, x='Machine', y='MTTR (hrs)',
                             color='Breakdown Type', barmode='group', text='MTTR (hrs)',
                             color_discrete_map={'Instrumentation': '#3498DB',
                                                 'Electrical': '#F39C12', 'Mechanical': '#E74C3C'})
                fig.update_layout(height=500, xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)
            with col2:
                fig = px.bar(type_stats_display, x='Machine', y='MTBF (hrs)',
                             color='Breakdown Type', barmode='group', text='MTBF (hrs)',
                             color_discrete_map={'Instrumentation': '#3498DB',
                                                 'Electrical': '#F39C12', 'Mechanical': '#E74C3C'})
                fig.update_layout(height=500, xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)

        st.divider()
        st.subheader("📊 Machine-wise MTBF & MTTR Analysis")
        chart_df = machine_stats_df[machine_stats_df['Events'] > 0].copy()
        if len(chart_df) > 0:
            st.markdown("### 🔧 MTTR per Machine — **High = Problematic**")
            mttr_chart = chart_df[chart_df['MTTR (hrs)'] != "—"].copy()
            if len(mttr_chart) > 0:
                mttr_chart['MTTR (hrs)'] = pd.to_numeric(mttr_chart['MTTR (hrs)'], errors='coerce').fillna(0)
                fig = px.bar(mttr_chart.sort_values('MTTR (hrs)', ascending=False),
                             x='Machine', y='MTTR (hrs)', color='MTTR (hrs)',
                             color_continuous_scale=['#27AE60', '#F39C12', '#E74C3C'], text_auto='.2f')
                fig.update_layout(height=500, xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)

            st.markdown("### ⏱️ MTBF per Machine — **Low = Problematic**")
            mtbf_chart = chart_df[chart_df['MTBF (hrs)'] != "—"].copy()
            if len(mtbf_chart) > 0:
                mtbf_chart['MTBF (hrs)'] = pd.to_numeric(mtbf_chart['MTBF (hrs)'], errors='coerce').fillna(0)
                fig = px.bar(mtbf_chart.sort_values('MTBF (hrs)', ascending=True),
                             x='Machine', y='MTBF (hrs)', color='MTBF (hrs)',
                             color_continuous_scale=['#E74C3C', '#F39C12', '#27AE60'], text_auto='.0f')
                fig.update_layout(height=500, xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)

    # ---------- TAB 6: MACHINE-WISE ----------
    with tab6:
        st.subheader("Machine-wise Breakdown Hours")
        mach = df.groupby('M/C Name').agg(Hours=('Total Hrs', 'sum'),
                                           Loss=('Production Loss', 'sum'),
                                           Events=('Total Hrs', 'count')).reset_index().sort_values(
            'Hours', ascending=False)
        st.dataframe(mach, use_container_width=True, hide_index=True)
        if len(mach) > 0:
            fig = px.bar(mach, x='M/C Name', y=['Hours', 'Loss'], barmode='group', text_auto='.1f')
            st.plotly_chart(fig, use_container_width=True)
        stack = df.groupby(['M/C Name', 'Breakdown Type'])['Total Hrs'].sum().reset_index()
        if len(stack) > 0:
            fig = px.bar(stack, x='M/C Name', y='Total Hrs', color='Breakdown Type', barmode='stack')
            st.plotly_chart(fig, use_container_width=True)
        shift_df = df.groupby(['Shift', 'Breakdown Type'])['Total Hrs'].sum().reset_index()
        if len(shift_df) > 0:
            fig = px.bar(shift_df, x='Shift', y='Total Hrs', color='Breakdown Type',
                         barmode='stack', text_auto='.1f')
            st.plotly_chart(fig, use_container_width=True)

    # ---------- TAB 7: CAUSES ----------
    with tab7:
        st.subheader("🔁 Top Repeated Causes")
        causes = df.groupby('Reason of Breakdown').agg(
            Frequency=('Reason of Breakdown', 'count'), Hours=('Total Hrs', 'sum'),
            Loss=('Production Loss', 'sum')).reset_index().sort_values('Frequency', ascending=False)
        st.dataframe(causes, use_container_width=True, hide_index=True)
        if len(causes) > 0:
            fig = px.bar(causes.head(10), x='Frequency', y='Reason of Breakdown',
                         orientation='h', color='Hours', text_auto=True)
            fig.update_layout(yaxis={'categoryorder': 'total ascending'}, height=500)
            st.plotly_chart(fig, use_container_width=True)
        if len(causes) > 0:
            top_causes = causes.head(15)['Reason of Breakdown'].tolist()
            pivot = df[df['Reason of Breakdown'].isin(top_causes)].pivot_table(
                index='M/C Name', columns='Reason of Breakdown',
                values='Total Hrs', aggfunc='count', fill_value=0)
            st.dataframe(pivot, use_container_width=True)
            # ---------- TAB 8: YTD (FISCAL YEAR) ----------
    with tab8:
       st.subheader("📆 YTD Analysis (Fiscal Year: April to March)")
    
    # Select Fiscal Year
    df['FY_Label'] = df['Date'].dt.year + (df['Date'].dt.month >=4).astype(int)
    df['FY_Label'] = df['FY_Label'].apply(lambda y: f"{int(y)-1}-{str(int(y))[-2:]}" if pd.notna(y) else "Unknown")
    fy_options = sorted(df['FY_Label'].dropna().unique())
    if len(fy_options) == 0:
        st.warning("No Fiscal Year data available. Make sure your Excel has a 'Date' column.")
    else:
        sel_fy = st.selectbox("Select Financial Year", fy_options, index=len(fy_options)-1)
        
        # Filter data for selected FY
        ytd_df = df[df['FY_Label'] == sel_fy].copy()
        
        # FY Month order (Apr to Mar)
        fy_months = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar']
        ytd_df['Month_Name'] = pd.Categorical(ytd_df['Month_Name'], categories=fy_months, ordered=True)
        
        # Group by month
        fy_trend = ytd_df.groupby('Month_Name', observed=False).agg(
            Hours=('Total Hrs', 'sum'),
            Loss=('Production Loss', 'sum'),
            Events=('Total Hrs', 'count')
        ).reset_index()
        
        # Drop months with zero events (future months)
        fy_trend = fy_trend[fy_trend['Events'] > 0]
        
        # Cumulative columns
        fy_trend['Cumulative_Hours'] = fy_trend['Hours'].cumsum()
        fy_trend['Cumulative_Loss'] = fy_trend['Loss'].cumsum()
        
        # KPI Cards for YTD
        st.markdown(f"### YTD Totals — FY {sel_fy}")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("YTD Breakdown Hours", f"{ytd_df['Total Hrs'].sum():,.1f}")
        c2.metric("YTD Production Loss", f"{ytd_df['Production Loss'].sum():,.0f} Kg")
        c3_val = ytd_df['Production Loss'].sum() * 21.0 # Using 21as default fallback
        c3.metric("YTD Cost of Loss", f"₹{c3_val:,.0f}")
        c4.metric("YTD Breakdown Events", len(ytd_df))
        
        st.divider()
        
        # Chart: Monthly + Cumulative
        st.subheader("📊 Monthly & Cumulative Trend (YTD)")
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=fy_trend['Month_Name'], y=fy_trend['Hours'],
            name='Monthly Hours', marker_color='#1F4E78',
            text=fy_trend['Hours'].round(1), textposition='outside'
        ))
        fig.add_trace(go.Scatter(
            x=fy_trend['Month_Name'], y=fy_trend['Cumulative_Hours'],
            name='Cumulative Hours', mode='lines+markers',
            line=dict(color='#E74C3C', width=3), yaxis='y2'
        ))
        fig.update_layout(
            yaxis=dict(title='Monthly Hours'),
            yaxis2=dict(title='Cumulative Hours', overlaying='y', side='right'),
            hovermode='x unified', height=450
        )
        st.plotly_chart(fig, use_container_width=True)
        
        # Table: Month-wise YTD
        st.subheader("📋 Month-wise YTD Breakdown")
        st.dataframe(fy_trend[['Month_Name', 'Hours', 'Loss', 'Events', 'Cumulative_Hours', 'Cumulative_Loss']], use_container_width=True, hide_index=True)
    # ============================================================
    # EXCEL EXPORT
    # ============================================================
    st.divider()
    st.subheader("⬇️ Export Full Analysis")

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(bold=True, color="FFFFFF")

    def style_sheet(ws):
        for column_cells in ws.columns:
            max_len = 0
            for c in column_cells:
                if c.value is not None:
                    max_len = max(max_len, len(str(c.value)))
            ws.column_dimensions[column_cells[0].column_letter].width = min(max_len + 3, 50)
        for cell in ws[1]:
            if cell.value:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

    def build_data_only_excel(df):
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            kpi = pd.DataFrame({
                "Metric": ["Total Hours", "Total Loss", "Events", "Avg Hours/Event",
                           "Periods", "Machines", "Files"],
                "Value": [round(df['Total Hrs'].sum(), 2), round(df['Production Loss'].sum(), 2),
                          len(df), round(df['Total Hrs'].mean(), 2) if len(df) else 0,
                          ", ".join(sorted(df['Period'].unique())), df['M/C Name'].nunique(),
                          ", ".join(sorted(df['Source File'].unique()))]})
            kpi.to_excel(writer, sheet_name="Overview", index=False)

            df.groupby('M/C Name').agg(Hours=('Total Hrs', 'sum'),
                                        Loss=('Production Loss', 'sum'),
                                        Events=('Total Hrs', 'count')).reset_index().sort_values(
                'Hours', ascending=False).to_excel(writer, sheet_name="Machine-wise", index=False)

            df.groupby('Reason of Breakdown').agg(Frequency=('Reason of Breakdown', 'count'),
                                                    Hours=('Total Hrs', 'sum'),
                                                    Loss=('Production Loss', 'sum')).reset_index().sort_values(
                'Frequency', ascending=False).to_excel(writer, sheet_name="Causes", index=False)

            df.groupby('Breakdown Type').agg(Hours=('Total Hrs', 'sum'),
                                              Loss=('Production Loss', 'sum'),
                                              Events=('Total Hrs', 'count')).reset_index().to_excel(
                writer, sheet_name="By Type", index=False)

            monthly_cmp = df.groupby('Period_Key').agg(
                Period=('Period', 'first'), Breakdown_Hours=('Total Hrs', 'sum'),
                Production_Loss=('Production Loss', 'sum'),
                Events=('Total Hrs', 'count')).reset_index().sort_values('Period_Key')
            monthly_cmp['Hours_Change_%'] = (monthly_cmp['Breakdown_Hours'].pct_change() * 100).round(2)
            monthly_cmp['Loss_Change_%'] = (monthly_cmp['Production_Loss'].pct_change() * 100).round(2)
            monthly_cmp.drop(columns=['Period_Key']).to_excel(
                writer, sheet_name="Monthly Comparison", index=False)

            yearly_cmp = df.groupby('Year').agg(
                Breakdown_Hours=('Total Hrs', 'sum'),
                Production_Loss=('Production Loss', 'sum'),
                Events=('Total Hrs', 'count'), Machines=('M/C Name', 'nunique'),
                Causes=('Reason of Breakdown', 'nunique')).reset_index().sort_values('Year')
            yearly_cmp['Hours_Change_%'] = (yearly_cmp['Breakdown_Hours'].pct_change() * 100).round(2)
            yearly_cmp['Loss_Change_%'] = (yearly_cmp['Production_Loss'].pct_change() * 100).round(2)
            yearly_cmp.to_excel(writer, sheet_name="Yearly Comparison", index=False)

            df.pivot_table(index='M/C Name', columns='Period', values='Total Hrs',
                           aggfunc='sum', fill_value=0).to_excel(writer, sheet_name="Machine x Period")
            df.to_excel(writer, sheet_name="Raw Data", index=False)

            for sheet in writer.sheets.values():
                style_sheet(sheet)
        buffer.seek(0)
        return buffer

    def build_full_excel(df, machine_stats_df, available_hours_per_machine,
                         plant_capacity, breakdown_hours, production_loss,
                         availability_pct, downtime_pct, mttr, mtbf,
                         total_machines, with_bd_count, without_bd_count, period_df):
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:

            # ---- Sheet 1: Overview ----
            kpi = pd.DataFrame({
                "Metric": ["Total Hours", "Total Loss", "Events", "Avg Hours/Event",
                           "Periods", "Machines", "Files"],
                "Value": [round(df['Total Hrs'].sum(), 2), round(df['Production Loss'].sum(), 2),
                          len(df), round(df['Total Hrs'].mean(), 2) if len(df) else 0,
                          ", ".join(sorted(df['Period'].unique())), df['M/C Name'].nunique(),
                          ", ".join(sorted(df['Source File'].unique()))]})
            kpi.to_excel(writer, sheet_name="1. Overview", index=False)
            ws = writer.sheets["1. Overview"]

            pie = df.groupby('Breakdown Type')['Total Hrs'].sum().reset_index()
            if len(pie) > 0:
                fig = px.pie(pie, names='Breakdown Type', values='Total Hrs', hole=0.4,
                             title='Breakdown Hours by Type',
                             color_discrete_sequence=px.colors.qualitative.Set2)
                insert_chart(ws, fig, 'D2', width=500, height=350)

            pie2 = df.groupby('Breakdown Type')['Production Loss'].sum().reset_index()
            if len(pie2) > 0:
                fig = px.pie(pie2, names='Breakdown Type', values='Production Loss', hole=0.4,
                             title='Production Loss by Type',
                             color_discrete_sequence=px.colors.qualitative.Pastel)
                insert_chart(ws, fig, 'K2', width=500, height=350)

            trend = df.groupby('Period_Key').agg(
                Period=('Period', 'first'), Hours=('Total Hrs', 'sum'),
                Loss=('Production Loss', 'sum')).reset_index().sort_values('Period_Key')
            if len(trend) > 0:
                fig = go.Figure()
                fig.add_trace(go.Bar(x=trend['Period'], y=trend['Hours'],
                                     name='Hours', marker_color='#1F4E78'))
                fig.add_trace(go.Scatter(x=trend['Period'], y=trend['Loss'],
                                         name='Loss', mode='lines+markers',
                                         line=dict(color='#E74C3C', width=3), yaxis='y2'))
                fig.update_layout(yaxis=dict(title='Hours'),
                                  yaxis2=dict(title='Loss', overlaying='y', side='right'),
                                  title='Monthly Trend', height=400)
                insert_chart(ws, fig, 'D21', width=700, height=400)

            # ---- Sheet 2: Capacity ----
            cap_kpi = pd.DataFrame({
                "Metric": ["Total Machines", "With Breakdown", "Without Breakdown",
                           "Available Hrs/Machine", "Plant Capacity", "Breakdown Hours",
                           "Availability %", "Downtime %", "MTBF (hrs)", "MTTR (hrs)",
                           "Production Loss", "Loss/Breakdown Hr"],
                "Value": [total_machines,
                          f"{with_bd_count} ({with_bd_count/total_machines*100:.1f}%)" if total_machines else "0",
                          f"{without_bd_count} ({without_bd_count/total_machines*100:.1f}%)" if total_machines else "0",
                          round(available_hours_per_machine, 0), round(plant_capacity, 0),
                          round(breakdown_hours, 1), f"{availability_pct:.2f}%",
                          f"{downtime_pct:.2f}%", round(mtbf, 1) if events else "—",
                          round(mttr, 2) if events else "—", round(production_loss, 0),
                          round(production_loss / breakdown_hours, 1) if breakdown_hours else "—"]})
            cap_kpi.to_excel(writer, sheet_name="2. Capacity", index=False)
            ws = writer.sheets["2. Capacity"]

            startrow = len(cap_kpi) + 3
            machine_stats_df.to_excel(writer, sheet_name="2. Capacity", startrow=startrow, index=False)

            chart_startrow = startrow + len(machine_stats_df) + 3
            pie_df = pd.DataFrame({"Status": ["With Breakdown", "Without Breakdown"],
                                   "Count": [with_bd_count, without_bd_count]})
            fig = px.pie(pie_df, names="Status", values="Count", hole=0.4,
                         title='Machines Status', color="Status",
                         color_discrete_map={"With Breakdown": "#E74C3C",
                                             "Without Breakdown": "#27AE60"})
            insert_chart(ws, fig, f'D{chart_startrow}', width=500, height=350)

            chart_df = machine_stats_df[machine_stats_df['Events'] > 0].copy()
            if len(chart_df) > 0:
                mttr_chart = chart_df[chart_df['MTTR (hrs)'] != "—"].copy()
                if len(mttr_chart) > 0:
                    mttr_chart['MTTR (hrs)'] = pd.to_numeric(mttr_chart['MTTR (hrs)'],
                                                              errors='coerce').fillna(0)
                    fig = px.bar(mttr_chart.sort_values('MTTR (hrs)', ascending=False),
                                 x='Machine', y='MTTR (hrs)', color='MTTR (hrs)',
                                 title='MTTR per Machine (High = Bad)',
                                 color_continuous_scale=['#27AE60', '#F39C12', '#E74C3C'])
                    fig.update_layout(xaxis_tickangle=-45, height=450)
                    insert_chart(ws, fig, f'D{chart_startrow + 20}', width=800, height=450)

                mtbf_chart = chart_df[chart_df['MTBF (hrs)'] != "—"].copy()
                if len(mtbf_chart) > 0:
                    mtbf_chart['MTBF (hrs)'] = pd.to_numeric(mtbf_chart['MTBF (hrs)'],
                                                              errors='coerce').fillna(0)
                    fig = px.bar(mtbf_chart.sort_values('MTBF (hrs)', ascending=True),
                                 x='Machine', y='MTBF (hrs)', color='MTBF (hrs)',
                                 title='MTBF per Machine (Low = Bad)',
                                 color_continuous_scale=['#E74C3C', '#F39C12', '#27AE60'])
                    fig.update_layout(xaxis_tickangle=-45, height=450)
                    insert_chart(ws, fig, f'D{chart_startrow + 45}', width=800, height=450)

            # ---- Sheet 3: Machine-wise ----
            mach = df.groupby('M/C Name').agg(Hours=('Total Hrs', 'sum'),
                                                Loss=('Production Loss', 'sum'),
                                                Events=('Total Hrs', 'count')).reset_index().sort_values(
                'Hours', ascending=False)
            mach.to_excel(writer, sheet_name="3. Machine-wise", index=False)
            ws = writer.sheets["3. Machine-wise"]
            if len(mach) > 0:
                fig = px.bar(mach.head(15), x='M/C Name', y=['Hours', 'Loss'], barmode='group',
                             title='Top 15 Machines')
                fig.update_layout(xaxis_tickangle=-45, height=450)
                insert_chart(ws, fig, f'A{len(mach) + 4}', width=900, height=450)

            # ---- Sheet 4: Causes ----
            causes = df.groupby('Reason of Breakdown').agg(
                Frequency=('Reason of Breakdown', 'count'), Hours=('Total Hrs', 'sum'),
                Loss=('Production Loss', 'sum')).reset_index().sort_values('Frequency', ascending=False)
            causes.to_excel(writer, sheet_name="4. Causes", index=False)
            ws = writer.sheets["4. Causes"]
            if len(causes) > 0:
                fig = px.bar(causes.head(10), x='Frequency', y='Reason of Breakdown',
                             orientation='h', color='Hours', title='Top 10 Causes')
                fig.update_layout(yaxis={'categoryorder': 'total ascending'}, height=500)
                insert_chart(ws, fig, f'A{len(causes) + 4}', width=800, height=500)

            # ---- Sheet 5: Monthly Comparison ----
            monthly_cmp = df.groupby('Period_Key').agg(
                Period=('Period', 'first'), Hours=('Total Hrs', 'sum'),
                Loss=('Production Loss', 'sum')).reset_index().sort_values('Period_Key')
            monthly_cmp['Hours_Change_%'] = (monthly_cmp['Hours'].pct_change() * 100).round(2)
            monthly_cmp.drop(columns=['Period_Key']).to_excel(
                writer, sheet_name="5. Monthly", index=False)
            ws = writer.sheets["5. Monthly"]
            if len(monthly_cmp) > 1:
                fig = px.bar(monthly_cmp, x='Period', y=['Hours', 'Loss'], barmode='group',
                             title='Monthly Comparison',
                             color_discrete_sequence=['#1F4E78', '#E74C3C'])
                insert_chart(ws, fig, f'A{len(monthly_cmp) + 4}', width=800, height=400)

            # ---- Sheet 6: Yearly Comparison ----
            yearly_cmp = df.groupby('Year').agg(
                Hours=('Total Hrs', 'sum'), Loss=('Production Loss', 'sum'),
                Events=('Total Hrs', 'count')).reset_index()
            yearly_cmp.to_excel(writer, sheet_name="6. Yearly", index=False)
            ws = writer.sheets["6. Yearly"]
            if len(yearly_cmp) > 0:
                fig = px.bar(yearly_cmp, x='Year', y=['Hours', 'Loss'], barmode='group',
                             title='Yearly Comparison',
                             color_discrete_sequence=['#1F4E78', '#E74C3C'])
                insert_chart(ws, fig, f'A{len(yearly_cmp) + 4}', width=800, height=400)

            # ---- Sheets 7-10: Pivots and Raw ----
            df.pivot_table(index='M/C Name', columns='Period', values='Total Hrs',
                           aggfunc='sum', fill_value=0).to_excel(writer, sheet_name="7. Machine x Period")
            df.pivot_table(index='Breakdown Type', columns='Year', values='Total Hrs',
                           aggfunc='sum', fill_value=0).to_excel(writer, sheet_name="8. Type x Year")
            df.pivot_table(index='M/C Name', columns='Reason of Breakdown', values='Total Hrs',
                           aggfunc='count', fill_value=0).to_excel(writer, sheet_name="9. Cause x Machine")
            df.to_excel(writer, sheet_name="10. Raw Data", index=False)

            for sheet in writer.sheets.values():
                style_sheet(sheet)
        buffer.seek(0)
        return buffer

    # ============================================================
    # POWERPOINT BUILDER
    # ============================================================
    def build_powerpoint(df, machine_stats_df, available_hours_per_machine,
                         plant_capacity, breakdown_hours, production_loss,
                         availability_pct, downtime_pct, mttr, mtbf,
                         total_machines, with_bd_count, without_bd_count):
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

        NAVY = RGBColor(0x1F, 0x4E, 0x78)
        RED = RGBColor(0xE7, 0x4C, 0x3C)
        GREEN = RGBColor(0x27, 0xAE, 0x60)
        ORANGE = RGBColor(0xF3, 0x9C, 0x12)
        WHITE = RGBColor(0xFF, 0xFF, 0xFF)
        DARK = RGBColor(0x2C, 0x3E, 0x50)

        BLANK = prs.slide_layouts[6]
        total_slides = 12

        def add_footer(slide, page_num):
            footer = slide.shapes.add_textbox(Inches(0.3), Inches(7.05),
                                               Inches(12.7), Inches(0.35))
            tf = footer.text_frame
            tf.text = f"Created by Vikas Singh | Instrumentation Dept.                              Slide {page_num}/{total_slides}"
            for p in tf.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(9)
                    r.font.color.rgb = RGBColor(0x88, 0x88, 0x88)

        def add_title_band(slide, title, subtitle=""):
            band = slide.shapes.add_shape(1, Inches(0), Inches(0),
                                           prs.slide_width, Inches(0.9))
            band.fill.solid()
            band.fill.fore_color.rgb = NAVY
            band.line.fill.background()

            tb = slide.shapes.add_textbox(Inches(0.4), Inches(0.15),
                                            Inches(12.5), Inches(0.7))
            tf = tb.text_frame
            tf.text = title
            p = tf.paragraphs[0]
            p.runs[0].font.size = Pt(26)
            p.runs[0].font.bold = True
            p.runs[0].font.color.rgb = WHITE
            if subtitle:
                sb = slide.shapes.add_textbox(Inches(0.4), Inches(0.62),
                                                Inches(12.5), Inches(0.3))
                tf2 = sb.text_frame
                tf2.text = subtitle
                p2 = tf2.paragraphs[0]
                p2.runs[0].font.size = Pt(12)
                p2.runs[0].font.color.rgb = RGBColor(0xCC, 0xDD, 0xEE)

        def add_kpi_card(slide, x, y, w, h, label, value, color):
            card = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
            card.fill.solid()
            card.fill.fore_color.rgb = color
            card.line.color.rgb = WHITE
            card.line.width = Pt(1.5)
            tf = card.text_frame
            tf.margin_left = Inches(0.1)
            tf.margin_right = Inches(0.1)
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            r = p.add_run()
            r.text = str(value)
            r.font.size = Pt(22)
            r.font.bold = True
            r.font.color.rgb = WHITE
            p2 = tf.add_paragraph()
            p2.alignment = PP_ALIGN.CENTER
            r2 = p2.add_run()
            r2.text = label
            r2.font.size = Pt(11)
            r2.font.color.rgb = WHITE

        def add_chart_image(slide, fig, x, y, w, h):
            png = fig_to_png(fig, width=900, height=500)
            if png is None:
                return
            slide.shapes.add_picture(io.BytesIO(png), Inches(x), Inches(y),
                                      width=Inches(w), height=Inches(h))

        # ---------- SLIDE 1: TITLE ----------
        s = prs.slides.add_slide(BLANK)
        bg = s.shapes.add_shape(1, Inches(0), Inches(0), prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = NAVY
        bg.line.fill.background()
        tb = s.shapes.add_textbox(Inches(1), Inches(2), Inches(11.3), Inches(1.5))
        tf = tb.text_frame
        tf.text = "🔧 Breakdown & Production Loss Analysis"
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.runs[0].font.size = Pt(42)
        p.runs[0].font.bold = True
        p.runs[0].font.color.rgb = WHITE
        tb2 = s.shapes.add_textbox(Inches(1), Inches(3.6), Inches(11.3), Inches(0.6))
        tf2 = tb2.text_frame
        tf2.text = ", ".join(sorted(df['Period'].unique()))
        p2 = tf2.paragraphs[0]
        p2.alignment = PP_ALIGN.CENTER
        p2.runs[0].font.size = Pt(22)
        p2.runs[0].font.color.rgb = RGBColor(0xCC, 0xDD, 0xEE)
        tb3 = s.shapes.add_textbox(Inches(1), Inches(5), Inches(11.3), Inches(1.5))
        tf3 = tb3.text_frame
        tf3.text = "Created by: Vikas Singh"
        p3 = tf3.paragraphs[0]
        p3.alignment = PP_ALIGN.CENTER
        p3.runs[0].font.size = Pt(18)
        p3.runs[0].font.bold = True
        p3.runs[0].font.color.rgb = WHITE
        p4 = tf3.add_paragraph()
        p4.alignment = PP_ALIGN.CENTER
        r4 = p4.add_run()
        r4.text = "Instrumentation Department"
        r4.font.size = Pt(14)
        r4.font.color.rgb = RGBColor(0xCC, 0xDD, 0xEE)

        # ---------- SLIDE 2: OVERVIEW KPIs ----------
        s = prs.slides.add_slide(BLANK)
        add_title_band(s, "Overview — Key Metrics", "Performance at a glance")
        add_kpi_card(s, 0.4, 1.2, 3.0, 1.2, "Breakdown Hours",
                     f"{df['Total Hrs'].sum():,.1f}", NAVY)
        add_kpi_card(s, 3.5, 1.2, 3.0, 1.2, "Production Loss",
                     f"{df['Production Loss'].sum():,.0f}", RED)
        add_kpi_card(s, 6.6, 1.2, 3.0, 1.2, "Events", f"{len(df)}", ORANGE)
        add_kpi_card(s, 9.7, 1.2, 3.2, 1.2, "Avg Hours/Event",
                     f"{df['Total Hrs'].mean():.2f}" if len(df) else "0", GREEN)
        pie = df.groupby('Breakdown Type')['Total Hrs'].sum().reset_index()
        if len(pie) > 0:
            fig = px.pie(pie, names='Breakdown Type', values='Total Hrs', hole=0.4,
                         title='Breakdown Hours by Type',
                         color_discrete_sequence=px.colors.qualitative.Set2)
            add_chart_image(s, fig, 3.5, 2.6, 6.3, 4.3)
        add_footer(s, 2)

        # ---------- SLIDE 3: PRODUCTION LOSS + TREND ----------
        s = prs.slides.add_slide(BLANK)
        add_title_band(s, "Production Loss & Monthly Trend")
        pie2 = df.groupby('Breakdown Type')['Production Loss'].sum().reset_index()
        if len(pie2) > 0:
            fig = px.pie(pie2, names='Breakdown Type', values='Production Loss',
                         hole=0.4, title='Production Loss by Type',
                         color_discrete_sequence=px.colors.qualitative.Pastel)
            add_chart_image(s, fig, 0.3, 1.1, 5.2, 5.8)
        trend = df.groupby('Period_Key').agg(
            Period=('Period', 'first'), Hours=('Total Hrs', 'sum'),
            Loss=('Production Loss', 'sum')).reset_index().sort_values('Period_Key')
        if len(trend) > 0:
            fig = go.Figure()
            fig.add_trace(go.Bar(x=trend['Period'], y=trend['Hours'],
                                 name='Hours', marker_color='#1F4E78'))
            fig.add_trace(go.Scatter(x=trend['Period'], y=trend['Loss'],
                                     name='Loss', mode='lines+markers',
                                     line=dict(color='#E74C3C', width=3), yaxis='y2'))
            fig.update_layout(yaxis=dict(title='Hours'),
                              yaxis2=dict(title='Loss', overlaying='y', side='right'),
                              title='Monthly Trend', height=500)
            add_chart_image(s, fig, 5.7, 1.1, 7.3, 5.8)
        add_footer(s, 3)

        # ---------- SLIDE 4: CAPACITY KPIs ----------
        s = prs.slides.add_slide(BLANK)
        add_title_band(s, "Plant Capacity Analysis")
        add_kpi_card(s, 0.4, 1.2, 3.0, 1.1, "Total Machines", total_machines, NAVY)
        add_kpi_card(s, 3.5, 1.2, 3.0, 1.1, "With Breakdown",
                     f"{with_bd_count} ({with_bd_count/total_machines*100:.0f}%)"
                     if total_machines else "0", RED)
        add_kpi_card(s, 6.6, 1.2, 3.0, 1.1, "Without Breakdown",
                     f"{without_bd_count} ({without_bd_count/total_machines*100:.0f}%)"
                     if total_machines else "0", GREEN)
        add_kpi_card(s, 9.7, 1.2, 3.2, 1.1, "Available Hrs",
                     f"{available_hours_per_machine:,.0f}", ORANGE)
        add_kpi_card(s, 0.4, 2.6, 3.0, 1.1, "Plant Capacity",
                     f"{plant_capacity:,.0f}", NAVY)
        add_kpi_card(s, 3.5, 2.6, 3.0, 1.1, "Availability %",
                     f"{availability_pct:.2f}%", GREEN)
        add_kpi_card(s, 6.6, 2.6, 3.0, 1.1, "Downtime %",
                     f"{downtime_pct:.2f}%", ORANGE)
        add_kpi_card(s, 9.7, 2.6, 3.2, 1.1, "Breakdown Hours",
                     f"{breakdown_hours:,.1f}", RED)
        pie_df = pd.DataFrame({"Status": ["With Breakdown", "Without Breakdown"],
                               "Count": [with_bd_count, without_bd_count]})
        fig = px.pie(pie_df, names="Status", values="Count", hole=0.4,
                     title='Machines — Breakdown Status', color="Status",
                     color_discrete_map={"With Breakdown": "#E74C3C",
                                         "Without Breakdown": "#27AE60"})
        add_chart_image(s, fig, 3.5, 3.9, 6.3, 3.2)
        add_footer(s, 4)

        # ---------- SLIDE 5: MTBF & MTTR ----------
        s = prs.slides.add_slide(BLANK)
        add_title_band(s, "Reliability Metrics")
        add_kpi_card(s, 0.4, 1.2, 4.0, 1.3, "MTBF (hrs)",
                     f"{mtbf:,.1f}" if events else "—", GREEN)
        add_kpi_card(s, 4.7, 1.2, 4.0, 1.3, "MTTR (hrs)",
                     f"{mttr:.2f}" if events else "—", RED)
        add_kpi_card(s, 9.0, 1.2, 3.9, 1.3, "Production Loss",
                     f"{production_loss:,.0f}", ORANGE)
        chart_df = machine_stats_df[machine_stats_df['Events'] > 0].copy()
        if len(chart_df) > 0:
            mttr_chart = chart_df[chart_df['MTTR (hrs)'] != "—"].copy()
            if len(mttr_chart) > 0:
                mttr_chart['MTTR (hrs)'] = pd.to_numeric(mttr_chart['MTTR (hrs)'],
                                                          errors='coerce').fillna(0)
                fig = px.bar(mttr_chart.sort_values('MTTR (hrs)', ascending=False),
                             x='Machine', y='MTTR (hrs)', color='MTTR (hrs)',
                             title='MTTR per Machine (High = Bad)',
                             color_continuous_scale=['#27AE60', '#F39C12', '#E74C3C'])
                fig.update_layout(xaxis_tickangle=-45, height=450)
                add_chart_image(s, fig, 0.3, 2.8, 6.4, 4.1)
            mtbf_chart = chart_df[chart_df['MTBF (hrs)'] != "—"].copy()
            if len(mtbf_chart) > 0:
                mtbf_chart['MTBF (hrs)'] = pd.to_numeric(mtbf_chart['MTBF (hrs)'],
                                                          errors='coerce').fillna(0)
                fig = px.bar(mtbf_chart.sort_values('MTBF (hrs)', ascending=True),
                             x='Machine', y='MTBF (hrs)', color='MTBF (hrs)',
                             title='MTBF per Machine (Low = Bad)',
                             color_continuous_scale=['#E74C3C', '#F39C12', '#27AE60'])
                fig.update_layout(xaxis_tickangle=-45, height=450)
                add_chart_image(s, fig, 6.8, 2.8, 6.2, 4.1)
        add_footer(s, 5)

        # ---------- SLIDE 6: MACHINE-WISE ----------
        s = prs.slides.add_slide(BLANK)
        add_title_band(s, "Machine-wise Breakdown")
        mach = df.groupby('M/C Name').agg(Hours=('Total Hrs', 'sum'),
                                            Loss=('Production Loss', 'sum')).reset_index().sort_values(
            'Hours', ascending=False).head(15)
        if len(mach) > 0:
            fig = px.bar(mach, x='M/C Name', y=['Hours', 'Loss'], barmode='group',
                         title='Top 15 Machines')
            fig.update_layout(xaxis_tickangle=-45, height=550)
            add_chart_image(s, fig, 0.3, 1.1, 12.7, 5.8)
        add_footer(s, 6)

        # ---------- SLIDE 7: CAUSES ----------
        s = prs.slides.add_slide(BLANK)
        add_title_band(s, "Top Repeated Causes")
        causes = df.groupby('Reason of Breakdown').agg(
            Frequency=('Reason of Breakdown', 'count'),
            Hours=('Total Hrs', 'sum')).reset_index().sort_values('Frequency', ascending=False)
        if len(causes) > 0:
            fig = px.bar(causes.head(10), x='Frequency', y='Reason of Breakdown',
                         orientation='h', color='Hours', title='Top 10 Causes',
                         color_continuous_scale='Reds')
            fig.update_layout(yaxis={'categoryorder': 'total ascending'}, height=550)
            add_chart_image(s, fig, 0.3, 1.1, 12.7, 5.8)
        add_footer(s, 7)

        # ---------- SLIDE 8: TYPE-WISE ----------
        s = prs.slides.add_slide(BLANK)
        add_title_band(s, "Breakdown by Type")
        tsum = df.groupby('Breakdown Type').agg(Hours=('Total Hrs', 'sum'),
                                                 Events=('Total Hrs', 'count')).reset_index()
        colors_map = {'Instrumentation': RGBColor(0x34, 0x98, 0xDB),
                      'Electrical': RGBColor(0xF3, 0x9C, 0x12),
                      'Mechanical': RGBColor(0xE7, 0x4C, 0x3C)}
        x_pos = 0.4
        for _, row in tsum.iterrows():
            color = colors_map.get(row['Breakdown Type'], NAVY)
            add_kpi_card(s, x_pos, 1.2, 4.0, 1.4, row['Breakdown Type'],
                         f"{row['Hours']:.1f} hrs | {row['Events']} events", color)
            x_pos += 4.3
        if len(tsum) > 0:
            fig = px.pie(tsum, names='Breakdown Type', values='Hours', hole=0.4,
                         color='Breakdown Type',
                         color_discrete_map={'Instrumentation': '#3498DB',
                                             'Electrical': '#F39C12',
                                             'Mechanical': '#E74C3C'},
                         title='Breakdown Hours by Type')
            add_chart_image(s, fig, 4.5, 3.0, 4.3, 4.0)
        add_footer(s, 8)

        # ---------- SLIDE 9: MONTHLY ----------
        s = prs.slides.add_slide(BLANK)
        add_title_band(s, "Monthly Comparison")
        monthly_cmp = df.groupby('Period_Key').agg(
            Period=('Period', 'first'), Hours=('Total Hrs', 'sum'),
            Loss=('Production Loss', 'sum')).reset_index().sort_values('Period_Key')
        if len(monthly_cmp) > 0:
            fig = px.bar(monthly_cmp, x='Period', y=['Hours', 'Loss'], barmode='group',
                         title='Month-over-Month',
                         color_discrete_sequence=['#1F4E78', '#E74C3C'])
            add_chart_image(s, fig, 0.3, 1.1, 12.7, 5.8)
        add_footer(s, 9)

        # ---------- SLIDE 10: YEARLY ----------
        s = prs.slides.add_slide(BLANK)
        add_title_band(s, "Yearly Comparison")
        yearly_cmp = df.groupby('Year').agg(Hours=('Total Hrs', 'sum'),
                                             Loss=('Production Loss', 'sum')).reset_index()
        if len(yearly_cmp) > 0:
            fig = px.bar(yearly_cmp, x='Year', y=['Hours', 'Loss'], barmode='group',
                         title='Year-over-Year',
                         color_discrete_sequence=['#1F4E78', '#E74C3C'])
            add_chart_image(s, fig, 0.3, 1.1, 12.7, 5.8)
        add_footer(s, 10)

        # ---------- SLIDE 11: TOP 10 TABLE ----------
        s = prs.slides.add_slide(BLANK)
        add_title_band(s, "Top 10 Machines — Detailed")
        top10 = machine_stats_df.sort_values('Breakdown Hours', ascending=False).head(10)
        cols = ['Machine', 'Breakdown Hours', 'Events', 'Availability %', 'MTTR (hrs)', 'MTBF (hrs)']
        top10 = top10[cols]
        rows = len(top10) + 1
        table_shape = s.shapes.add_table(rows, len(cols), Inches(0.3), Inches(1.2),
                                          Inches(12.7), Inches(5.5))
        table = table_shape.table
        for i, col in enumerate(cols):
            cell = table.cell(0, i)
            cell.text = col
            cell.fill.solid()
            cell.fill.fore_color.rgb = NAVY
            for p in cell.text_frame.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(11)
                    r.font.bold = True
                    r.font.color.rgb = WHITE
        for i, (_, row) in enumerate(top10.iterrows(), start=1):
            for j, col in enumerate(cols):
                cell = table.cell(i, j)
                cell.text = str(row[col])
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(0xF0, 0xF4, 0xF8) if i % 2 else WHITE
                for p in cell.text_frame.paragraphs:
                    for r in p.runs:
                        r.font.size = Pt(10)
                        r.font.color.rgb = DARK
        add_footer(s, 11)

        # ---------- SLIDE 12: KEY TAKEAWAYS ----------
        s = prs.slides.add_slide(BLANK)
        add_title_band(s, "Key Takeaways")
        worst_machine = machine_stats_df.iloc[0]['Machine'] if len(machine_stats_df) else "—"
        top_cause = df['Reason of Breakdown'].mode()[0] if len(df) else "—"
        box = s.shapes.add_textbox(Inches(0.8), Inches(1.3), Inches(11.7), Inches(5.5))
        tf = box.text_frame
        tf.word_wrap = True
        bullets = [
            f"📊 Total Breakdown Hours: {df['Total Hrs'].sum():,.1f} hrs across {len(df)} events",
            f"💰 Total Production Loss: {df['Production Loss'].sum():,.0f} units",
            f"🔧 Total Machines in Plant: {total_machines}",
            f"✅ Machines with Zero Breakdown: {without_bd_count} "
            f"({without_bd_count/total_machines*100:.1f}%)" if total_machines else "",
            f"⚠️ Machine with Lowest Availability: {worst_machine}",
            f"⏱️ MTBF: {mtbf:,.1f} hrs  |  MTTR: {mttr:.2f} hrs",
            f"🔥 Most Frequent Cause: {top_cause}",
            f"📈 Plant Availability: {availability_pct:.2f}%",
        ]
        for i, b in enumerate([x for x in bullets if x]):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            r = p.add_run()
            r.text = b
            r.font.size = Pt(18)
            r.font.color.rgb = DARK
            p.space_after = Pt(14)
        add_footer(s, 12)

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        return buf

    # ============================================================
    # EXPORT BUTTONS
    # ============================================================
    col1, col2 = st.columns(2)

    with col1:
        excel_data_only = build_data_only_excel(df)
        XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        st.download_button(
            "📥 Download Data Only (Fast, ~200 KB)",
            data=excel_data_only,
            file_name="Breakdown_Analysis_Data.xlsx",
            mime=XLSX_MIME,
            use_container_width=True)

    with col2:
        if st.button("📊 Prepare Full Report (with Charts)", use_container_width=True):
            with st.spinner("Rendering charts and building Excel... (~15 sec)"):
                try:
                    period_options = sorted(df['Period_Key'].unique())
                    sel_period_cap = period_options[-1]
                    period_df = df[df['Period_Key'] == sel_period_cap]
                    if sel_period_cap in calendar_data:
                        avail_hrs = calendar_data[sel_period_cap]
                    else:
                        avail_hrs = 20 * pd.Period(sel_period_cap).days_in_month
                    if master_machines:
                        all_m = master_machines
                    else:
                        all_m = sorted(period_df['M/C Name'].astype(str).unique())
                    machines_with_bd = set(period_df['M/C Name'].astype(str).unique())
                    tm = len(all_m)
                    wbd = len([m for m in all_m if m in machines_with_bd])
                    nobd = tm - wbd
                    pc = tm * avail_hrs
                    bh = period_df['Total Hrs'].sum()
                    pl = period_df['Production Loss'].sum()
                    ap = ((pc - bh) / pc * 100) if pc else 0
                    dp = (bh / pc * 100) if pc else 0
                    ev = len(period_df)
                    mt = (bh / ev) if ev else 0
                    mbf = ((pc - bh) / ev) if ev else 0
                    ms = []
                    for m in all_m:
                        m_df = period_df[period_df['M/C Name'].astype(str) == m]
                        h = m_df['Total Hrs'].sum()
                        e = len(m_df)
                        lo = m_df['Production Loss'].sum()
                        a = ((avail_hrs - h) / avail_hrs * 100) if avail_hrs else 100
                        bf = ((avail_hrs - h) / e) if e else None
                        tr = (h / e) if e else 0
                        ms.append({"Machine": m, "Breakdown Hours": round(h, 2),
                                   "Events": e, "Production Loss": round(lo, 1),
                                   "Availability %": round(a, 2),
                                   "MTBF (hrs)": round(bf, 1) if bf else "—",
                                   "MTTR (hrs)": round(tr, 2) if e else "—"})
                    msdf = pd.DataFrame(ms).sort_values('Availability %')
                    full_excel = build_full_excel(df, msdf, avail_hrs, pc, bh, pl,
                                                   ap, dp, mt, mbf, tm, wbd, nobd, period_df)
                    st.download_button("📥 Download Full Report (with Charts)",
                                       data=full_excel,
                                       file_name="Breakdown_Analysis_Full_Report.xlsx",
                                       mime=XLSX_MIME, use_container_width=True)
                except Exception as e:
                    st.error(f"❌ Full report error: {e}")

    # ---- PowerPoint Button ----
    st.markdown("")
    if st.button("🎨 Prepare PowerPoint Presentation", use_container_width=True):
        with st.spinner("Building colorful PowerPoint... (~20 sec)"):
            try:
                period_options = sorted(df['Period_Key'].unique())
                sel_period_cap = period_options[-1]
                period_df = df[df['Period_Key'] == sel_period_cap]
                if sel_period_cap in calendar_data:
                    avail_hrs = calendar_data[sel_period_cap]
                else:
                    avail_hrs = 20 * pd.Period(sel_period_cap).days_in_month
                if master_machines:
                    all_m = master_machines
                else:
                    all_m = sorted(period_df['M/C Name'].astype(str).unique())
                machines_with_bd = set(period_df['M/C Name'].astype(str).unique())
                tm = len(all_m)
                wbd = len([m for m in all_m if m in machines_with_bd])
                nobd = tm - wbd
                pc = tm * avail_hrs
                bh = period_df['Total Hrs'].sum()
                pl = period_df['Production Loss'].sum()
                ap = ((pc - bh) / pc * 100) if pc else 0
                dp = (bh / pc * 100) if pc else 0
                ev = len(period_df)
                mt = (bh / ev) if ev else 0
                mbf = ((pc - bh) / ev) if ev else 0
                ms = []
                for m in all_m:
                    m_df = period_df[period_df['M/C Name'].astype(str) == m]
                    h = m_df['Total Hrs'].sum()
                    e = len(m_df)
                    lo = m_df['Production Loss'].sum()
                    a = ((avail_hrs - h) / avail_hrs * 100) if avail_hrs else 100
                    bf = ((avail_hrs - h) / e) if e else None
                    tr = (h / e) if e else 0
                    ms.append({"Machine": m, "Breakdown Hours": round(h, 2),
                               "Events": e, "Production Loss": round(lo, 1),
                               "Availability %": round(a, 2),
                               "MTBF (hrs)": round(bf, 1) if bf else "—",
                               "MTTR (hrs)": round(tr, 2) if e else "—"})
                msdf = pd.DataFrame(ms).sort_values('Availability %')
                pptx_data = build_powerpoint(df, msdf, avail_hrs, pc, bh, pl,
                                              ap, dp, mt, mbf, tm, wbd, nobd)
                PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                st.download_button("📥 Download PowerPoint Presentation",
                                   data=pptx_data,
                                   file_name="Breakdown_Analysis_Presentation.pptx",
                                   mime=PPTX_MIME, use_container_width=True)
            except Exception as e:
                st.error(f"❌ PowerPoint error: {e}")

    else:
        st.info("👆 Upload at least one **Breakdown Excel** file to begin.")
    st.markdown("""
    ### 📁 Files you can upload

    | File | Required? | Format |
    |---|---|---|
    | **Breakdown Excel(s)** | ✅ Yes | Monthly files |
    | **Master Machine List** | Optional | `machines.xlsx` — column `Machine Name` |
    | **Production Calendar(s)** | Optional | `calendar_Jun-25.xlsx` — `Date`, `Available_Hours` |

    ### 📊 Exports available

    - **📥 Data Only** — Excel with 8 sheets, no charts, ~200 KB
    - **📊 Full Report** — Excel with 10 sheets + embedded chart images
    - **🎨 PowerPoint** — 12 colorful slides with charts (matches dashboard colors)
     | Team | Name | Department |
        |---|---|---|
        | **Team Leader** | J.C Roy | Mechanical |
        | **Team Member** | Ajit Basha | Electrical |
        | **Team Member** | Vikas Singh | Instrumentation |
        | **Team Member** | Bhaskar Ghormare | Dyeing |
    """)
