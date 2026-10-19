import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
import re
from openpyxl.styles import Font, PatternFill, Alignment

st.set_page_config(page_title="Breakdown Analysis", layout="wide")

st.title("🔧 Breakdown & Production Loss Analysis Dashboard")
st.markdown(
    """
    <div style="background-color:#1F4E78; padding:10px 15px; border-radius:8px; color:white; margin-bottom:15px;">
        <b>👤 Team Leader:</b> J.C Roy(Mechanical) &nbsp;|&nbsp;
        <b>👤 Team Member:</b> Ajit Basha(Electrical) &nbsp;|&nbsp;
        <b>👤 Team Member:</b> Vikas Singh(Instrumentation) &nbsp;|&nbsp;
        <b>👤 Team Member:</b> Bhaskar Ghormare(Dyeing) &nbsp;|&nbsp;
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
        # NEW: Calculate Fiscal Year (April to March)
        data['FY_Year'] = data['Date'].dt.year + (data['Date'].dt.month >= 4).astype(int)
        data['FY_Label'] = data['FY_Year'].apply(lambda y: f"{int(y)-1}-{str(int(y))[-2:]}" if pd.notna(y) else "Unknown")
    else:
        data['Date'] = pd.NaT
        data['FY_Label'] = "Unknown"
        
    data['Total Hrs'] = pd.to_numeric(data['Total Hrs'], errors='coerce').fillna(0)
    data['Production Loss'] = pd.to_numeric(data['Production Loss'], errors='coerce').fillna(0)
    data = data.dropna(subset=['M/C Name', 'Reason of Breakdown'])

    period_key, period_label = parse_period(uploaded.name)
    data['Period_Key'] = period_key
    data['Period'] = period_label
    data['Source File'] = uploaded.name
    return data


# ============================================================
# SIDEBAR UPLOADS
# ============================================================
st.sidebar.header("📤 Upload Files")

breakdown_files = st.sidebar.file_uploader(
    "Breakdown Excel file(s) — one per month",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
    key="breakdown"
)

machines_file = st.sidebar.file_uploader(
    "Master Machine List (machines.xlsx)",
    type=["xlsx", "xls"],
    key="machines"
)

calendar_files = st.sidebar.file_uploader(
    "Production Calendar(s) — one per month",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
    key="calendars"
)

# ---------- RESET BUTTON ----------
st.sidebar.markdown("---")
if st.sidebar.button("Clear All Uploaded Files", use_container_width=True):
    for key in ["breakdown", "machines", "calendars"]:
        if key in st.session_state:
            del st.session_state[key]
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.markdown("### 💰 Cost Settings")
cost_per_unit = st.sidebar.number_input("Cost per unit (₹)", min_value=0.0, value=50.0, step=1.0,help="Cost of producing one unit in ₹")
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

    data = pd.concat(all_data, ignore_index=True)
    data = data.sort_values('Period_Key')

    # ---- Master machine list ----
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

    # ---- Calendars ----
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
                totals = cdf.groupby('Period_Key')[hr_col].sum().to_dict()
                calendar_data.update(totals)
            except Exception as e:
                st.sidebar.warning(f"⚠️ Calendar {cf.name}: {e}")

    st.success(f"✅ Loaded {len(data)} records from {len(breakdown_files)} file(s): "
               f"{', '.join(sorted(data['Period'].unique()))}")

    # ---- SIDEBAR FILTERS ----
    st.sidebar.header("🔍 Filters")
    periods = sorted(data['Period'].unique())
    sel_periods = st.sidebar.multiselect("Period (Month-Year)", periods, default=periods)

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

    # ============================================================
    # TABS
    # ============================================================
    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
        "📊 Overview",
        "📅 Month Comparison",
        "📈 Year Comparison",
        "🏆 Yearly Analysis",
        "🏭 Capacity Analysis",
        "🔧 Machine-wise",
        "🔁 Causes",
        "📆 YTD (Fiscal Year)"
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
            Period=('Period', 'first'),
            Hours=('Total Hrs', 'sum'),
            Loss=('Production Loss', 'sum'),
            Events=('Total Hrs', 'count')
        ).reset_index().sort_values('Period_Key')

        if len(trend) > 0:
            fig = go.Figure()
            fig.add_trace(go.Bar(x=trend['Period'], y=trend['Hours'],
                                 name='Breakdown Hours', yaxis='y',
                                 marker_color='#1F4E78'))
            fig.add_trace(go.Scatter(x=trend['Period'], y=trend['Loss'],
                                     name='Production Loss', yaxis='y2',
                                     mode='lines+markers',
                                     line=dict(color='#E74C3C', width=3)))
            fig.update_layout(
                yaxis=dict(title='Breakdown Hours'),
                yaxis2=dict(title='Production Loss', overlaying='y', side='right'),
                hovermode='x unified', height=450
            )
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
            st.subheader("By Breakdown Type")
            cmp_type = df.groupby(['Period_Key', 'Period', 'Breakdown Type'])['Total Hrs'].sum().reset_index()
            if len(cmp_type) > 0:
                fig = px.bar(cmp_type, x='Breakdown Type', y='Total Hrs', color='Period',
                             barmode='group', text_auto='.1f')
                st.plotly_chart(fig, use_container_width=True)

            st.subheader("By Machine")
            cmp_mach = df.groupby(['Period_Key', 'Period', 'M/C Name'])['Total Hrs'].sum().reset_index()
            if len(cmp_mach) > 0:
                fig = px.bar(cmp_mach, x='M/C Name', y='Total Hrs', color='Period',
                             barmode='group', text_auto='.1f')
                st.plotly_chart(fig, use_container_width=True)

            st.subheader("Pivot — Machine × Month (Hours)")
            st.dataframe(df.pivot_table(index='M/C Name', columns='Period',
                                        values='Total Hrs', aggfunc='sum', fill_value=0),
                         use_container_width=True)

            st.subheader("Pivot — Type × Month (Hours)")
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
                Hours=('Total Hrs', 'sum'),
                Loss=('Production Loss', 'sum'),
                Events=('Total Hrs', 'count')
            ).reset_index()
            fig = px.bar(yr_totals, x='Year', y=['Hours', 'Loss'],
                         barmode='group', text_auto='.1f')
            st.plotly_chart(fig, use_container_width=True)

            st.subheader("Same Month — Different Years")
            monthly = df.groupby(['Year', 'Month_Name'])['Total Hrs'].sum().reset_index()
            fig = px.bar(monthly, x='Month_Name', y='Total Hrs', color='Year',
                         barmode='group', text_auto='.1f',
                         category_orders={'Month_Name': list(MONTH_NAMES.values())})
            st.plotly_chart(fig, use_container_width=True)

            st.subheader("Type × Year (Hours)")
            st.dataframe(df.pivot_table(index='Breakdown Type', columns='Year',
                                        values='Total Hrs', aggfunc='sum', fill_value=0),
                         use_container_width=True)

            st.subheader("Machine × Year (Hours)")
            st.dataframe(df.pivot_table(index='M/C Name', columns='Year',
                                        values='Total Hrs', aggfunc='sum', fill_value=0),
                         use_container_width=True)

    # ---------- TAB 4: YEARLY ANALYSIS ----------
    with tab4:
        years_available = sorted(df['Year'].unique())
        if len(years_available) == 0:
            st.info("Upload at least one file with a year in the filename.")
        else:
            selected_year = st.selectbox("Select Year", years_available,
                                         index=len(years_available) - 1)
            year_df = df[df['Year'] == selected_year]

            st.subheader(f"🏆 Total Year Summary — {selected_year}")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Total Breakdown Hours", f"{year_df['Total Hrs'].sum():,.1f}")
            c2.metric("Total Production Loss", f"{year_df['Production Loss'].sum():,.0f}")
            c3.metric("Total Events", len(year_df))
            c4.metric("Avg Hours/Event",
                      f"{year_df['Total Hrs'].mean():.2f}" if len(year_df) else "0")

            c5, c6, c7, c8 = st.columns(4)
            c5.metric("Machines Covered", year_df['M/C Name'].nunique())
            c6.metric("Distinct Causes", year_df['Reason of Breakdown'].nunique())
            c7.metric("Worst Machine",
                      year_df.groupby('M/C Name')['Total Hrs'].sum().idxmax()
                      if len(year_df) else "-")
            c8.metric("Top Type",
                      year_df.groupby('Breakdown Type')['Total Hrs'].sum().idxmax()
                      if len(year_df) else "-")

            st.divider()
            st.subheader(f"📅 Month-by-Month ({selected_year})")
            month_table = year_df.groupby(['Month_Num', 'Month_Name']).agg(
                Breakdown_Hours=('Total Hrs', 'sum'),
                Production_Loss=('Production Loss', 'sum'),
                Events=('Total Hrs', 'count')
            ).reset_index().sort_values('Month_Num')
            month_table['Cumulative_Hours'] = month_table['Breakdown_Hours'].cumsum()
            month_table['Cumulative_Loss'] = month_table['Production_Loss'].cumsum()
            st.dataframe(month_table.drop(columns=['Month_Num']),
                         use_container_width=True, hide_index=True)

            if len(month_table) > 0:
                fig = go.Figure()
                fig.add_trace(go.Bar(x=month_table['Month_Name'],
                                     y=month_table['Breakdown_Hours'],
                                     name='Monthly Hours', marker_color='#1F4E78'))
                fig.add_trace(go.Scatter(x=month_table['Month_Name'],
                                         y=month_table['Cumulative_Hours'],
                                         name='Cumulative Hours', mode='lines+markers',
                                         line=dict(color='#E74C3C', width=3),
                                         yaxis='y2'))
                fig.update_layout(
                    title=f"{selected_year} — Monthly & Cumulative Breakdown Hours",
                    yaxis=dict(title='Monthly Hours'),
                    yaxis2=dict(title='Cumulative Hours', overlaying='y', side='right'),
                    hovermode='x unified', height=450
                )
                st.plotly_chart(fig, use_container_width=True)

            st.subheader(f"📊 Month × Type ({selected_year})")
            month_type = year_df.pivot_table(index='Month_Name', columns='Breakdown Type',
                                             values='Total Hrs', aggfunc='sum', fill_value=0)
            month_order = [m for m in MONTH_NAMES.values() if m in month_type.index]
            month_type = month_type.reindex(month_order)
            st.dataframe(month_type, use_container_width=True)

            st.subheader(f"🏭 Month × Machine ({selected_year})")
            month_mach = year_df.pivot_table(index='M/C Name', columns='Month_Name',
                                             values='Total Hrs', aggfunc='sum', fill_value=0)
            month_mach = month_mach.reindex(columns=month_order, fill_value=0)
            st.dataframe(month_mach, use_container_width=True)

            st.divider()
            st.subheader("📈 Year-wise Total (All Years)")
            year_totals = df.groupby('Year').agg(
                Breakdown_Hours=('Total Hrs', 'sum'),
                Production_Loss=('Production Loss', 'sum'),
                Events=('Total Hrs', 'count'),
                Machines=('M/C Name', 'nunique'),
                Causes=('Reason of Breakdown', 'nunique')
            ).reset_index()
            st.dataframe(year_totals, use_container_width=True, hide_index=True)

            if len(year_totals) >= 2:
                year_totals = year_totals.sort_values('Year')
                year_totals['Hours_Change_%'] = (year_totals['Breakdown_Hours'].pct_change() * 100).round(2)
                year_totals['Loss_Change_%'] = (year_totals['Production_Loss'].pct_change() * 100).round(2)
                year_totals['Events_Change_%'] = (year_totals['Events'].pct_change() * 100).round(2)
                st.subheader("📉 Year-over-Year % Change")
                st.dataframe(year_totals[['Year', 'Hours_Change_%',
                                           'Loss_Change_%', 'Events_Change_%']],
                             use_container_width=True, hide_index=True)

            fig = px.bar(year_totals, x='Year', y=['Breakdown_Hours', 'Production_Loss'],
                         barmode='group', text_auto='.1f')
            st.plotly_chart(fig, use_container_width=True)

            st.subheader("🔁 Top Causes × Year")
            cause_year = df.pivot_table(index='Reason of Breakdown', columns='Year',
                                        values='Total Hrs', aggfunc='count', fill_value=0)
            cause_year['Total'] = cause_year.sum(axis=1)
            cause_year = cause_year.sort_values('Total', ascending=False).head(10)
            st.dataframe(cause_year.drop(columns=['Total']), use_container_width=True)

    # ---------- TAB 5: CAPACITY ANALYSIS ----------
    with tab5:
        st.subheader("🏭 Plant Capacity & Availability Analysis")

        if not master_machines:
            st.warning("⚠️ Upload **machines.xlsx** in the sidebar for full capacity analysis.")
            st.info("Without the master list, only machines with breakdowns are analyzed.")

        period_options = sorted(df['Period_Key'].unique())
        sel_period_cap = st.selectbox(
            "Select Period for Capacity Analysis",
            period_options,
            index=len(period_options) - 1,
            format_func=lambda x: df[df['Period_Key'] == x]['Period'].iloc[0]
            if len(df[df['Period_Key'] == x]) else x
        )
        period_df = df[df['Period_Key'] == sel_period_cap]

        # Available hours
        if sel_period_cap in calendar_data:
            available_hours_per_machine = calendar_data[sel_period_cap]
            st.success(f"✅ Calendar loaded for {sel_period_cap}: **{available_hours_per_machine} hrs/machine**")
        else:
            days_in_month = pd.Period(sel_period_cap).days_in_month
            available_hours_per_machine = 20 * days_in_month
            st.warning(f"⚠️ No calendar for {sel_period_cap}. Using default **{available_hours_per_machine} hrs/machine** "
                       f"(20 hrs/day × {days_in_month} days).")

        # Machine list
        if master_machines:
            all_machines = master_machines
        else:
            all_machines = sorted(period_df['M/C Name'].astype(str).unique())

        machines_with_bd = set(period_df['M/C Name'].astype(str).unique())
        machines_without_bd = [m for m in all_machines if m not in machines_with_bd]

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

        # KPI CARDS
        st.markdown("### 📊 Plant-wide KPIs")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Machines", total_machines)
        c2.metric("With Breakdown",
                  f"{with_bd_count} ({with_bd_count/total_machines*100:.1f}%)" if total_machines else "0")
        c3.metric("Without Breakdown",
                  f"{without_bd_count} ({without_bd_count/total_machines*100:.1f}%)" if total_machines else "0")
        c4.metric("Available Hrs/Machine", f"{available_hours_per_machine:,.0f}")

        st.markdown("### ⏱️ Availability & Loss")
        c5, c6, c7, c8 = st.columns(4)
        c5.metric("Plant Capacity (hrs)", f"{plant_capacity:,.0f}")
        c6.metric("Breakdown Hours", f"{breakdown_hours:,.1f}")
        c7.metric("Availability %", f"{availability_pct:.2f}%")
        c8.metric("Downtime %", f"{downtime_pct:.2f}%")

        c9, c10, c11, c12 = st.columns(4)
        c9.metric("MTBF (hrs)", f"{mtbf:,.1f}" if events else "—")
        c10.metric("MTTR (hrs)", f"{mttr:.2f}" if events else "—")
        c11.metric("Production Loss", f"{production_loss:,.0f}")
        c12.metric("Loss per Breakdown Hr",
                   f"{production_loss/breakdown_hours:.1f}" if breakdown_hours else "—")

        st.divider()

        # Two pies
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Machines — Breakdown Status")
            pie_df = pd.DataFrame({
                "Status": ["With Breakdown", "Without Breakdown"],
                "Count": [with_bd_count, without_bd_count]
            })
            fig = px.pie(pie_df, names="Status", values="Count", hole=0.4,
                         color="Status",
                         color_discrete_map={"With Breakdown": "#E74C3C",
                                             "Without Breakdown": "#27AE60"})
            st.plotly_chart(fig, use_container_width=True)
        with col2:
            st.subheader("Downtime Contribution")
            pie2 = period_df.groupby('Breakdown Type')['Total Hrs'].sum().reset_index()
            if len(pie2) > 0:
                fig = px.pie(pie2, names='Breakdown Type', values='Total Hrs', hole=0.4)
                st.plotly_chart(fig, use_container_width=True)

        # ---- Build per-machine stats ----
        machine_stats = []
        for m in all_machines:
            m_df = period_df[period_df['M/C Name'].astype(str) == m]
            m_hours = m_df['Total Hrs'].sum()
            m_events = len(m_df)
            m_loss = m_df['Production Loss'].sum()
            m_avail = ((available_hours_per_machine - m_hours) / available_hours_per_machine * 100) if available_hours_per_machine else 100
            m_mtbf = ((available_hours_per_machine - m_hours) / m_events) if m_events else None
            m_mttr = (m_hours / m_events) if m_events else 0

            machine_stats.append({
                "Machine": m,
                "Breakdown Hours": round(m_hours, 2),
                "Events": m_events,
                "Production Loss": round(m_loss, 1),
                "Availability %": round(m_avail, 2),
                "MTBF (hrs)": round(m_mtbf, 1) if m_mtbf else "—",
                "MTTR (hrs)": round(m_mttr, 2) if m_events else "—",
            })

        machine_stats_df = pd.DataFrame(machine_stats).sort_values('Availability %')

        st.subheader("📋 Per-Machine Availability (Overall)")
        st.dataframe(machine_stats_df, use_container_width=True, hide_index=True)

        # ---- SPLIT by type ----
        st.divider()
        st.subheader("🔬 MTBF & MTTR — Split by Breakdown Type (Machine-wise)")

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

                type_stats.append({
                    "Machine": m,
                    "Breakdown Type": btype,
                    "Events": t_events,
                    "Breakdown Hours": round(t_hours, 2),
                    "Production Loss": round(t_loss, 1),
                    "MTTR (hrs)": round(t_mttr, 2) if t_mttr else "—",
                    "MTBF (hrs)": round(t_mtbf, 1) if t_mtbf else "—",
                })

        type_stats_df = pd.DataFrame(type_stats)
        type_stats_display = type_stats_df[type_stats_df['Events'] > 0].copy()
        type_stats_display = type_stats_display.sort_values(
            ['Machine', 'Breakdown Type']).reset_index(drop=True)

        st.dataframe(type_stats_display, use_container_width=True, hide_index=True)

        # MTTR & MTBF charts split by type
        if len(type_stats_display) > 0:
            col1, col2 = st.columns(2)

            with col1:
                st.markdown("**🔧 MTTR (hrs) by Machine & Type**")
                fig = px.bar(
                    type_stats_display,
                    x='Machine', y='MTTR (hrs)',
                    color='Breakdown Type',
                    barmode='group',
                    text='MTTR (hrs)',
                    color_discrete_map={
                        'Instrumentation': '#3498DB',
                        'Electrical': '#F39C12',
                        'Mechanical': '#E74C3C'
                    }
                )
                fig.update_layout(height=500, xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)

            with col2:
                st.markdown("**⏱️ MTBF (hrs) by Machine & Type**")
                fig = px.bar(
                    type_stats_display,
                    x='Machine', y='MTBF (hrs)',
                    color='Breakdown Type',
                    barmode='group',
                    text='MTBF (hrs)',
                    color_discrete_map={
                        'Instrumentation': '#3498DB',
                        'Electrical': '#F39C12',
                        'Mechanical': '#E74C3C'
                    }
                )
                fig.update_layout(height=500, xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)

        # Pivots
        col3, col4 = st.columns(2)
        with col3:
            st.markdown("**📊 MTTR (hrs) — Pivot: Machine × Type**")
            mttr_pivot = type_stats_df.pivot_table(
                index='Machine', columns='Breakdown Type',
                values='MTTR (hrs)', aggfunc='first', fill_value=0
            )
            mttr_pivot = mttr_pivot.apply(pd.to_numeric, errors='coerce').fillna(0).round(2)
            st.dataframe(mttr_pivot, use_container_width=True)

        with col4:
            st.markdown("**📊 MTBF (hrs) — Pivot: Machine × Type**")
            mtbf_pivot = type_stats_df.pivot_table(
                index='Machine', columns='Breakdown Type',
                values='MTBF (hrs)', aggfunc='first', fill_value=0
            )
            mtbf_pivot = mtbf_pivot.apply(pd.to_numeric, errors='coerce').fillna(0).round(1)
            st.dataframe(mtbf_pivot, use_container_width=True)

        # Dominant type per machine
        st.subheader("🎯 Dominant Breakdown Type per Machine")
        if len(period_df) > 0:
            dominant = period_df.groupby(['M/C Name', 'Breakdown Type'])['Total Hrs'].sum().reset_index()
            idx = dominant.groupby('M/C Name')['Total Hrs'].idxmax()
            dominant_top = dominant.loc[idx].rename(columns={
                'M/C Name': 'Machine',
                'Breakdown Type': 'Dominant Type',
                'Total Hrs': 'Hours'
            }).reset_index(drop=True)
            st.dataframe(dominant_top, use_container_width=True, hide_index=True)

        # ============================================================
        # MACHINE-WISE MTBF & MTTR CHARTS
        # ============================================================
        st.divider()
        st.subheader("📊 Machine-wise MTBF & MTTR Analysis")

        chart_df = machine_stats_df[machine_stats_df['Events'] > 0].copy()

        if len(chart_df) > 0:

            # ---- MTTR chart (High = Problematic) ----
            st.markdown("### 🔧 MTTR per Machine — **High = Problematic**")
            st.caption("Higher MTTR means longer time to fix — prioritize machines with long repairs.")

            mttr_chart = chart_df[chart_df['MTTR (hrs)'] != "—"].copy()
            if len(mttr_chart) > 0:
                mttr_chart['MTTR (hrs)'] = pd.to_numeric(mttr_chart['MTTR (hrs)'],
                                                          errors='coerce').fillna(0)
                mttr_sorted = mttr_chart.sort_values('MTTR (hrs)', ascending=False)

                fig = px.bar(
                    mttr_sorted,
                    x='Machine',
                    y='MTTR (hrs)',
                    color='MTTR (hrs)',
                    color_continuous_scale=['#27AE60', '#F39C12', '#E74C3C'],
                    text_auto='.2f'
                )
                fig.update_layout(
                    height=500,
                    xaxis_tickangle=-45,
                    yaxis_title="MTTR (hrs)"
                )
                st.plotly_chart(fig, use_container_width=True)

            # ---- MTBF chart (Low = Problematic) ----
            st.markdown("### ⏱️ MTBF per Machine — **Low = Problematic**")
            st.caption("Lower MTBF means more frequent failures — reliability issue.")

            mtbf_chart = chart_df[chart_df['MTBF (hrs)'] != "—"].copy()
            if len(mtbf_chart) > 0:
                mtbf_chart['MTBF (hrs)'] = pd.to_numeric(mtbf_chart['MTBF (hrs)'],
                                                          errors='coerce').fillna(0)
                mtbf_sorted = mtbf_chart.sort_values('MTBF (hrs)', ascending=True)

                fig = px.bar(
                    mtbf_sorted,
                    x='Machine',
                    y='MTBF (hrs)',
                    color='MTBF (hrs)',
                    color_continuous_scale=['#E74C3C', '#F39C12', '#27AE60'],
                    text_auto='.0f'
                )
                fig.update_layout(
                    height=500,
                    xaxis_tickangle=-45,
                    yaxis_title="MTBF (hrs)"
                )
                st.plotly_chart(fig, use_container_width=True)

            # ---- Combined chart ----
            st.markdown("### 📊 Combined — MTBF & MTTR Side by Side")
            st.caption("Ideal machine: **High MTBF + Low MTTR**.")

            combined = chart_df.copy()
            combined['MTTR (hrs)'] = pd.to_numeric(combined['MTTR (hrs)'],
                                                     errors='coerce').fillna(0)
            combined['MTBF (hrs)'] = pd.to_numeric(combined['MTBF (hrs)'],
                                                     errors='coerce').fillna(0)
            combined = combined.sort_values('MTBF (hrs)', ascending=True)

            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=combined['Machine'],
                y=combined['MTBF (hrs)'],
                name='MTBF (hrs)',
                marker_color='#27AE60',
                text=combined['MTBF (hrs)'].round(0),
                textposition='outside',
                yaxis='y'
            ))
            fig.add_trace(go.Bar(
                x=combined['Machine'],
                y=combined['MTTR (hrs)'],
                name='MTTR (hrs)',
                marker_color='#E74C3C',
                text=combined['MTTR (hrs)'].round(2),
                textposition='outside',
                yaxis='y2'
            ))
            fig.update_layout(
                barmode='group',
                height=500,
                xaxis_tickangle=-45,
                yaxis=dict(title='MTBF (hrs)'),
                yaxis2=dict(title='MTTR (hrs)', overlaying='y', side='right'),
                legend=dict(orientation='h', yanchor='bottom', y=1.02,
                            xanchor='right', x=1),
                hovermode='x unified'
            )
            st.plotly_chart(fig, use_container_width=True)

            # ---- Top 10 worst by MTTR ----
            st.markdown("### ⚠️ Top 10 Worst Machines by MTTR")
            worst_mttr = chart_df.copy()
            worst_mttr['MTTR (hrs)'] = pd.to_numeric(worst_mttr['MTTR (hrs)'],
                                                       errors='coerce').fillna(0)
            worst_10 = worst_mttr.sort_values('MTTR (hrs)', ascending=False).head(10)
            st.dataframe(
                worst_10[['Machine', 'Events', 'Breakdown Hours',
                          'MTTR (hrs)', 'MTBF (hrs)', 'Availability %']],
                use_container_width=True, hide_index=True
            )

            # ---- Top 10 worst by MTBF (lowest) ----
            st.markdown("### ⚠️ Top 10 Worst Machines by MTBF (Lowest)")
            worst_mtbf = chart_df[chart_df['MTBF (hrs)'] != "—"].copy()
            if len(worst_mtbf) > 0:
                worst_mtbf['MTBF (hrs)'] = pd.to_numeric(worst_mtbf['MTBF (hrs)'],
                                                          errors='coerce').fillna(0)
                worst_10_mtbf = worst_mtbf.sort_values('MTBF (hrs)', ascending=True).head(10)
                st.dataframe(
                    worst_10_mtbf[['Machine', 'Events', 'Breakdown Hours',
                                   'MTBF (hrs)', 'MTTR (hrs)', 'Availability %']],
                    use_container_width=True, hide_index=True
                )

    # ---------- TAB 6: MACHINE-WISE ----------
    with tab6:
        st.subheader("Machine-wise Breakdown Hours")
        mach = df.groupby('M/C Name').agg(
            Hours=('Total Hrs', 'sum'),
            Loss=('Production Loss', 'sum'),
            Events=('Total Hrs', 'count')
        ).reset_index().sort_values('Hours', ascending=False)
        st.dataframe(mach, use_container_width=True, hide_index=True)

        if len(mach) > 0:
            fig = px.bar(mach, x='M/C Name', y=['Hours', 'Loss'],
                         barmode='group', text_auto='.1f')
            st.plotly_chart(fig, use_container_width=True)

        st.subheader("Stacked — Type per Machine")
        stack = df.groupby(['M/C Name', 'Breakdown Type'])['Total Hrs'].sum().reset_index()
        if len(stack) > 0:
            fig = px.bar(stack, x='M/C Name', y='Total Hrs',
                         color='Breakdown Type', barmode='stack')
            st.plotly_chart(fig, use_container_width=True)

        st.subheader("Shift-wise Breakdown Hours")
        shift_df = df.groupby(['Shift', 'Breakdown Type'])['Total Hrs'].sum().reset_index()
        if len(shift_df) > 0:
            fig = px.bar(shift_df, x='Shift', y='Total Hrs', color='Breakdown Type',
                         barmode='stack', text_auto='.1f')
            st.plotly_chart(fig, use_container_width=True)

    # ---------- TAB 7: CAUSES ----------
    with tab7:
        st.subheader("🔁 Top Repeated Causes")
        causes = df.groupby('Reason of Breakdown').agg(
            Frequency=('Reason of Breakdown', 'count'),
            Hours=('Total Hrs', 'sum'),
            Loss=('Production Loss', 'sum')
        ).reset_index().sort_values('Frequency', ascending=False)
        st.dataframe(causes, use_container_width=True, hide_index=True)

        if len(causes) > 0:
            fig = px.bar(causes.head(10), x='Frequency', y='Reason of Breakdown',
                         orientation='h', color='Hours', text_auto=True)
            fig.update_layout(yaxis={'categoryorder': 'total ascending'}, height=500)
            st.plotly_chart(fig, use_container_width=True)

        st.subheader("🔥 Cause × Machine Matrix (Top 15 Causes)")
        if len(causes) > 0:
            top_causes = causes.head(15)['Reason of Breakdown'].tolist()
            pivot = df[df['Reason of Breakdown'].isin(top_causes)].pivot_table(
                index='M/C Name', columns='Reason of Breakdown',
                values='Total Hrs', aggfunc='count', fill_value=0)
            st.dataframe(pivot, use_container_width=True)

        if len(sel_periods) >= 2 and len(causes) > 0:
            st.subheader("📅 Cause Comparison across Months")
            cause_cmp = df.groupby(['Period', 'Reason of Breakdown'])['Total Hrs'].sum().reset_index()
            top_causes = causes.head(8)['Reason of Breakdown'].tolist()
            cause_cmp_top = cause_cmp[cause_cmp['Reason of Breakdown'].isin(top_causes)]
            if len(cause_cmp_top) > 0:
                fig = px.bar(cause_cmp_top, x='Reason of Breakdown', y='Total Hrs',
                             color='Period', barmode='group', text_auto='.1f')
                fig.update_layout(xaxis_tickangle=-30, height=550)
                st.plotly_chart(fig, use_container_width=True)

    # ---------- TAB 8: YTD (FISCAL YEAR) ----------
    with tab8:
        st.subheader("📆 YTD Analysis (Fiscal Year: April to March)")
        
        fy_options = sorted(df['FY_Label'].dropna().unique())
        if len(fy_options) == 0:
            st.warning("No Fiscal Year data available. Make sure your Excel has a 'Date' column.")
        else:
            sel_fy = st.selectbox("Select Financial Year", fy_options, index=len(fy_options)-1)
            
            ytd_df = df[df['FY_Label'] == sel_fy].copy()
            
            fy_months = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar']
            ytd_df['Month_Name'] = pd.Categorical(ytd_df['Month_Name'], categories=fy_months, ordered=True)
            
            fy_trend = ytd_df.groupby('Month_Name', observed=False).agg(
                Hours=('Total Hrs', 'sum'),
                Loss=('Production Loss', 'sum'),
                Events=('Total Hrs', 'count')
            ).reset_index()
            
            fy_trend = fy_trend[fy_trend['Events'] > 0]
            
            fy_trend['Cumulative_Hours'] = fy_trend['Hours'].cumsum()
            fy_trend['Cumulative_Loss'] = fy_trend['Loss'].cumsum()
            
            st.markdown(f"### YTD Totals — FY {sel_fy}")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("YTD Breakdown Hours", f"{ytd_df['Total Hrs'].sum():,.1f}")
            c2.metric("YTD Production Loss", f"{ytd_df['Production Loss'].sum():,.0f} Kg")
            c3.metric("YTD Cost of Loss", f"₹{ytd_df['Production Loss'].sum() * cost_per_unit:,.0f}")
            c4.metric("YTD Breakdown Events", len(ytd_df))
            
            st.divider()
            
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
            
            st.subheader("📋 Month-wise YTD Breakdown")
            st.dataframe(fy_trend[['Month_Name', 'Hours', 'Loss', 'Events', 'Cumulative_Hours', 'Cumulative_Loss']], use_container_width=True, hide_index=True)


    # ============================================================
    # EXCEL EXPORT
    # ============================================================
    st.divider()
    st.subheader("⬇️ Export Full Analysis to Excel")

    def build_excel_report(df):
        header_fill = PatternFill("solid", fgColor="1F4E78")
        header_font = Font(bold=True, color="FFFFFF")
        buffer = io.BytesIO()

        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:

            kpi = pd.DataFrame({
                "Metric": ["Total Hours", "Total Loss", "Events", "Avg Hours/Event",
                           "Periods", "Machines", "Files"],
                "Value": [
                    round(df['Total Hrs'].sum(), 2),
                    round(df['Production Loss'].sum(), 2),
                    len(df),
                    round(df['Total Hrs'].mean(), 2) if len(df) else 0,
                    ", ".join(sorted(df['Period'].unique())),
                    df['M/C Name'].nunique(),
                    ", ".join(sorted(df['Source File'].unique()))
                ]
            })
            kpi.to_excel(writer, sheet_name="KPI Summary", index=False)

            monthly_cmp = df.groupby('Period_Key').agg(
                Period=('Period', 'first'),
                Breakdown_Hours=('Total Hrs', 'sum'),
                Production_Loss=('Production Loss', 'sum'),
                Events=('Total Hrs', 'count')
            ).reset_index().sort_values('Period_Key')
            monthly_cmp['Hours_Change_%'] = (monthly_cmp['Breakdown_Hours'].pct_change() * 100).round(2)
            monthly_cmp['Loss_Change_%'] = (monthly_cmp['Production_Loss'].pct_change() * 100).round(2)
            monthly_cmp['Cumulative_Hours'] = monthly_cmp['Breakdown_Hours'].cumsum()
            monthly_cmp.drop(columns=['Period_Key']).to_excel(
                writer, sheet_name="Monthly Comparison", index=False)

            yearly_cmp = df.groupby('Year').agg(
                Breakdown_Hours=('Total Hrs', 'sum'),
                Production_Loss=('Production Loss', 'sum'),
                Events=('Total Hrs', 'count'),
                Machines=('M/C Name', 'nunique'),
                Causes=('Reason of Breakdown', 'nunique')
            ).reset_index().sort_values('Year')
            yearly_cmp['Hours_Change_%'] = (yearly_cmp['Breakdown_Hours'].pct_change() * 100).round(2)
            yearly_cmp['Loss_Change_%'] = (yearly_cmp['Production_Loss'].pct_change() * 100).round(2)
            yearly_cmp.to_excel(writer, sheet_name="Yearly Comparison", index=False)

            df.groupby('Period').agg(
                Hours=('Total Hrs', 'sum'),
                Loss=('Production Loss', 'sum'),
                Events=('Total Hrs', 'count')
            ).reset_index().to_excel(writer, sheet_name="By Period", index=False)

            df.groupby('M/C Name').agg(
                Hours=('Total Hrs', 'sum'),
                Loss=('Production Loss', 'sum'),
                Events=('Total Hrs', 'count')
            ).reset_index().sort_values('Hours', ascending=False).to_excel(
                writer, sheet_name="By Machine", index=False)

            df.groupby('Breakdown Type').agg(
                Hours=('Total Hrs', 'sum'),
                Loss=('Production Loss', 'sum'),
                Events=('Total Hrs', 'count')
            ).reset_index().to_excel(writer, sheet_name="By Type", index=False)

            df.groupby('Reason of Breakdown').agg(
                Frequency=('Reason of Breakdown', 'count'),
                Hours=('Total Hrs', 'sum'),
                Loss=('Production Loss', 'sum')
            ).reset_index().sort_values('Frequency', ascending=False).to_excel(
                writer, sheet_name="By Cause", index=False)

            df.pivot_table(index='M/C Name', columns='Period',
                           values='Total Hrs', aggfunc='sum', fill_value=0).to_excel(
                writer, sheet_name="Machine x Period")

            df.pivot_table(index='Breakdown Type', columns='Year',
                           values='Total Hrs', aggfunc='sum', fill_value=0).to_excel(
                writer, sheet_name="Type x Year")

            df.pivot_table(index='M/C Name', columns='Reason of Breakdown',
                           values='Total Hrs', aggfunc='count', fill_value=0).to_excel(
                writer, sheet_name="Cause x Machine")

            df.to_excel(writer, sheet_name="Raw Data", index=False)

            for sheet in writer.sheets.values():
                for column_cells in sheet.columns:
                    max_len = max((len(str(c.value)) if c.value else 0) for c in column_cells)
                    sheet.column_dimensions[column_cells[0].column_letter].width = min(max_len + 3, 50)
                for cell in sheet[1]:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal="center", vertical="center")

        buffer.seek(0)
        return buffer

    excel_data = build_excel_report(df)
    XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    st.download_button(
        "📥 Download Full Multi-Month Analysis (.xlsx)",
        data=excel_data,
        file_name="Breakdown_Analysis_Report.xlsx",
        mime=XLSX_MIME
    )

else:
    st.info("👆 Upload at least one **Breakdown Excel** file to begin.")
    st.markdown("""
    ### 📁 Files you can upload

    | File | Required? | Format |
    |---|---|---|
    | **Breakdown Excel(s)** | ✅ Yes | Monthly breakdown files (e.g., `Break_Jun-25.xlsx`) |
    | **Master Machine List** | Optional | `machines.xlsx` — column `Machine Name` with all machines |
    | **Production Calendar(s)** | Optional | `calendar_Jun-25.xlsx` — columns `Date`, `Available_Hours` |

    ### 🎯 What you get

    - 📊 **Overview** — KPIs, pie charts, monthly trend
    - 📅 **Month Comparison** — MoM % change
    - 📈 **Year Comparison** — YoY trends
    - 🏆 **Yearly Analysis** — full-year totals + month-by-month
    - 🏭 **Capacity Analysis** — availability %, MTBF, MTTR (overall + split by type), machine-wise charts
    - 🔧 **Machine-wise** — per-machine rankings
    - 🔁 **Causes** — repeated causes, cross-month comparison
    - 📆 **YTD (Fiscal Year)** — April to March cumulative analysis
    """)


