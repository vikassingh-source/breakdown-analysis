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

st.title("Breakdown & Production Loss Analysis Dashboard")
st.markdown(
    "**Team Leader:** JC Roy (Mechanical) | "
    "**Team Member:** Vikas Singh (Instrumentation) | "
    "**Team Member:** Ajit Basha (Electrical) | "
    "**Team Member:** Bhaskar Ghormare (Dyeing) | "
    "**Version:** 2.0"
)
st.divider()

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

    date_col = None
    for c in data.columns:
        if 'date' in c.lower():
            date_col = c
            break
            
    if date_col:
        # Convert to string first to strip any hidden Excel number formats
        temp_dates = data[date_col].astype(str).str.strip()
        data['Date'] = pd.to_datetime(temp_dates, dayfirst=True, format='mixed' , errors='coerce').ffill()
        data['FY_Year'] = data['Date'].dt.year + (data['Date'].dt.month >= 4).astype(int)
        data['FY_Label'] = data['FY_Year'].apply(lambda y: f"{int(y)-1}-{str(int(y))[-2:]}" if pd.notna(y) else "Unknown")
    else:
        st.error(f"⚠️ No Date column found in {uploaded.name}")
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
st.sidebar.markdown("### 💰 Cost Settings")
cost_per_unit = st.sidebar.number_input(
    "Cost per unit (₹)",
    min_value=0.0, value=50.0, step=1.0,
    help="Cost of producing one unit in ₹."
)

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
        "🔧 Machine-wise", "🔁 Causes", "📆 YTD (Fiscal Year)"
    ])
# ---------- TAB 1: OVERVIEW ----------
    with tab1:
        total_loss_units = df['Production Loss'].sum()
        total_cost_inr = total_loss_units * cost_per_unit
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Breakdown Hours", f"{df['Total Hrs'].sum():,.1f}")
        c2.metric("Total Production Loss", f"{total_loss_units:,.0f} Kg")
        c3.metric("💸 Cost of Loss", f"₹{total_cost_inr:,.0f}")
        c4.metric("Total Breakdown Events", len(df))
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
            st.divider()
            month_table = year_df.groupby(['Month_Num', 'Month_Name']).agg(
                Breakdown_Hours=('Total Hrs', 'sum'),
                Production_Loss=('Production Loss', 'sum'),
                Events=('Total Hrs', 'count')).reset_index().sort_values('Month_Num')
            month_table['Cumulative_Hours'] = month_table['Breakdown_Hours'].cumsum()
            st.dataframe(month_table.drop(columns=['Month_Num']), use_container_width=True, hide_index=True)

    # ---------- TAB 5: CAPACITY ANALYSIS ----------
    with tab5:
        st.subheader("🏭 Plant Capacity & Availability Analysis")
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

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Machines", total_machines)
        c2.metric("With Breakdown", f"{with_bd_count}")
        c3.metric("Without Breakdown", f"{without_bd_count}")
        c4.metric("Available Hrs/Machine", f"{available_hours_per_machine:,.0f}")
        
        c5, c6, c7, c8 = st.columns(4)
        c5.metric("Availability %", f"{availability_pct:.2f}%")
        c6.metric("Downtime %", f"{downtime_pct:.2f}%")
        c7.metric("MTBF (hrs)", f"{mtbf:,.1f}" if events else "—")
        c8.metric("MTTR (hrs)", f"{mttr:.2f}" if events else "—")

        st.divider()
        chart_df = period_df.groupby('M/C Name')['Total Hrs'].sum().reset_index().sort_values('Total Hrs', ascending=True)
        st.markdown("### 🔧 Breakdown Hours per Machine")
        fig = px.bar(chart_df, y='M/C Name', x='Total Hrs', orientation='h',
                     color='Total Hrs', color_continuous_scale=['#27AE60', '#F39C12', '#E74C3C'])
        fig.update_layout(height=max(500, len(chart_df) * 20), yaxis_title="")
        st.plotly_chart(fig, use_container_width=True)

    # ---------- TAB 6: MACHINE-WISE ----------
    with tab6:
        st.subheader("Machine-wise Breakdown Hours")
        mach = df.groupby('M/C Name').agg(Hours=('Total Hrs', 'sum'),
                                            Loss=('Production Loss', 'sum'),
                                            Events=('Total Hrs', 'count')).reset_index().sort_values(
            'Hours', ascending=False)
        st.dataframe(mach, use_container_width=True, hide_index=True)

    # ---------- TAB 7: CAUSES ----------
    with tab7:
        st.subheader("🔁 Top Repeated Causes")
        causes = df.groupby('Reason of Breakdown').agg(
            Frequency=('Reason of Breakdown', 'count'), Hours=('Total Hrs', 'sum'),
            Loss=('Production Loss', 'sum')).reset_index().sort_values('Frequency', ascending=False)
        st.dataframe(causes, use_container_width=True, hide_index=True)

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
    # EXPORTS
    # ============================================================
    st.divider()
    st.subheader("⬇️ Export Full Analysis")
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("📊 Prepare Excel Report", use_container_width=True):
            with st.spinner("Generating Excel..."):
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df.groupby('M/C Name').agg(Hours=('Total Hrs', 'sum'), Loss=('Production Loss', 'sum')).to_excel(writer, sheet_name="By Machine")
                    df.groupby('Reason of Breakdown').agg(Freq=('Reason of Breakdown', 'count')).to_excel(writer, sheet_name="By Cause")
                    df.to_excel(writer, sheet_name="Raw Data")
                XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                st.download_button("📥 Download Excel Report", data=buffer.getvalue(), file_name="Breakdown_Report.xlsx", mime=XLSX_MIME, use_container_width=True)
    
    with col2:
        if st.button("🎨 Prepare PowerPoint", use_container_width=True):
            with st.spinner("Generating PowerPoint..."):
                prs = Presentation()
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1))
                tf = txBox.text_frame
                tf.text = "Breakdown & Production Loss Analysis"
                tf.paragraphs[0].font.size = Pt(32)
                tf.paragraphs[0].font.color.rgb = RGBColor(0x1F, 0x4E, 0x78)
                
                buffer = io.BytesIO()
                prs.save(buffer)
                buffer.seek(0)
                PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                st.download_button("📥 Download PowerPoint", data=buffer, file_name="Breakdown_Presentation.pptx", mime=PPTX_MIME, use_container_width=True)

else:
    st.info("👆 Upload at least one **Breakdown Excel** file to begin.")





