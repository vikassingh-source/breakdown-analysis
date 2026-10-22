import io
import re
import calendar

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.enum.shapes import MSO_SHAPE
    PPTX_OK = True
except ImportError:
    PPTX_OK = False

# =====================================================================
# 1. PAGE SETUP
# =====================================================================
st.set_page_config(page_title="Breakdown & Production Loss Analysis Dashboard", layout="wide")

st.markdown(
    """
    <div style="background-color:#1F3A5F;padding:16px 20px;border-radius:8px;color:white;margin-bottom:12px;">
        <h2 style="margin:0;color:white;">Breakdown &amp; Production Loss Analysis Dashboard</h2>
        <p style="margin:4px 0 0 0;color:white;">
            Created by: Vikas Singh | Department: Instrumentation | Version: 2.0
        </p>
        <p style="margin:4px 0 0 0;color:white;font-size:14px;">
            <b>CFT Members:</b> TL J.C Roy (Mechanical), Ajit Basha (Electrical),
            Vikas Singh (Instrumentation), Bhaskar Ghormare (Dyeing)
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)


# ---------- small UI helpers (work across Streamlit versions) ----------
def show_chart(fig):
    try:
        st.plotly_chart(fig, width="stretch")
    except TypeError:
        st.plotly_chart(fig, use_container_width=True)


def show_df(d, **kw):
    try:
        st.dataframe(d, width="stretch", **kw)
    except TypeError:
        st.dataframe(d, use_container_width=True, **kw)


# =====================================================================
# 3. DATA PROCESSING FUNCTIONS
# =====================================================================
MONTHS = {m.lower(): i for i, m in enumerate(calendar.month_abbr) if m}
FULL_MONTHS = {m.lower(): i for i, m in enumerate(calendar.month_name) if m}


def label_from_period(period):
    y, m = period.split("-")
    return f"{calendar.month_abbr[int(m)]}-{y[-2:]}"


def parse_period(filename):
    """Return ('YYYY-MM', 'Mon-YY') from a filename, or (None, stem)."""
    name = str(filename)
    m = re.search(r"(20\d{2})[-_ .]?(0[1-9]|1[0-2])(?!\d)", name)
    if m:
        period = f"{m.group(1)}-{m.group(2)}"
        return period, label_from_period(period)
    m = re.search(r"([A-Za-z]{3,9})[-_ .]*'?((?:20)?\d{2})(?!\d)", name)
    if m:
        key = m.group(1).lower()
        mon = MONTHS.get(key[:3]) if (key[:3] in MONTHS) else None
        if key in FULL_MONTHS:
            mon = FULL_MONTHS[key]
        if mon:
            yr = m.group(2)
            yr = int(yr) + 2000 if len(yr) == 2 else int(yr)
            period = f"{yr}-{mon:02d}"
            return period, label_from_period(period)
    return None, re.sub(r"\.[^.]+$", "", name)


def map_type(value):
    v = str(value).strip().lower()
    if v in ("nan", "", "none"):
        return "Other"
    if "inst" in v:
        return "Instrumentation"
    if "elec" in v:
        return "Electrical"
    if "mech" in v:
        return "Mechanical"
    return "Other"


def norm_name(x):
    return re.sub(r"\s+", " ", str(x)).strip().upper()


def find_col(cols, keys, exclude=()):
    for k in keys:
        for c in cols:
            lc = c.lower()
            if k in lc and not any(e in lc for e in exclude):
                return c
    return None


def read_breakdown_file(uploaded):
    """Read every sheet of a breakdown workbook and return a clean DataFrame."""
    period, _ = parse_period(uploaded.name)
    frames = []
    xl = pd.ExcelFile(uploaded)
    for sheet in xl.sheet_names:
        raw = xl.parse(sheet)
        if raw.empty:
            continue
        raw.columns = [re.sub(r"[.\s]+", "", str(c)) for c in raw.columns]
        cols = list(raw.columns)
        c_mc = find_col(cols, ["m/c", "mcname", "machine"])
        c_date = find_col(cols, ["date"])
        c_hrs = find_col(cols, ["totalhr", "totalhour", "hrs", "hours"])
        c_loss = find_col(cols, ["productionloss", "prodloss", "loss"])
        c_work = find_col(cols, ["workedby", "workby", "worked"])
        c_shift = find_col(cols, ["shift"])
        c_cause = find_col(cols, ["cause", "reason", "problem", "remark"])
        if c_mc is None or c_hrs is None:
            continue

        d = pd.DataFrame()
        mc = raw[c_mc].replace(r"^\s*$", pd.NA, regex=True).ffill()
        d["Machine"] = mc.astype(str).str.replace(r"\s+", " ", regex=True).str.strip()
        if c_date:
            d["Date"] = pd.to_datetime(raw[c_date].replace(r"^\s*$", pd.NA, regex=True).ffill(),
                                       errors="coerce", dayfirst=True)
        else:
            d["Date"] = pd.NaT
        d["Hours"] = pd.to_numeric(raw[c_hrs], errors="coerce")
        d["Loss"] = pd.to_numeric(raw[c_loss], errors="coerce").fillna(0) if c_loss else 0.0
        d["Type"] = raw[c_work].apply(map_type) if c_work else "Other"
        d["Shift"] = raw[c_shift].astype(str).str.strip().replace({"nan": "Unknown", "": "Unknown"}) if c_shift else "Unknown"
        d["Cause"] = raw[c_cause].astype(str).str.strip().replace({"nan": "Unknown", "": "Unknown"}) if c_cause else "Unknown"
        d["Sheet"] = sheet
        d["Source"] = uploaded.name

        d = d[d["Machine"].notna() & (d["Machine"].str.lower() != "nan")]
        d = d[~d["Machine"].str.lower().str.contains("total", na=False)]
        d = d[d["Hours"].notna()]
        if period:
            d["Period"] = period
        else:
            d["Period"] = d["Date"].dt.strftime("%Y-%m")
        d = d[d["Period"].notna()]
        frames.append(d)

    if not frames:
        return pd.DataFrame()
    out = pd.concat(frames, ignore_index=True)
    out["Year"] = out["Period"].str[:4].astype(int)
    out["Month"] = out["Period"].str[5:7].astype(int)
    out["Period_Label"] = out["Period"].apply(label_from_period)
    out["FY_Year"] = out.apply(lambda r: r["Year"] + 1 if r["Month"] >= 4 else r["Year"], axis=1)
    out["FY_Label"] = out["FY_Year"].apply(lambda y: f"{y - 1}-{str(y)[-2:]}")
    return out


def read_master(uploaded):
    xl = pd.ExcelFile(uploaded)
    d = xl.parse(xl.sheet_names[0])
    cols = [str(c) for c in d.columns]
    d.columns = cols
    col = find_col(cols, ["machine", "m/c", "name"]) or cols[0]
    names = d[col].dropna().astype(str).str.replace(r"\s+", " ", regex=True).str.strip()
    names = names[names != ""]
    return list(dict.fromkeys(names))


def read_calendars(files):
    """Return {'YYYY-MM': available hours}."""
    out = {}
    for f in files:
        xl = pd.ExcelFile(f)
        for sheet in xl.sheet_names:
            d = xl.parse(sheet)
            cols = [str(c) for c in d.columns]
            d.columns = cols
            c_date = find_col(cols, ["date"])
            c_hr = find_col(cols, ["hour", "avail"])
            if not c_date or not c_hr:
                continue
            tmp = pd.DataFrame({
                "d": pd.to_datetime(d[c_date], errors="coerce", dayfirst=True),
                "h": pd.to_numeric(d[c_hr], errors="coerce"),
            }).dropna()
            if tmp.empty:
                continue
            g = tmp.groupby(tmp["d"].dt.strftime("%Y-%m"))["h"].sum()
            for k, v in g.items():
                out[k] = out.get(k, 0) + float(v)
    return out


def pivot(d, index, columns, order=None, values="Hours"):
    p = d.pivot_table(index=index, columns=columns, values=values, aggfunc="sum", fill_value=0)
    if order is not None:
        p = p.reindex(columns=[c for c in order if c in p.columns])
    p["Total"] = p.sum(axis=1)
    return p.sort_values("Total", ascending=False).round(2)


def pct_change(cur, prev):
    if prev in (0, None) or pd.isna(prev):
        return None
    return (cur - prev) / prev * 100


# =====================================================================
# 5. EXCEL EXPORT
# =====================================================================
def build_excel_report(df):
    out = io.BytesIO()
    order = df.drop_duplicates("Period").sort_values("Period")["Period_Label"].tolist()

    kpi = pd.DataFrame({
        "Metric": ["Total Breakdown Hours", "Total Production Loss", "Total Events",
                   "Avg Hours / Event", "Machines Affected", "Periods Covered"],
        "Value": [round(df["Hours"].sum(), 2), round(df["Loss"].sum(), 2), len(df),
                  round(df["Hours"].mean(), 2), df["Machine"].nunique(), df["Period"].nunique()],
    })

    by_period = (df.groupby(["Period", "Period_Label"])
                 .agg(Hours=("Hours", "sum"), Loss=("Loss", "sum"), Events=("Hours", "count"))
                 .reset_index().sort_values("Period"))
    monthly = by_period.copy()
    monthly["Hours MoM %"] = monthly["Hours"].pct_change() * 100
    monthly["Loss MoM %"] = monthly["Loss"].pct_change() * 100
    monthly = monthly.round(2)

    yearly = df.groupby("Year").agg(Hours=("Hours", "sum"), Loss=("Loss", "sum"),
                                    Events=("Hours", "count")).reset_index()
    yearly["Hours YoY %"] = yearly["Hours"].pct_change() * 100
    yearly = yearly.round(2)

    by_machine = (df.groupby("Machine").agg(Hours=("Hours", "sum"), Loss=("Loss", "sum"),
                                            Events=("Hours", "count"))
                  .reset_index().sort_values("Hours", ascending=False))
    by_type = df.groupby("Type").agg(Hours=("Hours", "sum"), Loss=("Loss", "sum"),
                                     Events=("Hours", "count")).reset_index()
    by_cause = (df.groupby("Cause").agg(Hours=("Hours", "sum"), Loss=("Loss", "sum"),
                                        Events=("Hours", "count"))
                .reset_index().sort_values("Hours", ascending=False))

    m_p = pivot(df, "Machine", "Period_Label", order).reset_index()
    t_y = pivot(df, "Type", "Year").reset_index()
    c_m = pivot(df, "Cause", "Machine").reset_index()
    raw = df.copy()
    if "Date" in raw.columns:
        raw["Date"] = raw["Date"].dt.strftime("%Y-%m-%d")

    sheets = {
        "KPI Summary": kpi,
        "Monthly Comparison": monthly,
        "Yearly Comparison": yearly,
        "By Period": by_period.round(2),
        "By Machine": by_machine.round(2),
        "By Type": by_type.round(2),
        "By Cause": by_cause.round(2),
        "Machine x Period": m_p,
        "Type x Year": t_y,
        "Cause x Machine": c_m,
        "Raw Data": raw,
    }

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=name[:31], index=False)
        hfill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        hfont = Font(bold=True, color="FFFFFF")
        for ws in writer.book.worksheets:
            for cell in ws[1]:
                cell.fill = hfill
                cell.font = hfont
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            for col in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0 for c in col[:200])
                ws.column_dimensions[col[0].column_letter].width = min(max(width + 2, 10), 45)
            ws.freeze_panes = "A2"
    return out.getvalue()


# =====================================================================
# 6. POWERPOINT EXPORT (native, editable charts & tables - PowerPoint / WPS)
# =====================================================================
PPT_PALETTE = ["1F4E78", "E74C3C", "2ECC71", "F39C12", "8E44AD", "16A085", "D35400", "2980B9", "7F8C8D", "C0392B"]
PPT_SET2 = ["66C2A5", "FC8D62", "8DA0CB", "E78AC3", "A6D854", "FFD92F", "E5C494", "B3B3B3"]
PPT_PASTEL = ["FBB4AE", "B3CDE3", "CCEBC5", "DECBE4", "FED9A6", "FFFFCC", "E5D8BD", "FDDAEC"]


def _rgb(h):
    return RGBColor.from_string(h)


def _short(x, n=32):
    x = str(x)
    return x if len(x) <= n else x[: n - 1] + "…"


def _grad(vals, high_bad=True):
    """Green-Yellow-Red colours (hex) for a list of values."""
    vals = [0 if pd.isna(v) else float(v) for v in vals]
    lo, hi = (min(vals), max(vals)) if vals else (0, 1)
    out = []
    for v in vals:
        f = 0.5 if hi == lo else (v - lo) / (hi - lo)
        if not high_bad:
            f = 1 - f
        r = 255 if f > 0.5 else int(510 * f)
        g = 200 if f < 0.5 else int(200 * (1 - (f - 0.5) * 2))
        out.append(f"{r:02X}{g:02X}40")
    return out


def _ppt_text(slide, text, l, t, w, h, size=14, bold=False, color="000000", align=None):
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.text = str(text)
    for p in tf.paragraphs:
        p.font.size = Pt(size)
        p.font.bold = bold
        p.font.color.rgb = _rgb(color)
        if align is not None:
            p.alignment = align
    return tb


def _ppt_slide(prs, title):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.8))
    bar.fill.solid()
    bar.fill.fore_color.rgb = _rgb("1F4E78")
    bar.line.fill.background()
    tf = bar.text_frame
    tf.text = title
    tf.margin_left = Inches(0.4)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = _rgb("FFFFFF")
    p.alignment = PP_ALIGN.LEFT
    _ppt_text(s, "Created by: Vikas Singh | Department: Instrumentation | Version: 2.0",
              0.4, 7.1, 9, 0.3, 10, False, "7F7F7F")
    return s


def _ppt_kpis(slide, items, top=1.05, per_row=4, h=1.25):
    gap, total = 0.2, 12.5
    w = (total - gap * (per_row - 1)) / per_row
    for i, (label, value) in enumerate(items):
        r, c = divmod(i, per_row)
        l = 0.4 + c * (w + gap)
        t = top + r * (h + 0.15)
        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(h))
        box.fill.solid()
        box.fill.fore_color.rgb = _rgb("EAF1F8")
        box.line.color.rgb = _rgb("1F4E78")
        tf = box.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.text = str(label)
        p0 = tf.paragraphs[0]
        p0.font.size = Pt(12)
        p0.font.color.rgb = _rgb("595959")
        p0.alignment = PP_ALIGN.CENTER
        p1 = tf.add_paragraph()
        p1.text = str(value)
        p1.font.size = Pt(22)
        p1.font.bold = True
        p1.font.color.rgb = _rgb("1F4E78")
        p1.alignment = PP_ALIGN.CENTER


def _ppt_chart(slide, ctype, cats, series, title, l, t, w, h, colors=None,
               point_colors=None, labels=False, reverse=False, numfmt="#,##0.0"):
    C = XL_CHART_TYPE
    cd = CategoryChartData()
    cd.categories = [_short(c) for c in cats]
    for name, vals in series.items():
        cd.add_series(str(name), [None if pd.isna(v) else float(v) for v in vals])
    gf = slide.shapes.add_chart(ctype, Inches(l), Inches(t), Inches(w), Inches(h), cd)
    ch = gf.chart
    ch.has_title = True
    ch.chart_title.text_frame.text = title
    tp = ch.chart_title.text_frame.paragraphs[0]
    tp.font.size = Pt(14)
    tp.font.bold = True
    ch.has_legend = len(series) > 1 or ctype == C.PIE
    if ch.has_legend:
        ch.legend.position = XL_LEGEND_POSITION.BOTTOM
        ch.legend.include_in_layout = False
        ch.legend.font.size = Pt(10)
    colors = colors or PPT_PALETTE
    plot = ch.plots[0]

    if ctype == C.PIE:
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.show_percentage = True
        dl.show_value = False
        dl.number_format = "0.0%"
        dl.number_format_is_linked = False
        dl.font.size = Pt(11)
        for i in range(len(list(cats))):
            pt = plot.series[0].points[i]
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = _rgb(colors[i % len(colors)])
        return ch

    is_line = ctype == C.LINE_MARKERS
    for i, ser in enumerate(plot.series):
        col = _rgb(colors[i % len(colors)])
        if is_line:
            ser.format.line.color.rgb = col
            ser.format.line.width = Pt(3)
            ser.smooth = False
        else:
            ser.format.fill.solid()
            ser.format.fill.fore_color.rgb = col
    if point_colors and not is_line:
        for i, hx in enumerate(point_colors):
            pt = plot.series[0].points[i]
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = _rgb(hx)
    if not is_line:
        try:
            plot.gap_width = 60
        except Exception:  # noqa: BLE001
            pass
    if labels:
        plot.has_data_labels = True
        plot.data_labels.font.size = Pt(8)
        plot.data_labels.number_format = numfmt
        plot.data_labels.number_format_is_linked = False
    ch.category_axis.tick_labels.font.size = Pt(9)
    ch.value_axis.tick_labels.font.size = Pt(9)
    ch.value_axis.tick_labels.number_format = "#,##0"
    ch.value_axis.tick_labels.number_format_is_linked = False
    if reverse:
        ch.category_axis.reverse_order = True
    return ch


def _ppt_pivot_chart(slide, ctype, p, title, l, t, w, h, **kw):
    """p: index = categories, columns = series."""
    p = p.drop(columns="Total", errors="ignore")
    return _ppt_chart(slide, ctype, [str(i) for i in p.index],
                      {str(c): p[c].tolist() for c in p.columns}, title, l, t, w, h, **kw)


def _fmt_cell(v):
    if pd.isna(v):
        return "-"
    if pd.api.types.is_integer(v):
        return f"{int(v):,}"
    if pd.api.types.is_float(v):
        return f"{v:,.2f}"
    return _short(v, 45)


def _ppt_table(slide, d, l, t, w, h, fs=10):
    rows, cols = d.shape
    shape = slide.shapes.add_table(rows + 1, cols, Inches(l), Inches(t), Inches(w), Inches(h))
    tbl = shape.table
    first_w = w * 0.25 if cols > 3 else w / cols
    other_w = (w - first_w) / (cols - 1) if cols > 1 else w
    tbl.columns[0].width = Inches(first_w)
    for j in range(1, cols):
        tbl.columns[j].width = Inches(other_w)
    row_h = Inches(h / (rows + 1))
    for i in range(rows + 1):
        tbl.rows[i].height = row_h
    for j, c in enumerate(d.columns):
        cell = tbl.cell(0, j)
        cell.text = _short(c, 25)
        cell.fill.solid()
        cell.fill.fore_color.rgb = _rgb("1F4E78")
        p = cell.text_frame.paragraphs[0]
        p.font.size = Pt(fs)
        p.font.bold = True
        p.font.color.rgb = _rgb("FFFFFF")
        p.alignment = PP_ALIGN.CENTER
    for i in range(rows):
        for j in range(cols):
            cell = tbl.cell(i + 1, j)
            cell.text = _fmt_cell(d.iat[i, j])
            cell.fill.solid()
            cell.fill.fore_color.rgb = _rgb("F2F6FA" if i % 2 else "FFFFFF")
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(fs)
            p.font.color.rgb = _rgb("000000")
            p.alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.RIGHT
            cell.margin_top = cell.margin_bottom = Inches(0.02)
    return tbl


def _ppt_table_slides(prs, title, d, rows_per=14, max_cols=9, fs=10):
    if d is None or d.empty:
        return
    d = d.copy()
    cols = list(d.columns)
    first, rest = cols[0], cols[1:]
    per = max(max_cols - 1, 1)
    col_chunks = [rest[i:i + per] for i in range(0, len(rest), per)] or [[]]
    row_chunks = [d.iloc[i:i + rows_per] for i in range(0, len(d), rows_per)]
    total, n = len(col_chunks) * len(row_chunks), 0
    for cc in col_chunks:
        for rc in row_chunks:
            n += 1
            s = _ppt_slide(prs, title + (f" ({n}/{total})" if total > 1 else ""))
            _ppt_table(s, rc[[first] + cc], 0.4, 1.05, 12.5, min(0.36 * (len(rc) + 1), 5.9), fs)


def _agg(d, by):
    return (d.groupby(by).agg(Hours=("Hours", "sum"), Loss=("Loss", "sum"), Events=("Hours", "count")))


def build_ppt_report(df, cost_per_unit, cal_hours, master_list, period_order, label_to_code, cap_label):
    C = XL_CHART_TYPE
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    types = sorted(df["Type"].unique())
    years = sorted(df["Year"].unique())
    month_abbr = [calendar.month_abbr[i] for i in range(1, 13)]
    LW, RL = 6.1, 6.8   # half-width chart size / right column left

    # ------------------------------------------------ title slide
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = _rgb("1F4E78")
    bg.line.fill.background()
    _ppt_text(s, "Breakdown & Production Loss Analysis", 0.8, 1.9, 11.7, 1.2, 38, True, "FFFFFF")
    _ppt_text(s, "Created by: Vikas Singh | Department: Instrumentation | Version: 2.0", 0.8, 3.3, 11.7, 0.5, 18, False, "FFFFFF")
    _ppt_text(s, "CFT Members: TL J.C Ray (Mechanical), Ajit Dasha (Electrical), "
                 "Vikas Singh (Instrumentation), Bhaskar Ghormare (Dyeing)", 0.8, 3.9, 11.7, 0.9, 16, False, "FFFFFF")
    _ppt_text(s, f"Data: {period_order[0]} to {period_order[-1]}  |  {len(df):,} events  |  "
                 f"Generated {datetime.now():%d-%b-%Y}", 0.8, 5.0, 11.7, 0.5, 14, False, "DDEBF7")

    # ================================================ TAB 1 - OVERVIEW
    bt = _agg(df, "Type").reset_index().round(2)
    s = _ppt_slide(prs, "1. Overview - KPIs")
    _ppt_kpis(s, [("Total Hours", f"{df['Hours'].sum():,.1f}"), ("Total Loss", f"{df['Loss'].sum():,.1f}"),
                  ("Total Events", f"{len(df):,}"), ("Avg Hours/Event", f"{df['Hours'].mean():,.2f}")])
    _ppt_table(s, bt, 0.4, 2.8, 12.5, 0.4 * (len(bt) + 1), 12)

    s = _ppt_slide(prs, "1. Overview - Distribution by Breakdown Type")
    if bt["Hours"].sum() > 0:
        _ppt_chart(s, C.PIE, bt["Type"], {"Hours": bt["Hours"]}, "Breakdown Hours by Type", 0.4, 1.0, LW, 5.9, colors=PPT_SET2)
    if bt["Loss"].sum() > 0:
        _ppt_chart(s, C.PIE, bt["Type"], {"Loss": bt["Loss"]}, "Production Loss by Type", RL, 1.0, LW, 5.9, colors=PPT_PASTEL)

    mon = _agg(df, "Period_Label").reindex(period_order)
    s = _ppt_slide(prs, "1. Overview - Monthly Hours & Production Loss")
    _ppt_chart(s, C.COLUMN_CLUSTERED, mon.index, {"Hours": mon["Hours"]}, "Monthly Breakdown Hours", 0.4, 1.0, LW, 5.9,
               colors=["1F4E78"], labels=True)
    _ppt_chart(s, C.LINE_MARKERS, mon.index, {"Production Loss": mon["Loss"]}, "Monthly Production Loss", RL, 1.0, LW, 5.9,
               colors=["E74C3C"])

    # ================================================ TAB 2 - MONTH COMPARISON
    mp = mon.copy()
    for c in ["Hours", "Loss", "Events"]:
        mp[f"{c} MoM %"] = mp[c].pct_change() * 100
    if len(period_order) >= 2:
        s = _ppt_slide(prs, f"2. Month Comparison - {period_order[-1]} vs {period_order[-2]}")
        items = []
        for c in ["Hours", "Loss", "Events"]:
            ch = mp[f"{c} MoM %"].iloc[-1]
            items.append((f"{c} ({period_order[-1]})",
                          f"{mp[c].iloc[-1]:,.1f}  ({ch:+.1f}%)" if pd.notna(ch) else f"{mp[c].iloc[-1]:,.1f}  (n/a)"))
        _ppt_kpis(s, items, per_row=3)
    _ppt_table_slides(prs, "2. Month Comparison - Monthly Summary & MoM %", mp.reset_index().rename(
        columns={"Period_Label": "Period"}).round(2), rows_per=14, max_cols=7)

    pv = df.pivot_table(index="Period_Label", columns="Type", values="Hours", aggfunc="sum", fill_value=0).reindex(period_order)
    s = _ppt_slide(prs, "2. Month Comparison - Hours by Breakdown Type")
    _ppt_pivot_chart(s, C.COLUMN_CLUSTERED, pv, "Hours by Breakdown Type per Month", 0.4, 1.0, 12.5, 5.9)

    top15 = df.groupby("Machine")["Hours"].sum().nlargest(15).index.tolist()
    last6 = period_order[-6:]
    pv = (df[df["Machine"].isin(top15)].pivot_table(index="Machine", columns="Period_Label", values="Hours",
                                                    aggfunc="sum", fill_value=0).reindex(top15))
    pv = pv[[c for c in last6 if c in pv.columns]]
    s = _ppt_slide(prs, "2. Month Comparison - Hours by Machine (Top 15 machines, last 6 months)")
    _ppt_pivot_chart(s, C.COLUMN_CLUSTERED, pv, "Hours by Machine per Month", 0.4, 1.0, 12.5, 5.9)

    _ppt_table_slides(prs, "2. Month Comparison - Machine x Month (Hours)",
                      pivot(df, "Machine", "Period_Label", period_order).reset_index(), rows_per=13, max_cols=9, fs=9)
    _ppt_table_slides(prs, "2. Month Comparison - Breakdown Type x Month (Hours)",
                      pivot(df, "Type", "Period_Label", period_order).reset_index(), rows_per=13, max_cols=9, fs=9)

    # ================================================ TAB 3 - YEAR COMPARISON
    ty = df.pivot_table(index="Year", columns="Type", values="Hours", aggfunc="sum", fill_value=0)
    s = _ppt_slide(prs, "3. Year Comparison - Hours by Year")
    _ppt_pivot_chart(s, C.COLUMN_CLUSTERED, ty, "Hours by Year & Breakdown Type", 0.4, 1.0, LW, 5.9)
    my = (df.pivot_table(index="Month", columns="Year", values="Hours", aggfunc="sum", fill_value=0)
          .reindex(range(1, 13)).fillna(0))
    my.index = month_abbr
    _ppt_pivot_chart(s, C.COLUMN_CLUSTERED, my, "Hours by Month & Year", RL, 1.0, LW, 5.9)
    _ppt_table_slides(prs, "3. Year Comparison - Breakdown Type x Year (Hours)", pivot(df, "Type", "Year").reset_index(), max_cols=9)
    _ppt_table_slides(prs, "3. Year Comparison - Machine x Year (Hours)", pivot(df, "Machine", "Year").reset_index(),
                      rows_per=13, max_cols=9, fs=9)

    # ================================================ TAB 4 - YEARLY ANALYSIS
    for yr_ in years:
        dy = df[df["Year"] == yr_]
        mt = _agg(dy, ["Month", "Period_Label"]).reset_index().sort_values("Month")
        mt["Cumulative Hours"] = mt["Hours"].cumsum()
        s = _ppt_slide(prs, f"4. Yearly Analysis - {yr_}")
        _ppt_kpis(s, [("Total Hours", f"{dy['Hours'].sum():,.1f}"), ("Total Loss", f"{dy['Loss'].sum():,.1f}"),
                      ("Total Events", f"{len(dy):,}"), ("Avg Hours/Event", f"{dy['Hours'].mean():,.2f}")], h=1.1)
        _ppt_chart(s, C.COLUMN_CLUSTERED, mt["Period_Label"], {"Monthly Hours": mt["Hours"]}, "Monthly Hours",
                   0.4, 2.45, LW, 4.5, colors=["1F4E78"], labels=True)
        _ppt_chart(s, C.LINE_MARKERS, mt["Period_Label"], {"Cumulative Hours": mt["Cumulative Hours"]}, "Cumulative Hours",
                   RL, 2.45, LW, 4.5, colors=["E74C3C"])
        _ppt_table_slides(prs, f"4. Yearly Analysis - {yr_} Month-by-Month",
                          mt.drop(columns="Month").rename(columns={"Period_Label": "Month"}).round(2))

    # ================================================ TAB 5 - CAPACITY
    code = label_to_code[cap_label]
    yy, mm = int(code[:4]), int(code[5:7])
    days = calendar.monthrange(yy, mm)[1]
    if code in cal_hours:
        avail_hrs, src = cal_hours[code], "production calendar"
    else:
        avail_hrs, src = 20 * days, f"default 20 hrs x {days} days"
    universe = master_list if master_list else sorted(df["Machine"].unique())
    dfp = df[df["Period_Label"] == cap_label]

    rows = []
    for mc in universe:
        dm = dfp[dfp["Machine"] == mc]
        hrs, ev = dm["Hours"].sum(), len(dm)
        up = max(avail_hrs - hrs, 0)
        rows.append({"Machine": mc, "Hours": round(hrs, 2), "Events": ev, "Loss": round(dm["Loss"].sum(), 2),
                     "Availability %": round(up / avail_hrs * 100, 2) if avail_hrs else None,
                     "MTBF (hrs)": round(up / ev, 2) if ev else None,
                     "MTTR (hrs)": round(hrs / ev, 2) if ev else None})
    ms_all = pd.DataFrame(rows)
    n_m = len(universe)
    with_bd = int((ms_all["Events"] > 0).sum())
    cap = n_m * avail_hrs
    bd_h = ms_all["Hours"].sum()
    tot_ev = int(ms_all["Events"].sum())
    av = (cap - bd_h) / cap * 100 if cap else 0
    mtbf = (cap - bd_h) / tot_ev if tot_ev else None
    mttr = bd_h / tot_ev if tot_ev else None

    s = _ppt_slide(prs, f"5. Capacity Analysis - {cap_label}")
    _ppt_kpis(s, [("Total Machines", n_m), ("With Breakdown", with_bd), ("Without Breakdown", n_m - with_bd),
                  ("Plant Capacity (hrs)", f"{cap:,.0f}"), ("Breakdown Hours", f"{bd_h:,.1f}"),
                  ("Availability %", f"{av:.2f}%"), ("Downtime %", f"{100 - av:.2f}%"), ("Events", tot_ev),
                  ("MTBF - Overall (hrs)", f"{mtbf:,.2f}" if mtbf else "n/a"),
                  ("MTTR - Overall (hrs)", f"{mttr:,.2f}" if mttr else "n/a")], per_row=5, h=1.3)
    _ppt_text(s, f"Available hours per machine: {avail_hrs:,.1f} (source: {src})", 0.4, 4.2, 12.5, 0.4, 13, False, "595959")

    s = _ppt_slide(prs, f"5. Capacity Analysis - Machines & Downtime ({cap_label})")
    _ppt_chart(s, C.PIE, ["With Breakdown", "Without Breakdown"], {"Machines": [with_bd, n_m - with_bd]},
               "Machines With / Without Breakdown", 0.4, 1.0, LW, 5.9, colors=["E74C3C", "2ECC71"])
    dtp = dfp.groupby("Type")["Hours"].sum()
    if dtp.sum() > 0:
        _ppt_chart(s, C.PIE, dtp.index, {"Hours": dtp.values}, "Downtime by Type", RL, 1.0, LW, 5.9, colors=PPT_SET2)

    _ppt_table_slides(prs, f"5. Capacity Analysis - Machine Statistics ({cap_label})", ms_all, rows_per=14, max_cols=7)

    trows = []
    for mc in universe:
        for tp in types:
            dm = dfp[(dfp["Machine"] == mc) & (dfp["Type"] == tp)]
            if len(dm):
                hrs = dm["Hours"].sum()
                trows.append({"Machine": mc, "Type": tp, "Hours": hrs, "MTTR": hrs / len(dm),
                              "MTBF": max(avail_hrs - hrs, 0) / len(dm)})
    tdf = pd.DataFrame(trows)
    if not tdf.empty:
        topm = tdf.groupby("Machine")["Hours"].sum().nlargest(15).index.tolist()
        for metric, ttl in [("MTTR", "MTTR by Machine & Type (hrs)"), ("MTBF", "MTBF by Machine & Type (hrs)")]:
            pvt = tdf[tdf["Machine"].isin(topm)].pivot_table(index="Machine", columns="Type", values=metric,
                                                              aggfunc="mean", fill_value=0).reindex(topm)
            s = _ppt_slide(prs, f"5. Capacity Analysis - {ttl} ({cap_label}, top 15 machines)")
            _ppt_pivot_chart(s, C.COLUMN_CLUSTERED, pvt, ttl, 0.4, 1.0, 12.5, 5.9)

    ms = ms_all.dropna(subset=["MTTR (hrs)", "MTBF (hrs)"])
    if not ms.empty:
        a = ms.sort_values("MTTR (hrs)", ascending=False).head(15)
        b = ms.sort_values("MTBF (hrs)", ascending=True).head(15)
        s = _ppt_slide(prs, f"5. Capacity Analysis - Machine-wise MTTR & MTBF ({cap_label}, worst 15)")
        _ppt_chart(s, C.BAR_CLUSTERED, a["Machine"], {"MTTR (hrs)": a["MTTR (hrs)"]}, "MTTR (highest to lowest)",
                   0.4, 1.0, LW, 5.9, point_colors=_grad(a["MTTR (hrs)"], True), reverse=True, labels=True, numfmt="0.0")
        _ppt_chart(s, C.BAR_CLUSTERED, b["Machine"], {"MTBF (hrs)": b["MTBF (hrs)"]}, "MTBF (lowest to highest)",
                   RL, 1.0, LW, 5.9, point_colors=_grad(b["MTBF (hrs)"], False), reverse=True, labels=True, numfmt="0.0")

        s = _ppt_slide(prs, f"5. Capacity Analysis - MTTR vs MTBF ({cap_label}, top 15 by MTTR)")
        _ppt_chart(s, C.COLUMN_CLUSTERED, a["Machine"], {"MTTR (hrs)": a["MTTR (hrs)"]}, "MTTR (hrs)",
                   0.4, 1.0, LW, 5.9, colors=["E74C3C"])
        _ppt_chart(s, C.LINE_MARKERS, a["Machine"], {"MTBF (hrs)": a["MTBF (hrs)"]}, "MTBF (hrs)",
                   RL, 1.0, LW, 5.9, colors=["1F4E78"])

        s = _ppt_slide(prs, f"5. Capacity Analysis - Top 10 Worst Machines ({cap_label})")
        _ppt_text(s, "By MTTR (highest)", 0.4, 0.95, LW, 0.35, 14, True, "1F4E78")
        w1 = ms.nlargest(10, "MTTR (hrs)")[["Machine", "Events", "Hours", "MTTR (hrs)"]]
        _ppt_table(s, w1, 0.4, 1.4, LW, 0.4 * (len(w1) + 1), 10)
        _ppt_text(s, "By MTBF (lowest)", RL, 0.95, LW, 0.35, 14, True, "1F4E78")
        w2 = ms.nsmallest(10, "MTBF (hrs)")[["Machine", "Events", "Hours", "MTBF (hrs)"]]
        _ppt_table(s, w2, RL, 1.4, LW, 0.4 * (len(w2) + 1), 10)

    trend = []
    for p in period_order:
        pc = label_to_code[p]
        av_h = cal_hours.get(pc, 20 * calendar.monthrange(int(pc[:4]), int(pc[5:7]))[1])
        pcap = n_m * av_h
        ph = df[df["Period_Label"] == p]["Hours"].sum()
        trend.append({"Period": p, "Available hrs/machine": round(av_h, 1), "Plant Capacity": round(pcap, 1),
                      "Breakdown Hours": round(ph, 2), "Availability %": round((pcap - ph) / pcap * 100, 2) if pcap else None})
    trend = pd.DataFrame(trend)
    s = _ppt_slide(prs, "5. Capacity Analysis - Plant Availability Trend (all periods)")
    _ppt_chart(s, C.LINE_MARKERS, trend["Period"], {"Availability %": trend["Availability %"]}, "Plant Availability %",
               0.4, 1.0, 12.5, 5.9, colors=["2ECC71"], labels=True, numfmt="0.0")
    _ppt_table_slides(prs, "5. Capacity Analysis - Availability by Period", trend, max_cols=6)

    # ================================================ TAB 6 - MACHINE-WISE
    mw = _agg(df, "Machine").sort_values("Hours", ascending=False).reset_index().round(2)
    _ppt_table_slides(prs, "6. Machine-wise - Hours, Loss & Events", mw, rows_per=14, max_cols=5)
    top20 = mw.head(20)
    s = _ppt_slide(prs, "6. Machine-wise - Hours vs Production Loss (Top 20 machines)")
    _ppt_chart(s, C.COLUMN_CLUSTERED, top20["Machine"], {"Hours": top20["Hours"]}, "Breakdown Hours", 0.4, 1.0, LW, 5.9, colors=["1F4E78"])
    _ppt_chart(s, C.COLUMN_CLUSTERED, top20["Machine"], {"Loss": top20["Loss"]}, "Production Loss", RL, 1.0, LW, 5.9, colors=["E74C3C"])
    names20 = top20["Machine"].tolist()
    pv = df[df["Machine"].isin(names20)].pivot_table(index="Machine", columns="Type", values="Hours", aggfunc="sum", fill_value=0).reindex(names20)
    s = _ppt_slide(prs, "6. Machine-wise - Breakdown Type per Machine (Top 20)")
    _ppt_pivot_chart(s, C.COLUMN_STACKED, pv, "Hours by Type per Machine", 0.4, 1.0, 12.5, 5.9)
    pv = df[df["Machine"].isin(names20)].pivot_table(index="Machine", columns="Shift", values="Hours", aggfunc="sum", fill_value=0).reindex(names20)
    s = _ppt_slide(prs, "6. Machine-wise - Shift-wise Breakdown Hours (Top 20)")
    _ppt_pivot_chart(s, C.COLUMN_STACKED, pv, "Hours by Shift per Machine", 0.4, 1.0, 12.5, 5.9)

    # ================================================ TAB 7 - CAUSES
    cz = _agg(df, "Cause").sort_values("Hours", ascending=False).reset_index().round(2)
    _ppt_table_slides(prs, "7. Causes - Top Causes", cz, rows_per=14, max_cols=5)
    t10 = cz.head(10)
    s = _ppt_slide(prs, "7. Causes - Top 10 Causes by Hours")
    _ppt_chart(s, C.BAR_CLUSTERED, t10["Cause"], {"Hours": t10["Hours"]}, "Top 10 Causes (Hours)", 0.4, 1.0, 12.5, 5.9,
               point_colors=_grad(t10["Hours"], True), reverse=True, labels=True)
    top15c = cz.head(15)["Cause"].tolist()
    cm = pivot(df[df["Cause"].isin(top15c)], "Cause", "Machine")
    keep = cm.drop(columns="Total").sum().nlargest(8).index.tolist()
    cm = cm[keep + ["Total"]].reset_index()
    _ppt_table_slides(prs, "7. Causes - Cause x Machine (Top 15 causes, top 8 machines, Hours)", cm, rows_per=15, max_cols=10, fs=9)
    t10c = cz.head(10)["Cause"].tolist()
    pv = (df[df["Cause"].isin(t10c)].pivot_table(index="Cause", columns="Period_Label", values="Hours", aggfunc="sum", fill_value=0)
          .reindex(t10c))
    pv = pv[[c for c in period_order[-6:] if c in pv.columns]]
    s = _ppt_slide(prs, "7. Causes - Top 10 Causes across Months (last 6 months)")
    _ppt_pivot_chart(s, C.COLUMN_CLUSTERED, pv, "Hours by Cause per Month", 0.4, 1.0, 12.5, 5.9)

    # ================================================ TAB 8 - YTD
    fy_order = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]
    for fy_ in sorted(df["FY_Label"].unique()):
        dfy = df[df["FY_Label"] == fy_]
        mf = _agg(dfy, "Month").reindex(fy_order).dropna(how="all").fillna(0).reset_index()
        mf["Month"] = mf["Month"].apply(lambda m: calendar.month_abbr[int(m)])
        mf["Cumulative Hours"] = mf["Hours"].cumsum()
        mf["Cost (Rs)"] = mf["Loss"] * cost_per_unit
        s = _ppt_slide(prs, f"8. YTD (Fiscal Year) - FY {fy_}")
        _ppt_kpis(s, [("YTD Hours", f"{dfy['Hours'].sum():,.1f}"), ("YTD Loss", f"{dfy['Loss'].sum():,.1f}"),
                      ("YTD Cost (Rs)", f"{dfy['Loss'].sum() * cost_per_unit:,.0f}"), ("YTD Events", f"{len(dfy):,}")], h=1.1)
        _ppt_chart(s, C.COLUMN_CLUSTERED, mf["Month"], {"Monthly Hours": mf["Hours"]}, "Monthly Hours (Apr - Mar)",
                   0.4, 2.45, LW, 4.5, colors=["1F4E78"], labels=True)
        _ppt_chart(s, C.LINE_MARKERS, mf["Month"], {"Cumulative Hours": mf["Cumulative Hours"]}, "Cumulative Hours",
                   RL, 2.45, LW, 4.5, colors=["E74C3C"])
        _ppt_table_slides(prs, f"8. YTD (Fiscal Year) - FY {fy_} Monthly Detail", mf.round(2), max_cols=7)

    out = io.BytesIO()
    prs.save(out)
    return out.getvalue()


# =====================================================================
# 2. SIDEBAR
# =====================================================================
st.sidebar.header("📁 Upload Files")
bd_files = st.sidebar.file_uploader("Breakdown Excel files", type=["xlsx", "xls"],
                                    accept_multiple_files=True, key="bd_files")
master_file = st.sidebar.file_uploader("Master Machine List (machines.xlsx)", type=["xlsx", "xls"],
                                       key="master_file")
cal_files = st.sidebar.file_uploader("Production Calendar Excel files", type=["xlsx", "xls"],
                                     accept_multiple_files=True, key="cal_files")

if st.sidebar.button("🗑️ Clear All Uploaded Files"):
    st.session_state.clear()
    st.rerun()

st.sidebar.header("💰 Cost Settings")
cost_per_unit = st.sidebar.number_input("Cost per unit (₹)", min_value=0.0, value=50.0, step=1.0)

if not bd_files:
    st.info("👈 Upload one or more Breakdown Excel files from the sidebar to begin. "
            "File names should contain the period, e.g. `Breakdown_2025-06.xlsx` or `Jun-25.xlsx`.")
    st.stop()

# ---------- load data ----------
frames, errors = [], []
for f in bd_files:
    try:
        d = read_breakdown_file(f)
        if d.empty:
            errors.append(f"{f.name}: no usable rows (check 'M/C Name' and 'Total Hrs' columns).")
        else:
            frames.append(d)
    except Exception as e:  # noqa: BLE001
        errors.append(f"{f.name}: {e}")
for e in errors:
    st.sidebar.warning(e)
if not frames:
    st.error("No valid breakdown data could be read.")
    st.stop()

df_all = pd.concat(frames, ignore_index=True)

master_list = []
if master_file is not None:
    try:
        master_list = read_master(master_file)
    except Exception as e:  # noqa: BLE001
        st.sidebar.warning(f"Master list error: {e}")

if master_list:
    mmap = {norm_name(n): n for n in master_list}
    df_all["Machine"] = df_all["Machine"].apply(lambda x: mmap.get(norm_name(x), x))

cal_hours = {}
if cal_files:
    try:
        cal_hours = read_calendars(cal_files)
    except Exception as e:  # noqa: BLE001
        st.sidebar.warning(f"Calendar error: {e}")

# ---------- period ordering ----------
label_to_code = dict(zip(df_all["Period_Label"], df_all["Period"]))
all_periods = sorted(label_to_code, key=lambda l: label_to_code[l])

# ---------- filters ----------
st.sidebar.header("🔎 Filters")
sel_period = st.sidebar.multiselect("Period", all_periods, default=all_periods)
all_machines = sorted(df_all["Machine"].unique())
sel_machine = st.sidebar.multiselect("Machine", all_machines, default=all_machines)
all_types = sorted(df_all["Type"].unique())
sel_type = st.sidebar.multiselect("Breakdown Type", all_types, default=all_types)
all_shifts = sorted(df_all["Shift"].unique())
sel_shift = st.sidebar.multiselect("Shift", all_shifts, default=all_shifts)

df = df_all[
    df_all["Period_Label"].isin(sel_period)
    & df_all["Machine"].isin(sel_machine)
    & df_all["Type"].isin(sel_type)
    & df_all["Shift"].isin(sel_shift)
].copy()

if df.empty:
    st.warning("No data matches the selected filters.")
    st.stop()

period_order = [p for p in all_periods if p in set(df["Period_Label"])]
years = sorted(df["Year"].unique())
MONTH_ABBR = [calendar.month_abbr[i] for i in range(1, 13)]

# =====================================================================
# 4. TABS
# =====================================================================
tabs = st.tabs(["📊 Overview", "📅 Month Comparison", "📈 Year Comparison", "🏆 Yearly Analysis",
                "🏭 Capacity Analysis", "🔧 Machine-wise", "🔁 Causes", "📆 YTD (Fiscal Year)"])

# ---------------------------------------------------------------- Tab 1
with tabs[0]:
    st.subheader("Overview")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Hours", f"{df['Hours'].sum():,.1f}")
    c2.metric("Total Loss", f"{df['Loss'].sum():,.1f}")
    c3.metric("Total Events", f"{len(df):,}")
    c4.metric("Avg Hours/Event", f"{df['Hours'].mean():,.2f}")

    by_type = df.groupby("Type").agg(Hours=("Hours", "sum"), Loss=("Loss", "sum")).reset_index()
    p1, p2 = st.columns(2)
    with p1:
        f = px.pie(by_type, names="Type", values="Hours", title="Breakdown Hours by Type",
                   color_discrete_sequence=px.colors.qualitative.Set2)
        show_chart(f)
    with p2:
        f = px.pie(by_type, names="Type", values="Loss", title="Production Loss by Type",
                   color_discrete_sequence=px.colors.qualitative.Pastel)
        show_chart(f)

    mon = (df.groupby("Period_Label").agg(Hours=("Hours", "sum"), Loss=("Loss", "sum"))
           .reindex(period_order).reset_index())
    fig = go.Figure()
    fig.add_bar(x=mon["Period_Label"], y=mon["Hours"], name="Hours", marker_color="#1F4E78")
    fig.add_trace(go.Scatter(x=mon["Period_Label"], y=mon["Loss"], name="Production Loss",
                             mode="lines+markers", line=dict(color="#E74C3C", width=3), yaxis="y2"))
    fig.update_layout(title="Monthly Breakdown Hours vs Production Loss",
                      yaxis=dict(title="Hours"),
                      yaxis2=dict(title="Production Loss", overlaying="y", side="right"),
                      legend=dict(orientation="h", y=1.12))
    show_chart(fig)

# ---------------------------------------------------------------- Tab 2
with tabs[1]:
    st.subheader("Month Comparison")
    if len(period_order) >= 2:
        cur_l, prev_l = period_order[-1], period_order[-2]
        cur, prev = df[df["Period_Label"] == cur_l], df[df["Period_Label"] == prev_l]
        st.caption(f"Latest month **{cur_l}** vs previous month **{prev_l}**")
        m1, m2, m3 = st.columns(3)
        for col, name, a, b in [
            (m1, "Hours", cur["Hours"].sum(), prev["Hours"].sum()),
            (m2, "Production Loss", cur["Loss"].sum(), prev["Loss"].sum()),
            (m3, "Events", len(cur), len(prev)),
        ]:
            ch = pct_change(a, b)
            col.metric(name, f"{a:,.1f}", f"{ch:+.1f}%" if ch is not None else "n/a", delta_color="inverse")
    else:
        st.info("Select at least two periods to see month-over-month change.")

    g = df.groupby(["Period_Label", "Type"])["Hours"].sum().reset_index()
    f = px.bar(g, x="Period_Label", y="Hours", color="Type", barmode="group",
               category_orders={"Period_Label": period_order}, title="Hours by Breakdown Type per Month")
    show_chart(f)

    top_m = df.groupby("Machine")["Hours"].sum().nlargest(15).index
    g = df[df["Machine"].isin(top_m)].groupby(["Period_Label", "Machine"])["Hours"].sum().reset_index()
    f = px.bar(g, x="Machine", y="Hours", color="Period_Label", barmode="group",
               category_orders={"Period_Label": period_order}, title="Hours by Machine per Month (Top 15 machines)")
    show_chart(f)

    st.markdown("**Machine × Month (Hours)**")
    show_df(pivot(df, "Machine", "Period_Label", period_order))
    st.markdown("**Breakdown Type × Month (Hours)**")
    show_df(pivot(df, "Type", "Period_Label", period_order))

# ---------------------------------------------------------------- Tab 3
with tabs[2]:
    st.subheader("Year Comparison")
    g = df.groupby(["Year", "Type"])["Hours"].sum().reset_index()
    g["Year"] = g["Year"].astype(str)
    show_chart(px.bar(g, x="Year", y="Hours", color="Type", barmode="group", title="Hours by Year & Type"))

    g = df.groupby(["Month", "Year"])["Hours"].sum().reset_index()
    g["Month Name"] = g["Month"].apply(lambda m: calendar.month_abbr[m])
    g["Year"] = g["Year"].astype(str)
    show_chart(px.bar(g, x="Month Name", y="Hours", color="Year", barmode="group",
                      category_orders={"Month Name": MONTH_ABBR}, title="Hours by Month & Year"))

    st.markdown("**Breakdown Type × Year (Hours)**")
    show_df(pivot(df, "Type", "Year"))
    st.markdown("**Machine × Year (Hours)**")
    show_df(pivot(df, "Machine", "Year"))

# ---------------------------------------------------------------- Tab 4
with tabs[3]:
    st.subheader("Yearly Analysis")
    yr = st.selectbox("Select Year", years, index=len(years) - 1, key="yr_sel")
    dy = df[df["Year"] == yr]
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Total Hours", f"{dy['Hours'].sum():,.1f}")
    k2.metric("Total Loss", f"{dy['Loss'].sum():,.1f}")
    k3.metric("Total Events", f"{len(dy):,}")
    k4.metric("Avg Hours/Event", f"{dy['Hours'].mean():,.2f}")

    mt = (dy.groupby(["Month", "Period_Label"])
          .agg(Hours=("Hours", "sum"), Loss=("Loss", "sum"), Events=("Hours", "count"))
          .reset_index().sort_values("Month"))
    mt["Cumulative Hours"] = mt["Hours"].cumsum()
    show_df(mt.drop(columns="Month").rename(columns={"Period_Label": "Month"}).round(2), hide_index=True)

    fig = go.Figure()
    fig.add_bar(x=mt["Period_Label"], y=mt["Hours"], name="Monthly Hours", marker_color="#1F4E78")
    fig.add_trace(go.Scatter(x=mt["Period_Label"], y=mt["Cumulative Hours"], name="Cumulative Hours",
                             mode="lines+markers", line=dict(color="#E74C3C", width=3), yaxis="y2"))
    fig.update_layout(title=f"Monthly & Cumulative Hours - {yr}", yaxis=dict(title="Hours"),
                      yaxis2=dict(title="Cumulative Hours", overlaying="y", side="right"),
                      legend=dict(orientation="h", y=1.12))
    show_chart(fig)

# ---------------------------------------------------------------- Tab 5
with tabs[4]:
    st.subheader("Capacity Analysis")
    cap_label = st.selectbox("Select Period", period_order, index=len(period_order) - 1, key="cap_sel")
    code = label_to_code[cap_label]
    yy, mm = int(code[:4]), int(code[5:7])
    days = calendar.monthrange(yy, mm)[1]
    if code in cal_hours:
        avail_hrs = cal_hours[code]
        st.success(f"Using production calendar: {avail_hrs:,.1f} available hours per machine for {cap_label}.")
    else:
        avail_hrs = 20 * days
        st.warning(f"No calendar found for {cap_label}. Using default 20 hrs × {days} days = {avail_hrs:,.0f} hrs per machine.")

    dfp = df[df["Period_Label"] == cap_label]
    if master_list:
        universe = master_list
    else:
        universe = sorted(df_all["Machine"].unique())
        st.warning("No master machine list uploaded - using machines found in breakdown data.")
    outside = set(dfp["Machine"]) - set(universe)
    if outside:
        st.caption(f"{len(outside)} machine(s) in breakdown data are not in the master list and are excluded from machine counts.")

    # ---- per-machine stats ----
    rows = []
    for mc in universe:
        dm = dfp[dfp["Machine"] == mc]
        hrs, ev, loss = dm["Hours"].sum(), len(dm), dm["Loss"].sum()
        uptime = max(avail_hrs - hrs, 0)
        rows.append({
            "Machine": mc, "Hours": round(hrs, 2), "Events": ev, "Loss": round(loss, 2),
            "Availability %": round(uptime / avail_hrs * 100, 2) if avail_hrs else None,
            "MTBF (hrs)": round(uptime / ev, 2) if ev else None,
            "MTTR (hrs)": round(hrs / ev, 2) if ev else None,
        })
    machine_stats_df = pd.DataFrame(rows)

    total_m = len(universe)
    with_bd = int((machine_stats_df["Events"] > 0).sum())
    without_bd = total_m - with_bd
    plant_cap = total_m * avail_hrs
    bd_hours = machine_stats_df["Hours"].sum()
    avail_pct = (plant_cap - bd_hours) / plant_cap * 100 if plant_cap else 0
    down_pct = 100 - avail_pct
    tot_events = int(machine_stats_df["Events"].sum())
    mtbf_all = (plant_cap - bd_hours) / tot_events if tot_events else None
    mttr_all = bd_hours / tot_events if tot_events else None

    a1, a2, a3, a4 = st.columns(4)
    a1.metric("Total Machines", total_m)
    a2.metric("With Breakdown", with_bd)
    a3.metric("Without Breakdown", without_bd)
    a4.metric("Plant Capacity (hrs)", f"{plant_cap:,.0f}")
    b1, b2, b3, b4 = st.columns(4)
    b1.metric("Breakdown Hours", f"{bd_hours:,.1f}")
    b2.metric("Availability %", f"{avail_pct:.2f}%")
    b3.metric("Downtime %", f"{down_pct:.2f}%")
    b4.metric("Events", tot_events)
    c1, c2 = st.columns(2)
    c1.metric("MTBF - Overall (hrs)", f"{mtbf_all:,.2f}" if mtbf_all else "n/a")
    c2.metric("MTTR - Overall (hrs)", f"{mttr_all:,.2f}" if mttr_all else "n/a")

    p1, p2 = st.columns(2)
    with p1:
        f = px.pie(names=["With Breakdown", "Without Breakdown"], values=[with_bd, without_bd],
                   title="Machines With / Without Breakdown", color=["With Breakdown", "Without Breakdown"],
                   color_discrete_map={"With Breakdown": "#E74C3C", "Without Breakdown": "#2ECC71"})
        show_chart(f)
    with p2:
        dt = dfp.groupby("Type")["Hours"].sum().reset_index()
        if dt.empty:
            st.info("No downtime recorded in this period.")
        else:
            show_chart(px.pie(dt, names="Type", values="Hours", title="Downtime by Type",
                              color_discrete_sequence=px.colors.qualitative.Set2))

    st.markdown("**Machine-wise Statistics**")
    show_df(machine_stats_df, hide_index=True)

    # ---- split by type ----
    trows = []
    for mc in universe:
        for tp in sorted(df_all["Type"].unique()):
            dm = dfp[(dfp["Machine"] == mc) & (dfp["Type"] == tp)]
            ev = len(dm)
            if ev == 0:
                continue
            hrs = dm["Hours"].sum()
            trows.append({"Machine": mc, "Type": tp, "Events": ev, "Hours": round(hrs, 2),
                          "MTTR (hrs)": round(hrs / ev, 2),
                          "MTBF (hrs)": round(max(avail_hrs - hrs, 0) / ev, 2)})
    type_stats_df = pd.DataFrame(trows)

    st.markdown("### MTTR & MTBF by Breakdown Type")
    if type_stats_df.empty:
        st.info("No breakdowns for the selected period.")
    else:
        show_chart(px.bar(type_stats_df, x="Machine", y="MTTR (hrs)", color="Type", barmode="group",
                          title="MTTR by Machine & Type"))
        show_chart(px.bar(type_stats_df, x="Machine", y="MTBF (hrs)", color="Type", barmode="group",
                          title="MTBF by Machine & Type"))

    # ---- machine-wise MTTR / MTBF ----
    st.markdown("### Machine-wise MTTR & MTBF")
    ms = machine_stats_df.dropna(subset=["MTTR (hrs)", "MTBF (hrs)"])
    if ms.empty:
        st.info("No machines with breakdowns in this period.")
    else:
        h = max(400, 22 * len(ms))
        mttr_s = ms.sort_values("MTTR (hrs)", ascending=False)
        f = px.bar(mttr_s, x="MTTR (hrs)", y="Machine", orientation="h", color="MTTR (hrs)",
                   color_continuous_scale="RdYlGn_r", title="MTTR by Machine (highest to lowest)", height=h)
        f.update_yaxes(autorange="reversed")
        show_chart(f)

        mtbf_s = ms.sort_values("MTBF (hrs)", ascending=True)
        f = px.bar(mtbf_s, x="MTBF (hrs)", y="Machine", orientation="h", color="MTBF (hrs)",
                   color_continuous_scale="RdYlGn", title="MTBF by Machine (lowest to highest)", height=h)
        f.update_yaxes(autorange="reversed")
        show_chart(f)

        comb = ms.sort_values("MTTR (hrs)", ascending=False)
        fig = go.Figure()
        fig.add_bar(x=comb["Machine"], y=comb["MTTR (hrs)"], name="MTTR (hrs)", marker_color="#E74C3C")
        fig.add_trace(go.Scatter(x=comb["Machine"], y=comb["MTBF (hrs)"], name="MTBF (hrs)",
                                 mode="lines+markers", line=dict(color="#1F4E78", width=3), yaxis="y2"))
        fig.update_layout(title="MTTR vs MTBF by Machine", yaxis=dict(title="MTTR (hrs)"),
                          yaxis2=dict(title="MTBF (hrs)", overlaying="y", side="right"),
                          legend=dict(orientation="h", y=1.12))
        show_chart(fig)

        t1, t2 = st.columns(2)
        with t1:
            st.markdown("**Top 10 Worst Machines by MTTR**")
            show_df(ms.nlargest(10, "MTTR (hrs)")[["Machine", "Events", "Hours", "MTTR (hrs)"]], hide_index=True)
        with t2:
            st.markdown("**Top 10 Worst Machines by MTBF**")
            show_df(ms.nsmallest(10, "MTBF (hrs)")[["Machine", "Events", "Hours", "MTBF (hrs)"]], hide_index=True)

# ---------------------------------------------------------------- Tab 6
with tabs[5]:
    st.subheader("Machine-wise Analysis")
    mw = (df.groupby("Machine").agg(Hours=("Hours", "sum"), Loss=("Loss", "sum"), Events=("Hours", "count"))
          .reset_index().sort_values("Hours", ascending=False).round(2))
    show_df(mw, hide_index=True)

    f = px.bar(mw.melt(id_vars="Machine", value_vars=["Hours", "Loss"], var_name="Metric", value_name="Value"),
               x="Machine", y="Value", color="Metric", barmode="group", title="Hours vs Production Loss by Machine",
               color_discrete_map={"Hours": "#1F4E78", "Loss": "#E74C3C"})
    show_chart(f)

    g = df.groupby(["Machine", "Type"])["Hours"].sum().reset_index()
    show_chart(px.bar(g, x="Machine", y="Hours", color="Type", barmode="stack",
                      category_orders={"Machine": mw["Machine"].tolist()}, title="Breakdown Hours by Type per Machine"))

    g = df.groupby(["Machine", "Shift"])["Hours"].sum().reset_index()
    show_chart(px.bar(g, x="Machine", y="Hours", color="Shift", barmode="stack",
                      category_orders={"Machine": mw["Machine"].tolist()}, title="Shift-wise Breakdown Hours per Machine"))

# ---------------------------------------------------------------- Tab 7
with tabs[6]:
    st.subheader("Cause Analysis")
    cz = (df.groupby("Cause").agg(Hours=("Hours", "sum"), Loss=("Loss", "sum"), Events=("Hours", "count"))
          .reset_index().sort_values("Hours", ascending=False).round(2))
    st.markdown("**Top Causes**")
    show_df(cz, hide_index=True)

    top10 = cz.head(10)
    f = px.bar(top10, x="Hours", y="Cause", orientation="h", color="Hours",
               color_continuous_scale="Reds", title="Top 10 Causes by Hours")
    f.update_yaxes(autorange="reversed")
    show_chart(f)

    top15 = cz.head(15)["Cause"].tolist()
    st.markdown("**Cause × Machine Matrix (Top 15 causes, Hours)**")
    show_df(pivot(df[df["Cause"].isin(top15)], "Cause", "Machine"))

    top10c = cz.head(10)["Cause"].tolist()
    g = df[df["Cause"].isin(top10c)].groupby(["Period_Label", "Cause"])["Hours"].sum().reset_index()
    show_chart(px.bar(g, x="Cause", y="Hours", color="Period_Label", barmode="group",
                      category_orders={"Period_Label": period_order, "Cause": top10c},
                      title="Top 10 Causes - Comparison Across Months"))

# ---------------------------------------------------------------- Tab 8
with tabs[7]:
    st.subheader("YTD (Fiscal Year: April - March)")
    fys = sorted(df["FY_Label"].unique())
    fy = st.selectbox("Select Fiscal Year", fys, index=len(fys) - 1, key="fy_sel")
    dfy = df[df["FY_Label"] == fy]
    ytd_hours, ytd_loss = dfy["Hours"].sum(), dfy["Loss"].sum()
    y1, y2, y3, y4 = st.columns(4)
    y1.metric("YTD Hours", f"{ytd_hours:,.1f}")
    y2.metric("YTD Loss", f"{ytd_loss:,.1f}")
    y3.metric("YTD Cost (₹)", f"{ytd_loss * cost_per_unit:,.0f}")
    y4.metric("YTD Events", f"{len(dfy):,}")

    fy_order = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]
    mon = dfy.groupby("Month")["Hours"].sum().reindex(fy_order).dropna().reset_index()
    mon["Month Name"] = mon["Month"].apply(lambda m: calendar.month_abbr[m])
    mon["Cumulative Hours"] = mon["Hours"].cumsum()
    fig = go.Figure()
    fig.add_bar(x=mon["Month Name"], y=mon["Hours"], name="Monthly Hours", marker_color="#1F4E78")
    fig.add_trace(go.Scatter(x=mon["Month Name"], y=mon["Cumulative Hours"], name="Cumulative Hours",
                             mode="lines+markers", line=dict(color="#E74C3C", width=3), yaxis="y2"))
    fig.update_layout(title=f"FY {fy} - Monthly & Cumulative Hours",
                      xaxis=dict(categoryorder="array", categoryarray=[calendar.month_abbr[m] for m in fy_order]),
                      yaxis=dict(title="Hours"),
                      yaxis2=dict(title="Cumulative Hours", overlaying="y", side="right"),
                      legend=dict(orientation="h", y=1.12))
    show_chart(fig)

# =====================================================================
# 5. EXPORTS
# =====================================================================
st.markdown("---")
st.subheader("📥 Export Report")
try:
    st.download_button(
        label="Download Excel Report",
        data=build_excel_report(df),
        file_name="Breakdown_Analysis_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
except Exception as e:  # noqa: BLE001
    st.error(f"Could not build Excel report: {e}")

st.markdown("**PowerPoint report** (editable charts & tables - opens in Microsoft PowerPoint and WPS Presentation)")
if not PPTX_OK:
    st.info("To enable PowerPoint export run: `pip install python-pptx` and restart the app.")
else:
    st.caption("Covers all 8 tabs for the current filters. Capacity slides use the period selected in the "
               "Capacity tab; regenerate after changing filters.")
    if st.button("Generate PowerPoint Report"):
        with st.spinner("Building presentation..."):
            try:
                st.session_state["ppt_bytes"] = build_ppt_report(
                    df, cost_per_unit, cal_hours, master_list, period_order, label_to_code, cap_label)
            except Exception as e:  # noqa: BLE001
                st.session_state.pop("ppt_bytes", None)
                st.error(f"Could not build PowerPoint: {e}")
    if "ppt_bytes" in st.session_state:
        st.download_button(
            label="Download PowerPoint (.pptx)",
            data=st.session_state["ppt_bytes"],
            file_name="Breakdown_Analysis_Report.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
