import streamlit as st
import pandas as pd
import io
import zipfile
import re
import datetime

st.set_page_config(
    page_title="Videmi – Booking to CSV",
    page_icon="🏡",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── CSS ────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* ── global ── */
    [data-testid="stAppViewContainer"] { background:#f8faf8; }
    [data-testid="stHeader"] { background:transparent; }

    /* ── header ── */
    .videmi-header {
        display:flex; align-items:center; gap:12px;
        padding:1.2rem 0 0.2rem;
    }
    .videmi-logo {
        background:#2c5f2e; color:white; border-radius:10px;
        width:46px; height:46px; display:flex; align-items:center;
        justify-content:center; font-size:1.5rem; flex-shrink:0;
    }
    .videmi-title { font-size:1.9rem; font-weight:800; color:#2c5f2e; margin:0; line-height:1.1; }
    .videmi-sub   { font-size:0.88rem; color:#888; margin:0; }

    /* ── upload zone ── */
    .upload-hint {
        text-align:center; padding:2.5rem 1rem;
        border:2px dashed #b5d4b5; border-radius:12px;
        background:#f0f7f0; color:#555; margin:1rem 0;
    }
    .upload-hint h3 { color:#2c5f2e; margin-bottom:0.4rem; }

    /* ── client card ── */
    .client-card {
        background:white; border:1px solid #d4e8d4;
        border-top:4px solid #2c5f2e;
        border-radius:10px; padding:1.2rem 1.4rem 0.8rem;
        margin-bottom:1rem;
        box-shadow:0 2px 6px rgba(44,95,46,0.08);
    }
    .client-name { font-size:1.25rem; font-weight:700; color:#2c5f2e; margin:0 0 2px; }
    .client-file { font-size:0.78rem; color:#999; margin:0; }

    /* ── stat boxes ── */
    .stat-row { display:flex; gap:10px; margin:1rem 0; flex-wrap:wrap; }
    .stat-box {
        flex:1; min-width:90px; background:white;
        border:1px solid #e8f0e8; border-radius:8px;
        padding:0.6rem 0.8rem; text-align:center;
    }
    .stat-num { font-size:1.5rem; font-weight:700; color:#2c5f2e; line-height:1.1; }
    .stat-lbl { font-size:0.7rem; color:#999; text-transform:uppercase; letter-spacing:0.03em; }

    /* ── property cards ── */
    .prop-grid { display:flex; gap:10px; flex-wrap:wrap; margin:0.5rem 0; }
    .prop-card {
        flex:1; min-width:200px; background:#f6fbf6;
        border:1px solid #c8e4c8; border-radius:8px;
        padding:0.7rem 1rem; font-size:0.87rem;
    }
    .prop-name { font-weight:700; color:#1a3d1a; margin-bottom:4px; }
    .prop-detail { color:#555; line-height:1.6; }

    /* ── badges ── */
    .badge {
        display:inline-block; padding:3px 10px; border-radius:20px;
        font-size:0.78rem; font-weight:600; margin:2px 3px;
    }
    .badge-green  { background:#d4edda; color:#155724; }
    .badge-yellow { background:#fff3cd; color:#856404; }
    .badge-red    { background:#f8d7da; color:#721c24; }
    .badge-gray   { background:#e9ecef; color:#495057; }
    .badge-blue   { background:#cce5ff; color:#004085; }

    /* ── section headers ── */
    .section-title {
        font-size:0.88rem; font-weight:700; color:#2c5f2e;
        text-transform:uppercase; letter-spacing:0.06em;
        border-bottom:2px solid #d4e8d4; padding-bottom:5px;
        margin:1rem 0 0.6rem;
    }

    /* ── settings row ── */
    .settings-row { display:flex; gap:20px; flex-wrap:wrap; margin:0.8rem 0; }
    .setting-item { display:flex; align-items:center; gap:6px; font-size:0.88rem; }
    .setting-label { color:#666; }

    /* ── preview table header ── */
    .preview-bar {
        background:#2c5f2e; color:white; border-radius:8px 8px 0 0;
        padding:0.5rem 1rem; font-weight:600; font-size:0.9rem;
        display:flex; align-items:center; gap:8px;
    }

    /* ── export buttons ── */
    .export-section {
        background:#f0f7f0; border-radius:0 0 8px 8px;
        padding:1rem; border:1px solid #d4e8d4; border-top:none;
    }
    .export-label { font-size:0.8rem; font-weight:600; color:#555;
                    text-transform:uppercase; letter-spacing:0.04em; margin-bottom:6px; }

    /* ── master export ── */
    .master-export {
        background:white; border:2px solid #2c5f2e; border-radius:10px;
        padding:1.2rem 1.5rem; margin:1rem 0;
        text-align:center;
    }

    /* ── info boxes ── */
    .info-box {
        background:#e8f4fd; border-left:4px solid #1a73e8;
        border-radius:0 6px 6px 0; padding:0.7rem 1rem;
        font-size:0.88rem; color:#1a3d5c; margin:0.5rem 0;
    }

    /* ── divider ── */
    .client-divider { border:none; border-top:2px dashed #d4e8d4; margin:1.5rem 0; }
</style>
""", unsafe_allow_html=True)

# ── constants ──────────────────────────────────────────────────────────────────
OUT_COLS = [
    "Client Name:", "Property list", "Reservation ID",
    "DATE:", "VILLA:", "TYPE CLEAN:", "PAX:", "START TIME:", "END TIME:",
    "STATUS:", "LAUNDRY :", "Key:", "Code:", "Ameneties:",
    "COMMENTS:", "QB SHIFT ID", "LAST SYNC",
]

STATUS_BADGE = {
    "SCHEDULED": "badge-green",
    "CANCELED":  "badge-red",
    "CANCELLED": "badge-red",
    "NONE":      "badge-gray",
    "RESERVED":  "badge-yellow",
    "UPDATE":    "badge-blue",
}


# ── helpers ────────────────────────────────────────────────────────────────────
def cs(v):
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    s = str(v).strip()
    return "" if s.lower() in ("nan", "nat", "none", "nat") else s


def scan_label(raw, keyword):
    kw = keyword.lower()
    for _, row in raw.iterrows():
        if len(row) < 2:
            continue
        if kw in cs(row.iloc[0]).lower():
            val = cs(row.iloc[1])
            if val:
                return val
    return ""


def find_info_sheet(xl):
    for name in xl.sheet_names:
        if any(k in name.lower() for k in ("client", "profile", "info", "general")):
            return name
    return None


def find_header_row(raw):
    for i, row in raw.iterrows():
        vals = [cs(v).upper().rstrip(":") for v in row if cs(v)]
        if "DATE" in vals:
            return i
    return None


def is_month_sheet(name):
    months = ["jan", "feb", "mar", "mrt", "apr", "apl", "may", "mei",
              "jun", "jul", "aug", "sep", "oct", "okt", "nov", "dec"]
    return any(name.strip().lower().startswith(m) for m in months)


def fmt_date(val):
    s = cs(val)
    if not s:
        return ""
    try:
        dt = pd.to_datetime(val)
        return f"{dt.day}/{dt.month}/{dt.year}"
    except Exception:
        return s


def fmt_time(val):
    s = cs(val)
    if not s:
        return ""
    try:
        if isinstance(val, datetime.time):
            return val.strftime("%H:%M")
        return pd.to_datetime(val).strftime("%H:%M")
    except Exception:
        return s


def get_col(row, idx):
    if idx is None or idx >= len(row):
        return ""
    return cs(row.iloc[idx])


def badge_html(label, cls="badge-gray"):
    return f'<span class="badge {cls}">{label}</span>'


def yn_badge(val):
    v = str(val).strip().lower()
    cls = "badge-green" if v in ("yes", "y", "true", "1") else \
          "badge-red"   if v in ("no", "n", "false", "0") else "badge-gray"
    return badge_html(val or "—", cls)


def safe_name(s, fallback="client"):
    result = re.sub(r"[^\w\-]", "_", s or fallback).strip("_")
    return result or fallback


# ── rich client info extraction ────────────────────────────────────────────────
def read_client_info(xl):
    info = {
        "client_name":   "",
        "properties":    [],
        "checkout_time": "",
        "checkin_time":  "",
        "keys":          "",
        "codes":         "",
        "amenities":     "",
        "laundry":       "",
        "clean_types":   [],
        "amenity_items": [],
        "linen_items":   [],
    }

    sheet = find_info_sheet(xl)
    if not sheet:
        return info

    raw = xl.parse(sheet, header=None)

    # client name
    for _, row in raw.iterrows():
        for v in row:
            s = cs(v)
            if s and "@" not in s:
                name = re.sub(r"(client\s*(profile|info).*)", "", s, flags=re.IGNORECASE).strip()
                if name:
                    info["client_name"] = name
                    break
        if info["client_name"]:
            break

    # scalar settings
    info["checkout_time"] = fmt_time(scan_label(raw, "check-out"))
    info["checkin_time"]  = fmt_time(scan_label(raw, "check-in"))
    info["keys"]          = scan_label(raw, "key")
    info["codes"]         = scan_label(raw, "code")
    info["amenities"]     = scan_label(raw, "amenities")
    info["laundry"]       = scan_label(raw, "laundry")

    # clean types — in columns 4-6
    in_clean = False
    for _, row in raw.iterrows():
        label = cs(row.iloc[4]).lower() if len(row) > 4 else ""
        if not in_clean:
            if "type of clean" in label:
                in_clean = True
            continue
        name  = cs(row.iloc[4]) if len(row) > 4 else ""
        code  = cs(row.iloc[5]) if len(row) > 5 else ""
        desc  = cs(row.iloc[6]) if len(row) > 6 else ""
        if not name or not code:
            break
        info["clean_types"].append({"code": code, "name": name, "description": desc})

    # properties
    in_props = False
    for _, row in raw.iterrows():
        label = cs(row.iloc[0]).lower()
        if not in_props:
            if any(k in label for k in ("villas", "appartment", "apartment", "properties")):
                if any(k in label for k in ("name", "list")) or label.endswith(":"):
                    in_props = True
            continue
        p = cs(row.iloc[0])
        if not p:
            break
        if any(k in p.lower() for k in ("name", "address", "item", "service")):
            continue
        info["properties"].append({
            "name":    re.sub(r"\bApp\b", "Apartment", p),
            "address": cs(row.iloc[1]) if len(row) > 1 else "",
            "hours":   cs(row.iloc[2]) if len(row) > 2 else "",
            "so_hrs":  cs(row.iloc[3]) if len(row) > 3 else "",
        })

    # amenity items
    in_amen = False
    for _, row in raw.iterrows():
        label = cs(row.iloc[0]).lower()
        if not in_amen:
            if "list of amenities" in label or label == "item":
                in_amen = True
            continue
        item = cs(row.iloc[0])
        qty  = cs(row.iloc[1]) if len(row) > 1 else ""
        if not item:
            break
        if any(k in item.lower() for k in ("linen", "service", "item", "laundry")):
            break
        info["amenity_items"].append({"item": item, "qty": qty if qty else "0"})

    # linen items
    in_linen = False
    for _, row in raw.iterrows():
        label = cs(row.iloc[0]).lower()
        if not in_linen:
            if "linen" in label or "service/item" in label:
                in_linen = True
            continue
        item = cs(row.iloc[0])
        per  = cs(row.iloc[1]) if len(row) > 1 else ""
        if not item:
            break
        info["linen_items"].append({"item": item, "per": per})

    return info


# ── sheet → dataframe ──────────────────────────────────────────────────────────
def parse_sheet(xl, sheet_name, client):
    raw = xl.parse(sheet_name, header=None)
    hi  = find_header_row(raw)
    if hi is None:
        return pd.DataFrame(columns=OUT_COLS)

    hv = [cs(v).upper().rstrip(":") for v in raw.iloc[hi]]

    def ci(kw):
        for i, c in enumerate(hv):
            if kw in c:
                return i
        return None

    idx = {
        "date":   ci("DATE"),    "villa":  ci("VILLA"),
        "type":   ci("TYPE CLEAN"), "pax": ci("PAX"),
        "start":  ci("START TIME"), "end":  ci("END TIME"),
        "status": ci("RESERVATION STATUS"),
        "laund":  ci("LAUNDRY"),    "comm": ci("COMMENTS"),
    }

    rows = []
    for _, row in raw.iloc[hi + 1:].reset_index(drop=True).iterrows():
        status = get_col(row, idx["status"]).upper() or "NONE"
        villa  = re.sub(r"\bApp\b", "Apartment", get_col(row, idx["villa"]))
        rows.append({
            "Client Name:":   "",
            "Property list":  "",
            "Reservation ID": "",
            "DATE:":          fmt_date(get_col(row, idx["date"])),
            "VILLA:":         villa,
            "TYPE CLEAN:":    get_col(row, idx["type"]),
            "PAX:":           get_col(row, idx["pax"]),
            "START TIME:":    get_col(row, idx["start"]),
            "END TIME:":      get_col(row, idx["end"]),
            "STATUS:":        status,
            "LAUNDRY :":      get_col(row, idx["laund"]),
            "Key:":           client["keys"],
            "Code:":          client["codes"],
            "Ameneties:":     client["amenities"],
            "COMMENTS:":      get_col(row, idx["comm"]),
            "QB SHIFT ID":    "",
            "LAST SYNC":      "",
        })

    if not rows:
        return pd.DataFrame(columns=OUT_COLS)

    df = pd.DataFrame(rows, columns=OUT_COLS)
    df.at[0, "Client Name:"] = client["client_name"]
    for i, prop in enumerate(client["properties"]):
        if i < len(df):
            df.at[i, "Property list"] = prop["name"]
    return df


def process_file(f):
    xl     = pd.ExcelFile(f)
    client = read_client_info(xl)
    months = [s for s in xl.sheet_names if is_month_sheet(s)]
    frames = {s: parse_sheet(xl, s, client) for s in months}
    return client, months, frames


def build_client_zip(cname, frames):
    safe = safe_name(cname)
    buf  = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for sheet, df in frames.items():
            zf.writestr(f"Website_Reservation_{safe}_{sheet}.csv", df.to_csv(index=False))
    buf.seek(0)
    return buf


def build_master_zip(all_results):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname, (client, _, frames) in all_results.items():
            safe = safe_name(client["client_name"] or fname.replace(".xlsx", ""))
            for sheet, df in frames.items():
                zf.writestr(
                    f"{safe}/Website_Reservation_{safe}_{sheet}.csv",
                    df.to_csv(index=False),
                )
    buf.seek(0)
    return buf


# ═══════════════════════════════════════════════════════════════════════════════
# UI
# ═══════════════════════════════════════════════════════════════════════════════

# Header
st.markdown("""
<div class="videmi-header">
  <div class="videmi-logo">🏡</div>
  <div>
    <p class="videmi-title">Videmi – Booking to CSV</p>
    <p class="videmi-sub">Convert client reservation spreadsheets into website-ready import CSVs</p>
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown("---")

# ── file uploader ──────────────────────────────────────────────────────────────
uploaded_files = st.file_uploader(
    "Upload client .xlsx files",
    type=["xlsx"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)

if not uploaded_files:
    st.markdown("""
    <div class="upload-hint">
        <h3>📂 Drop one or more client <code>.xlsx</code> files here</h3>
        <p>Each file is processed independently — client name, properties, settings and bookings<br>
        are all auto-detected from the spreadsheet structure.</p>
    </div>
    """, unsafe_allow_html=True)

    with st.expander("ℹ️ How it works / Expected file format"):
        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown("""
**Auto-detected from each file**
- Client name (from profile/info sheet)
- Properties — name, address, clean hours
- Check-in & Check-out times
- Keys, Codes, Amenities, Laundry settings
- Clean type legend (CI, SO, CO/CI, FU, DC, COC)
- Amenity items & linen/towel details
- All monthly booking sheets
            """)
        with col_b:
            st.markdown("""
**Output CSV columns**
`Client Name:` · `Property list` · `Reservation ID`
`DATE:` · `VILLA:` · `TYPE CLEAN:` · `PAX:`
`START TIME:` · `END TIME:` · `STATUS:`
`LAUNDRY :` · `Key:` · `Code:` · `Ameneties:`
`COMMENTS:` · `QB SHIFT ID` · `LAST SYNC`

**Supported month sheet names**
Jan, Feb, Mar/Mrt, Apr/Apl, May/Mei,
Jun, Jul, Aug, Sep, Oct/Okt, Nov, Dec
+ 2-digit year suffix (e.g. `Jan26`)
            """)
    st.stop()

# ── process files ──────────────────────────────────────────────────────────────
all_results = {}
for f in uploaded_files:
    try:
        client, months, frames = process_file(f)
        all_results[f.name] = (client, months, frames)
    except Exception as e:
        st.error(f"❌ Could not read **{f.name}**: {e}")

if not all_results:
    st.stop()

# ── per-client section ─────────────────────────────────────────────────────────
for fname, (client, months, frames) in all_results.items():

    cname = client["client_name"] or fname.replace(".xlsx", "")
    safe  = safe_name(cname)

    total_bookings = sum(len(df[df["STATUS:"] != "NONE"]) for df in frames.values())
    total_rows     = sum(len(df) for df in frames.values())
    canceled       = sum(len(df[df["STATUS:"].str.upper().isin(["CANCELED","CANCELLED"])]) for df in frames.values())

    st.markdown(f'<div class="client-card">', unsafe_allow_html=True)

    # name + file
    st.markdown(f"""
    <p class="client-name">🏡 {cname}</p>
    <p class="client-file">📄 {fname}</p>
    """, unsafe_allow_html=True)

    # stat boxes
    st.markdown(f"""
    <div class="stat-row">
      <div class="stat-box"><div class="stat-num">{len(months)}</div><div class="stat-lbl">Months</div></div>
      <div class="stat-box"><div class="stat-num">{len(client["properties"])}</div><div class="stat-lbl">Properties</div></div>
      <div class="stat-box"><div class="stat-num">{total_bookings}</div><div class="stat-lbl">Bookings</div></div>
      <div class="stat-box"><div class="stat-num" style="color:#c0392b">{canceled}</div><div class="stat-lbl">Canceled</div></div>
      <div class="stat-box"><div class="stat-num">{total_rows}</div><div class="stat-lbl">CSV Rows</div></div>
      <div class="stat-box"><div class="stat-num">{client["checkout_time"] or "—"}</div><div class="stat-lbl">Check-Out</div></div>
      <div class="stat-box"><div class="stat-num">{client["checkin_time"] or "—"}</div><div class="stat-lbl">Check-In</div></div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

    # ── tabs ───────────────────────────────────────────────────────────────────
    tab_props, tab_clean, tab_amen, tab_preview = st.tabs([
        "🏠 Properties & Settings",
        "🧹 Clean Types",
        "🛒 Amenities & Linens",
        "👁️ Preview & Export",
    ])

    # ── TAB: Properties & Settings ─────────────────────────────────────────────
    with tab_props:
        if client["properties"]:
            st.markdown('<p class="section-title">Properties</p>', unsafe_allow_html=True)
            prop_cols = st.columns(min(len(client["properties"]), 3))
            for i, prop in enumerate(client["properties"]):
                with prop_cols[i % 3]:
                    st.markdown(f"""
                    <div class="prop-card">
                        <div class="prop-name">{prop["name"]}</div>
                        <div class="prop-detail">
                            📍 {prop.get("address") or "—"}<br>
                            ⏱ Full clean: <b>{prop.get("hours") or "—"}h</b>
                            &nbsp;|&nbsp; Stay-over: <b>{prop.get("so_hrs") or "—"}h</b>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
        else:
            st.info("No property list found in this file.")

        st.markdown('<p class="section-title" style="margin-top:1.2rem">Service Settings</p>', unsafe_allow_html=True)

        set_c1, set_c2, set_c3, set_c4, set_c5, set_c6 = st.columns(6)
        set_c1.markdown(f"**🔑 Keys**<br>{yn_badge(client['keys'])}", unsafe_allow_html=True)
        set_c2.markdown(f"**🔢 Codes**<br>{yn_badge(client['codes'])}", unsafe_allow_html=True)
        set_c3.markdown(f"**🎁 Amenities**<br>{yn_badge(client['amenities'])}", unsafe_allow_html=True)
        set_c4.markdown(f"**🧺 Laundry**<br>{yn_badge(client['laundry'])}", unsafe_allow_html=True)
        set_c5.markdown(f"**🕐 Check-Out**<br><span class='badge badge-blue'>{client['checkout_time'] or '—'}</span>", unsafe_allow_html=True)
        set_c6.markdown(f"**🕐 Check-In**<br><span class='badge badge-blue'>{client['checkin_time'] or '—'}</span>", unsafe_allow_html=True)

    # ── TAB: Clean Types ───────────────────────────────────────────────────────
    with tab_clean:
        if client["clean_types"]:
            st.markdown('<p class="section-title">Clean Type Legend</p>', unsafe_allow_html=True)
            ct_df = pd.DataFrame(client["clean_types"]).rename(columns={
                "code": "Code", "name": "Type", "description": "Description"
            })
            st.dataframe(ct_df, use_container_width=True, hide_index=True,
                         column_config={
                             "Code":        st.column_config.TextColumn("Code", width="small"),
                             "Type":        st.column_config.TextColumn("Type", width="medium"),
                             "Description": st.column_config.TextColumn("Description", width="large"),
                         })
        else:
            st.info("No clean type legend found in this file.")

    # ── TAB: Amenities & Linens ────────────────────────────────────────────────
    with tab_amen:
        a_col, l_col = st.columns(2)

        with a_col:
            st.markdown('<p class="section-title">🛒 Amenity Items</p>', unsafe_allow_html=True)
            if client["amenity_items"]:
                a_df = pd.DataFrame(client["amenity_items"]).rename(
                    columns={"item": "Item", "qty": "Quantity"}
                )
                st.dataframe(a_df, use_container_width=True, hide_index=True)
            else:
                st.caption("No amenity items found.")

        with l_col:
            st.markdown('<p class="section-title">🧺 Linen & Towels</p>', unsafe_allow_html=True)
            if client["linen_items"]:
                l_df = pd.DataFrame(client["linen_items"]).rename(
                    columns={"item": "Item", "per": "Per Guest / Per Clean"}
                )
                st.dataframe(l_df, use_container_width=True, hide_index=True)
            else:
                st.caption("No linen items found.")

    # ── TAB: Preview & Export ──────────────────────────────────────────────────
    with tab_preview:
        if not frames:
            st.warning("No monthly sheets detected in this file.")
        else:
            # Controls row
            ctrl_a, ctrl_b = st.columns([2, 2])
            with ctrl_a:
                sel_month = st.selectbox(
                    "Month to preview",
                    list(frames.keys()),
                    key=f"month_{fname}",
                )
            with ctrl_b:
                show_none = st.checkbox(
                    "Include empty (NONE) rows",
                    value=False,
                    key=f"show_none_{fname}",
                )

            preview_df  = frames[sel_month].copy()
            display_df  = preview_df if show_none else preview_df[preview_df["STATUS:"] != "NONE"]

            # Status badge row
            status_counts = preview_df["STATUS:"].value_counts()
            badges = " ".join(
                badge_html(f"{s}: {c}", STATUS_BADGE.get(s.upper(), "badge-gray"))
                for s, c in status_counts.items()
            )
            st.markdown(badges + f'&nbsp; <span style="color:#999;font-size:0.8rem">({len(display_df)} rows shown)</span>', unsafe_allow_html=True)

            # Table
            st.markdown(f'<div class="preview-bar">📅 {sel_month} — {cname}</div>', unsafe_allow_html=True)
            st.dataframe(display_df, use_container_width=True, hide_index=True,
                         column_config={
                             "DATE:":   st.column_config.TextColumn("Date",   width="small"),
                             "VILLA:":  st.column_config.TextColumn("Villa",  width="medium"),
                             "STATUS:": st.column_config.TextColumn("Status", width="small"),
                             "TYPE CLEAN:": st.column_config.TextColumn("Type", width="small"),
                             "PAX:":    st.column_config.TextColumn("PAX",   width="small"),
                         })

            # Export
            st.markdown('<div class="export-section">', unsafe_allow_html=True)
            st.markdown('<div class="export-label">⬇️ Download</div>', unsafe_allow_html=True)

            e1, e2, e3 = st.columns(3)

            with e1:
                csv_single = frames[sel_month].to_csv(index=False).encode("utf-8")
                st.download_button(
                    f"📄 {sel_month} only",
                    data=csv_single,
                    file_name=f"Website_Reservation_{safe}_{sel_month}.csv",
                    mime="text/csv",
                    use_container_width=True,
                    key=f"dl_single_{fname}_{sel_month}",
                )

            with e2:
                combined = pd.concat(frames.values(), ignore_index=True)
                combined["Client Name:"]  = ""
                combined["Property list"] = ""
                combined.at[0, "Client Name:"] = client["client_name"]
                for i, prop in enumerate(client["properties"]):
                    if i < len(combined):
                        combined.at[i, "Property list"] = prop["name"]
                csv_all = combined.to_csv(index=False).encode("utf-8")
                st.download_button(
                    f"📦 All months (1 CSV)",
                    data=csv_all,
                    file_name=f"Website_Reservation_{safe}_All.csv",
                    mime="text/csv",
                    use_container_width=True,
                    key=f"dl_all_{fname}",
                )

            with e3:
                zip_data = build_client_zip(cname, frames)
                st.download_button(
                    f"🗜️ All months ZIP ({len(frames)} files)",
                    data=zip_data,
                    file_name=f"Website_Reservation_{safe}.zip",
                    mime="application/zip",
                    use_container_width=True,
                    key=f"dl_zip_{fname}",
                )

            st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<hr class="client-divider">', unsafe_allow_html=True)


# ── master export ──────────────────────────────────────────────────────────────
if len(all_results) > 1:
    total_files = sum(len(frames) for _, frames in [(c, f) for _, (c, _, f) in all_results.items()])
    st.markdown(f"""
    <div class="master-export">
        <h4 style="margin:0 0 6px;color:#2c5f2e">🗂️ Master Export — All {len(all_results)} Clients</h4>
        <p style="color:#666;font-size:0.88rem;margin:0 0 12px">
            Downloads a single ZIP with subfolders for each client
            ({total_files} CSV files total)
        </p>
    </div>
    """, unsafe_allow_html=True)

    master_zip = build_master_zip(all_results)
    st.download_button(
        f"📦 Download master ZIP — all clients ({total_files} files)",
        data=master_zip,
        file_name="Videmi_All_Clients.zip",
        mime="application/zip",
        use_container_width=True,
        key="master_zip_dl",
        type="primary",
    )
