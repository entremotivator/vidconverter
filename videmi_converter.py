import streamlit as st
import pandas as pd
import io
import zipfile
import re

st.set_page_config(
    page_title="Videmi – Booking to CSV",
    page_icon="🏡",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── CSS ────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-header { font-size:2rem; font-weight:700; color:#2c5f2e; margin-bottom:0; }
    .sub-header  { color:#666; font-size:0.95rem; margin-top:0; margin-bottom:1.5rem; }
    .client-card {
        background:#f0f7f0; border-left:4px solid #2c5f2e;
        border-radius:8px; padding:1rem 1.2rem; margin-bottom:0.5rem;
    }
    .client-name { font-size:1.3rem; font-weight:700; color:#2c5f2e; margin:0; }
    .prop-card {
        background:#fff; border:1px solid #d4e8d4;
        border-radius:6px; padding:0.6rem 0.9rem; margin:0.3rem 0;
        font-size:0.88rem;
    }
    .stat-box {
        background:#fff; border:1px solid #e0e0e0;
        border-radius:8px; padding:0.7rem 1rem; text-align:center;
    }
    .stat-num  { font-size:1.6rem; font-weight:700; color:#2c5f2e; }
    .stat-lbl  { font-size:0.78rem; color:#888; }
    .section-title { font-size:1rem; font-weight:600; color:#444;
                     border-bottom:1px solid #e0e0e0; padding-bottom:4px; margin:0.8rem 0 0.5rem; }
    .badge {
        display:inline-block; padding:2px 8px; border-radius:12px;
        font-size:0.78rem; font-weight:600; margin:2px;
    }
    .badge-green  { background:#d4edda; color:#155724; }
    .badge-yellow { background:#fff3cd; color:#856404; }
    .badge-red    { background:#f8d7da; color:#721c24; }
    .badge-gray   { background:#e9ecef; color:#495057; }
    .preview-header { background:#2c5f2e; color:white; padding:0.4rem 0.8rem;
                      border-radius:6px 6px 0 0; font-weight:600; font-size:0.9rem; }
</style>
""", unsafe_allow_html=True)

# ── header ─────────────────────────────────────────────────────────────────────
st.markdown('<p class="main-header">🏡 Videmi – Booking to CSV</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Upload one or more client <code>.xlsx</code> booking files and export clean CSVs ready for the website import.</p>', unsafe_allow_html=True)

# ── constants ──────────────────────────────────────────────────────────────────
OUT_COLS = [
    "Client Name:", "Property list", "Reservation ID",
    "DATE:", "VILLA:", "TYPE CLEAN:", "PAX:", "START TIME:", "END TIME:",
    "STATUS:", "LAUNDRY :", "Key:", "Code:", "Ameneties:",
    "COMMENTS:", "QB SHIFT ID", "LAST SYNC",
]

STATUS_COLORS = {
    "SCHEDULED": "badge-green",
    "CANCELED":  "badge-red",
    "NONE":      "badge-gray",
    "RESERVED":  "badge-yellow",
}

# ── parsing helpers ────────────────────────────────────────────────────────────
def cs(v):
    """Clean string from cell value."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("nan", "nat", "none") else s


def scan_label(raw, keyword):
    kw = keyword.lower()
    for _, row in raw.iterrows():
        if len(row) < 2:
            continue
        label = cs(row.iloc[0]).lower()
        if kw in label:
            return cs(row.iloc[1])
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
    months = ["jan","feb","mar","mrt","apr","apl","may","mei","jun",
              "jul","aug","sep","oct","okt","nov","dec"]
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
        import datetime
        if isinstance(val, datetime.time):
            return val.strftime("%H:%M")
        dt = pd.to_datetime(val)
        return dt.strftime("%H:%M")
    except Exception:
        return s


def get_col(row, idx):
    if idx is None or idx >= len(row):
        return ""
    return cs(row.iloc[idx])


# ── rich client info extraction ────────────────────────────────────────────────
def read_client_info(xl):
    """Extract everything useful from the client info sheet."""
    info = {
        "client_name":   "",
        "properties":    [],        # list of dicts: name, address, hours, so_hours
        "checkout_time": "",
        "checkin_time":  "",
        "keys":          "",
        "codes":         "",
        "amenities":     "",
        "laundry":       "",
        "clean_types":   [],        # list of dicts: code, name, description
        "amenity_items": [],        # list of dicts: item, quantity
        "linen_items":   [],        # list of dicts: item, per
    }

    sheet = find_info_sheet(xl)
    if sheet is None:
        return info

    raw = xl.parse(sheet, header=None)

    # ── client name: first non-empty cell ──────────────────────────────────────
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

    # ── scalar settings ────────────────────────────────────────────────────────
    info["checkout_time"] = fmt_time(scan_label(raw, "check-out"))
    info["checkin_time"]  = fmt_time(scan_label(raw, "check-in"))
    info["keys"]          = scan_label(raw, "key")
    info["codes"]         = scan_label(raw, "code")
    info["amenities"]     = scan_label(raw, "amenities")
    info["laundry"]       = scan_label(raw, "laundry")

    # ── clean types table (cols: Type label | Code | Description) ─────────────
    in_clean = False
    for _, row in raw.iterrows():
        label = cs(row.iloc[4]) if len(row) > 4 else ""
        if not in_clean:
            if "type of clean" in label.lower():
                in_clean = True
            continue
        type_name = cs(row.iloc[4]) if len(row) > 4 else ""
        code      = cs(row.iloc[5]) if len(row) > 5 else ""
        desc      = cs(row.iloc[6]) if len(row) > 6 else ""
        if not type_name or not code:
            break
        info["clean_types"].append({"name": type_name, "code": code, "description": desc})

    # ── properties table ───────────────────────────────────────────────────────
    in_props = False
    for _, row in raw.iterrows():
        label = cs(row.iloc[0]).lower()
        if not in_props:
            if any(k in label for k in ("villas", "appartment", "apartment", "properties")):
                if any(k in label for k in ("name", "list")) or label.endswith(":"):
                    in_props = True
            continue
        p_name = cs(row.iloc[0])
        if not p_name:
            break
        if any(k in p_name.lower() for k in ("name", "address", "item", "service")):
            continue
        info["properties"].append({
            "name":    re.sub(r"\bApp\b", "Apartment", p_name),
            "address": cs(row.iloc[1]) if len(row) > 1 else "",
            "hours":   cs(row.iloc[2]) if len(row) > 2 else "",
            "so_hrs":  cs(row.iloc[3]) if len(row) > 3 else "",
        })

    # ── amenity items ──────────────────────────────────────────────────────────
    in_amen = False
    for _, row in raw.iterrows():
        label = cs(row.iloc[0]).lower()
        if not in_amen:
            if "list of amenities" in label or (label == "item"):
                in_amen = True
            continue
        item = cs(row.iloc[0])
        qty  = cs(row.iloc[1]) if len(row) > 1 else ""
        if not item:
            break
        if any(k in item.lower() for k in ("linen", "service", "item", "laundry")):
            break
        info["amenity_items"].append({"item": item, "qty": qty if qty else "0"})

    # ── linen items ────────────────────────────────────────────────────────────
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


# ── sheet parser ───────────────────────────────────────────────────────────────
def parse_sheet(xl, sheet_name, client):
    raw = xl.parse(sheet_name, header=None)
    header_idx = find_header_row(raw)
    if header_idx is None:
        return pd.DataFrame(columns=OUT_COLS)

    hv = [cs(v).upper().rstrip(":") for v in raw.iloc[header_idx]]

    def ci(keyword):
        for i, c in enumerate(hv):
            if keyword in c:
                return i
        return None

    idx = {
        "date":   ci("DATE"),   "villa": ci("VILLA"),
        "type":   ci("TYPE CLEAN"), "pax": ci("PAX"),
        "start":  ci("START TIME"), "end": ci("END TIME"),
        "status": ci("RESERVATION STATUS"),
        "laund":  ci("LAUNDRY"),    "comm": ci("COMMENTS"),
    }

    rows = []
    for _, row in raw.iloc[header_idx + 1:].reset_index(drop=True).iterrows():
        status = get_col(row, idx["status"]).upper() or "NONE"
        villa  = re.sub(r"\bApp\b", "Apartment", get_col(row, idx["villa"]))
        rows.append({
            "Client Name:": "", "Property list": "", "Reservation ID": "",
            "DATE:":        fmt_date(get_col(row, idx["date"])),
            "VILLA:":       villa,
            "TYPE CLEAN:":  get_col(row, idx["type"]),
            "PAX:":         get_col(row, idx["pax"]),
            "START TIME:":  get_col(row, idx["start"]),
            "END TIME:":    get_col(row, idx["end"]),
            "STATUS:":      status,
            "LAUNDRY :":    get_col(row, idx["laund"]),
            "Key:":         client["keys"],
            "Code:":        client["codes"],
            "Ameneties:":   client["amenities"],
            "COMMENTS:":    get_col(row, idx["comm"]),
            "QB SHIFT ID":  "", "LAST SYNC": "",
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


# ── UI ─────────────────────────────────────────────────────────────────────────
uploaded_files = st.file_uploader(
    "Drop client `.xlsx` files here",
    type=["xlsx"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)

if not uploaded_files:
    st.markdown("### 👆 Upload one or more client `.xlsx` booking files to get started")
    with st.expander("ℹ️ What this app does"):
        st.markdown("""
        **Videmi – Booking to CSV** reads any client reservation spreadsheet and converts it
        into the exact website import format.

        **Auto-detected from each file:**
        - Client name, properties (name + address), check-in/out times
        - Keys, Codes, Amenities, Laundry settings
        - Clean type legend (CI, SO, CO/CI, FU, DC, COC)
        - Amenity items & linen details

        **Output columns:**
        `Client Name:` · `Property list` · `Reservation ID` · `DATE:` · `VILLA:` · `TYPE CLEAN:` · `PAX:` · `START TIME:` · `END TIME:` · `STATUS:` · `LAUNDRY :` · `Key:` · `Code:` · `Ameneties:` · `COMMENTS:` · `QB SHIFT ID` · `LAST SYNC`
        """)
    st.stop()

# ── process all uploaded files ─────────────────────────────────────────────────
all_results = {}
for f in uploaded_files:
    try:
        client, months, frames = process_file(f)
        all_results[f.name] = (client, months, frames)
    except Exception as e:
        st.error(f"❌ Could not read **{f.name}**: {e}")

if not all_results:
    st.stop()

# ── per-client panels ──────────────────────────────────────────────────────────
for fname, (client, months, frames) in all_results.items():

    cname = client["client_name"] or fname.replace(".xlsx", "")

    # Count real bookings (not NONE)
    total_bookings = sum(
        len(df[df["STATUS:"] != "NONE"])
        for df in frames.values()
    )
    total_rows = sum(len(df) for df in frames.values())

    with st.container():
        st.markdown(f"""
        <div class="client-card">
            <p class="client-name">🏡 {cname}</p>
            <span style="color:#666;font-size:0.85rem;">📄 {fname}</span>
        </div>
        """, unsafe_allow_html=True)

        # ── stats row ─────────────────────────────────────────────────────────
        c1, c2, c3, c4, c5, c6 = st.columns(6)
        c1.markdown(f'<div class="stat-box"><div class="stat-num">{len(months)}</div><div class="stat-lbl">Month Sheets</div></div>', unsafe_allow_html=True)
        c2.markdown(f'<div class="stat-box"><div class="stat-num">{len(client["properties"])}</div><div class="stat-lbl">Properties</div></div>', unsafe_allow_html=True)
        c3.markdown(f'<div class="stat-box"><div class="stat-num">{total_bookings}</div><div class="stat-lbl">Bookings</div></div>', unsafe_allow_html=True)
        c4.markdown(f'<div class="stat-box"><div class="stat-num">{total_rows}</div><div class="stat-lbl">Total CSV Rows</div></div>', unsafe_allow_html=True)
        c5.markdown(f'<div class="stat-box"><div class="stat-num">{client["checkin_time"] or "—"}</div><div class="stat-lbl">Check-In</div></div>', unsafe_allow_html=True)
        c6.markdown(f'<div class="stat-box"><div class="stat-num">{client["checkout_time"] or "—"}</div><div class="stat-lbl">Check-Out</div></div>', unsafe_allow_html=True)

        st.markdown("")

        # ── detail tabs ────────────────────────────────────────────────────────
        tab_props, tab_clean, tab_amen, tab_preview = st.tabs([
            "🏠 Properties", "🧹 Clean Types", "🛒 Amenities & Linens", "👁️ Preview & Export"
        ])

        # ── Properties tab ─────────────────────────────────────────────────────
        with tab_props:
            if client["properties"]:
                pcols = st.columns(min(len(client["properties"]), 3))
                for i, prop in enumerate(client["properties"]):
                    with pcols[i % 3]:
                        addr = prop.get("address", "") or "—"
                        hrs  = prop.get("hours", "")  or "—"
                        so   = prop.get("so_hrs", "") or "—"
                        st.markdown(f"""
                        <div class="prop-card">
                            <b>{prop['name']}</b><br>
                            📍 {addr}<br>
                            ⏱ Clean: <b>{hrs}h</b> &nbsp;|&nbsp; Stay-over: <b>{so}h</b>
                        </div>
                        """, unsafe_allow_html=True)

            scol1, scol2, scol3, scol4 = st.columns(4)
            def badge(val):
                v = str(val).strip().lower()
                cls = "badge-green" if v in ("yes","y","true","1") else "badge-red" if v in ("no","n","false","0") else "badge-gray"
                return f'<span class="badge {cls}">{val}</span>'

            scol1.markdown(f"🔑 **Keys** {badge(client['keys'])}", unsafe_allow_html=True)
            scol2.markdown(f"🔢 **Codes** {badge(client['codes'])}", unsafe_allow_html=True)
            scol3.markdown(f"🎁 **Amenities** {badge(client['amenities'])}", unsafe_allow_html=True)
            scol4.markdown(f"🧺 **Laundry** {badge(client['laundry'])}", unsafe_allow_html=True)

        # ── Clean types tab ────────────────────────────────────────────────────
        with tab_clean:
            if client["clean_types"]:
                ct_df = pd.DataFrame(client["clean_types"]).rename(
                    columns={"code": "Code", "name": "Type", "description": "Description"}
                )
                st.dataframe(ct_df, use_container_width=True, hide_index=True)
            else:
                st.info("No clean type legend found in this file.")

        # ── Amenities & Linens tab ─────────────────────────────────────────────
        with tab_amen:
            acol, lcol = st.columns(2)
            with acol:
                st.markdown('<p class="section-title">🛒 Amenity Items</p>', unsafe_allow_html=True)
                if client["amenity_items"]:
                    a_df = pd.DataFrame(client["amenity_items"]).rename(
                        columns={"item": "Item", "qty": "Quantity"}
                    )
                    st.dataframe(a_df, use_container_width=True, hide_index=True)
                else:
                    st.caption("None found.")
            with lcol:
                st.markdown('<p class="section-title">🧺 Linen & Towels</p>', unsafe_allow_html=True)
                if client["linen_items"]:
                    l_df = pd.DataFrame(client["linen_items"]).rename(
                        columns={"item": "Item", "per": "Per Guest / Per Clean"}
                    )
                    st.dataframe(l_df, use_container_width=True, hide_index=True)
                else:
                    st.caption("None found.")

        # ── Preview & Export tab ───────────────────────────────────────────────
        with tab_preview:
            if not frames:
                st.warning("No monthly sheets found.")
            else:
                # Month selector
                month_options = list(frames.keys())
                sel_month = st.selectbox(
                    "Select month to preview",
                    month_options,
                    key=f"month_{fname}",
                )

                preview_df = frames[sel_month].copy()

                # Status summary badges
                status_counts = preview_df["STATUS:"].value_counts()
                badge_html = ""
                for status, count in status_counts.items():
                    cls = STATUS_COLORS.get(status.upper(), "badge-gray")
                    badge_html += f'<span class="badge {cls}">{status}: {count}</span> '
                st.markdown(badge_html, unsafe_allow_html=True)

                # Only show real booking rows in preview (hide NONE rows toggle)
                show_all = st.checkbox("Show empty (NONE) rows", value=False, key=f"shownone_{fname}")
                display_df = preview_df if show_all else preview_df[preview_df["STATUS:"] != "NONE"]

                st.markdown(f'<div class="preview-header">📅 {sel_month} — {len(display_df)} rows shown</div>', unsafe_allow_html=True)
                st.dataframe(display_df, use_container_width=True, hide_index=True)

                # ── export controls ────────────────────────────────────────────
                st.markdown("---")
                st.markdown("#### ⬇️ Export")
                exp_col1, exp_col2, exp_col3 = st.columns(3)

                # Single month CSV
                with exp_col1:
                    csv_month = frames[sel_month].to_csv(index=False).encode("utf-8")
                    safe = re.sub(r"[^\w\-]", "_", cname)
                    st.download_button(
                        f"📄 This month ({sel_month})",
                        data=csv_month,
                        file_name=f"Website_Reservation_{safe}_{sel_month}.csv",
                        mime="text/csv",
                        use_container_width=True,
                        key=f"dl_month_{fname}_{sel_month}",
                    )

                # All months combined CSV for this client
                with exp_col2:
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

                # ZIP of all months
                with exp_col3:
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                        for sheet, df in frames.items():
                            zf.writestr(
                                f"Website_Reservation_{safe}_{sheet}.csv",
                                df.to_csv(index=False),
                            )
                    zip_buf.seek(0)
                    st.download_button(
                        f"🗜 All months (ZIP, {len(frames)} files)",
                        data=zip_buf,
                        file_name=f"Website_Reservation_{safe}.zip",
                        mime="application/zip",
                        use_container_width=True,
                        key=f"dl_zip_{fname}",
                    )

    st.markdown("---")

# ── master export (multi-client) ───────────────────────────────────────────────
if len(all_results) > 1:
    st.markdown("### 🗂 Master Export — All Clients")
    if st.button("⬇️ Download master ZIP (all clients × all months)", use_container_width=True, type="primary"):
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for fname, (client, _, frames) in all_results.items():
                safe = re.sub(r"[^\w\-]", "_", client["client_name"] or fname.replace(".xlsx", ""))
                for sheet, df in frames.items():
                    zf.writestr(f"{safe}/Website_Reservation_{safe}_{sheet}.csv", df.to_csv(index=False))
        zip_buf.seek(0)
        st.download_button(
            "📦 Download master ZIP",
            data=zip_buf,
            file_name="Videmi_All_Clients.zip",
            mime="application/zip",
            use_container_width=True,
            key="master_zip",
        )
