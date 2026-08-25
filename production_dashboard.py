import io
import random
import re
import shutil
import tempfile
import zipfile
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ---- Canonical stage names used everywhere downstream ----
STAGES = [
    "Chassis",
    "Air leak",
    "Insulation",
    "Calibration",
    "Lat",
    "Comm",
    "EOL",
    "PDI",
]

RESULTS = ["Pass", "Fail", "Pending"]

# Optional columns that may or may not be present in a real export
DATE_COL_CANDIDATES = ["Date", "date", "Entry_Date", "Received_Date"]
WORK_ORDER_COL_CANDIDATES = ["Work_Order", "work_order", "WO", "Work Order"]
# NOTE: real exports use all kinds of names for the serial column
# ("Module Serial Number" in the user's sheet, "battery_serial" in the
# sample data, etc.) so we try several candidates in order.
SERIAL_COL_CANDIDATES = [
    "battery_serial",
    "Battery_Serial",
    "Module Serial Number",
    "Module_Serial_Number",
    "Serial Number",
    "Serial_Number",
    "Serial",
    "serial",
]


# =============================================================================
# Sample / plain-file data (unchanged simple mode)
# =============================================================================

def generate_sample_data(n_packs: int = 30) -> pd.DataFrame:
    """Builds demo data in memory — no file I/O, so it can't silently fail
    due to folder permissions, OneDrive sync locks, etc."""
    rows = []
    work_orders = ["WO-101", "WO-102", "WO-103", "WO-104"]
    for i in range(1, n_packs + 1):
        row = {"battery_serial": f"BP-LFP-2026-{1000 + i}"}
        row["Work_Order"] = random.choice(work_orders)
        row["Date"] = (datetime(2026, 6, 1) + timedelta(days=random.randint(0, 50))).strftime("%Y-%m-%d")
        stopped = False
        for stage in STAGES:
            if stopped:
                row[stage] = "Pending"
                continue
            roll = random.random()
            if roll < 0.80:
                row[stage] = "Pass"
            elif roll < 0.90:
                row[stage] = "Fail"
                stopped = True
            else:
                row[stage] = "Pending"
                stopped = True
        row["IB"] = random.randint(0, 100)
        row["DOM"] = random.randint(0, 800)
        row["Total"] = row["IB"] + row["DOM"]
        rows.append(row)
    return pd.DataFrame(rows)


def _read_one(uploaded_file) -> pd.DataFrame:
    if uploaded_file.name.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    return pd.read_excel(uploaded_file)


def load_data(uploaded_files) -> pd.DataFrame:
    if uploaded_files:
        frames = [_read_one(f) for f in uploaded_files]
        return pd.concat(frames, ignore_index=True)

    st.sidebar.info("No file uploaded — using sample data")
    return generate_sample_data()


def find_col(df: pd.DataFrame, candidates: list):
    for c in candidates:
        if c in df.columns:
            return c
    return None


def _normalize(name: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(name).lower())


def match_stage_columns(df: pd.DataFrame) -> dict:
    """Maps canonical STAGES names -> whatever column actually holds that
    stage's data in this dataframe, so real-world spreadsheets don't have to
    use our exact stage spelling.

    Handles cases like:
      "Air leak"     <-> "Airleak"          (spacing/case differences: exact
                                              normalized match)
      "Lat"          <-> "Latlong "         (abbreviation is a substring of
                                              the full column name)
      "Comm"         <-> "Communication"    (same idea)
    """
    normalized_cols = {c: _normalize(c) for c in df.columns}
    mapping = {}
    used_cols = set()

    # Pass 1: exact normalized match (handles "Airleak" vs "Air leak")
    for stage in STAGES:
        stage_norm = _normalize(stage)
        for col, col_norm in normalized_cols.items():
            if col in used_cols:
                continue
            if col_norm == stage_norm:
                mapping[stage] = col
                used_cols.add(col)
                break

    # Pass 2: substring match for abbreviations (handles "Lat" vs "Latlong",
    # "Comm" vs "Communication")
    for stage in STAGES:
        if stage in mapping:
            continue
        stage_norm = _normalize(stage)
        for col, col_norm in normalized_cols.items():
            if col in used_cols:
                continue
            if stage_norm and (stage_norm in col_norm or col_norm in stage_norm):
                mapping[stage] = col
                used_cols.add(col)
                break

    return mapping


def apply_stage_aliases(df: pd.DataFrame) -> pd.DataFrame:
    """Renames whatever stage-like columns exist to the canonical STAGES
    names, and fills blank/NaN cells in those columns with 'Pending' so they
    aren't silently dropped from the Pass/Fail/Pending counts."""
    mapping = match_stage_columns(df)
    rename_map = {actual: canonical for canonical, actual in mapping.items() if actual != canonical}
    if rename_map:
        df = df.rename(columns=rename_map)
    for stage in STAGES:
        if stage in df.columns:
            df[stage] = df[stage].fillna("Pending")
            # also normalize blank strings / whitespace-only cells
            df[stage] = df[stage].replace(r"^\s*$", "Pending", regex=True)
    return df


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return buf.getvalue()


def excel_style_column_filters(df: pd.DataFrame, key_prefix: str, max_unique: int = 100) -> pd.DataFrame:
    """Renders one 'Filter by <column>' multiselect per column (like Excel's
    AutoFilter dropdowns) and returns the dataframe filtered by whatever the
    user picked. Columns with too many unique values (e.g. free-text notes)
    are skipped to avoid an unusable giant dropdown."""
    if df.empty:
        return df

    filterable_cols = [c for c in df.columns if df[c].nunique(dropna=True) <= max_unique]
    if not filterable_cols:
        return df

    with st.expander("Filter columns (like Excel's column filters)", expanded=False):
        st.caption("Pick values for any column to narrow down the table. Leave a column empty to include everything.")
        n_cols = 3
        cols_ui = st.columns(n_cols)
        filtered = df
        for i, col in enumerate(filterable_cols):
            options = sorted(df[col].dropna().astype(str).unique().tolist())
            with cols_ui[i % n_cols]:
                picked = st.multiselect(
                    f"{col}",
                    options=options,
                    default=[],
                    key=f"{key_prefix}_filter_{col}",
                )
            if picked:
                filtered = filtered[filtered[col].astype(str).isin(picked)]
        return filtered


# =============================================================================
# Folder-structured upload (zip):
#   uploaded_folder.zip
#     Chassis/
#       2026-07-20/  data.xlsx
#       2026-07-21/  data.xlsx
#     Air leak/
#       2026-07-20/  data.xlsx
#       ...
#
# We now read EVERY day folder (not just the latest) so we can:
#   - search "which batteries were pending/waiting on date X"
#   - show a battery's full date/month history
#   - avoid re-counting a battery that already completed on an earlier day
#     (its current status = its most recent recorded result, per stage)
# =============================================================================

def _find_stage_dirs(root: Path) -> dict:
    stage_lookup = {_normalize(s): s for s in STAGES}
    found = {}
    for p in root.rglob("*"):
        if p.is_dir():
            key = _normalize(p.name)
            if key in stage_lookup and stage_lookup[key] not in found:
                found[stage_lookup[key]] = p
    # also allow substring matches for folder names, same as column matching
    for p in root.rglob("*"):
        if p.is_dir():
            key = _normalize(p.name)
            for stage_norm, stage in stage_lookup.items():
                if stage in found:
                    continue
                if stage_norm and (stage_norm in key or key in stage_norm):
                    found[stage] = p
    return found


def _parse_date(name: str):
    parsed = pd.to_datetime(name, errors="coerce", dayfirst=False)
    if pd.isna(parsed):
        cleaned = name.replace("_", "-").replace(".", "-")
        parsed = pd.to_datetime(cleaned, errors="coerce", dayfirst=True)
    return parsed


def _all_date_subdirs(stage_dir: Path) -> list:
    """All valid (date, dir) pairs under a stage folder, sorted oldest first."""
    found = []
    for p in stage_dir.iterdir():
        if not p.is_dir():
            continue
        parsed = _parse_date(p.name)
        if pd.isna(parsed):
            continue
        found.append((parsed, p))
    found.sort(key=lambda t: t[0])
    return found


def _first_data_file(folder: Path):
    for p in sorted(folder.iterdir()):
        if p.suffix.lower() in (".xlsx", ".xls", ".csv"):
            return p
    return None


def load_history_from_zip(uploaded_zip) -> tuple:
    """Walks every Stage -> every Date -> Excel/CSV inside the zip.

    Returns:
      history_df: long format, one row per (stage, date, battery_serial)
      latest_wide_df: one row per battery_serial, one column per stage, using
                  the MOST RECENT recorded result for that battery/stage
      stage_info: latest date/file used per stage, for the summary table
      debug_info: dict with 'all_dirs' (every folder name found in the zip)
                  and 'unparsed_date_dirs' (sub-folders under a matched stage
                  whose name did NOT parse as a date) — used to explain
                  failures instead of just returning nothing
    """
    tmp_dir = Path(tempfile.mkdtemp())
    history_rows = []
    stage_info = []
    debug_info = {"all_dirs": [], "unparsed_date_dirs": []}

    try:
        with zipfile.ZipFile(uploaded_zip) as zf:
            zf.extractall(tmp_dir)
            mtimes = {}
            for zi in zf.infolist():
                mtimes[zi.filename] = datetime(*zi.date_time)

        debug_info["all_dirs"] = sorted(
            str(p.relative_to(tmp_dir)) for p in tmp_dir.rglob("*") if p.is_dir()
        )

        stage_dirs = _find_stage_dirs(tmp_dir)
        if not stage_dirs:
            return pd.DataFrame(), pd.DataFrame(), [], debug_info

        for stage, sdir in stage_dirs.items():
            date_dirs = _all_date_subdirs(sdir)
            for p in sdir.iterdir():
                if p.is_dir() and pd.isna(_parse_date(p.name)):
                    debug_info["unparsed_date_dirs"].append(f"{stage}/{p.name}")
            if not date_dirs:
                stage_info.append({"stage": stage, "latest_date": None, "file": None})
                continue

            for parsed_date, ddir in date_dirs:
                data_file = _first_data_file(ddir)
                if data_file is None:
                    continue
                frame = (
                    pd.read_csv(data_file)
                    if data_file.suffix.lower() == ".csv"
                    else pd.read_excel(data_file)
                )

                serial_col = find_col(frame, SERIAL_COL_CANDIDATES)
                if serial_col is None:
                    serial_col = frame.columns[0]

                result_col = None
                for cand in [stage, "Result", "result", "Status", "status"]:
                    if cand in frame.columns and cand != serial_col:
                        result_col = cand
                        break
                if result_col is None:
                    other_cols = [c for c in frame.columns if c != serial_col]
                    result_col = other_cols[0] if other_cols else None
                if result_col is None:
                    continue

                wo_col = find_col(frame, WORK_ORDER_COL_CANDIDATES)

                rel_path = str(data_file.relative_to(tmp_dir))
                file_updated = mtimes.get(rel_path)

                for _, r in frame.iterrows():
                    result_val = r[result_col]
                    if pd.isna(result_val) or str(result_val).strip() == "":
                        result_val = "Pending"
                    history_rows.append({
                        "stage": stage,
                        "date": parsed_date.normalize(),
                        "battery_serial": r[serial_col],
                        "result": result_val,
                        "work_order": r[wo_col] if wo_col else None,
                        "file_updated": file_updated,
                    })

            latest_date, latest_dir = date_dirs[-1]
            latest_file = _first_data_file(latest_dir)
            stage_info.append({
                "stage": stage,
                "latest_date": str(latest_date.date()),
                "file": latest_file.name if latest_file else None,
                "days_available": len(date_dirs),
            })

        history_df = pd.DataFrame(history_rows)
        if history_df.empty:
            return history_df, pd.DataFrame(), stage_info, debug_info

        # Current status per stage = most recent recorded result for that
        # battery in that stage (so completed batteries aren't "lost" just
        # because they don't reappear in later day-folders).
        latest_idx = history_df.sort_values("date").groupby(
            ["battery_serial", "stage"]
        )["date"].idxmax()
        latest_records = history_df.loc[latest_idx]
        latest_wide_df = latest_records.pivot(
            index="battery_serial", columns="stage", values="result"
        ).reset_index()

        for stage in STAGES:
            if stage not in latest_wide_df.columns:
                latest_wide_df[stage] = "Pending"
            else:
                latest_wide_df[stage] = latest_wide_df[stage].fillna("Pending")

        # Attach most recent work_order per battery, if present anywhere
        wo_map = (
            history_df.dropna(subset=["work_order"])
            .sort_values("date")
            .groupby("battery_serial")["work_order"]
            .last()
        )
        if not wo_map.empty:
            latest_wide_df["Work_Order"] = latest_wide_df["battery_serial"].map(wo_map)

        return history_df, latest_wide_df, stage_info, debug_info
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# =============================================================================
# Shared computations
# =============================================================================

def stage_counts(df: pd.DataFrame) -> pd.DataFrame:
    counts = {}
    for stage in STAGES:
        if stage in df.columns:
            # Belt-and-suspenders: treat any remaining blank/NaN cell as
            # Pending so it isn't silently excluded from value_counts().
            col = df[stage].fillna("Pending")
            counts[stage] = col.value_counts().reindex(RESULTS).fillna(0).astype(int)
    summary = pd.DataFrame(counts).T
    summary = summary[RESULTS]
    return summary


def render_chart(summary: pd.DataFrame):
    colors = {"Pass": "#008300", "Fail": "#e34948", "Pending": "#eda100"}
    fig = go.Figure()
    for result in RESULTS:
        fig.add_trace(go.Bar(
            y=summary.index,
            x=summary[result],
            name=result,
            orientation="h",
            marker_color=colors[result],
        ))
    fig.update_layout(
        barmode="stack",
        height=450,
        margin=dict(l=10, r=10, t=10, b=10),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        yaxis=dict(autorange="reversed"),
    )
    st.plotly_chart(fig, width="stretch")


def current_stage(row, stage_cols):
    for stage in stage_cols:
        if row[stage] in ("Fail", "Pending"):
            return stage, row[stage]
    return "Complete", "Pass"


def build_waiting_report(df: pd.DataFrame, stage_cols: list, serial_col: str) -> pd.DataFrame:
    rows = []
    for _, r in df.iterrows():
        stage, status = current_stage(r, stage_cols)
        rows.append({
            "battery_serial": r[serial_col],
            "waiting_at_stage": stage,
            "status": status,
        })
    return pd.DataFrame(rows)


def completion_pct(row, stage_cols):
    passed = sum(1 for s in stage_cols if row[s] == "Pass")
    return round(100 * passed / len(stage_cols), 1)


# =============================================================================
# App
# =============================================================================

def main():
    st.set_page_config(page_title="Production dashboard", layout="wide")
    st.title("Production dashboard")

    st.sidebar.subheader("Data source")
    source_mode = st.sidebar.radio(
        "How is your data organized?",
        ["Stage/Date folders (zip)", "Plain Excel/CSV file(s)"],
        help=(
            "Use 'Stage/Date folders' if your data looks like "
            "Chassis/2026-07-21/data.xlsx, Air leak/2026-07-21/data.xlsx, etc. "
            "Zip that top-level folder and upload it here."
        ),
    )

    history_df = pd.DataFrame()
    stage_info = []

    if source_mode == "Stage/Date folders (zip)":
        uploaded_zip = st.sidebar.file_uploader(
            "Upload the zipped folder (one sub-folder per stage, "
            "each with one sub-folder per day)",
            type=["zip"],
        )
        if uploaded_zip is not None:
            history_df, df, stage_info, debug_info = load_history_from_zip(uploaded_zip)

            if df.empty and not stage_info:
                st.error(
                    "Couldn't find any stage folders matching "
                    f"{STAGES} inside that zip."
                )
                st.write("**Folders actually found inside your zip:**")
                if debug_info["all_dirs"]:
                    st.write(debug_info["all_dirs"])
                else:
                    st.write("No sub-folders found at all — the zip may be flat "
                              "(files directly inside, no stage folders) or empty.")
                st.info(
                    "Fix: rename your top-level stage folders to match one of "
                    f"{STAGES} (case/spaces/underscores/hyphens are OK), rezip, "
                    "and re-upload."
                )
                return

            if df.empty and stage_info:
                st.error(
                    "Found stage folders, but no valid date sub-folders (or no "
                    "readable Excel/CSV) inside them."
                )
                st.write("**Stage folders found:**", [s["stage"] for s in stage_info])
                if debug_info["unparsed_date_dirs"]:
                    st.write(
                        "**These sub-folders were found but their names didn't "
                        "parse as dates:**"
                    )
                    st.write(debug_info["unparsed_date_dirs"])
                st.info(
                    "Fix: day-folder names must look like a date, e.g. "
                    "'2026-07-21', '21-07-2026', or '2026_07_21'. Each day "
                    "folder also needs an .xlsx/.xls/.csv file directly inside it."
                )
                return
        else:
            df = generate_sample_data()
            st.sidebar.info("No zip uploaded — using sample data")
    else:
        uploaded_files = st.sidebar.file_uploader(
            "Upload Excel/CSV file(s) — select all files to combine them",
            type=["xlsx", "csv"],
            accept_multiple_files=True,
        )
        df = load_data(uploaded_files)

    if df.empty:
        st.warning("No data found.")
        return

    # ---- Map real-world column names onto our canonical stage names, and
    # ---- make sure blank cells count as "Pending" instead of vanishing.
    df = apply_stage_aliases(df)

    serial_col = find_col(df, SERIAL_COL_CANDIDATES)
    stage_cols = [s for s in STAGES if s in df.columns]
    date_col = find_col(df, DATE_COL_CANDIDATES)
    wo_col = find_col(df, WORK_ORDER_COL_CANDIDATES)

    if date_col:
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce")

    # ---- Work order: listing (always, from full folder) + filter ----
    if wo_col:
        st.sidebar.subheader("Work orders in this dataset")
        wo_list = sorted(df[wo_col].dropna().astype(str).unique().tolist())
        st.sidebar.write(f"{len(wo_list)} work order(s) found")
        with st.sidebar.expander("View work order list"):
            st.write(pd.DataFrame({"Work_Order": wo_list}))

        selected_wos = st.sidebar.multiselect(
            "Choose work order number(s)",
            options=wo_list,
            default=[],
            help="Leave empty to show all work orders, or pick one/more numbers to filter.",
        )
        if selected_wos:
            df = df[df[wo_col].astype(str).isin(selected_wos)]
            if not history_df.empty:
                history_df = history_df[history_df["work_order"].astype(str).isin(selected_wos)]

    st.subheader("Raw data")
    if serial_col:
        cols = [serial_col] + [c for c in df.columns if c != serial_col]
        df = df[cols]
    df_view = excel_style_column_filters(df, key_prefix="raw")
    st.dataframe(df_view, width="stretch")
    st.download_button(
        "Download this table (Excel)",
        data=to_excel_bytes(df_view),
        file_name="production_data_export.xlsx",
    )

    # ---- Latest date + last-updated info per stage (from folder mode) ----
    if stage_info:
        st.subheader("Latest data per stage")
        info_df = pd.DataFrame(stage_info)
        st.dataframe(info_df, width="stretch", hide_index=True)

    # ---- Oldest entry (plain-file mode) ----
    if date_col and df[date_col].notna().any():
        oldest_row = df.loc[df[date_col].idxmin()]
        st.info(
            f"Oldest entry: **{oldest_row.get(serial_col, 'N/A')}** "
            f"dated **{oldest_row[date_col].strftime('%Y-%m-%d')}**"
        )

    st.subheader("Stage-wise Pass / Fail / Pending counts")
    summary = stage_counts(df)
    st.dataframe(summary, width="stretch")
    render_chart(summary)

    # ---- WIP / waiting-stage report ----
    if serial_col and stage_cols:
        st.subheader("Work in progress — which stage each pack is waiting on")
        waiting_df = build_waiting_report(df, stage_cols, serial_col)

        total = len(waiting_df)
        complete = (waiting_df["waiting_at_stage"] == "Complete").sum()
        wip = total - complete
        avg_completion = df.apply(lambda r: completion_pct(r, stage_cols), axis=1).mean() if total else 0

        c1, c2, c3 = st.columns(3)
        c1.metric("Total packs", total)
        c2.metric("Total WIP (not yet complete)", wip)
        c3.metric("Avg. completion", f"{avg_completion:.1f}%")

        wip_by_stage = (
            waiting_df[waiting_df["waiting_at_stage"] != "Complete"]
            .groupby(["waiting_at_stage", "status"])
            .size()
            .reset_index(name="count")
            .sort_values("waiting_at_stage")
        )
        st.dataframe(wip_by_stage, width="stretch", hide_index=True)

        stage_filter_options = [s for s in stage_cols if s in waiting_df["waiting_at_stage"].values]
        if stage_filter_options:
            wip_stage_choice = st.selectbox(
                "See battery numbers waiting at a specific stage", stage_filter_options
            )
            serials_at_stage = waiting_df.loc[
                waiting_df["waiting_at_stage"] == wip_stage_choice, ["battery_serial", "status"]
            ]
            st.dataframe(serials_at_stage, width="stretch", hide_index=True)
            st.download_button(
                f"Download batteries waiting at {wip_stage_choice} (Excel)",
                data=to_excel_bytes(serials_at_stage),
                file_name=f"waiting_at_{wip_stage_choice}.xlsx".replace(" ", "_"),
                key="dl_wip_stage",
            )
    elif stage_cols and not serial_col:
        st.info(
            "No serial/ID column recognized in this dataset, so the WIP report "
            "can't be built. Add a column like 'battery_serial' or "
            "'Module Serial Number' to enable it."
        )

    # ---- Date-wise pending/waiting search (needs folder history) ----
    if not history_df.empty:
        st.subheader("Search pending / waiting batteries by date")
        available_dates = sorted(history_df["date"].dt.date.unique())
        selected_date = st.date_input(
            "Pick a date",
            value=available_dates[-1],
            min_value=available_dates[0],
            max_value=available_dates[-1],
        )
        day_records = history_df[history_df["date"].dt.date == selected_date]
        pending_that_day = day_records[day_records["result"].isin(["Pending", "Fail"])]

        if pending_that_day.empty:
            st.info(f"No Pending/Fail records found on {selected_date}.")
        else:
            c1, c2 = st.columns(2)
            c1.metric("Pending/Fail records that day", len(pending_that_day))
            c2.metric("Stages affected", pending_that_day["stage"].nunique())
            display_cols = ["stage", "battery_serial", "result"]
            if pending_that_day["work_order"].notna().any():
                display_cols.append("work_order")
            st.dataframe(
                pending_that_day[display_cols].sort_values(["stage", "battery_serial"]),
                width="stretch",
                hide_index=True,
            )
            st.download_button(
                f"Download pending/waiting list for {selected_date} (Excel)",
                data=to_excel_bytes(pending_that_day[display_cols]),
                file_name=f"pending_{selected_date}.xlsx",
                key="dl_date_pending",
            )

        # Day-over-day new completions (avoids re-counting yesterday's passes as new)
        st.markdown("**New completions on this date (per stage)**")
        passed_today = day_records[day_records["result"] == "Pass"]
        prior = history_df[history_df["date"].dt.date < selected_date]
        already_passed_before = set(
            zip(prior[prior["result"] == "Pass"]["battery_serial"], prior[prior["result"] == "Pass"]["stage"])
        )
        newly_passed = passed_today[
            ~passed_today.apply(lambda r: (r["battery_serial"], r["stage"]) in already_passed_before, axis=1)
        ]
        if newly_passed.empty:
            st.write("No new passes on this date (or all had already passed earlier).")
        else:
            new_counts = newly_passed.groupby("stage").size().reset_index(name="newly_passed")
            st.dataframe(new_counts, width="stretch", hide_index=True)

    st.subheader("Battery numbers behind each stage")
    if serial_col:
        stage_choice = st.selectbox("Pick a stage", stage_cols)
        c1, c2, c3 = st.columns(3)
        colors = {"Pass": "success", "Fail": "error", "Pending": "warning"}
        for col, result in zip([c1, c2, c3], RESULTS):
            serials = df.loc[df[stage_choice] == result, serial_col].astype(str).tolist()
            with col:
                getattr(st, colors[result])(f"{result}: {len(serials)}")
                st.dataframe(
                    pd.DataFrame({"battery_serial": serials}),
                    width="stretch",
                    hide_index=True,
                )
    else:
        st.info("No serial/ID column found in the data, so individual battery numbers can't be listed.")

    if {"IB", "DOM", "Total"}.issubset(df.columns):
        st.subheader("IB / DOM / Total")
        totals = df[["IB", "DOM", "Total"]].sum()
        c1, c2, c3 = st.columns(3)
        c1.metric("IB", int(totals["IB"]))
        c2.metric("DOM", int(totals["DOM"]))
        c3.metric("Total", int(totals["Total"]))

    # ---- Battery search: search by serial OR by date, with download ----
    st.subheader("Search by battery pack serial")
    if serial_col:
        search_mode = st.radio(
            "Search by",
            ["Serial number", "Date"],
            horizontal=True,
            key="battery_search_mode",
        )

        # ---------------- Search by serial number ----------------
        if search_mode == "Serial number":
            search = st.text_input("Serial number")
            if search:
                match_df = df[df[serial_col].astype(str).str.contains(search, case=False, na=False)]
                st.dataframe(match_df, width="stretch", hide_index=True)
                if not match_df.empty:
                    st.download_button(
                        "Download this battery's current record (Excel)",
                        data=to_excel_bytes(match_df),
                        file_name=f"battery_{search}_current.xlsx",
                        key="dl_battery_current",
                    )

                if not history_df.empty:
                    hist_match = history_df[
                        history_df["battery_serial"].astype(str).str.contains(search, case=False, na=False)
                    ].sort_values("date")
                    if not hist_match.empty:
                        view_mode = st.radio(
                            "View history by", ["Date-wise", "Month-wise"], horizontal=True, key="hist_view_mode"
                        )
                        if view_mode == "Date-wise":
                            show = hist_match[["date", "stage", "result", "work_order"]]
                            st.dataframe(show, width="stretch", hide_index=True)
                        else:
                            hist_match["month"] = hist_match["date"].dt.to_period("M").astype(str)
                            show = (
                                hist_match.groupby(["month", "stage"])["result"]
                                .agg(lambda s: s.value_counts().idxmax())
                                .reset_index()
                            )
                            st.dataframe(show, width="stretch", hide_index=True)
                        st.download_button(
                            "Download full history for this battery (Excel)",
                            data=to_excel_bytes(hist_match),
                            file_name=f"battery_{search}_history.xlsx",
                            key="dl_battery_history",
                        )
                    else:
                        st.caption("No day-by-day history found for this serial in the uploaded folder.")

        # ---------------- Search by date ----------------
        else:
            if not history_df.empty:
                # Folder mode: full per-stage day-by-day records are available
                available_dates = sorted(history_df["date"].dt.date.unique())
                picked_date = st.date_input(
                    "Pick a date",
                    value=available_dates[-1],
                    min_value=available_dates[0],
                    max_value=available_dates[-1],
                    key="battery_search_date",
                )
                on_date = history_df[history_df["date"].dt.date == picked_date].sort_values(
                    ["stage", "battery_serial"]
                )
                if on_date.empty:
                    st.info(f"No records found on {picked_date}.")
                else:
                    display_cols = ["stage", "battery_serial", "result"]
                    if on_date["work_order"].notna().any():
                        display_cols.append("work_order")
                    st.dataframe(on_date[display_cols], width="stretch", hide_index=True)
                    st.download_button(
                        f"Download all battery info for {picked_date} (Excel)",
                        data=to_excel_bytes(on_date[display_cols]),
                        file_name=f"batteries_on_{picked_date}.xlsx",
                        key="dl_date_all_batteries",
                    )
            elif date_col:
                # Plain-file mode: use whatever date column the sheet has
                valid_dates = df[date_col].dropna()
                if valid_dates.empty:
                    st.info("No valid dates found in the data.")
                else:
                    available_dates = sorted(valid_dates.dt.date.unique())
                    picked_date = st.date_input(
                        "Pick a date",
                        value=available_dates[-1],
                        min_value=available_dates[0],
                        max_value=available_dates[-1],
                        key="battery_search_date_plain",
                    )
                    on_date = df[df[date_col].dt.date == picked_date]
                    if on_date.empty:
                        st.info(f"No records found on {picked_date}.")
                    else:
                        st.dataframe(on_date, width="stretch", hide_index=True)
                        st.download_button(
                            f"Download all battery info for {picked_date} (Excel)",
                            data=to_excel_bytes(on_date),
                            file_name=f"batteries_on_{picked_date}.xlsx",
                            key="dl_date_all_batteries_plain",
                        )
            else:
                st.info(
                    "No date information available in this dataset — upload a Stage/Date "
                    "folder zip, or a file with a Date column, to search by date."
                )


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        st.error("Something went wrong while rendering the dashboard:")
        st.exception(e)