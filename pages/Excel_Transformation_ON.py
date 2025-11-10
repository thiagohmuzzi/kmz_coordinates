# Excel Coordinate Transformation — Ontario grid only
# Input rows may contain either NAD83 Geographic (lat/long) or NAD27 / UTM Zone 17N (N/E).
# For each row:
#   - If lat/long present -> compute N/E (NAD27 / UTM17) using ON27CSv1.gsb.
#   - If N/E present -> compute lat/long (NAD83 geographic) using ON27CSv1.gsb.
# Output: the SAME sheet (all original columns preserved), with both coordinate pairs filled and grid_used.

import os
from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st
from pyproj import Transformer

# --- Toggle this page on/off easily ---
PAGE_ENABLED = False  # Change between True/False to enable/disable this page
if not PAGE_ENABLED:
    st.stop()

st.set_page_config(page_title="Excel Coordinate Transformation (Ontario only)", page_icon="🧭")
st.title("Excel Coordinate Transformation")

# -------- repo root & Ontario grid only --------
THIS = Path(__file__).resolve()
ROOT = THIS.parents[1] if THIS.parent.name == "pages" else THIS.parent
ON_GRID = ROOT / "ON27CSv1.gsb"  # Ontario NTv2 grid

# Let PROJ find our local grids; disable network
os.environ["PROJ_DATA"] = str(ROOT.resolve())
os.environ["PROJ_NETWORK"] = "OFF"

st.caption(
    "Provide either **NAD83 Geographic** (lat/long) or **NAD27 / UTM Zone 17N** (N/E). "
    "All transformations use the **Ontario NTv2 grid (ON27CSv1.gsb)**."
)

# -------- template download --------
tpl_cols = ["folder", "feature_name", "lat", "long", "N", "E", "elevation"]
tpl_df = pd.DataFrame(columns=tpl_cols)
tpl_buf = BytesIO()
with pd.ExcelWriter(tpl_buf, engine="openpyxl") as xw:
    tpl_df.to_excel(xw, index=False, sheet_name="Template")
tpl_buf.seek(0)
st.download_button(
    "Download Excel template",
    data=tpl_buf.getvalue(),
    file_name="Excel_Coordinate_Transformation_template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# -------- upload --------
up = st.file_uploader(
    "Upload Excel (columns may include: folder, feature_name, lat, long, N, E, elevation — others are preserved)",
    type=["xlsx", "xls"],
)

# -------- helpers --------
def pick(df: pd.DataFrame, options) -> str | None:
    for o in options:
        if o in df.columns:
            return o
    return None

def invalid_ll(lon: pd.Series, lat: pd.Series) -> pd.Series:
    bad = ~np.isfinite(lon) | ~np.isfinite(lat)
    bad |= ~lon.between(-180.0, 180.0) | ~lat.between(-90.0, 90.0)
    return bad

def invalid_utm(e: pd.Series, n: pd.Series) -> pd.Series:
    bad = ~np.isfinite(e) | ~np.isfinite(n)
    # loose UTM sanity (zone 17N plausible bounds; keep wide)
    bad |= ~e.between(100_000, 900_000) | ~n.between(4_000_000, 6_000_000)
    return bad

def tr_to_ll(grid_path: Path) -> Transformer:
    # NAD27/UTM17N (m) -> NAD83 geographic (deg) using ON27CSv1.gsb
    pipe = (
        f"+proj=pipeline "
        f"+step +inv +proj=utm +zone=17 +datum=NAD27 "
        f"+step +proj=hgridshift +grids={grid_path} "
        f"+step +proj=unitconvert +xy_in=rad +xy_out=deg"
    )
    return Transformer.from_pipeline(pipe)

def tr_to_utm(grid_path: Path) -> Transformer:
    # NAD83 geographic (deg) -> NAD27/UTM17N (m) using ON27CSv1.gsb
    pipe = (
        f"+proj=pipeline "
        f"+step +proj=unitconvert +xy_in=deg +xy_out=rad "
        f"+step +proj=hgridshift +grids={grid_path} +inv "
        f"+step +proj=utm +zone=17 +datum=NAD27"
    )
    return Transformer.from_pipeline(pipe)

# Build transformers (Ontario only)
TR_ON_TO_LL   = tr_to_ll(ON_GRID)   if ON_GRID.exists()  else None
TR_ON_TO_UTM  = tr_to_utm(ON_GRID)  if ON_GRID.exists()  else None

# -------- session to hold output so multiple downloads work if needed --------
for k in ("out_xlsx_bytes", "out_filename"):
    if k not in st.session_state:
        st.session_state[k] = None

if up and st.button("Transform"):
    df0 = pd.read_excel(up)
    if df0.empty:
        st.error("No rows found.")
    else:
        df = df0.copy()
        df.columns = [str(c).strip().lower().replace(" ", "_") for c in df.columns]

        col_lat  = pick(df, ["lat", "latitude"])
        col_lon  = pick(df, ["long", "lon", "longitude"])
        col_n    = pick(df, ["n","northing","utm_n","utm_northing","y"])
        col_e    = pick(df, ["e","easting","utm_e","utm_easting","x"])

        # Ensure output columns exist in the same dataframe (do not drop other columns)
        if col_lat is None: col_lat = "lat";  df[col_lat] = np.nan
        if col_lon is None: col_lon = "long"; df[col_lon] = np.nan
        if col_n is None:   col_n   = "N";    df[col_n]   = np.nan
        if col_e is None:   col_e   = "E";    df[col_e]   = np.nan

        # Keep/ensure grid_used column
        if "grid_used" not in df.columns:
            df["grid_used"] = ""

        # numeric coercion on coord columns
        for c in (col_lat, col_lon, col_n, col_e):
            df[c] = pd.to_numeric(df[c], errors="coerce")

        # Row masks
        has_latlon = df[col_lat].notna() & df[col_lon].notna()
        has_utm    = df[col_n].notna() & df[col_e].notna()

        # ---- Case 1: N/E -> lat/long (Ontario grid only) ----
        if has_utm.any() and TR_ON_TO_LL is not None:
            idx = df.index[has_utm]
            lon_o, lat_o = TR_ON_TO_LL.transform(df.loc[idx, col_e].to_numpy(),
                                                 df.loc[idx, col_n].to_numpy())
            lon_o = pd.Series(lon_o, index=idx, dtype="float64")
            lat_o = pd.Series(lat_o, index=idx, dtype="float64")

            good = ~invalid_ll(lon_o, lat_o)
            df.loc[good.index[good], col_lon] = lon_o[good]
            df.loc[good.index[good], col_lat] = lat_o[good]
            df.loc[idx, "grid_used"] = "ON27CSv1.gsb"

        # ---- Case 2: lat/long -> N/E (Ontario grid only) ----
        if has_latlon.any() and TR_ON_TO_UTM is not None:
            idx = df.index[has_latlon]
            e_o, n_o = TR_ON_TO_UTM.transform(df.loc[idx, col_lon].to_numpy(),
                                              df.loc[idx, col_lat].to_numpy())
            e_o = pd.Series(e_o, index=idx, dtype="float64")
            n_o = pd.Series(n_o, index=idx, dtype="float64")

            good = ~invalid_utm(e_o, n_o)
            df.loc[good.index[good], col_e] = e_o[good]
            df.loc[good.index[good], col_n] = n_o[good]
            # preserve any previous tag, otherwise set to Ontario
            df.loc[idx, "grid_used"] = df.loc[idx, "grid_used"].replace("", "ON27CSv1.gsb")

        # round for neatness
        df[col_lat] = df[col_lat].round(9)
        df[col_lon] = df[col_lon].round(9)
        df[col_e]   = df[col_e].round(3)
        df[col_n]   = df[col_n].round(3)

        # pack Excel (same sheet)
        out = BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as xw:
            df.to_excel(xw, index=False, sheet_name="Transformed")
        out.seek(0)

        st.session_state.out_xlsx_bytes = out.getvalue()
        st.session_state.out_filename = f"{Path(up.name).stem}_coords_filled_ON27.xlsx"

# -------- download --------
if st.session_state.out_xlsx_bytes:
    st.success("Transformed Excel is ready.")
    st.download_button(
        "Download transformed Excel",
        data=st.session_state.out_xlsx_bytes,
        file_name=st.session_state.out_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_out",
    )
