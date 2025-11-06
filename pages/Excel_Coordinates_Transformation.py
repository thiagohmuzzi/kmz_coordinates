# Excel Coordinate Transformation
# Input rows may contain either NAD83 Geographic (lat/long) or NAD27 / UTM Zone 17N (N/E).
# For each row:
#   - If lat/long present -> compute N/E (NAD27 / UTM17).
#   - If N/E present -> compute lat/long (NAD83 geographic).
# Toronto grid (TO27CSv1.gsb) is used by default; if result invalid, fallback to ON27CSv1.gsb.
# Output: the SAME sheet (all original columns preserved), plus filled lat/long & N/E, and grid_used.

import os
from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st
from pyproj import Transformer

st.set_page_config(page_title="Excel Coordinates Transformation", page_icon="📄")
st.title("Excel Coordinate Transformation")

# -------- repo root & grids --------
THIS = Path(__file__).resolve()
ROOT = THIS.parents[1] if THIS.parent.name == "pages" else THIS.parent
TOR_GRID = ROOT / "TO27CSv1.gsb"
ON_GRID  = ROOT / "ON27CSv1.gsb"

# let PROJ find our local grids; disable network
os.environ["PROJ_DATA"] = str(ROOT.resolve())
os.environ["PROJ_NETWORK"] = "OFF"

st.caption(
    "Input coordinates to the excel template provided. Rows may contain either **NAD83 Geographic** (lat/long) or **NAD27 / UTM Zone 17N** (N/E). "
    "Folder information is optional and can be provided if desired to have the features nested in separate folders for Google Earth. "
    "Coordinate transformation applies the Toronto grid (TO27CSv1.gsb) by default; Ontario grid (ON27CSv1.gsb) is used as fallback if coordinates fall outside the default grid coverage."
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
    # NAD27/UTM17N (m) -> NAD83 geographic (deg)
    pipe = (
        f"+proj=pipeline "
        f"+step +inv +proj=utm +zone=17 +datum=NAD27 "
        f"+step +proj=hgridshift +grids={grid_path} "
        f"+step +proj=unitconvert +xy_in=rad +xy_out=deg"
    )
    return Transformer.from_pipeline(pipe)

def tr_to_utm(grid_path: Path) -> Transformer:
    # NAD83 geographic (deg) -> NAD27/UTM17N (m)
    pipe = (
        f"+proj=pipeline "
        f"+step +proj=unitconvert +xy_in=deg +xy_out=rad "
        f"+step +proj=hgridshift +grids={grid_path} +inv "
        f"+step +proj=utm +zone=17 +datum=NAD27"
    )
    return Transformer.from_pipeline(pipe)

# build transformers when grids exist
TR_TOR_TO_LL  = tr_to_ll(TOR_GRID)  if TOR_GRID.exists() else None
TR_ON_TO_LL   = tr_to_ll(ON_GRID)   if ON_GRID.exists()  else None
TR_TOR_TO_UTM = tr_to_utm(TOR_GRID) if TOR_GRID.exists() else None
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
        if col_e is None:   col_e    = "E";    df[col_e]    = np.nan

        # Keep a grid_used column
        if "grid_used" not in df.columns:
            df["grid_used"] = ""

        # numeric coercion on coord columns
        for c in (col_lat, col_lon, col_n, col_e):
            df[c] = pd.to_numeric(df[c], errors="coerce")

        # Row masks
        has_latlon = df[col_lat].notna() & df[col_lon].notna()
        has_utm    = df[col_n].notna() & df[col_e].notna()

        # ---- Case 1: N/E -> lat/long (Toronto first, Ontario fallback) ----
        if has_utm.any() and TR_TOR_TO_LL is not None:
            idx = df.index[has_utm]
            lon_t, lat_t = TR_TOR_TO_LL.transform(df.loc[idx, col_e].to_numpy(),
                                                  df.loc[idx, col_n].to_numpy())
            lon_t = pd.Series(lon_t, index=idx, dtype="float64")
            lat_t = pd.Series(lat_t, index=idx, dtype="float64")
            bad = invalid_ll(lon_t, lat_t)
            # fallback for bad rows
            if bad.any() and TR_ON_TO_LL is not None:
                idx_bad = bad.index[bad]
                lon_o, lat_o = TR_ON_TO_LL.transform(df.loc[idx_bad, col_e].to_numpy(),
                                                     df.loc[idx_bad, col_n].to_numpy())
                lon_t.loc[idx_bad] = lon_o
                lat_t.loc[idx_bad] = lat_o
                # mark grids
                df.loc[idx, "grid_used"] = "TO27CSv1.gsb"
                df.loc[idx_bad, "grid_used"] = "ON27CSv1.gsb"
                # still-bad rows remain unmodified; caller keeps whatever was there
            else:
                df.loc[idx, "grid_used"] = "TO27CSv1.gsb"

            # write NAD83 lat/lon
            good = ~invalid_ll(lon_t, lat_t)
            df.loc[good.index[good], col_lon] = lon_t[good]
            df.loc[good.index[good], col_lat] = lat_t[good]

        # ---- Case 2: lat/long -> N/E (Toronto first, Ontario fallback) ----
        if has_latlon.any() and TR_TOR_TO_UTM is not None:
            idx = df.index[has_latlon]
            e_t, n_t = TR_TOR_TO_UTM.transform(df.loc[idx, col_lon].to_numpy(),
                                               df.loc[idx, col_lat].to_numpy())
            e_t = pd.Series(e_t, index=idx, dtype="float64")
            n_t = pd.Series(n_t, index=idx, dtype="float64")
            bad = invalid_utm(e_t, n_t)
            if bad.any() and TR_ON_TO_UTM is not None:
                idx_bad = bad.index[bad]
                e_o, n_o = TR_ON_TO_UTM.transform(df.loc[idx_bad, col_lon].to_numpy(),
                                                  df.loc[idx_bad, col_lat].to_numpy())
                e_t.loc[idx_bad] = e_o
                n_t.loc[idx_bad] = n_o
                # mark grids
                # (don't overwrite grid_used if it was already set by the other branch)
                df.loc[idx, "grid_used"] = df.loc[idx, "grid_used"].replace("", "TO27CSv1.gsb")
                df.loc[idx_bad, "grid_used"] = "ON27CSv1.gsb"
            else:
                df.loc[idx, "grid_used"] = df.loc[idx, "grid_used"].replace("", "TO27CSv1.gsb")

            good = ~invalid_utm(e_t, n_t)
            df.loc[good.index[good], col_e] = e_t[good]
            df.loc[good.index[good], col_n] = n_t[good]

        # round for neatness
        df[col_lat] = df[col_lat].round(9)
        df[col_lon] = df[col_lon].round(9)
        df[col_e]   = df[col_e].round(3)
        df[col_n]   = df[col_n].round(3)

        # pack Excel (same sheet name as input’s first sheet)
        out = BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as xw:
            df.to_excel(xw, index=False, sheet_name="Transformed")
        out.seek(0)

        st.session_state.out_xlsx_bytes = out.getvalue()
        st.session_state.out_filename = f"{Path(up.name).stem}_coords_filled.xlsx"

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
