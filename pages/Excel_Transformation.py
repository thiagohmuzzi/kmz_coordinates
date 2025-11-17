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
    """
    Input coordinates to the excel template provided. 
    Rows may contain either **NAD83 Geographic** (lat/long) or **NAD27 / UTM Zone 17N** (N/E). 
    Folder information and elevations are optional. \n
    **Note:** Always confirm samples of the converted coordinates with the [NRCan NTv2 website](https://webapp.csrs-scrs.nrcan-rncan.gc.ca/geod/tools-outils/ntv2.php).
    """
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
    "Upload Excel (columns may include: folder, feature_name, lat, long, N, E, elevation)",
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

def invalid_utm(E: pd.Series, N: pd.Series) -> pd.Series:
    bad = ~np.isfinite(E) | ~np.isfinite(N)
    bad |= ~E.between(100_000, 900_000) | ~N.between(4_000_000, 6_000_000)
    return bad

def tr_to_ll(grid_path: Path) -> Transformer:
    """
    NAD27 / UTM Zone 17N (m) -> NAD83 geographic (deg)
    """
    pipe = (
        f"+proj=pipeline "
        f"+step +inv +proj=utm +zone=17 +datum=NAD27 "
        f"+step +proj=hgridshift +grids={grid_path} "
        f"+step +proj=unitconvert +xy_in=rad +xy_out=deg"
    )
    return Transformer.from_pipeline(pipe)

def tr_to_utm(grid_path: Path) -> Transformer:
    """
    NAD83 geographic (deg) -> NAD27 / UTM Zone 17N (m)
    """
    pipe = (
        f"+proj=pipeline "
        f"+step +proj=unitconvert +xy_in=deg +xy_out=rad "
        f"+step +inv +proj=hgridshift +grids={grid_path} "
        f"+step +proj=utm +zone=17 +datum=NAD27"
    )
    return Transformer.from_pipeline(pipe)

# build transformers when grids exist
TR_TOR_TO_LL  = tr_to_ll(TOR_GRID)  if TOR_GRID.exists() else None
TR_ON_TO_LL   = tr_to_ll(ON_GRID)   if ON_GRID.exists()  else None
TR_TOR_TO_UTM = tr_to_utm(TOR_GRID) if TOR_GRID.exists() else None
TR_ON_TO_UTM  = tr_to_utm(ON_GRID)  if ON_GRID.exists()  else None

# -------- session to hold output --------
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

        col_lat = pick(df, ["lat", "latitude"])
        col_lon = pick(df, ["long", "lon", "longitude"])
        col_N   = pick(df, ["n","northing","utm_n","utm_northing","y"])
        col_E   = pick(df, ["e","easting","utm_e","utm_easting","x"])

        if col_lat is None: col_lat = "lat";  df[col_lat] = np.nan
        if col_lon is None: col_lon = "long"; df[col_lon] = np.nan
        if col_N is None:   col_N   = "N";    df[col_N]   = np.nan
        if col_E is None:   col_E   = "E";    df[col_E]   = np.nan

        if "grid_used" not in df.columns:
            df["grid_used"] = ""

        for c in (col_lat, col_lon, col_N, col_E):
            df[c] = pd.to_numeric(df[c], errors="coerce")

        has_latlon = df[col_lat].notna() & df[col_lon].notna()
        has_utm    = df[col_N].notna() & df[col_E].notna()

        # ---- Case 1: N/E -> lat/long ----
        if has_utm.any() and TR_TOR_TO_LL is not None:
            idx = df.index[has_utm]
            lon_t, lat_t = TR_TOR_TO_LL.transform(
                df.loc[idx, col_E].to_numpy(),
                df.loc[idx, col_N].to_numpy()
            )
            lon_t = pd.Series(lon_t, index=idx, dtype="float64")
            lat_t = pd.Series(lat_t, index=idx, dtype="float64")
            bad = invalid_ll(lon_t, lat_t)

            if bad.any() and TR_ON_TO_LL is not None:
                idx_bad = bad.index[bad]
                lon_o, lat_o = TR_ON_TO_LL.transform(
                    df.loc[idx_bad, col_E].to_numpy(),
                    df.loc[idx_bad, col_N].to_numpy()
                )
                lon_t.loc[idx_bad] = lon_o
                lat_t.loc[idx_bad] = lat_o

                df.loc[idx, "grid_used"] = "TO27CSv1.gsb"
                df.loc[idx_bad, "grid_used"] = "ON27CSv1.gsb"
            else:
                df.loc[idx, "grid_used"] = "TO27CSv1.gsb"

            good = ~invalid_ll(lon_t, lat_t)
            df.loc[good.index[good], col_lon] = lon_t[good]
            df.loc[good.index[good], col_lat] = lat_t[good]

        # ---- Case 2: lat/long -> N/E ----
        if has_latlon.any() and TR_TOR_TO_UTM is not None:
            idx = df.index[has_latlon]
            E_t, N_t = TR_TOR_TO_UTM.transform(
                df.loc[idx, col_lon].to_numpy(),
                df.loc[idx, col_lat].to_numpy()
            )
            E_t = pd.Series(E_t, index=idx, dtype="float64")
            N_t = pd.Series(N_t, index=idx, dtype="float64")
            bad = invalid_utm(E_t, N_t)

            if bad.any() and TR_ON_TO_UTM is not None:
                idx_bad = bad.index[bad]
                E_o, N_o = TR_ON_TO_UTM.transform(
                    df.loc[idx_bad, col_lon].to_numpy(),
                    df.loc[idx_bad, col_lat].to_numpy()
                )
                E_t.loc[idx_bad] = E_o
                N_t.loc[idx_bad] = N_o

                df.loc[idx, "grid_used"] = df.loc[idx, "grid_used"].replace("", "TO27CSv1.gsb")
                df.loc[idx_bad, "grid_used"] = "ON27CSv1.gsb"
            else:
                df.loc[idx, "grid_used"] = df.loc[idx, "grid_used"].replace("", "TO27CSv1.gsb")

            good = ~invalid_utm(E_t, N_t)
            df.loc[good.index[good], col_E] = E_t[good]
            df.loc[good.index[good], col_N] = N_t[good]

        df[col_lat] = df[col_lat].round(9)
        df[col_lon] = df[col_lon].round(9)
        df[col_E]   = df[col_E].round(3)
        df[col_N]   = df[col_N].round(3)

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
