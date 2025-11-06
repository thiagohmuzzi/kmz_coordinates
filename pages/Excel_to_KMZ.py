# Excel → KMZ (NAD83 lat/long OR NAD27/UTM17N)
# Toronto grid default with Ontario fallback per-row; POINTS ONLY; KMZ + Validation Excel; persistent downloads
import os
import zipfile
from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st
from pyproj import Transformer
from xml.sax.saxutils import escape

st.set_page_config(page_title="Excel → KMZ", page_icon="🌎")
st.title("Excel to KMZ")

# ---------- Resolve grids in REPO ROOT ----------
THIS = Path(__file__).resolve()
ROOT = THIS.parents[1] if THIS.parent.name == "pages" else THIS.parent
TOR_GRID = ROOT / "TO27CSv1.gsb"
ON_GRID  = ROOT / "ON27CSv1.gsb"

# Make grids discoverable and disable network
os.environ["PROJ_DATA"] = str(ROOT.resolve())
os.environ["PROJ_NETWORK"] = "OFF"

st.caption(
    "Rows may contain either **NAD83 Geographic** (lat/long) or **NAD27 / UTM Zone 17N** (N/E). "
    "UTM rows try **Toronto grid (TO27CSv1.gsb)** first; if outside coverage, they fall back to **ON27CSv1.gsb**."
)

# ---------------------- Template download ----------------------
tpl_cols = ["folder", "feature_name", "lat", "long", "N", "E", "elevation"]
tpl_df = pd.DataFrame(columns=tpl_cols)
tpl_buf = BytesIO()
with pd.ExcelWriter(tpl_buf, engine="openpyxl") as xw:
    tpl_df.to_excel(xw, index=False, sheet_name="Template")
tpl_buf.seek(0)
st.download_button(
    "Download Excel template",
    data=tpl_buf.getvalue(),
    file_name="Excel_to_KMZ_template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# ---------------------- File upload ----------------------
up = st.file_uploader(
    "Upload Excel (columns: folder, feature_name, lat, long, N, E, elevation)",
    type=["xlsx", "xls"],
)

# ---------- helpers ----------
def pick(df: pd.DataFrame, options) -> str | None:
    for o in options:
        if o in df.columns:
            return o
    return None

def invalid_ll(lon_series: pd.Series, lat_series: pd.Series) -> pd.Series:
    """Invalid if NaN, ±inf, or out of global lat/lon bounds."""
    bad = ~np.isfinite(lon_series) | ~np.isfinite(lat_series)
    bad |= ~lon_series.between(-180.0, 180.0) | ~lat_series.between(-90.0, 90.0)
    return bad

def transformer_for_nad27utm17_to_nad83_ll(grid_path: Path) -> Transformer:
    """
    Input: NAD27 / UTM Zone 17N (meters)
    Output: NAD83 geographic (degrees)
    Steps:
      1) inverse UTM17 (to NAD27 geographic, radians)
      2) apply NTv2 grid shift (NAD27->NAD83), radians
      3) convert to degrees
    """
    pipe = (
        f"+proj=pipeline "
        f"+step +inv +proj=utm +zone=17 +datum=NAD27 "
        f"+step +proj=hgridshift +grids={grid_path} "
        f"+step +proj=unitconvert +xy_in=rad +xy_out=deg"
    )
    return Transformer.from_pipeline(pipe)

def build_kmz_points(doc_name: str, rows: pd.DataFrame) -> bytes:
    """
    rows expected: folder, feature_name, lon, lat, elevation(optional)
    Each row -> Point Placemark. Blank folder => placed at Document level.
    """
    def kml_header(name):
        return ('<?xml version="1.0" encoding="UTF-8"?>\n'
                '<kml xmlns="http://www.opengis.net/kml/2.2">\n'
                f'  <Document><name>{escape(name)}</name>\n')
    def kml_footer(): return '  </Document>\n</kml>\n'
    def kml_folder(name): return f'    <Folder><name>{escape(name)}</name>\n'
    def kml_folder_end(): return '    </Folder>\n'
    def pm_point(name, lon, lat, elev=None):
        coord = f"{lon:.9f},{lat:.9f}" + (f",{float(elev):.2f}" if elev is not None and not np.isnan(elev) else "")
        return ("      <Placemark>\n"
                f"        <name>{escape(str(name))}</name>\n"
                f"        <Point><coordinates>{coord}</coordinates></Point>\n"
                "      </Placemark>\n")

    kml = BytesIO()
    kml.write(kml_header(doc_name).encode("utf-8"))

    folders = rows["folder"].fillna("").astype(str) if "folder" in rows.columns else pd.Series([""]*len(rows), index=rows.index)
    elev = rows["elevation"] if "elevation" in rows.columns else pd.Series([np.nan]*len(rows), index=rows.index)

    for folder_name, g in rows.groupby(folders):
        in_folder = bool(folder_name)
        if in_folder:
            kml.write(kml_folder(folder_name).encode("utf-8"))
        for i, r in g.iterrows():
            kml.write(pm_point(r["feature_name"], float(r["lon"]), float(r["lat"]), elev.get(i, np.nan)).encode("utf-8"))
        if in_folder:
            kml.write(kml_folder_end().encode("utf-8"))

    kml.write(kml_footer().encode("utf-8"))
    kml.seek(0)

    kmz = BytesIO()
    with zipfile.ZipFile(kmz, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("doc.kml", kml.getvalue())
    kmz.seek(0)
    return kmz.getvalue()

# ---------------------- session persistence ----------------------
for key in ("kmz_bytes", "validation_xlsx", "base_name"):
    if key not in st.session_state:
        st.session_state[key] = None

convert_clicked = st.button("Convert") if up else False

if convert_clicked and up:
    try:
        df0 = pd.read_excel(up)
        if df0.empty:
            st.error("No rows found.")
        else:
            # Normalize headers
            df = df0.copy()
            df.columns = [str(c).strip().lower().replace(" ", "_") for c in df.columns]

            col_folder = pick(df, ["folder","group","layer"])
            col_name   = pick(df, ["feature_name","name","label","id","title"])
            col_lat    = pick(df, ["lat","latitude"])
            col_lon    = pick(df, ["long","lon","longitude"])
            col_n      = pick(df, ["n","northing","utm_n","utm_northing","y"])
            col_e      = pick(df, ["e","easting","utm_e","utm_easting","x"])
            col_z      = pick(df, ["elevation","elev","z","altitude","height","elevation_m"])

            if col_name is None:
                st.error("Missing required column: feature_name")
            else:
                keep = [c for c in [col_folder,col_name,col_lat,col_lon,col_n,col_e,col_z] if c]
                df = df[keep].copy()

                # Numeric coercion
                for c in [col_lat, col_lon, col_n, col_e, col_z]:
                    if c and c in df.columns:
                        df[c] = pd.to_numeric(df[c], errors="coerce")

                # Transformers (build only if grids exist)
                tr_tor = transformer_for_nad27utm17_to_nad83_ll(TOR_GRID) if TOR_GRID.exists() else None
                tr_on  = transformer_for_nad27utm17_to_nad83_ll(ON_GRID)  if ON_GRID.exists()  else None

                # Output frame
                out = pd.DataFrame(index=df.index)
                out["folder"] = df[col_folder] if col_folder in df.columns else ""
                out["feature_name"] = df[col_name]
                out["elevation"] = df[col_z] if col_z in df.columns else np.nan
                out["N"] = df[col_n] if col_n in df.columns else np.nan
                out["E"] = df[col_e] if col_e in df.columns else np.nan
                out["lat"] = np.nan
                out["lon"] = np.nan
                out["grid_used"] = ""
                out["input_type"] = ""

                # A) explicit NAD83 lat/long
                if col_lat and col_lon:
                    m_latlon = df[col_lat].notna() & df[col_lon].notna()
                    out.loc[m_latlon, "lat"] = df.loc[m_latlon, col_lat]
                    out.loc[m_latlon, "lon"] = df.loc[m_latlon, col_lon]
                    out.loc[m_latlon, "grid_used"] = "input_latlon"
                    out.loc[m_latlon, "input_type"] = "latlon"

                # B) NAD27/UTM17 → NAD83 lat/lon with Toronto-first & Ontario fallback
                if col_n and col_e and tr_tor:
                    m_utm = df[col_n].notna() & df[col_e].notna()
                    if m_utm.any():
                        lon_t, lat_t = tr_tor.transform(
                            df.loc[m_utm, col_e].to_numpy(),
                            df.loc[m_utm, col_n].to_numpy()
                        )
                        lon_t = pd.Series(lon_t, index=df.loc[m_utm].index, dtype="float64")
                        lat_t = pd.Series(lat_t, index=df.loc[m_utm].index, dtype="float64")

                        bad = invalid_ll(lon_t, lat_t)

                        if bad.any() and tr_on:
                            lon_o, lat_o = tr_on.transform(
                                df.loc[bad, col_e].to_numpy(),
                                df.loc[bad, col_n].to_numpy()
                            )
                            lon_o = pd.Series(lon_o, index=df.loc[bad].index, dtype="float64")
                            lat_o = pd.Series(lat_o, index=df.loc[bad].index, dtype="float64")
                            lon_t.loc[bad] = lon_o
                            lat_t.loc[bad] = lat_o

                            bad_after = invalid_ll(lon_t.loc[bad], lat_t.loc[bad])
                            out.loc[m_utm, "grid_used"] = "TO27CSv1.gsb"
                            out.loc[bad, "grid_used"] = "ON27CSv1.gsb"
                            if bad_after.any():
                                out.loc[bad_after.index[bad_after], "grid_used"] = "None"
                        else:
                            out.loc[m_utm, "grid_used"] = "TO27CSv1.gsb"

                        out.loc[m_utm, "lat"] = lat_t
                        out.loc[m_utm, "lon"] = lon_t
                        out.loc[m_utm, "input_type"] = "utm"

                # Remove any invalid rows
                bad_final = invalid_ll(out["lon"], out["lat"])
                if bad_final.any():
                    st.warning(f"{bad_final.sum()} row(s) had invalid coordinates after fallback and were skipped.")
                    out = out[~bad_final].copy()

                if out.empty:
                    st.error("No valid coordinates were found after processing.")
                else:
                    # Round for validation
                    out["lat"] = out["lat"].round(9)
                    out["lon"] = out["lon"].round(9)
                    if "N" in out.columns: out["N"] = out["N"].round(3)
                    if "E" in out.columns: out["E"] = out["E"].round(3)

                    base = Path(up.name).stem

                    # KMZ (POINTS ONLY)
                    kmz_bytes = build_kmz_points(f"{base} — NAD83 Geographic", out[["folder","feature_name","lon","lat","elevation"]])

                    # Validation Excel
                    valid_cols = ["folder","feature_name","lat","lon","N","E","elevation","grid_used","input_type"]
                    for c in valid_cols:
                        if c not in out.columns:
                            out[c] = np.nan if c not in ["folder","feature_name","grid_used","input_type"] else ""
                    out_valid = out[valid_cols].copy()
                    xbuf = BytesIO()
                    with pd.ExcelWriter(xbuf, engine="openpyxl") as xw:
                        readme = pd.DataFrame({
                            "Validation": [
                                "All exported coordinates are NAD83 Geographic (lat/lon).",
                                "UTM inputs assumed NAD27 / UTM Zone 17N; Toronto grid default with Ontario fallback per-row.",
                            ]
                        })
                        readme.to_excel(xw, index=False, sheet_name="README")
                        out_valid.to_excel(xw, index=False, sheet_name="Validation")
                    xbuf.seek(0)

                    # Persist in session so both downloads remain available
                    st.session_state.kmz_bytes = kmz_bytes
                    st.session_state.validation_xlsx = xbuf.getvalue()
                    st.session_state.base_name = base

    except Exception as e:
        st.error("Conversion failed.")
        st.exception(e)

# ---------------------- Downloads (persisted) ----------------------
if st.session_state.kmz_bytes and st.session_state.validation_xlsx:
    st.success("KMZ and Validation Excel are ready.")
    st.download_button(
        "Download KMZ",
        data=st.session_state.kmz_bytes,
        file_name=f"{st.session_state.base_name}_NAD83.kmz",
        mime="application/vnd.google-earth.kmz",
        key="dl_kmz",
    )
    st.download_button(
        "Download Validation Excel",
        data=st.session_state.validation_xlsx,
        file_name=f"{st.session_state.base_name}_Validation.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_xlsx",
    )

    # Queue a rerun instead of calling st.rerun() inside the callback
    def _queue_refresh():
        for k in ("kmz_bytes", "validation_xlsx", "base_name"):
            st.session_state.pop(k, None)
        st.session_state["_do_rerun"] = True

    st.button("Refresh / New conversion", on_click=_queue_refresh, type="secondary")

# Perform the rerun outside of any callback (prevents the yellow warning)
if st.session_state.get("_do_rerun"):
    st.session_state.pop("_do_rerun", None)
    st.rerun()

