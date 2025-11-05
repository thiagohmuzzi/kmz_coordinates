# Excel → KMZ (mixed inputs: NAD83 lat/long OR NAD27/UTM17N)
# Toronto grid default with Ontario fallback per-row; closed polylines by feature_name
import os
import zipfile
from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st
from pyproj import Transformer
from xml.sax.saxutils import escape


st.set_page_config(page_title="Excel → KMZ", page_icon="🗂️")
st.title("Excel to KMZ")

# --- Grids: resolve robustly (works whether file is at repo root or alongside this page) ---
THIS = Path(__file__).resolve()
ROOT = THIS.parents[1] if THIS.parent.name == "pages" else THIS.parent

TOR_GRID = next((p for p in [
    ROOT / "TO27CSv1.gsb",
    THIS.parent / "TO27CSv1.gsb",
    Path.cwd() / "TO27CSv1.gsb",
] if p.exists()), None)

ON_GRID = next((p for p in [
    ROOT / "ON27CSv1.gsb",
    THIS.parent / "ON27CSv1.gsb",
    Path.cwd() / "ON27CSv1.gsb",
] if p.exists()), None)

st.caption(
    "Input rows may contain either **NAD83 Geographic** (lat/long) or **NAD27 / UTM Zone 17N** (N/E). "
    "For UTM rows, the app applies **Toronto grid (TO27CSv1.gsb)** by default; if a point is outside that grid, "
    "**Ontario grid (ON27CSv1.gsb)** is used as fallback."
)

# Ensure PROJ can find the grids; disable network grids
for g in [TOR_GRID, ON_GRID]:
    if g:
        os.environ["PROJ_DATA"] = str(g.parent.resolve())
os.environ["PROJ_NETWORK"] = "OFF"

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
    # Optional: tighten to Ontario vicinity
    # bad |= ~lon_series.between(-96.0, -72.0) | ~lat_series.between(41.0, 57.0)
    return bad

def transformer_for(grid_path: Path) -> Transformer:
    # NAD27/UTM17N -> NAD83 geographic (degrees)
    pipe = (
        f"+proj=pipeline "
        f"+step +inv +proj=utm +zone=17 +datum=NAD27 "
        f"+step +proj=hgridshift +grids={grid_path} "
        f"+step +proj=unitconvert +xy_in=rad +xy_out=deg"
    )
    return Transformer.from_pipeline(pipe)

# ---------------------- main ----------------------
if not up:
    st.stop()

if st.button("Convert"):
    try:
        df0 = pd.read_excel(up)
        if df0.empty:
            st.error("No rows found."); st.stop()

        # Normalize headers to snake_case
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
            st.error("Missing required column: feature_name"); st.stop()

        # Keep relevant columns only
        keep = [c for c in [col_folder,col_name,col_lat,col_lon,col_n,col_e,col_z] if c]
        df = df[keep].copy()

        # Numeric coercion
        for c in [col_lat, col_lon, col_n, col_e, col_z]:
            if c and c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce")

        # Transformers
        tr_tor = transformer_for(TOR_GRID) if TOR_GRID else None
        tr_on  = transformer_for(ON_GRID)  if ON_GRID  else None

        # Prepare output frame
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

        # Case A: explicit NAD83 lat/long
        if col_lat and col_lon:
            m_latlon = df[col_lat].notna() & df[col_lon].notna()
            out.loc[m_latlon, "lat"] = df.loc[m_latlon, col_lat]
            out.loc[m_latlon, "lon"] = df.loc[m_latlon, col_lon]
            out.loc[m_latlon, "grid_used"] = "input_latlon"
            out.loc[m_latlon, "input_type"] = "latlon"

        # Case B: UTM N/E (NAD27/UTM17N) → NAD83 lat/lon
        if col_n and col_e and tr_tor:
            m_utm = df[col_n].notna() & df[col_e].notna()
            if m_utm.any():
                # Toronto first
                lon_t, lat_t = tr_tor.transform(
                    df.loc[m_utm, col_e].to_numpy(),
                    df.loc[m_utm, col_n].to_numpy()
                )
                lon_t = pd.Series(lon_t, index=df.loc[m_utm].index, dtype="float64")
                lat_t = pd.Series(lat_t, index=df.loc[m_utm].index, dtype="float64")

                bad = invalid_ll(lon_t, lat_t)

                # Fallback to Ontario for invalid rows
                if bad.any() and tr_on:
                    lon_o, lat_o = tr_on.transform(
                        df.loc[bad, col_e].to_numpy(),
                        df.loc[bad, col_n].to_numpy()
                    )
                    lon_o = pd.Series(lon_o, index=df.loc[bad].index, dtype="float64")
                    lat_o = pd.Series(lat_o, index=df.loc[bad].index, dtype="float64")

                    lon_t.loc[bad] = lon_o
                    lat_t.loc[bad] = lat_o

                    # mark grids
                    out.loc[m_utm & ~bad, "grid_used"] = "TO27CSv1.gsb"
                    out.loc[bad, "grid_used"] = "ON27CSv1.gsb"

                    # If some still invalid after fallback, mark them None
                    bad_after = invalid_ll(lon_t.loc[bad], lat_t.loc[bad])
                    if bad_after.any():
                        out.loc[bad_after.index[bad_after], "grid_used"] = "None"
                else:
                    out.loc[m_utm, "grid_used"] = "TO27CSv1.gsb"

                out.loc[m_utm, "lat"] = lat_t
                out.loc[m_utm, "lon"] = lon_t
                out.loc[m_utm, "input_type"] = "utm"

        # Remove rows that are still invalid
        bad_final = invalid_ll(out["lon"], out["lat"])
        if bad_final.any():
            st.warning(f"{bad_final.sum()} row(s) had invalid coordinates after fallback and were skipped.")
            out = out[~bad_final].copy()

        if out.empty:
            st.error("No valid coordinates were found after processing."); st.stop()

        # Round for display/validation
        out["lat"] = out["lat"].round(9)
        out["lon"] = out["lon"].round(9)
        if "N" in out.columns: out["N"] = out["N"].round(3)
        if "E" in out.columns: out["E"] = out["E"].round(3)

        # ---------------------- KML (for KMZ packaging) ----------------------
        def kml_header(doc_name):
            return ('<?xml version="1.0" encoding="UTF-8"?>\n'
                    '<kml xmlns="http://www.opengis.net/kml/2.2">\n'
                    f'  <Document><name>{escape(doc_name)}</name>\n')

        def kml_footer():
            return '  </Document>\n</kml>\n'

        def kml_folder(name):
            return f'    <Folder><name>{escape(name)}</name>\n'

        def kml_folder_end():
            return '    </Folder>\n'

        def pm_point(name, lon, lat, elev=None):
            coord = f"{lon:.9f},{lat:.9f}" + (f",{float(elev):.2f}" if elev is not None and not np.isnan(elev) else "")
            return ("      <Placemark>\n"
                    f"        <name>{escape(str(name))}</name>\n"
                    f"        <Point><coordinates>{coord}</coordinates></Point>\n"
                    "      </Placemark>\n")

        def pm_linestring(name, coords):
            # Close polyline by repeating first vertex
            if len(coords) >= 2 and (coords[0][0] != coords[-1][0] or coords[0][1] != coords[-1][1]):
                coords.append(coords[0])
            coord_txt = " ".join([(",".join([f"{c[0]:.9f}", f"{c[1]:.9f}"] + ([f"{float(c[2]):.2f}"] if len(c) > 2 and c[2] is not None and not np.isnan(c[2]) else []))) for c in coords])
            return ("      <Placemark>\n"
                    f"        <name>{escape(str(name))}</name>\n"
                    f"        <LineString><tessellate>1</tessellate><coordinates>{coord_txt}</coordinates></LineString>\n"
                    "      </Placemark>\n")

        base = Path(up.name).stem
        kml = BytesIO()
        kml.write(kml_header(f"{base} — NAD83 Geographic").encode("utf-8"))

        folder_series = out["folder"].fillna("").astype(str) if "folder" in out.columns else pd.Series([""]*len(out), index=out.index)
        elev_series = out["elevation"] if "elevation" in out.columns else pd.Series([np.nan]*len(out), index=out.index)

        # Group: folder → feature_name; build closed polylines for multi-vertex features
        for folder_name, g_folder in out.groupby(folder_series):
            if folder_name:
                kml.write(kml_folder(folder_name).encode("utf-8"))
            for feat, g_feat in g_folder.groupby(out["feature_name"]):
                g_feat = g_feat.sort_index()
                coords = [
                    (float(r["lon"]), float(r["lat"]), float(elev_series.loc[i]))
                    if pd.notna(elev_series.loc[i])
                    else (float(r["lon"]), float(r["lat"]))
                    for i, r in g_feat.iterrows()
                ]
                if len(coords) >= 2:
                    kml.write(pm_linestring(feat, coords.copy()).encode("utf-8"))
                else:
                    lon1, lat1 = coords[0][0], coords[0][1]
                    elev1 = coords[0][2] if len(coords[0]) > 2 else None
                    kml.write(pm_point(feat, lon1, lat1, elev1).encode("utf-8"))
            if folder_name:
                kml.write(kml_folder_end().encode("utf-8"))

        kml.write(kml_footer().encode("utf-8"))
        kml.seek(0)

        # KMZ packaging (KMZ only)
        kmz = BytesIO()
        with zipfile.ZipFile(kmz, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("doc.kml", kml.getvalue())
        kmz.seek(0)

        # ---------------------- Validation Excel ----------------------
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

        # UI outputs
        st.success("KMZ and Validation Excel created.")
        st.download_button(
            "Download KMZ",
            data=kmz.getvalue(),
            file_name=f"{base}_NAD83.kmz",
            mime="application/vnd.google-earth.kmz",
        )
        st.download_button(
            "Download Validation Excel",
            data=xbuf.getvalue(),
            file_name=f"{base}_Validation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.dataframe(out_valid, use_container_width=True)

    except Exception as e:
        st.error("Conversion failed.")
        st.exception(e)
