# Excel → KMZ
# Input rows may contain either:
#   - WGS84 Geographic: lat / long
#   - WGS84 UTM Zone 17N: N / E
#
# Output:
#   - KMZ with Google Earth-ready WGS84 geographic coordinates
#   - Validation Excel showing final lat/long used for plotting
#
# No NAD27, no NAD83, no NTv2 grids, no GTAA survey transformation.

import os
import zipfile
from io import BytesIO
from pathlib import Path
from xml.sax.saxutils import escape

import numpy as np
import pandas as pd
import streamlit as st
from pyproj import Transformer


st.set_page_config(page_title="Excel to KMZ", page_icon="🌎")
st.title("Excel to KMZ")

st.caption(
    """
    Input coordinates to the Excel template provided.  
    Rows may contain either **WGS84 Geographic** (`lat`, `long`) or **WGS84 UTM Zone 17N** (`N`, `E`).    
    Folder information is optional and can be used to nest features into separate Google Earth folders. Elevation is optional.  

    **Note:** Visually confirm the geographic placement of points in the new file.
    """
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

# ---------------------- Helpers ----------------------
def pick(df: pd.DataFrame, options) -> str | None:
    for o in options:
        if o in df.columns:
            return o
    return None


def invalid_ll(long_series: pd.Series, lat_series: pd.Series) -> pd.Series:
    """Invalid if NaN, inf, or outside global lat/long bounds."""
    bad = ~np.isfinite(long_series) | ~np.isfinite(lat_series)
    bad |= ~long_series.between(-180.0, 180.0)
    bad |= ~lat_series.between(-90.0, 90.0)
    return bad


def invalid_wgs84_utm17(e_series: pd.Series, n_series: pd.Series) -> pd.Series:
    """
    Broad sanity check for WGS84 UTM Zone 17N.
    Keeps wide bounds because this tool may be used outside Pearson.
    """
    bad = ~np.isfinite(e_series) | ~np.isfinite(n_series)
    bad |= ~e_series.between(100_000, 900_000)
    bad |= ~n_series.between(0, 10_000_000)
    return bad


def transformer_wgs84_utm17_to_wgs84_geo() -> Transformer:
    """
    WGS84 / UTM Zone 17N meters -> WGS84 Geographic degrees.
    EPSG:32617 -> EPSG:4326
    """
    return Transformer.from_crs("EPSG:32617", "EPSG:4326", always_xy=True)


def build_kmz_points(doc_name: str, rows: pd.DataFrame) -> bytes:
    """
    rows expected:
      folder, feature_name, long, lat, elevation(optional)

    KML coordinate order is always:
      longitude,latitude[,elevation]
    """

    def kml_header(name):
        return (
            '<?xml version="1.0" encoding="UTF-8"?>\n'
            '<kml xmlns="http://www.opengis.net/kml/2.2">\n'
            f'  <Document><name>{escape(name)}</name>\n'
        )

    def kml_footer():
        return "  </Document>\n</kml>\n"

    def kml_folder(name):
        return f'    <Folder><name>{escape(name)}</name>\n'

    def kml_folder_end():
        return "    </Folder>\n"

    def pm_point(name, long_val, lat_val, elev=None):
        coord = f"{long_val:.9f},{lat_val:.9f}"
        if elev is not None and np.isfinite(elev):
            coord += f",{float(elev):.2f}"

        return (
            "      <Placemark>\n"
            f"        <name>{escape(str(name))}</name>\n"
            f"        <Point><coordinates>{coord}</coordinates></Point>\n"
            "      </Placemark>\n"
        )

    kml = BytesIO()
    kml.write(kml_header(doc_name).encode("utf-8"))

    if "folder" in rows.columns:
        folders = rows["folder"].fillna("").astype(str)
    else:
        folders = pd.Series([""] * len(rows), index=rows.index)

    if "elevation" in rows.columns:
        elev = rows["elevation"]
    else:
        elev = pd.Series([np.nan] * len(rows), index=rows.index)

    for folder_name, group in rows.groupby(folders):
        in_folder = bool(folder_name)

        if in_folder:
            kml.write(kml_folder(folder_name).encode("utf-8"))

        for i, r in group.iterrows():
            kml.write(
                pm_point(
                    r["feature_name"],
                    float(r["long"]),
                    float(r["lat"]),
                    elev.get(i, np.nan),
                ).encode("utf-8")
            )

        if in_folder:
            kml.write(kml_folder_end().encode("utf-8"))

    kml.write(kml_footer().encode("utf-8"))
    kml.seek(0)

    kmz = BytesIO()
    with zipfile.ZipFile(kmz, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("doc.kml", kml.getvalue())

    kmz.seek(0)
    return kmz.getvalue()


# ---------------------- Session persistence ----------------------
for key in ("kmz_bytes", "validation_xlsx", "base_name"):
    if key not in st.session_state:
        st.session_state[key] = None

convert_clicked = st.button("Convert") if up else False

# ---------------------- Main conversion ----------------------
if convert_clicked and up:
    try:
        df0 = pd.read_excel(up)

        if df0.empty:
            st.error("No rows found.")
        else:
            # Normalize headers
            df = df0.copy()
            df.columns = [str(c).strip().lower().replace(" ", "_") for c in df.columns]

            col_folder = pick(df, ["folder", "group", "layer"])
            col_name = pick(df, ["feature_name", "name", "label", "id", "title"])
            col_lat = pick(df, ["lat", "latitude"])
            col_long = pick(df, ["long", "lon", "longitude"])
            col_n = pick(df, ["n", "northing", "utm_n", "utm_northing", "y"])
            col_e = pick(df, ["e", "easting", "utm_e", "utm_easting", "x"])
            col_z = pick(df, ["elevation", "elev", "z", "altitude", "height", "elevation_m"])

            if col_name is None:
                st.error("Missing required column: feature_name")
            else:
                keep = [c for c in [col_folder, col_name, col_lat, col_long, col_n, col_e, col_z] if c]
                df = df[keep].copy()

                # Numeric coercion
                for c in [col_lat, col_long, col_n, col_e, col_z]:
                    if c and c in df.columns:
                        df[c] = pd.to_numeric(df[c], errors="coerce")

                # Output frame
                out = pd.DataFrame(index=df.index)
                out["folder"] = df[col_folder] if col_folder and col_folder in df.columns else ""
                out["feature_name"] = df[col_name]
                out["elevation"] = df[col_z] if col_z and col_z in df.columns else np.nan

                out["lat"] = np.nan
                out["long"] = np.nan
                out["N"] = df[col_n] if col_n and col_n in df.columns else np.nan
                out["E"] = df[col_e] if col_e and col_e in df.columns else np.nan
                out["input_type"] = ""
                out["conversion_note"] = ""

                # A) WGS84 geographic input: pass through directly
                has_latlong = False
                if col_lat and col_long:
                    m_latlong = df[col_lat].notna() & df[col_long].notna()
                    has_latlong = m_latlong.any()

                    out.loc[m_latlong, "lat"] = df.loc[m_latlong, col_lat]
                    out.loc[m_latlong, "long"] = df.loc[m_latlong, col_long]
                    out.loc[m_latlong, "input_type"] = "wgs84_geographic"
                    out.loc[m_latlong, "conversion_note"] = "plotted_from_lat_long"

                # B) WGS84 UTM Zone 17N input: convert to WGS84 geographic
                if col_n and col_e:
                    m_utm = df[col_n].notna() & df[col_e].notna()

                    # If both lat/long and UTM are present on the same row, prefer lat/long.
                    if col_lat and col_long:
                        m_both = (
                            df[col_lat].notna()
                            & df[col_long].notna()
                            & df[col_n].notna()
                            & df[col_e].notna()
                        )
                        if m_both.any():
                            out.loc[m_both, "input_type"] = "wgs84_geographic_preferred"
                            out.loc[m_both, "conversion_note"] = "both_lat_long_and_utm_present_lat_long_used"

                        m_utm = m_utm & ~m_both

                    if m_utm.any():
                        bad_utm = invalid_wgs84_utm17(df.loc[m_utm, col_e], df.loc[m_utm, col_n])

                        if bad_utm.any():
                            st.warning(
                                f"{bad_utm.sum()} UTM row(s) failed basic WGS84 UTM Zone 17N range checks and were skipped."
                            )

                        good_utm_idx = bad_utm.index[~bad_utm]

                        if len(good_utm_idx) > 0:
                            tr = transformer_wgs84_utm17_to_wgs84_geo()
                            long_vals, lat_vals = tr.transform(
                                df.loc[good_utm_idx, col_e].to_numpy(),
                                df.loc[good_utm_idx, col_n].to_numpy(),
                            )

                            out.loc[good_utm_idx, "long"] = long_vals
                            out.loc[good_utm_idx, "lat"] = lat_vals
                            out.loc[good_utm_idx, "input_type"] = "wgs84_utm_zone_17n"
                            out.loc[good_utm_idx, "conversion_note"] = "converted_from_wgs84_utm_zone_17n"

                # Final coordinate validation
                bad_final = invalid_ll(out["long"], out["lat"])

                if bad_final.any():
                    st.warning(
                        f"{bad_final.sum()} row(s) had invalid or missing final WGS84 geographic coordinates and were skipped."
                    )
                    out = out[~bad_final].copy()

                if out.empty:
                    st.error("No valid coordinates were found after processing.")
                else:
                    # Round validation output
                    out["lat"] = out["lat"].round(9)
                    out["long"] = out["long"].round(9)
                    out["N"] = pd.to_numeric(out["N"], errors="coerce").round(3)
                    out["E"] = pd.to_numeric(out["E"], errors="coerce").round(3)

                    base = Path(up.name).stem

                    # KMZ: Google Earth WGS84 geographic points only
                    kmz_bytes = build_kmz_points(
                        f"{base} — WGS84 Geographic",
                        out[["folder", "feature_name", "long", "lat", "elevation"]],
                    )

                    # Validation Excel
                    valid_cols = [
                        "folder",
                        "feature_name",
                        "lat",
                        "long",
                        "N",
                        "E",
                        "elevation",
                        "input_type",
                        "conversion_note",
                    ]

                    for c in valid_cols:
                        if c not in out.columns:
                            out[c] = np.nan if c not in ["folder", "feature_name", "input_type", "conversion_note"] else ""

                    out_valid = out[valid_cols].copy()

                    xbuf = BytesIO()
                    with pd.ExcelWriter(xbuf, engine="openpyxl") as xw:
                        readme = pd.DataFrame(
                            {
                                "Validation": [
                                    "All exported KMZ coordinates are WGS84 Geographic (lat/long).",
                                    "Rows with lat/long are plotted directly.",
                                    "Rows with N/E are interpreted as WGS84 / UTM Zone 17N (Google Earth may show this as Zone 17T) and converted to WGS84 Geographic.",
                                    "If both lat/long and N/E are provided on the same row, lat/long is used.",
                                    "No NAD27, NAD83, NTv2 grid, or GTAA survey transformation is applied in this tool.",
                                ]
                            }
                        )
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

# ---------------------- Downloads ----------------------
if st.session_state.kmz_bytes and st.session_state.validation_xlsx:
    st.success("KMZ and Validation Excel are ready.")

    st.download_button(
        "Download KMZ",
        data=st.session_state.kmz_bytes,
        file_name=f"{st.session_state.base_name}_WGS84.kmz",
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

    def _queue_refresh():
        for k in ("kmz_bytes", "validation_xlsx", "base_name"):
            st.session_state.pop(k, None)
        st.session_state["_do_rerun"] = True

    st.button("Refresh / New conversion", on_click=_queue_refresh, type="secondary")

if st.session_state.get("_do_rerun"):
    st.session_state.pop("_do_rerun", None)
    st.rerun()
