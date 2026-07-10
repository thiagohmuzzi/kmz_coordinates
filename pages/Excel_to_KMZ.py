# Excel → KMZ
# Input rows may contain either:
#   - WGS84 Geographic: lat / long
#   - WGS84 UTM Zone 17N: N / E
#
# Template displays "WGS84 UTM 17T" because Google Earth may show the latitude band as 17T.
# For pyproj, conversion is handled as WGS84 / UTM Zone 17N = EPSG:32617.
#
# No NAD27, no NAD83, no NTv2 grids, no GTAA survey transformation.

import zipfile
from io import BytesIO
from pathlib import Path
from xml.sax.saxutils import escape

import numpy as np
import pandas as pd
import streamlit as st
from pyproj import Transformer

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill


# -------------------------------------------------------------------
# Streamlit page
# -------------------------------------------------------------------
st.set_page_config(page_title="Excel to KMZ", page_icon="🌎")
st.title("Excel to KMZ")

st.caption(
    """
    Input coordinates to the Excel template provided.  
    Rows may contain either **WGS84 Geographic** (`lat`, `long`) or **WGS84 UTM 17T** (`N`, `E`).  
    Folder and subfolder information are optional and can be used to nest features in Google Earth. Elevation is optional.  

    **Note:** Visually confirm the geographic placement of points in the new file.
    """
)


# -------------------------------------------------------------------
# Template download - generated in code
# -------------------------------------------------------------------
def build_excel_template() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Template"

    # Group headers
    ws.merge_cells("D1:E1")
    ws.merge_cells("F1:G1")
    ws["D1"] = "WGS84 Geographic"
    ws["F1"] = "WGS84 UTM 17T"

    # Column headers
    headers = [
        "folder",
        "subfolder",
        "feature_name",
        "lat",
        "long",
        "N",
        "E",
        "elevation (optional)",
    ]

    for col_idx, header in enumerate(headers, start=1):
        ws.cell(row=2, column=col_idx).value = header

    # Formatting
    thin = Side(style="thin", color="808080")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    group_fill = PatternFill("solid", fgColor="EDEDED")
    header_fill = PatternFill("solid", fgColor="EDEDED")

    for cell_ref in ["D1", "F1"]:
        cell = ws[cell_ref]
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = group_fill
        cell.border = border

    for row in ws.iter_rows(min_row=1, max_row=2, min_col=1, max_col=8):
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border

    for cell in ws[2]:
        cell.font = Font(bold=True)
        cell.fill = header_fill

    # Column widths
    widths = {
        "A": 18,  # folder
        "B": 18,  # subfolder
        "C": 26,  # feature_name
        "D": 16,  # lat
        "E": 16,  # long
        "F": 16,  # N
        "G": 16,  # E
        "H": 22,  # elevation optional
    }

    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 22
    ws.freeze_panes = "A3"

    # Number formats for user-entry area
    for row in range(3, 1003):
        ws[f"D{row}"].number_format = "0.000000000"
        ws[f"E{row}"].number_format = "0.000000000"
        ws[f"F{row}"].number_format = "0.000"
        ws[f"G{row}"].number_format = "0.000"
        ws[f"H{row}"].number_format = "0.00"

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


st.download_button(
    "Download Excel template",
    data=build_excel_template(),
    file_name="Excel_to_KMZ_template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)


# -------------------------------------------------------------------
# File upload
# -------------------------------------------------------------------
up = st.file_uploader(
    "Upload Excel template",
    type=["xlsx", "xls"],
)


# -------------------------------------------------------------------
# Helpers
# -------------------------------------------------------------------
def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [
        str(c).strip().lower().replace("\n", " ").replace(" ", "_")
        for c in df.columns
    ]
    return df


def pick(df: pd.DataFrame, options) -> str | None:
    for o in options:
        if o in df.columns:
            return o
    return None


def read_excel_template_aware(file_bytes: bytes) -> pd.DataFrame:
    """
    New template uses row 2 as actual headers, so first try header=1.
    Fallback to header=0 to remain compatible with older simple templates.
    """
    for header_row in [1, 0]:
        temp = pd.read_excel(BytesIO(file_bytes), header=header_row)
        temp = normalize_columns(temp)

        has_name = pick(temp, ["feature_name", "name", "label", "id", "title"]) is not None
        has_ll = (
            pick(temp, ["lat", "latitude"]) is not None
            and pick(temp, ["long", "lon", "longitude"]) is not None
        )
        has_utm = (
            pick(temp, ["n", "northing", "utm_n", "utm_northing", "y"]) is not None
            and pick(temp, ["e", "easting", "utm_e", "utm_easting", "x"]) is not None
        )

        if has_name and (has_ll or has_utm):
            return temp

    # Return row-2 version anyway so error messages are based on expected template
    return normalize_columns(pd.read_excel(BytesIO(file_bytes), header=1))


def invalid_ll(long_series: pd.Series, lat_series: pd.Series) -> pd.Series:
    """Invalid if NaN, inf, or outside global WGS84 lat/long bounds."""
    bad = ~np.isfinite(long_series) | ~np.isfinite(lat_series)
    bad |= ~long_series.between(-180.0, 180.0)
    bad |= ~lat_series.between(-90.0, 90.0)
    return bad


def invalid_wgs84_utm17(e_series: pd.Series, n_series: pd.Series) -> pd.Series:
    """
    Broad sanity check for WGS84 UTM Zone 17N.
    Google Earth may show Pearson/Toronto as UTM 17T, but conversion uses EPSG:32617.
    """
    bad = ~np.isfinite(e_series) | ~np.isfinite(n_series)
    bad |= ~e_series.between(100_000, 900_000)
    bad |= ~n_series.between(0, 10_000_000)
    return bad


def transformer_wgs84_utm17_to_wgs84_geo() -> Transformer:
    """
    WGS84 / UTM Zone 17N meters -> WGS84 Geographic degrees.
    EPSG:32617 -> EPSG:4326.
    """
    return Transformer.from_crs("EPSG:32617", "EPSG:4326", always_xy=True)


def clean_folder_value(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def build_kmz_points(doc_name: str, rows: pd.DataFrame) -> bytes:
    """
    rows expected:
      folder, subfolder, feature_name, long, lat, elevation

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

    def kml_folder(name, indent="    "):
        return f'{indent}<Folder><name>{escape(name)}</name>\n'

    def kml_folder_end(indent="    "):
        return f"{indent}</Folder>\n"

    def pm_point(name, long_val, lat_val, elev=None, indent="      "):
        coord = f"{long_val:.9f},{lat_val:.9f}"
        if elev is not None and np.isfinite(elev):
            coord += f",{float(elev):.2f}"

        return (
            f"{indent}<Placemark>\n"
            f"{indent}  <name>{escape(str(name))}</name>\n"
            f"{indent}  <Point><coordinates>{coord}</coordinates></Point>\n"
            f"{indent}</Placemark>\n"
        )

    def write_points(kml_buf, point_rows, indent="      "):
        for _, r in point_rows.iterrows():
            kml_buf.write(
                pm_point(
                    r["feature_name"],
                    float(r["long"]),
                    float(r["lat"]),
                    r.get("elevation", np.nan),
                    indent=indent,
                ).encode("utf-8")
            )

    work = rows.copy()
    work["_folder"] = work["folder"].map(clean_folder_value) if "folder" in work.columns else ""
    work["_subfolder"] = work["subfolder"].map(clean_folder_value) if "subfolder" in work.columns else ""

    kml = BytesIO()
    kml.write(kml_header(doc_name).encode("utf-8"))

    # Rows without folder or subfolder: document-level placemarks
    doc_level = work[(work["_folder"] == "") & (work["_subfolder"] == "")]
    write_points(kml, doc_level, indent="    ")

    # Rows with no main folder but with subfolder: subfolder becomes document-level folder
    subfolder_only = work[(work["_folder"] == "") & (work["_subfolder"] != "")]
    for subfolder_name, sg in subfolder_only.groupby("_subfolder", sort=False):
        kml.write(kml_folder(subfolder_name, indent="    ").encode("utf-8"))
        write_points(kml, sg, indent="      ")
        kml.write(kml_folder_end(indent="    ").encode("utf-8"))

    # Rows with main folder
    with_folder = work[work["_folder"] != ""]
    for folder_name, fg in with_folder.groupby("_folder", sort=False):
        kml.write(kml_folder(folder_name, indent="    ").encode("utf-8"))

        # Directly under folder
        direct = fg[fg["_subfolder"] == ""]
        write_points(kml, direct, indent="      ")

        # Nested subfolders
        nested = fg[fg["_subfolder"] != ""]
        for subfolder_name, sg in nested.groupby("_subfolder", sort=False):
            kml.write(kml_folder(subfolder_name, indent="      ").encode("utf-8"))
            write_points(kml, sg, indent="        ")
            kml.write(kml_folder_end(indent="      ").encode("utf-8"))

        kml.write(kml_folder_end(indent="    ").encode("utf-8"))

    kml.write(kml_footer().encode("utf-8"))
    kml.seek(0)

    kmz = BytesIO()
    with zipfile.ZipFile(kmz, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("doc.kml", kml.getvalue())

    kmz.seek(0)
    return kmz.getvalue()


# -------------------------------------------------------------------
# Session persistence
# -------------------------------------------------------------------
for key in ("kmz_bytes", "validation_xlsx", "base_name"):
    if key not in st.session_state:
        st.session_state[key] = None

convert_clicked = st.button("Convert") if up else False


# -------------------------------------------------------------------
# Main conversion
# -------------------------------------------------------------------
if convert_clicked and up:
    try:
        df0 = read_excel_template_aware(up.getvalue())

        if df0.empty:
            st.error("No rows found.")
        else:
            df = df0.copy()

            col_folder = pick(df, ["folder", "group", "layer"])
            col_subfolder = pick(df, ["subfolder", "sub_folder"])
            col_name = pick(df, ["feature_name", "name", "label", "id", "title"])
            col_lat = pick(df, ["lat", "latitude"])
            col_long = pick(df, ["long", "lon", "longitude"])
            col_n = pick(df, ["n", "northing", "utm_n", "utm_northing", "y"])
            col_e = pick(df, ["e", "easting", "utm_e", "utm_easting", "x"])
            col_z = pick(
                df,
                [
                    "elevation",
                    "elevation_(optional)",
                    "elevation_optional",
                    "elev",
                    "z",
                    "altitude",
                    "height",
                    "elevation_m",
                ],
            )

            if col_name is None:
                st.error("Missing required column: feature_name")
            elif not ((col_lat and col_long) or (col_n and col_e)):
                st.error("Missing coordinate columns. Provide either lat/long or N/E.")
            else:
                keep = [
                    c for c in [
                        col_folder,
                        col_subfolder,
                        col_name,
                        col_lat,
                        col_long,
                        col_n,
                        col_e,
                        col_z,
                    ]
                    if c
                ]
                df = df[keep].copy()

                # Remove fully blank rows from the template body
                important_cols = [c for c in [col_name, col_lat, col_long, col_n, col_e] if c]
                df = df[~df[important_cols].isna().all(axis=1)].copy()

                if df.empty:
                    st.error("No filled rows found.")
                    st.stop()

                # Numeric coercion
                for c in [col_lat, col_long, col_n, col_e, col_z]:
                    if c and c in df.columns:
                        df[c] = pd.to_numeric(df[c], errors="coerce")

                # Output frame
                out = pd.DataFrame(index=df.index)
                out["folder"] = df[col_folder] if col_folder and col_folder in df.columns else ""
                out["subfolder"] = df[col_subfolder] if col_subfolder and col_subfolder in df.columns else ""

                out["feature_name"] = df[col_name].fillna("").astype(str).str.strip()
                blank_names = out["feature_name"] == ""
                for seq, idx in enumerate(out.index[blank_names], start=1):
                    out.loc[idx, "feature_name"] = f"Point {seq}"

                out["elevation"] = df[col_z] if col_z and col_z in df.columns else np.nan

                out["lat"] = np.nan
                out["long"] = np.nan
                out["N"] = df[col_n] if col_n and col_n in df.columns else np.nan
                out["E"] = df[col_e] if col_e and col_e in df.columns else np.nan

                out["input_type"] = ""
                out["conversion_note"] = ""

                # A) WGS84 geographic input: pass through directly
                if col_lat and col_long:
                    m_latlong = df[col_lat].notna() & df[col_long].notna()

                    out.loc[m_latlong, "lat"] = df.loc[m_latlong, col_lat]
                    out.loc[m_latlong, "long"] = df.loc[m_latlong, col_long]
                    out.loc[m_latlong, "input_type"] = "wgs84_geographic"
                    out.loc[m_latlong, "conversion_note"] = "plotted_from_lat_long"

                # B) WGS84 UTM 17T / Zone 17N input: convert to WGS84 geographic
                if col_n and col_e:
                    m_utm = df[col_n].notna() & df[col_e].notna()

                    # If both lat/long and N/E are present, use lat/long.
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
                        utm_idx = df.index[m_utm]
                        bad_utm = invalid_wgs84_utm17(
                            df.loc[utm_idx, col_e],
                            df.loc[utm_idx, col_n],
                        )

                        if bad_utm.any():
                            st.warning(
                                f"{bad_utm.sum()} UTM row(s) failed WGS84 UTM 17T range checks and were skipped."
                            )

                        good_utm_idx = bad_utm.index[~bad_utm.to_numpy()]

                        if len(good_utm_idx) > 0:
                            tr = transformer_wgs84_utm17_to_wgs84_geo()
                            long_vals, lat_vals = tr.transform(
                                df.loc[good_utm_idx, col_e].to_numpy(),
                                df.loc[good_utm_idx, col_n].to_numpy(),
                            )

                            out.loc[good_utm_idx, "long"] = long_vals
                            out.loc[good_utm_idx, "lat"] = lat_vals
                            out.loc[good_utm_idx, "input_type"] = "wgs84_utm_17t"
                            out.loc[good_utm_idx, "conversion_note"] = "converted_from_wgs84_utm_zone_17n"

                # Final coordinate validation
                bad_final = invalid_ll(out["long"], out["lat"])

                if bad_final.any():
                    st.warning(
                        f"{bad_final.sum()} row(s) had invalid or missing final WGS84 coordinates and were skipped."
                    )
                    out = out[~bad_final].copy()

                if out.empty:
                    st.error("No valid coordinates were found after processing.")
                else:
                    # Round validation output
                    out["lat"] = pd.to_numeric(out["lat"], errors="coerce").round(9)
                    out["long"] = pd.to_numeric(out["long"], errors="coerce").round(9)
                    out["N"] = pd.to_numeric(out["N"], errors="coerce").round(3)
                    out["E"] = pd.to_numeric(out["E"], errors="coerce").round(3)
                    out["elevation"] = pd.to_numeric(out["elevation"], errors="coerce")

                    base = Path(up.name).stem

                    # KMZ: Google Earth WGS84 geographic points
                    kmz_bytes = build_kmz_points(
                        f"{base} — WGS84 Geographic",
                        out[[
                            "folder",
                            "subfolder",
                            "feature_name",
                            "long",
                            "lat",
                            "elevation",
                        ]],
                    )

                    # Validation Excel
                    valid_cols = [
                        "folder",
                        "subfolder",
                        "feature_name",
                        "lat",
                        "long",
                        "N",
                        "E",
                        "elevation",
                        "input_type",
                        "conversion_note",
                    ]

                    out_valid = out[valid_cols].copy()

                    xbuf = BytesIO()
                    with pd.ExcelWriter(xbuf, engine="openpyxl") as xw:
                        readme = pd.DataFrame(
                            {
                                "Validation": [
                                    "All exported KMZ coordinates are WGS84 Geographic longitude/latitude.",
                                    "Rows with lat/long are plotted directly.",
                                    "Rows with N/E are interpreted as WGS84 UTM 17T / UTM Zone 17N and converted to WGS84 Geographic.",
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


# -------------------------------------------------------------------
# Downloads
# -------------------------------------------------------------------
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
