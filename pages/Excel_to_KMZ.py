import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
from pyproj import Transformer
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import os
from pathlib import Path
import numpy as np

# -------------------------------------------------------------------
# Paths and PROJ setup
# -------------------------------------------------------------------
THIS = Path(__file__).resolve()
ROOT = THIS.parent  # KMZ_Coordinates.py lives in repo root

TOR_GRID = ROOT / "TO27CSv1.gsb"   # default (Toronto grid)
ON_GRID  = ROOT / "ON27CSv1.gsb"   # fallback (Ontario-wide)

os.environ["PROJ_DATA"] = str(ROOT.resolve())
os.environ["PROJ_NETWORK"] = "OFF"

# -------------------------------------------------------------------
# Streamlit UI
# -------------------------------------------------------------------
st.set_page_config(page_title="KMZ Coordinates Extraction", page_icon="🧭")
st.title("KMZ Coordinates to Excel – NAD83 Geographic / NAD27 UTM Zone 17N")
st.caption(
    """
    Input coordinates to the excel template provided. 
    Rows may contain either **NAD83 Geographic** (lat/long) or **NAD27 / UTM Zone 17N** (N/E). 
    Folder information is optional and can be provided if desired to have the features nested in separate folders for Google Earth. Elevation is optional. \n
    **Note:** Always visually confirm the geographic placement of the points in the new file.

    """
)

up = st.file_uploader("Upload KMZ or KML", type=["kmz", "kml"])


# -------------------------------------------------------------------
# Helpers
# -------------------------------------------------------------------
def parse_kml_bytes(kml_bytes: bytes) -> pd.DataFrame:
    """
    Parse KML bytes and return DataFrame:
    feature_name, vertex_index, lat_4269, lon_4269, elevation_m
    (lat/long assumed NAD83, as displayed by Google Earth default).
    """
    ns = {"kml": "http://www.opengis.net/kml/2.2"}
    root = ET.fromstring(kml_bytes)
    rows = []
    for pm in root.findall(".//kml:Placemark", ns):
        name_el = pm.find("kml:name", ns)
        name = name_el.text if name_el is not None else "Unnamed"
        for ct in pm.findall(".//kml:coordinates", ns):
            text = (ct.text or "").strip()
            if not text:
                continue
            for idx, c in enumerate(text.split()):
                parts = c.split(",")
                if len(parts) < 2:
                    continue
                try:
                    lon = float(parts[0])
                    lat = float(parts[1])
                except ValueError:
                    continue
                elev = None
                if len(parts) > 2 and parts[2]:
                    try:
                        elev = float(parts[2])
                    except ValueError:
                        elev = None
                rows.append(
                    {
                        "feature_name": name,
                        "vertex_index": idx + 1,
                        "lat_4269": lat,
                        "lon_4269": lon,
                        "elevation_m": elev,
                    }
                )
    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.sort_values(["feature_name", "vertex_index"])
    return df


def transformer_nad83_ll_to_nad27_utm(grid_path: Path) -> Transformer:
    """
    NAD83 geographic (degrees) -> NAD27 / UTM Zone 17N (meters)

    This matches the canonical pipeline we used in Excel_Transformation:
      1) degrees -> radians
      2) inverse NTv2 grid shift NAD83 -> NAD27 (radians)
      3) UTM 17N (NAD27) -> easting/northing
    """
    pipe = (
        f"+proj=pipeline "
        f"+step +proj=unitconvert +xy_in=deg +xy_out=rad "
        f"+step +inv +proj=hgridshift +grids={grid_path} "
        f"+step +proj=utm +zone=17 +datum=NAD27"
    )
    return Transformer.from_pipeline(pipe)


def invalid_utm(e: pd.Series, n: pd.Series) -> pd.Series:
    """
    Flag obviously invalid UTM17 coordinates.
    Keep bounds broad to avoid falsely rejecting valid airport points.
    """
    m = e.isna() | n.isna()
    m |= ~e.between(100_000, 900_000)
    m |= ~n.between(4_000_000, 6_000_000)
    return m


# -------------------------------------------------------------------
# Main
# -------------------------------------------------------------------
if up and st.button("Convert"):
    try:
        # Ensure grids exist
        if not TOR_GRID.exists():
            st.error(f"Toronto grid not found at: {TOR_GRID}")
            st.stop()
        if not ON_GRID.exists():
            st.info(f"Ontario grid not found at: {ON_GRID}. Fallback will not be used.")

        # Read KML bytes from KMZ or KML
        if up.name.lower().endswith(".kmz"):
            with zipfile.ZipFile(up) as z:
                kml_name = next(n for n in z.namelist() if n.lower().endswith(".kml"))
                kml_bytes = z.read(kml_name)
        else:
            kml_bytes = up.read()

        df = parse_kml_bytes(kml_bytes)
        if df.empty:
            st.error("No coordinates found in KML/KMZ.")
            st.stop()

        # Round lat/long slightly for readability
        df["lat_4269"] = df["lat_4269"].round(9)
        df["lon_4269"] = df["lon_4269"].round(9)

        # ---- NAD83 → NAD27 / UTM 17N with Toronto + Ontario fallback ----
        # Toronto transform
        t_tor = transformer_nad83_ll_to_nad27_utm(TOR_GRID)
        E_tor, N_tor = t_tor.transform(df["lon_4269"].to_numpy(), df["lat_4269"].to_numpy())

        E = pd.Series(E_tor, index=df.index, dtype="float64")
        N = pd.Series(N_tor, index=df.index, dtype="float64")
        grid_used = np.full(len(df), "TO27CSv1.gsb", dtype=object)

        bad = invalid_utm(E, N)

        # Ontario fallback (row-by-row where Toronto failed)
        if bad.any() and ON_GRID.exists():
            t_on = transformer_nad83_ll_to_nad27_utm(ON_GRID)
            E_on, N_on = t_on.transform(
                df.loc[bad, "lon_4269"].to_numpy(),
                df.loc[bad, "lat_4269"].to_numpy()
            )
            E.loc[bad] = E_on
            N.loc[bad] = N_on

            still_bad = invalid_utm(E.loc[bad], N.loc[bad])
            grid_used[bad.values] = "ON27CSv1.gsb"

            if still_bad.any():
                idxs = still_bad[still_bad].index
                grid_used[idxs] = "None"
                st.warning(f"{still_bad.sum()} point(s) remain invalid after fallback.")

            st.info(f"Applied fallback grid (ON27CSv1.gsb) for {bad.sum()} point(s).")

        elif bad.any():
            st.warning(
                f"{bad.sum()} point(s) outside expected UTM17 bounds and Ontario grid "
                "is not available – N/E may be invalid."
            )

        # Final N/E, rounded
        df["N_26717"] = N.round(3)
        df["E_26717"] = E.round(3)
        df["grid_used"] = grid_used

        # ----------------------------------------------------------------
        # Excel output
        # ----------------------------------------------------------------
        wb = Workbook()
        ws = wb.active
        ws.title = "Coordinates"

        # Datum header row
        ws.append([
            "",
            "",
            "NAD83 Geographic",
            "NAD83 Geographic",
            "NAD27 / UTM Zone 17N",
            "NAD27 / UTM Zone 17N",
            "",
            "",
        ])

        # Column header row
        ws.append([
            "feature_name",
            "vertex_index",
            "lat_4269",
            "lon_4269",
            "N_26717",
            "E_26717",
            "elevation_m",
            "grid_used",
        ])

        out = df[[
            "feature_name",
            "vertex_index",
            "lat_4269",
            "lon_4269",
            "N_26717",
            "E_26717",
            "elevation_m",
            "grid_used",
        ]].copy()

        for r in dataframe_to_rows(out, index=False, header=False):
            ws.append(r)

        bio = BytesIO()
        wb.save(bio)
        bio.seek(0)

        base = os.path.splitext(up.name)[0]
        out_name = f"{base}_coordinates.xlsx"

        st.success("Extraction and transformation complete.")
        st.download_button(
            "Download XLSX",
            data=bio,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error("KMZ → Excel transformation failed.")
        st.exception(e)
