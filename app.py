import streamlit as st, zipfile, xml.etree.ElementTree as ET, pandas as pd
from pyproj import Transformer
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import os
from pathlib import Path
import numpy as np

# --- Resolve absolute paths to grid files ---
APP_DIR = Path(__file__).parent.resolve()
TOR_GRID = APP_DIR / "TO27CSv1.gsb"   # default (Toronto)
ON_GRID  = APP_DIR / "ON27CSv1.gsb"   # fallback (Ontario-wide)

# Ensure PROJ uses local grids only
os.environ["PROJ_DATA"] = str(APP_DIR)
os.environ["PROJ_NETWORK"] = "OFF"

st.set_page_config(page_title="Toronto Grid Coordinate Converter", page_icon="📍")
st.title("KMZ Coordinates.xlsx NAD83 Geographic → NAD27 / UTM 17N")
st.caption("Default grid: TO27CSv1.gsb  •  Fallback on outside coverage: ON27CSv1.gsb")

up = st.file_uploader("Upload KMZ or KML", type=["kmz","kml"])

def parse_kml_bytes(kml_bytes):
    ns = {"kml":"http://www.opengis.net/kml/2.2"}
    root = ET.fromstring(kml_bytes)
    rows = []
    for pm in root.findall(".//kml:Placemark", ns):
        name_el = pm.find("kml:name", ns)
        name = name_el.text if name_el is not None else "Unnamed"
        for ct in pm.findall(".//kml:coordinates", ns):
            text = (ct.text or "").strip()
            for idx, c in enumerate(text.split()):
                parts = c.split(",")
                if len(parts) >= 2:
                    lon, lat = float(parts[0]), float(parts[1])
                    elev = float(parts[2]) if len(parts) > 2 and parts[2] else None
                    rows.append({
                        "feature_name": name,
                        "vertex_index": idx + 1,
                        "lon_4269": lon,
                        "lat_4269": lat,
                        "elevation_m": elev
                    })
    return pd.DataFrame(rows).sort_values(["feature_name", "vertex_index"])

def transformer_for(grid_path: Path) -> Transformer:
    # NAD83 (geog) → NAD27/UTM17N with explicit NTv2 grid
    pipeline = (
        f"+proj=pipeline "
        f"+step +proj=unitconvert +xy_in=deg +xy_out=rad "
        f"+step +inv +proj=hgridshift +grids={grid_path} "
        f"+step +proj=utm +zone=17 +datum=NAD27"
    )
    return Transformer.from_pipeline(pipeline)

if up and st.button("Convert"):
    try:
        # Read KML bytes
        if up.name.lower().endswith(".kmz"):
            with zipfile.ZipFile(up) as z:
                kml_name = next(n for n in z.namelist() if n.lower().endswith(".kml"))
                kml_bytes = z.read(kml_name)
        else:
            kml_bytes = up.read()

        df = parse_kml_bytes(kml_bytes)
        if df.empty:
            st.error("No coordinates found."); st.stop()

        # Default transform with Toronto grid
        t_tor = transformer_for(TOR_GRID)
        E_tor, N_tor = t_tor.transform(df["lon_4269"].to_numpy(), df["lat_4269"].to_numpy())
        E = pd.Series(E_tor, dtype="float64")
        N = pd.Series(N_tor, dtype="float64")

        # Identify rows outside coverage (pyproj returns NaN when grid not applicable)
        mask_missing = E.isna() | N.isna()

        used_grid = np.where(mask_missing, "ON27CSv1.gsb", "TO27CSv1.gsb")

        # Fallback only for the rows that failed on Toronto grid
        if mask_missing.any():
            t_on = transformer_for(ON_GRID)
            E_on, N_on = t_on.transform(
                df.loc[mask_missing, "lon_4269"].to_numpy(),
                df.loc[mask_missing, "lat_4269"].to_numpy()
            )
            # insert fallback results
            E.loc[mask_missing] = E_on
            N.loc[mask_missing] = N_on
            st.warning(f"Applied fallback grid (ON27CSv1.gsb) for {mask_missing.sum()} point(s).")

        # Assemble output
        out = df[["feature_name","vertex_index","lat_4269","lon_4269"]].copy()
        out["N_26717"] = N
        out["E_26717"] = E
        # rounding
        out["lat_4269"] = out["lat_4269"].round(6)
        out["lon_4269"] = out["lon_4269"].round(6)
        out["N_26717"] = out["N_26717"].round(3)
        out["E_26717"] = out["E_26717"].round(3)
        # keep elevation and add grid_used AFTER elevation
        out["elevation_m"] = df["elevation_m"]
        out["grid_used"] = used_grid

        # Excel output
        wb = Workbook(); ws = wb.active; ws.title = "Coordinates"
        ws.append(["", "",
                   "NAD 83 Geographic", "NAD 83 Geographic",
                   "NAD 27 / UTM Zone 17N", "NAD 27 / UTM Zone 17N",
                   "", ""])
        ws.append(["feature_name","vertex_index","lat_4269","lon_4269",
                   "N_26717","E_26717","elevation_m","grid_used"])
        for r in dataframe_to_rows(out, index=False, header=False):
            ws.append(r)

        bio = BytesIO(); wb.save(bio); bio.seek(0)
        base = os.path.splitext(up.name)[0]
        out_name = f"{base}_coordinates.xlsx"

        st.success("Done.")
        st.download_button("Download XLSX", data=bio,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error("Transformation failed. Ensure grid files exist in repo root (TO27CSv1.gsb, ON27CSv1.gsb).")
        st.exception(e)
