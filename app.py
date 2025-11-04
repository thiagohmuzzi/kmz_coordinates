import streamlit as st, zipfile, xml.etree.ElementTree as ET, pandas as pd
from pyproj import Transformer
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import os
from pathlib import Path

# --- Resolve absolute path to the Toronto grid file ---
APP_DIR = Path(__file__).parent.resolve()
GRID_PATH = str(APP_DIR / "TO27CSv1.gsb")

# Make sure PROJ looks in the app folder (belt & suspenders)
os.environ["PROJ_DATA"] = str(APP_DIR)
os.environ["PROJ_NETWORK"] = "OFF"  # avoid network fetch; we have the file

st.set_page_config(page_title="Toronto Grid Coordinate Converter", page_icon="📍")
st.title("KMZ Coordinates.xlsx NAD83 Geographic / NAD27 UTM 17N Toronto")

up = st.file_uploader("Upload KMZ or KML", type=["kmz","kml"])

def parse_kml_bytes(kml_bytes):
    ns = {"kml":"http://www.opengis.net/kml/2.2"}
    root = ET.fromstring(kml_bytes)
    rows=[]
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
                    rows.append({"feature_name": name, "vertex_index": idx+1,
                                 "lon_4269": lon, "lat_4269": lat, "elevation_m": elev})
    return pd.DataFrame(rows).sort_values(["feature_name","vertex_index"])

if up and st.button("Convert"):
    try:
        # Read KML bytes from KMZ or KML
        if up.name.lower().endswith(".kmz"):
            z = zipfile.ZipFile(up)
            kml_name = next(n for n in z.namelist() if n.lower().endswith(".kml"))
            kml_bytes = z.read(kml_name)
        else:
            kml_bytes = up.read()

        df = parse_kml_bytes(kml_bytes)
        if df.empty:
            st.error("No coordinates found."); st.stop()

        # Critical: use absolute grid path
        pipeline = (
            f"+proj=pipeline "
            f"+step +proj=unitconvert +xy_in=deg +xy_out=rad "
            f"+step +inv +proj=hgridshift +grids={GRID_PATH} "
            f"+step +proj=utm +zone=17 +datum=NAD27"
        )
        transformer = Transformer.from_pipeline(pipeline)

        E, N = transformer.transform(df["lon_4269"].values, df["lat_4269"].values)
        df["E_26717"], df["N_26717"] = E, N

        out = df[["feature_name","vertex_index","lat_4269","lon_4269","N_26717","E_26717","elevation_m"]].copy()
        out["lat_4269"]=out["lat_4269"].round(6); out["lon_4269"]=out["lon_4269"].round(6)
        out["N_26717"]=out["N_26717"].round(3); out["E_26717"]=out["E_26717"].round(3)

        wb = Workbook(); ws = wb.active; ws.title="Coordinates"
        ws.append(["","", "NAD 83 Geographic","NAD 83 Geographic",
                   "NAD 27 / UTM Zone 17N (Grid: TO27CSv1.gsb)",
                   "NAD 27 / UTM Zone 17N (Grid: TO27CSv1.gsb)", ""])
        ws.append(["feature_name","vertex_index","lat_4269","lon_4269","N_26717","E_26717","elevation_m"])
        for r in dataframe_to_rows(out, index=False, header=False): ws.append(r)
        bio = BytesIO(); wb.save(bio); bio.seek(0)

        base = os.path.splitext(up.name)[0]
        out_name = f"{base}_coordinates.xlsx"

        st.success("Done.")
        st.download_button("Download XLSX", data=bio,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        st.error("Transformation failed. Check that TO27CSv1.gsb exists in the repo root.")
        st.exception(e)
