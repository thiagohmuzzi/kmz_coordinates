import streamlit as st, zipfile, xml.etree.ElementTree as ET, pandas as pd
from pyproj import Transformer
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import os
from pathlib import Path
import numpy as np

APP_DIR = Path(__file__).parent.resolve()
TOR_GRID = APP_DIR / "TO27CSv1.gsb"   # default (Toronto)
ON_GRID  = APP_DIR / "ON27CSv1.gsb"   # fallback (Ontario-wide)

# Let PROJ see local grids; network off
os.environ["PROJ_DATA"] = str(APP_DIR)
os.environ["PROJ_NETWORK"] = "OFF"

st.set_page_config(page_title="KMZ Coordinates Extraction", page_icon="🧭")
st.title("KMZ Coordinates to Excel - NAD83 Geographic / NAD27 UTM Zone 17N")
st.caption("Default grid: TO27CSv1.gsb • Fallback if outside coverage: ON27CSv1.gsb")

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
                        "lon": lon,
                        "lat": lat,
                        "elevation_m": elev
                    })
    return pd.DataFrame(rows).sort_values(["feature_name", "vertex_index"])

def transformer_nad83_ll_to_nad27_utm(grid_path: Path) -> Transformer:
    """
    NAD83 geographic (deg) -> NAD27 / UTM Zone 17N (m)
    (same pipeline as Excel_Transformation.tr_to_utm)
    """
    pipe = (
        f"+proj=pipeline "
        f"+step +proj=unitconvert +xy_in=deg +xy_out=rad "
        f"+step +inv +proj=hgridshift +grids={grid_path} "
        f"+step +proj=utm +zone=17 +datum=NAD27"
    )
    return Transformer.from_pipeline(pipe)

def invalid_mask(E: pd.Series, N: pd.Series) -> pd.Series:
    m = E.isna() | N.isna()
    m |= ~E.between(100_000, 900_000)
    m |= ~N.between(4_000_000, 6_000_000)
    return m

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

        # Toronto first
        if not TOR_GRID.exists():
            st.error("TO27CSv1.gsb not found next to KMZ_Coordinates.py.")
            st.stop()

        t_tor = transformer_nad83_ll_to_nad27_utm(TOR_GRID)
        E_tor, N_tor = t_tor.transform(df["lon"].to_numpy(), df["lat"].to_numpy())
        E = pd.Series(E_tor, dtype="float64")
        N = pd.Series(N_tor, dtype="float64")
        used_grid = np.full(len(df), "TO27CSv1.gsb", dtype=object)

        bad = invalid_mask(E, N)

        # Ontario fallback per bad row (if ON grid exists)
        if bad.any() and ON_GRID.exists():
            t_on = transformer_nad83_ll_to_nad27_utm(ON_GRID)
            E_on, N_on = t_on.transform(
                df.loc[bad, "lon"].to_numpy(),
                df.loc[bad, "lat"].to_numpy()
            )
            E.loc[bad] = E_on
            N.loc[bad] = N_on

            still_bad = invalid_mask(E.loc[bad], N.loc[bad])
            used_grid[bad.values] = "ON27CSv1.gsb"
            if still_bad.any():
                idxs = still_bad[still_bad].index
                used_grid[idxs] = "None"
                st.warning(f"{still_bad.sum()} point(s) remain invalid after fallback.")
            st.info(f"Applied fallback grid (ON27CSv1.gsb) for {bad.sum()} point(s).")
        elif bad.any():
            st.warning(
                f"{bad.sum()} point(s) flagged invalid and ON27CSv1.gsb not available; "
                "grid_used will remain TO27CSv1.gsb but N/E may be out of expected bounds."
            )

        # Assemble output (clean headers)
        out = df[["feature_name", "vertex_index", "lat", "lon"]].copy()
        out["N_26717"] = N.round(3)
        out["E_26717"] = E.round(3)
        out["elevation_m"] = df["elevation_m"]
        out["grid_used"] = used_grid

        # Excel output with header row defining datums
        wb = Workbook(); ws = wb.active; ws.title = "Coordinates"
        ws.append([
            "",
            "",
            "NAD83 Geographic",
            "NAD83 Geographic",
            "NAD27 / UTM Zone 17N",
            "NAD27 / UTM Zone 17N",
            "",
            ""
        ])
        ws.append(["feature_name","vertex_index","lat_4269","lon_4269","N_26717","E_26717","elevation_m","grid_used"])

        out_for_xlsx = out.rename(
            columns={
                "lat": "lat_4269",
                "lon": "lon_4269"
            }
        )

        for r in dataframe_to_rows(out_for_xlsx, index=False, header=False):
            ws.append(r)

        bio = BytesIO(); wb.save(bio); bio.seek(0)
        base = os.path.splitext(up.name)[0]
        out_name = f"{base}_coordinates.xlsx"

        st.success("Conversion complete.")
        st.download_button(
            "Download XLSX",
            data=bio,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error("Transformation failed. Check grid files exist next to KMZ_Coordinates.py (TO27CSv1.gsb, ON27CSv1.gsb).")
        st.exception(e)
