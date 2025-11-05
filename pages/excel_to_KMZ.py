import streamlit as st, pandas as pd, numpy as np, zipfile, os
from io import BytesIO
from pyproj import Transformer
from xml.sax.saxutils import escape
from pathlib import Path

# --- Locate Toronto grid file robustly ---
from pathlib import Path
import os
import streamlit as st

APP_DIR = Path(__file__).parent.resolve()

# Search common locations (repo root, /pages/, working dir)
CANDIDATES = [
    APP_DIR / "TO27CSv1.gsb",
    APP_DIR.parent / "TO27CSv1.gsb",
    Path.cwd() / "TO27CSv1.gsb",
    APP_DIR.parent.parent / "TO27CSv1.gsb",
]

GRID = next((p for p in CANDIDATES if p.exists()), None)
if GRID is None:
    st.error("TO27CSv1.gsb not found. Place it at repo root or next to this app file.")
    st.stop()

# Register for PROJ
os.environ["PROJ_DATA"] = str(GRID.parent)
os.environ["PROJ_NETWORK"] = "OFF"
GRID_PATH = str(GRID)


st.title("Excel (UTM) to KMZ")
st.caption("Input: NAD27 UTM Zone 17N • Grid used: TO27CSv1.gsb")

GRID = Path(__file__).parent / "TO27CSv1.gsb"
if not GRID.exists():
    st.error("TO27CSv1.gsb not found in repo root."); st.stop()

up = st.file_uploader("Upload Excel", type=["xlsx","xls"])
btn = st.button("Convert")

def pick(cols, opts):
    for o in opts:
        if o in cols: return o
    return None

if up and btn:
    df0 = pd.read_excel(up)
    df = df0.copy()
    df.columns = [str(c).strip().lower().replace(" ", "_") for c in df.columns]

    folder_col = pick(df.columns, ["folder","group","layer"])
    name_col   = pick(df.columns, ["feature_name","name","label","id","title"])
    e_col      = pick(df.columns, ["easting","e","utm_e","utm_easting","east","x"])
    n_col      = pick(df.columns, ["northing","n","utm_n","utm_northing","north","y"])
    z_col      = pick(df.columns, ["elevation","elev","elevation_m","z","altitude","height"])

    missing = [k for k,v in {"folder":folder_col,"feature_name":name_col,"easting":e_col,"northing":n_col}.items() if v is None]
    if missing:
        st.error(f"Missing required columns: {missing}. Found: {list(df.columns)}"); st.stop()

    df[e_col] = pd.to_numeric(df[e_col], errors="coerce")
    df[n_col] = pd.to_numeric(df[n_col], errors="coerce")
    if z_col: df[z_col] = pd.to_numeric(df[z_col], errors="coerce")
    df = df.dropna(subset=[e_col,n_col]).copy()

    pipeline = (
        f"+proj=pipeline "
        f"+step +inv +proj=utm +zone=17 +datum=NAD27 "
        f"+step +proj=hgridshift +grids={GRID_PATH} "
        f"+step +proj=unitconvert +xy_in=rad +xy_out=deg"
    )
    transformer = Transformer.from_pipeline(pipeline)

    lon, lat = tr.transform(df[e_col].to_numpy(), df[n_col].to_numpy())
    df["lon_4269"] = lon; df["lat_4269"] = lat
    df["grid_status"] = np.where(df[["lon_4269","lat_4269"]].isna().any(axis=1),
                                "no_shift (outside TO27CSv1.gsb or invalid)", "shift_ok (TO27CSv1.gsb)")

    def kml_header(name):
        return ('<?xml version="1.0" encoding="UTF-8"?>\n'
                '<kml xmlns="http://www.opengis.net/kml/2.2">\n'
                f'  <Document><name>{escape(name)}</name>\n')
    def kml_footer(): return '  </Document>\n</kml>\n'
    def kml_folder(name): return f'    <Folder><name>{escape(name)}</name>\n'
    def kml_folder_end(): return '    </Folder>\n'
    def kml_placemark(name, lon, lat, elev=None, desc=None):
        coords = f"{lon:.6f},{lat:.6f}" + (f",{elev:.2f}" if elev is not None and not np.isnan(elev) else "")
        dtag = f"        <description>{escape(desc)}</description>\n" if desc else ""
        return ("      <Placemark>\n"
                f"        <name>{escape(name)}</name>\n"
                f"{dtag}        <Point><coordinates>{coords}</coordinates></Point>\n"
                "      </Placemark>\n")

    kml_bio = BytesIO()
    kml_bio.write(kml_header("Excel → KMZ — NAD83 Geographic (Toronto grid)").encode("utf-8"))
    for folder, g in df.groupby(folder_col):
        kml_bio.write(kml_folder(str(folder)).encode("utf-8"))
        for _, r in g.iterrows():
            nm = str(r[name_col]) if pd.notna(r[name_col]) else "Unnamed"
            elev = r[z_col] if z_col else None
            if not np.isnan(r["lon_4269"]) and not np.isnan(r["lat_4269"]):
                kml_bio.write(kml_placemark(nm, r["lon_4269"], r["lat_4269"], elev,
                                            f"{r[folder_col]} | {r['grid_status']}").encode("utf-8"))
        kml_bio.write(kml_folder_end().encode("utf-8"))
    kml_bio.write(kml_footer().encode("utf-8"))
    kml_bio.seek(0)

    kmz_bio = BytesIO()
    with zipfile.ZipFile(kmz_bio, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr("doc.kml", kml_bio.getvalue())
    kmz_bio.seek(0)

    base = os.path.splitext(up.name)[0]
    st.download_button("Download KML", data=kml_bio.getvalue(), file_name=f"{base}_converted.kml", mime="application/vnd.google-earth.kml+xml")
    st.download_button("Download KMZ", data=kmz_bio.getvalue(), file_name=f"{base}_converted.kmz", mime="application/vnd.google-earth.kmz")
