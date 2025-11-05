# excel_to_kmz_app.py  — Excel (NAD27 / UTM 17N) -> KML/KMZ (NAD83 geo), grouped by folder
import streamlit as st, pandas as pd, numpy as np, zipfile, os
from io import BytesIO
from pathlib import Path
from pyproj import Transformer
from xml.sax.saxutils import escape

st.set_page_config(page_title="Excel → KMZ (Toronto grid)", page_icon="🗂️")
st.title("Excel to KMZ")

# --- Locate Toronto grid file robustly ---
APP_DIR = Path(__file__).parent.resolve()
CANDIDATES = [
    APP_DIR / "TO27CSv1.gsb",          # same folder as this file (works for single-file app)
    APP_DIR.parent / "TO27CSv1.gsb",   # repo root when this file is under /pages
    Path.cwd() / "TO27CSv1.gsb",       # Streamlit working dir
    APP_DIR.parent.parent / "TO27CSv1.gsb",
]
GRID = next((p for p in CANDIDATES if p.exists()), None)

st.caption(
    "Input: NAD27 UTM Zone 17N • Grid used: TO27CSv1.gsb "
    "• Output grouped by ‘folder’ column (one Folder per unique value)."
)

# Optional: allow a one-time grid upload if not found in repo
if GRID is None:
    st.warning("TO27CSv1.gsb not found. Upload it once here or place it at repo root.")
    grid_up = st.file_uploader("Upload TO27CSv1.gsb", type=["gsb"])
    if grid_up is not None:
        tmp = Path("TO27CSv1.gsb")
        tmp.write_bytes(grid_up.read())
        GRID = tmp

# Stop if no grid is available
if GRID is None or not GRID.exists():
    st.error("TO27CSv1.gsb not found in repo or upload. Place it at repo root or upload it above.")
    st.stop()

# Tell PROJ where to find the grid (belt & suspenders)
os.environ["PROJ_DATA"] = str(GRID.parent.resolve())
os.environ["PROJ_NETWORK"] = "OFF"
GRID_PATH = str(GRID.resolve())

# --- Upload Excel ---
up = st.file_uploader("Upload Excel (columns: folder, feature_name, easting, northing; optional elevation)", type=["xlsx","xls"])
if not up:
    st.stop()

if st.button("Convert"):
    try:
        # -------- Read Excel --------
        df0 = pd.read_excel(up)
        if df0.empty:
            st.error("No rows found in the uploaded Excel."); st.stop()

        df = df0.copy()
        df.columns = [str(c).strip().lower().replace(" ", "_") for c in df.columns]

        # Column detection helpers
        def pick(cols, opts):
            for o in opts:
                if o in cols: return o
            return None

        folder_col = pick(df.columns, ["folder","group","layer"])
        name_col   = pick(df.columns, ["feature_name","name","label","id","title"])
        e_col      = pick(df.columns, ["easting","e","utm_e","utm_easting","east","x"])
        n_col      = pick(df.columns, ["northing","n","utm_n","utm_northing","north","y"])
        z_col      = pick(df.columns, ["elevation","elev","elevation_m","z","altitude","height"])

        missing = [k for k,v in {
            "folder":folder_col, "feature_name":name_col, "easting":e_col, "northing":n_col
        }.items() if v is None]
        if missing:
            st.error(f"Missing required columns: {missing}. Found: {list(df.columns)}"); st.stop()

        # Clean numeric and drop empties
        df[e_col] = pd.to_numeric(df[e_col], errors="coerce")
        df[n_col] = pd.to_numeric(df[n_col], errors="coerce")
        if z_col: df[z_col] = pd.to_numeric(df[z_col], errors="coerce")
        df = df.dropna(subset=[e_col, n_col]).copy()
        if df.empty:
            st.error("All rows are missing easting/northing after cleaning."); st.stop()

        # -------- Transform: NAD27/UTM17N -> NAD83 geographic using Toronto grid --------
        pipeline = (
            f"+proj=pipeline "
            f"+step +inv +proj=utm +zone=17 +datum=NAD27 "
            f"+step +proj=hgridshift +grids={GRID_PATH} "
            f"+step +proj=unitconvert +xy_in=rad +xy_out=deg"
        )
        tr = Transformer.from_pipeline(pipeline)
        lon, lat = tr.transform(df[e_col].to_numpy(), df[n_col].to_numpy())
        df["lon_4269"] = lon
        df["lat_4269"] = lat
        df["grid_status"] = np.where(
            df[["lon_4269","lat_4269"]].isna().any(axis=1),
            "no_shift (outside TO27CSv1.gsb or invalid)", "shift_ok (TO27CSv1.gsb)"
        )

        # -------- Build KML in-memory (group by folder) --------
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

        base = os.path.splitext(up.name)[0]
        kml_bio = BytesIO()
        kml_bio.write(kml_header(f"{base} — NAD83 Geographic (Toronto grid)").encode("utf-8"))

        for folder, g in df.groupby(folder_col):
            kml_bio.write(kml_folder(str(folder)).encode("utf-8"))
            for _, r in g.iterrows():
                nm = str(r[name_col]) if pd.notna(r[name_col]) else "Unnamed"
                elev = r[z_col] if z_col else None
                if not np.isnan(r["lon_4269"]) and not np.isnan(r["lat_4269"]):
                    desc = f"{r[folder_col]} | {r['grid_status']}"
                    kml_bio.write(
                        kml_placemark(nm, r["lon_4269"], r["lat_4269"], elev, desc).encode("utf-8")
                    )
            kml_bio.write(kml_folder_end().encode("utf-8"))

        kml_bio.write(kml_footer().encode("utf-8"))
        kml_bio.seek(0)

        # Package as KMZ
        kmz_bio = BytesIO()
        with zipfile.ZipFile(kmz_bio, "w", compression=zipfile.ZIP_DEFLATED) as z:
            z.writestr("doc.kml", kml_bio.getvalue())
        kmz_bio.seek(0)

        # Downloads
        st.success(f"Done. Grid file used: {GRID.name}")
        st.download_button("Download KML", data=kml_bio.getvalue(),
                           file_name=f"{base}_converted.kml",
                           mime="application/vnd.google-earth.kml+xml")
        st.download_button("Download KMZ", data=kmz_bio.getvalue(),
                           file_name=f"{base}_converted.kmz",
                           mime="application/vnd.google-earth.kmz")

    except Exception as e:
        st.error("Conversion failed.")
        st.exception(e)
