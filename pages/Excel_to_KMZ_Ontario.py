# pages/Excel_to_KMZ_Ontario.py
import streamlit as st, pandas as pd, numpy as np, zipfile, os
from io import BytesIO
from pathlib import Path
from pyproj import Transformer
from xml.sax.saxutils import escape

st.set_page_config(page_title="Excel → KMZ (Ontario NTv2 grids)", page_icon="🗂️")
st.title("Excel → KMZ — NAD27 UTM 17N → NAD83 Geographic (Ontario NTv2)")

st.subheader("1) Upload Ontario grid (.gsb)")
st.caption("Most common for province-wide NAD27→NAD83: **ON27CSv1.gsb**. You may also try ON76CSv1.gsb if legacy datasets were used. ON83CSv1.gsb is generally not for 27→83 forward shifts.")
grid_up = st.file_uploader("Upload Ontario NTv2 grid (.gsb)", type=["gsb"])

GRID_PATH = None
if grid_up:
    tmp = Path("uploaded_ontario_grid.gsb")
    tmp.write_bytes(grid_up.read())
    GRID_PATH = str(tmp.resolve())
    # Tell PROJ where to find the grid
    os.environ["PROJ_DATA"] = str(tmp.parent.resolve())
    os.environ["PROJ_NETWORK"] = "OFF"

if GRID_PATH is None:
    st.stop()

st.subheader("2) Upload Excel")
up = st.file_uploader("Upload Excel (columns: folder, feature_name, easting, northing; optional elevation)", type=["xlsx","xls"])
if not up:
    st.stop()

# Read & normalize
df0 = pd.read_excel(up)
if df0.empty:
    st.error("Excel has no rows."); st.stop()

df = df0.copy()
df.columns = [str(c).strip().lower().replace(" ", "_") for c in df.columns]
cols = list(df.columns)

st.subheader("3) Map columns (avoid Easting/Northing swap)")
def pick(label, options, default=None):
    idx = options.index(default) if default in options else 0
    return st.selectbox(label, options, index=idx)

folder_col = pick("Folder column", cols, "folder")
name_col   = pick("Name column", cols, "feature_name")
e_col      = pick("Easting (UTM m, NAD27 Zone 17N)", cols)
n_col      = pick("Northing (UTM m, NAD27 Zone 17N)", cols)
z_sel      = pick("Elevation (optional)", ["(none)"] + cols, "(none)")
z_col      = None if z_sel == "(none)" else z_sel

# Coerce numeric and drop empties
for c in [e_col, n_col] + ([z_col] if z_col else []):
    df[c] = pd.to_numeric(df[c], errors="coerce")
df = df.dropna(subset=[e_col, n_col]).copy()
if df.empty:
    st.error("All rows missing Easting/Northing after cleaning."); st.stop()

# Basic sanity for GTA area (adjust if your AOI differs)
warn = []
if not df[e_col].between(550000, 700000).any():
    warn.append("Easting outside typical UTM17N range near GTA (~610k).")
if not df[n_col].between(4_700_000, 5_400_000).any():
    warn.append("Northing outside typical GTA range (~4.83–4.84M).")
if warn: st.warning(" | ".join(warn))

st.subheader("4) Convert & Download")
if st.button("Convert to KML/KMZ"):
    try:
        # NAD27 / UTM 17N → NAD83 (geographic) with selected Ontario grid
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
            np.isnan(lon) | np.isnan(lat),
            "no_shift (outside grid/invalid)",
            "shift_ok"
        )

        # KML helpers
        def kml_header(name):
            return ('<?xml version="1.0" encoding="UTF-8"?>\n'
                    '<kml xmlns="http://www.opengis.net/kml/2.2">\n'
                    f'  <Document><name>{escape(name)}</name>\n')
        def kml_footer(): return '  </Document>\n</kml>\n'
        def kml_folder(name): return f'    <Folder><name>{escape(name)}</name>\n'
        def kml_folder_end(): return '    </Folder>\n'
        def kml_placemark(name, lon, lat, elev=None, desc=None):
            coords = f"{lon:.6f},{lat:.6f}" + (f",{elev:.2f}" if (elev is not None and not np.isnan(elev)) else "")
            dtag = f"        <description>{escape(desc)}</description>\n" if desc else ""
            return ("      <Placemark>\n"
                    f"        <name>{escape(name)}</name>\n"
                    f"{dtag}        <Point><coordinates>{coords}</coordinates></Point>\n"
                    "      </Placemark>\n")

        base = os.path.splitext(up.name)[0]
        kml = BytesIO(); kml.write(kml_header(f"{base} — NAD83 (Ontario grid: {Path(GRID_PATH).name})").encode("utf-8"))

        for folder, G in df.groupby(folder_col):
            kml.write(kml_folder(str(folder)).encode("utf-8"))
            for _, r in G.iterrows():
                nm = str(r[name_col]) if pd.notna(r[name_col]) else "Unnamed"
                elev = r[z_col] if z_col else None
                if not np.isnan(r["lon_4269"]) and not np.isnan(r["lat_4269"]):
                    desc = f"{r[folder_col]} | {r['grid_status']}"
                    kml.write(kml_placemark(nm, r["lon_4269"], r["lat_4269"], elev, desc).encode("utf-8"))
            kml.write(kml_folder_end().encode("utf-8"))

        kml.write(kml_footer().encode("utf-8")); kml.seek(0)

        kmz = BytesIO()
        with zipfile.ZipFile(kmz, "w", compression=zipfile.ZIP_DEFLATED) as z:
            z.writestr("doc.kml", kml.getvalue())
        kmz.seek(0)

        st.success(f"Done. Grid used: {Path(GRID_PATH).name}")
        st.download_button("Download KML", data=kml.getvalue(),
                           file_name=f"{base}_converted.kml",
                           mime="application/vnd.google-earth.kml+xml")
        st.download_button("Download KMZ", data=kmz.getvalue(),
                           file_name=f"{base}_converted.kmz",
                           mime="application/vnd.google-earth.kmz")
    except Exception as e:
        st.error("Conversion failed.")
        st.exception(e)
