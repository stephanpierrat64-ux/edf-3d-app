"""
Préparation des données pour le déploiement Streamlit exploitants agricoles.
À lancer une fois en local après chaque mise à jour des données sources.

Source de vérité : C:/Users/steph/Documents/3D EDF/xls/Agriculteurs EDF.xlsx
RP  : C:/Users/steph/Documents/3D EDF/S3/All_20260425.kmz
SP  : C:/Users/steph/Documents/3D EDF/S3/Preplan V6 geosources.shp

Génère dans data/ :
  exploitants.csv          – fiche contact des 81 exploitants (depuis l'Excel)
  receivers_agri.csv       – points RP avec coordonnées WGS84
  sources_agri.csv         – points SP avec coordonnées WGS84
  parcelles_agri.geojson   – polygones parcelles des 81 exploitants
"""

import os, re, unicodedata, zipfile, tempfile, glob, warnings
import xml.etree.ElementTree as ET
import pandas as pd
import geopandas as gpd
from shapely.geometry import Point

os.makedirs("data", exist_ok=True)

BASE_EDF      = "C:/Users/steph/Documents/3D EDF"
EXCEL_AGRI    = f"{BASE_EDF}/xls/Agriculteurs EDF.xlsx"
RCV_KMZ       = f"{BASE_EDF}/S3/All_20260425.kmz"
SOURCES_SHP   = f"{BASE_EDF}/S3/Preplan V6 geosources.shp"
PERMIT_ZIP    = "Permit avancement (44).zip"
INTERSECT_CSV = "intersection_rcv_permit_2004.csv"


# ── Utilitaires ───────────────────────────────────────────────────────────────

def normalize(s):
    if not s:
        return ""
    s = str(s).strip().upper()
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = re.sub(r"[\-\.\,\;\:\'\"]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def make_key(nom, prenom):
    return f"{normalize(nom)}|{normalize(prenom)}"


def format_phone(val):
    if pd.isna(val) or str(val).strip() in ("", "nan"):
        return ""
    s = str(val).split(".")[0].strip()
    if len(s) == 9:
        s = "0" + s
    return " ".join(s[i:i+2] for i in range(0, len(s), 2))


def parse_kml_desc(text):
    """Extrait les paires clé/valeur du tableau HTML d'une description KML."""
    if not text:
        return {}
    pairs = re.findall(r"<td>([^<]+)</td><td>([^<]*)</td>", text)
    return dict(pairs)


# ── 0. EXCEL AGRICULTEURS ─────────────────────────────────────────────────────
print("=" * 55)
print("0. EXCEL AGRICULTEURS")
print("=" * 55)

xl = pd.read_excel(EXCEL_AGRI)
xl.columns = [c.strip() for c in xl.columns]
phone_col = next(
    (c for c in xl.columns if "l" in c.lower() and
     any(w in c.lower() for w in ("phone", "l phone", "phon"))),
    next((c for c in xl.columns if c.lower() in
          ("telephone", "téléphone", "tel", "tél")), None)
)

xl["agri_key"]     = xl.apply(lambda r: make_key(r["NOM"], r["PRENOM"]), axis=1)
xl["agri_display"] = xl.apply(lambda r: f"{r['NOM'].strip()} {r['PRENOM'].strip()}", axis=1)
xl["TEL_FORMAT"]   = xl[phone_col].apply(format_phone) if phone_col else ""

exp_cols = ["NOM", "PRENOM", "agri_key", "agri_display",
            "Cadagri_26", "CULTURES 2", "STATUT", "STATUT DET",
            "CONSIGNES", "CONSIGNE_1", "surface_2", "agric", "TEL_FORMAT"]
exp_cols = [c for c in exp_cols if c in xl.columns]
xl[exp_cols].to_csv("data/exploitants.csv", index=False)
print(f"   {len(xl)} exploitants -> data/exploitants.csv")

agri_exact  = set(xl["agri_key"].tolist())
agri_by_nom = {}
for _, row in xl.iterrows():
    agri_by_nom.setdefault(normalize(row["NOM"]), []).append(normalize(row["PRENOM"]))

cadagri_to_key = {}
if "Cadagri_26" in xl.columns:
    for _, row in xl.iterrows():
        cad = str(row["Cadagri_26"]).strip()
        if cad and cad != "nan":
            cadagri_to_key[cad] = row["agri_key"]

key_to_display = dict(zip(xl["agri_key"], xl["agri_display"]))


def match_key(nom, prenom, cadagri=None):
    if cadagri:
        cad = str(cadagri).strip()
        if cad in cadagri_to_key:
            return cadagri_to_key[cad]
    nn, pp = normalize(nom), normalize(prenom)
    key = f"{nn}|{pp}"
    if key in agri_exact:
        return key
    if nn in agri_by_nom:
        for pref in agri_by_nom[nn]:
            if pref == pp or pref in pp or pp in pref:
                return f"{nn}|{pref}"
            if set(pref.split()) == set(pp.split()) and pref.split():
                return f"{nn}|{pref}"
    return None


# ── 1. PARCELLES (nécessaire avant les joins spatiaux) ────────────────────────
print("=" * 55)
print("1. PARCELLES")
print("=" * 55)

with zipfile.ZipFile(PERMIT_ZIP) as z:
    tmp_permit = tempfile.mkdtemp()
    z.extractall(tmp_permit)
shp_files = glob.glob(f"{tmp_permit}/*.shp")
if not shp_files:
    raise FileNotFoundError(f"Aucun .shp dans {PERMIT_ZIP}")

with warnings.catch_warnings():
    warnings.simplefilter("ignore")
    parcelles = gpd.read_file(shp_files[0])

print(f"   {len(parcelles)} parcelles chargees")
parcelles_wgs = parcelles.to_crs("EPSG:4326")

idx_agri = []
for idx, r in parcelles_wgs.iterrows():
    key = match_key(r.get("NOM","") or "", r.get("PRENOM","") or "",
                    r.get("Cadagri_26","") or "")
    if key:
        idx_agri.append(idx)

parcelles_agri = parcelles_wgs.loc[idx_agri].copy()
parcelles_agri["agri_key"] = parcelles_agri.apply(
    lambda r: match_key(r.get("NOM","") or "", r.get("PRENOM","") or "",
                        r.get("Cadagri_26","") or ""), axis=1)
parcelles_agri["agri_display"] = parcelles_agri["agri_key"].map(key_to_display).fillna("")

keep = [c for c in ["NOM","PRENOM","Cadagri_26","NOM_COM","STATUT","STATUT DET",
                     "CULTURES 2","CONSIGNES","CONSIGNE_1","agri_key","agri_display",
                     "geometry"] if c in parcelles_agri.columns]
parcelles_agri[keep].to_file("data/parcelles_agri.geojson", driver="GeoJSON")
n_exp_p = parcelles_agri["agri_key"].nunique()
print(f"   -> {len(parcelles_agri)} parcelles pour {n_exp_p} exploitants  ->  data/parcelles_agri.geojson")

# Préparer pour les joins spatiaux (géométries valides)
parcelles_join = parcelles_agri[["agri_key","agri_display","NOM","PRENOM",
                                  "Cadagri_26","STATUT","STATUT DET",
                                  "CULTURES 2","CONSIGNES","geometry"]].copy()
parcelles_join["geometry"] = parcelles_join.geometry.buffer(0)


# ── 2. RECEIVERS RP (depuis KMZ) ──────────────────────────────────────────────
print("=" * 55)
print("2. RECEIVERS RP (KMZ)")
print("=" * 55)

tmp_kmz = tempfile.mkdtemp()
with zipfile.ZipFile(RCV_KMZ) as z:
    z.extractall(tmp_kmz)
kml_path = os.path.join(tmp_kmz, "doc.kml")

tree = ET.parse(kml_path)
root = tree.getroot()
ns = {"kml": "http://www.opengis.net/kml/2.2"}

rows_rp = []
for pm in root.findall(".//kml:Placemark", ns):
    desc = pm.find("kml:description", ns)
    pt   = pm.find(".//kml:Point/kml:coordinates", ns)
    if pt is None:
        continue
    coords = pt.text.strip().split(",")
    lon, lat = float(coords[0]), float(coords[1])
    attrs = parse_kml_desc(desc.text if desc is not None else "")
    rows_rp.append({
        "station":       attrs.get("station", ""),
        "line":          attrs.get("line", ""),
        "point":         attrs.get("point", ""),
        "state":         attrs.get("state", ""),
        "rs_code":       attrs.get("rs_code", ""),
        "receiver_type": attrs.get("receiver_type", ""),
        "lat": lat,
        "lon": lon,
    })

rp_df = pd.DataFrame(rows_rp)
print(f"   {len(rp_df)} RP extraits du KMZ")

# Convertir en GeoDataFrame WGS84 pour la jointure spatiale
rp_gdf = gpd.GeoDataFrame(
    rp_df,
    geometry=[Point(r.lon, r.lat) for _, r in rp_df.iterrows()],
    crs="EPSG:4326",
)

# Jointure spatiale RP × parcelles agriculteurs
rp_joined = gpd.sjoin(
    rp_gdf,
    parcelles_join,
    how="inner",
    predicate="within",
)
rp_joined = rp_joined.drop(columns=["geometry","index_right"], errors="ignore")
rp_joined["agri_display"] = rp_joined["agri_key"].map(key_to_display).fillna("")
rp_joined.to_csv("data/receivers_agri.csv", index=False)
n_exp_r = rp_joined["agri_key"].nunique()
print(f"   -> {len(rp_joined)} RP pour {n_exp_r} exploitants  ->  data/receivers_agri.csv")


# ── 3. SOURCES SP (nouveau shapefile) ─────────────────────────────────────────
print("=" * 55)
print("3. SOURCES SP (Preplan V6)")
print("=" * 55)

sources = gpd.read_file(SOURCES_SHP)
print(f"   {len(sources)} sources chargees (CRS {sources.crs})")
sources_wgs = sources.to_crs("EPSG:4326")
sources_wgs["lat"] = sources_wgs.geometry.y
sources_wgs["lon"] = sources_wgs.geometry.x

keep_sp = [c for c in ["id","type","station","line","point","source_typ",
                        "rs_code","status","densite","commune","lat","lon",
                        "geometry"] if c in sources_wgs.columns]

src_joined = gpd.sjoin(
    sources_wgs[keep_sp],
    parcelles_join[["agri_key","agri_display","geometry"]],
    how="inner",
    predicate="within",
)
src_joined = src_joined.drop(columns=["geometry","index_right"], errors="ignore")
src_joined.to_csv("data/sources_agri.csv", index=False)
n_exp_s = src_joined["agri_key"].nunique()
print(f"   -> {len(src_joined)} SP pour {n_exp_s} exploitants  ->  data/sources_agri.csv")


# ── Rapport final ─────────────────────────────────────────────────────────────
print("\n" + "=" * 55)
print("RESUME")
print("=" * 55)
all_keys  = set(xl["agri_key"].tolist())
rp_keys   = set(rp_joined["agri_key"].unique()) if not rp_joined.empty else set()
missing   = all_keys - rp_keys
if missing:
    print(f"[!] {len(missing)} exploitants sans RP :")
    for k in sorted(missing):
        d = xl[xl["agri_key"]==k]["agri_display"].iloc[0] if k in xl["agri_key"].values else k
        print(f"    {d}")
else:
    print("[OK] Les 81 exploitants ont au moins un RP")

print(f"\nFichiers generes dans data/")
print(f"  exploitants.csv       : {len(xl)} lignes")
print(f"  receivers_agri.csv    : {len(rp_joined)} lignes")
print(f"  sources_agri.csv      : {len(src_joined)} lignes")
print(f"  parcelles_agri.geojson: {len(parcelles_agri)} entites")
