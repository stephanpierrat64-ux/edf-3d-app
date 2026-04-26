"""
Préparation des données pour le déploiement Streamlit exploitants agricoles.
À lancer une fois en local après chaque mise à jour des données sources.

Source de vérité : C:/Users/steph/Documents/3D EDF/xls/Agriculteurs EDF.xlsx

Génère dans data/ :
  exploitants.csv          – fiche contact des 81 exploitants (depuis l'Excel)
  receivers_agri.csv       – points RP avec coordonnées WGS84
  sources_agri.csv         – points SP avec coordonnées WGS84
  parcelles_agri.geojson   – polygones parcelles des 81 exploitants
"""

import os, re, unicodedata, zipfile, tempfile, glob, warnings
import pandas as pd
import geopandas as gpd

os.makedirs("data", exist_ok=True)

BASE_EDF      = "C:/Users/steph/Documents/3D EDF"
EXCEL_AGRI    = f"{BASE_EDF}/xls/Agriculteurs EDF.xlsx"
RCV_GPKG      = f"{BASE_EDF}/S3/rcv_summary_2004.gpkg"
SOURCES_SHP   = "Sources Filtrées.shp"
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
    """695440179 -> '06 95 44 01 79'"""
    if pd.isna(val) or str(val).strip() in ("", "nan"):
        return ""
    s = str(val).split(".")[0].strip()   # enlever .0 si float
    if len(s) == 9:
        s = "0" + s
    return " ".join(s[i:i+2] for i in range(0, len(s), 2))


# ── 0. LECTURE DE L'EXCEL ─────────────────────────────────────────────────────
print("=" * 55)
print("0. EXCEL AGRICULTEURS")
print("=" * 55)

xl = pd.read_excel(EXCEL_AGRI)
# Normaliser le nom de colonne téléphone (accent possible)
xl.columns = [c.strip() for c in xl.columns]
phone_col = next((c for c in xl.columns if "l" in c.lower() and "phone" in c.lower()
                  or c.lower() in ("telephone", "téléphone", "tel", "tél")), None)
if phone_col is None:
    # Chercher par position ou contenu
    for c in xl.columns:
        if "l" in c.lower():
            phone_col = c
            break

xl["agri_key"]     = xl.apply(lambda r: make_key(r["NOM"], r["PRENOM"]), axis=1)
xl["agri_display"] = xl.apply(lambda r: f"{r['NOM'].strip()} {r['PRENOM'].strip()}", axis=1)
if phone_col:
    xl["TEL_FORMAT"] = xl[phone_col].apply(format_phone)
else:
    xl["TEL_FORMAT"] = ""

# Sauvegarder la fiche exploitants
exp_cols = ["NOM", "PRENOM", "agri_key", "agri_display",
            "Cadagri_26", "CULTURES 2", "STATUT", "STATUT DET",
            "CONSIGNES", "CONSIGNE_1", "surface_2", "agric", "TEL_FORMAT"]
exp_cols = [c for c in exp_cols if c in xl.columns]
xl[exp_cols].to_csv("data/exploitants.csv", index=False)
print(f"   {len(xl)} exploitants -> data/exploitants.csv")
if phone_col:
    print(f"   Téléphone lu depuis colonne : '{phone_col}'")

# Index de matching
agri_exact  = set(xl["agri_key"].tolist())
agri_by_nom = {}
for _, row in xl.iterrows():
    nn = normalize(row["NOM"])
    pp = normalize(row["PRENOM"])
    agri_by_nom.setdefault(nn, []).append(pp)

# Dictionnaire Cadagri_26 -> agri_key (pour jointure directe)
cadagri_to_key = {}
if "Cadagri_26" in xl.columns:
    for _, row in xl.iterrows():
        cad = str(row["Cadagri_26"]).strip()
        if cad and cad != "nan":
            cadagri_to_key[cad] = row["agri_key"]


def match_key(nom, prenom, cadagri=None):
    """Retourne agri_key si exploitant reconnu, sinon None.
    Essaie d'abord par Cadagri_26, puis NOM/PRENOM."""
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


# ── 1. RECEIVERS ─────────────────────────────────────────────────────────────
print("=" * 55)
print("1. RECEIVERS")
print("=" * 55)

rcv = gpd.read_file(RCV_GPKG)
print(f"   {len(rcv)} points charges (CRS {rcv.crs})")
rcv_wgs = rcv.to_crs("EPSG:4326")
rcv_wgs["lat"] = rcv_wgs.geometry.y
rcv_wgs["lon"]  = rcv_wgs.geometry.x
rcv_coords = rcv_wgs[["station", "lat", "lon"]].copy()
rcv_coords["station"] = rcv_coords["station"].astype(str).str.strip()

join_df = pd.read_csv(INTERSECT_CSV, dtype=str).fillna("")
join_df.columns = [c.lstrip("﻿").strip() for c in join_df.columns]
join_df["station"] = join_df["station"].str.strip()

merged = rcv_coords.merge(join_df, on="station", how="inner")
print(f"   {len(merged)} lignes apres jointure sur station")

rows = []
for _, r in merged.iterrows():
    key = match_key(r.get("NOM", ""), r.get("PRENOM", ""), r.get("Cadagri_26", ""))
    if key:
        d = r.to_dict()
        d["agri_key"] = key
        # Récupérer le display depuis l'Excel
        exp_row = xl[xl["agri_key"] == key]
        d["agri_display"] = exp_row["agri_display"].iloc[0] if not exp_row.empty else f"{r.get('NOM','')} {r.get('PRENOM','')}".strip()
        rows.append(d)

receivers_agri = pd.DataFrame(rows)
receivers_agri.to_csv("data/receivers_agri.csv", index=False)
n_exp = receivers_agri["agri_key"].nunique() if not receivers_agri.empty else 0
print(f"   -> {len(receivers_agri)} RP pour {n_exp} exploitants  ->  data/receivers_agri.csv")


# ── 2. PARCELLES ─────────────────────────────────────────────────────────────
print("=" * 55)
print("2. PARCELLES")
print("=" * 55)

with zipfile.ZipFile(PERMIT_ZIP) as z:
    tmp = tempfile.mkdtemp()
    z.extractall(tmp)
shp_files = glob.glob(f"{tmp}/*.shp")
if not shp_files:
    raise FileNotFoundError(f"Aucun .shp dans {PERMIT_ZIP}")

with warnings.catch_warnings():
    warnings.simplefilter("ignore")
    parcelles = gpd.read_file(shp_files[0])

print(f"   {len(parcelles)} parcelles chargees")
parcelles_wgs = parcelles.to_crs("EPSG:4326")

rows_p = []
for idx, r in parcelles_wgs.iterrows():
    key = match_key(r.get("NOM", "") or "", r.get("PRENOM", "") or "", r.get("Cadagri_26", "") or "")
    if key:
        rows_p.append(idx)

parcelles_agri = parcelles_wgs.loc[rows_p].copy()
parcelles_agri["agri_key"] = parcelles_agri.apply(
    lambda r: match_key(r.get("NOM", "") or "", r.get("PRENOM", "") or "", r.get("Cadagri_26", "") or ""),
    axis=1,
)
parcelles_agri["agri_display"] = parcelles_agri["agri_key"].map(
    dict(zip(xl["agri_key"], xl["agri_display"]))
).fillna("")

keep = [c for c in [
    "NOM", "PRENOM", "Cadagri_26", "NOM_COM", "STATUT", "STATUT DET",
    "CULTURES 2", "CONSIGNES", "CONSIGNE_1", "agri_key", "agri_display", "geometry"
] if c in parcelles_agri.columns]

parcelles_agri[keep].to_file("data/parcelles_agri.geojson", driver="GeoJSON")
n_exp_p = parcelles_agri["agri_key"].nunique()
print(f"   -> {len(parcelles_agri)} parcelles pour {n_exp_p} exploitants  ->  data/parcelles_agri.geojson")


# ── 3. SOURCES (SP) ───────────────────────────────────────────────────────────
print("=" * 55)
print("3. SOURCES")
print("=" * 55)
sources = gpd.read_file(SOURCES_SHP)
print(f"   {len(sources)} sources chargees (CRS {sources.crs})")

sources = sources.copy()
sources["geometry"] = sources.geometry.centroid
sources_wgs = sources.to_crs("EPSG:4326")
sources_wgs["lat"] = sources_wgs.geometry.y
sources_wgs["lon"]  = sources_wgs.geometry.x

parcelles_join = parcelles_agri[["agri_key", "agri_display", "geometry"]].copy()
parcelles_join["geometry"] = parcelles_join.geometry.buffer(0)

src_joined = gpd.sjoin(
    sources_wgs[["id", "type", "Densite", "semaine", "lat", "lon", "geometry"]],
    parcelles_join,
    how="inner",
    predicate="within",
)
src_joined = src_joined.drop(columns=["geometry", "index_right"], errors="ignore")
src_joined.to_csv("data/sources_agri.csv", index=False)
n_exp_s = src_joined["agri_key"].nunique() if not src_joined.empty else 0
print(f"   -> {len(src_joined)} SP pour {n_exp_s} exploitants  ->  data/sources_agri.csv")


# ── Rapport final ─────────────────────────────────────────────────────────────
print("\n" + "=" * 55)
print("RESUME")
print("=" * 55)
if not receivers_agri.empty:
    matched = set(receivers_agri["agri_key"].unique())
    all_keys = set(xl["agri_key"].tolist())
    missing = all_keys - matched
    if missing:
        print(f"[!] {len(missing)} exploitants sans RP :")
        for k in sorted(missing):
            display = xl[xl["agri_key"] == k]["agri_display"].iloc[0] if k in xl["agri_key"].values else k
            print(f"    {display}")
    else:
        print("[OK] Les 81 exploitants ont au moins un RP")

print(f"\nFichiers generes dans data/")
print(f"  exploitants.csv       : {len(xl)} lignes")
print(f"  receivers_agri.csv    : {len(receivers_agri)} lignes")
print(f"  sources_agri.csv      : {len(src_joined)} lignes")
print(f"  parcelles_agri.geojson: {len(parcelles_agri)} entites")
