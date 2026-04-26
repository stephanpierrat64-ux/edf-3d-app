"""
Génère all_exploitants.kml : tous les 81 exploitants, leurs parcelles,
leurs points RP et SP — prêt pour import dans Google MyMaps.

Nécessite : pip install simplekml
Lancer depuis le dossier Avancement/ (les fichiers data/ doivent être présents).
"""

import json
import pandas as pd
import simplekml

# ── Chargement des données ────────────────────────────────────────────────────
rcv_df = pd.read_csv("data/receivers_agri.csv", dtype=str).fillna("")
src_df = pd.read_csv("data/sources_agri.csv",   dtype=str).fillna("")
exp_df = pd.read_csv("data/exploitants.csv",     dtype=str).fillna("")

with open("data/parcelles_agri.geojson", encoding="utf-8") as f:
    parcelles_json = json.load(f)

exp_lookup = {row["agri_key"]: row for _, row in exp_df.iterrows()}

print(f"Exploitants  : {len(exp_df)}")
print(f"Parcelles    : {len(parcelles_json.get('features', []))}")
print(f"Receivers RP : {len(rcv_df)}")
print(f"Sources SP   : {len(src_df)}")

# ── KML ───────────────────────────────────────────────────────────────────────
kml = simplekml.Kml()
kml.document.name = "Exploitants agricoles EDF 3D 2026"

# Styles polygones par statut
def poly_style(statut):
    FILL = {
        "OK":  simplekml.Color.changealpha("70", simplekml.Color.green),
        "PVG": simplekml.Color.changealpha("70", simplekml.Color.yellow),
        "PG":  simplekml.Color.changealpha("70", simplekml.Color.yellow),
        "NC":  simplekml.Color.changealpha("70", simplekml.Color.red),
        "PC":  simplekml.Color.changealpha("70", simplekml.Color.cyan),
        "R":   simplekml.Color.changealpha("70", simplekml.Color.purple),
        "NR":  simplekml.Color.changealpha("70", simplekml.Color.lightgrey),
    }
    LINE = {
        "OK":  simplekml.Color.darkgreen,
        "PVG": simplekml.Color.orange,
        "PG":  simplekml.Color.orange,
        "NC":  simplekml.Color.red,
        "PC":  simplekml.Color.blue,
        "R":   simplekml.Color.purple,
        "NR":  simplekml.Color.grey,
    }
    s = simplekml.Style()
    s.polystyle.color = FILL.get(statut, simplekml.Color.changealpha("70", simplekml.Color.yellow))
    s.linestyle.color = LINE.get(statut, simplekml.Color.orange)
    s.linestyle.width = 2
    return s

# Style RP (petit point bleu)
sty_rp = simplekml.Style()
sty_rp.iconstyle.color = simplekml.Color.blue
sty_rp.iconstyle.scale = 0.7
sty_rp.iconstyle.icon.href = (
    "http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png"
)
sty_rp.labelstyle.scale = 0

# Style SP (petit triangle rouge)
sty_sp = simplekml.Style()
sty_sp.iconstyle.color = simplekml.Color.red
sty_sp.iconstyle.scale = 0.7
sty_sp.iconstyle.icon.href = (
    "http://maps.google.com/mapfiles/kml/shapes/triangle.png"
)
sty_sp.labelstyle.scale = 0

# ── Dossier Parcelles ─────────────────────────────────────────────────────────
fld_parc = kml.newfolder(name="Parcelles")
n_parc = 0

for feat in parcelles_json.get("features", []):
    geom  = feat.get("geometry", {})
    props = feat.get("properties", {})
    gtype = geom.get("type", "")
    if gtype not in ("Polygon", "MultiPolygon"):
        continue

    agri_key = props.get("agri_key", "")
    exp = exp_lookup.get(agri_key, {})

    nom    = props.get("NOM", "").strip()
    prenom = props.get("PRENOM", "").strip()
    statut = str(props.get("STATUT", "")).strip().upper()
    tel    = str(exp.get("TEL_FORMAT", "")).strip()

    desc_lines = [
        f"Exploitant : {nom} {prenom}",
    ]
    if tel and tel != "nan":
        desc_lines.append(f"Tel : {tel}")
    for key, label in [
        ("NOM_COM",    "Commune"),
        ("CULTURES 2", "Culture"),
        ("STATUT DET", "Statut"),
        ("CONSIGNES",  "Consignes"),
        ("Cadagri_26", "Cadagri"),
    ]:
        val = str(props.get(key, "")).strip()
        if val and val != "nan":
            desc_lines.append(f"{label} : {val}")

    description = "\n".join(desc_lines)
    sty = poly_style(statut)

    coords_raw = geom.get("coordinates", [])
    polys = [coords_raw] if gtype == "Polygon" else coords_raw

    for poly_coords in polys:
        outer = poly_coords[0] if poly_coords else []
        pol = fld_parc.newpolygon(
            name=f"{nom} {prenom}",
            outerboundaryis=[(c[0], c[1]) for c in outer],
        )
        pol.style = sty
        pol.description = description
        n_parc += 1

# ── Dossier Receivers RP ──────────────────────────────────────────────────────
fld_rp = kml.newfolder(name="Receivers RP")
n_rp = 0

for _, row in rcv_df.iterrows():
    try:
        lat, lon = float(row["lat"]), float(row["lon"])
    except (ValueError, KeyError):
        continue

    agri_key  = row.get("agri_key", "")
    exp       = exp_lookup.get(agri_key, {})
    nom_pren  = str(exp.get("agri_display", "")).strip() or agri_key
    tel       = str(exp.get("TEL_FORMAT", "")).strip()

    desc_lines = [f"Exploitant : {nom_pren}"]
    if tel and tel != "nan":
        desc_lines.append(f"Tel : {tel}")
    for key, label in [
        ("line",      "Ligne"),
        ("point",     "Point"),
        ("state",     "Etat"),
        ("STATUT DET","Statut"),
        ("CULTURES 2","Culture"),
        ("CONSIGNES", "Consignes"),
    ]:
        val = str(row.get(key, "")).strip()
        if val and val != "nan":
            desc_lines.append(f"{label} : {val}")

    pt = fld_rp.newpoint(
        name=str(row.get("station", "")),
        coords=[(lon, lat)],
    )
    pt.style = sty_rp
    pt.description = "\n".join(desc_lines)
    n_rp += 1

# ── Dossier Sources SP ────────────────────────────────────────────────────────
fld_sp = kml.newfolder(name="Sources SP")
n_sp = 0

for _, row in src_df.iterrows():
    try:
        lat, lon = float(row["lat"]), float(row["lon"])
    except (ValueError, KeyError):
        continue

    agri_key = row.get("agri_key", "")
    exp      = exp_lookup.get(agri_key, {})
    nom_pren = str(exp.get("agri_display", "")).strip() or agri_key

    desc_lines = [f"Exploitant : {nom_pren}"]
    for key, label in [
        ("source_typ", "Type"),
        ("status",     "Statut"),
        ("commune",    "Commune"),
    ]:
        val = str(row.get(key, "")).strip()
        if val and val != "nan":
            desc_lines.append(f"{label} : {val}")

    station = str(row.get("station", row.get("id", ""))).strip()
    pt = fld_sp.newpoint(
        name=f"SP {station}",
        coords=[(lon, lat)],
    )
    pt.style = sty_sp
    pt.description = "\n".join(desc_lines)
    n_sp += 1

# ── Sauvegarde ────────────────────────────────────────────────────────────────
OUTPUT = "all_exploitants.kml"
kml.save(OUTPUT)

print(f"\nKML genere : {OUTPUT}")
print(f"  {n_parc} polygones parcelles")
print(f"  {n_rp} points RP")
print(f"  {n_sp} points SP")
print(f"\nImporter dans MyMaps :")
print(f"  1. Ouvrir maps.google.com > MyMaps > Creer une carte")
print(f"  2. Importer > selectionner {OUTPUT}")
print(f"  3. Choisir le champ de nom : NOM ou PRENOM")
