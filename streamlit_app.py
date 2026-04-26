import json, io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="EDF 3D 2026",
    page_icon="📡",
    layout="wide",
)

# ── Navigation ────────────────────────────────────────────────────────────────
page = st.sidebar.radio(
    "Navigation",
    ["📡 Recherche station", "🌾 Exploitants agricoles"],
)

# ── Données partagées ────────────────────────────────────────────────────────

@st.cache_data
def load_station_data():
    df = pd.read_csv("intersection_rcv_permit_2004.csv", dtype=str).fillna("")
    df.columns = [c.lstrip("﻿").strip() for c in df.columns]
    return df


@st.cache_data
def load_agri_data():
    rcv = pd.read_csv("data/receivers_agri.csv", dtype=str).fillna("")
    src = pd.read_csv("data/sources_agri.csv", dtype=str).fillna("")
    exp = pd.read_csv("data/exploitants.csv", dtype=str).fillna("")
    with open("data/parcelles_agri.geojson", encoding="utf-8") as f:
        parcelles = json.load(f)
    return rcv, src, exp, parcelles


# ═════════════════════════════════════════════════════════════════════════════
# PAGE 1 – RECHERCHE STATION (existante, inchangée)
# ═════════════════════════════════════════════════════════════════════════════

LABELS = {
    "station": "Station", "line": "Ligne", "point": "Point",
    "Cadagri_26": "N° Cadagri", "NOM_COM": "Commune",
    "SECTION": "Section", "NUMERO": "N° cadastral",
    "CONTENANCE": "Surface (m²)", "NOM": "Nom propriétaire",
    "PRENOM": "Prénom", "ADRESSE": "Adresse", "CP": "Code postal",
    "VILLE": "Ville", "DATE": "Date contact", "TELEPHONE": "Téléphone",
    "EMAIL": "Email", "CULTURES 2": "Culture", "agric": "Culture agricole",
    "STATUT": "Statut", "STATUT DET": "Statut détaillé",
    "CONSIGNES": "Consignes", "CONSIGNE_1": "Consignes (suite)",
    "REMARQUES": "Remarques",
}

STATUT_COLORS = {
    "OK": "#d4edda", "PVG": "#fff3cd", "PG": "#fff3cd",
    "NC": "#f8d7da", "PV": "#d1ecf1", "NR": "#e2e3e5",
}

SECTIONS = [
    ("📡 Point sismique",  ["station", "line", "point"]),
    ("👤 Propriétaire",    ["Cadagri_26", "NOM_COM", "SECTION", "NUMERO",
                            "CONTENANCE", "NOM", "PRENOM", "ADRESSE", "CP", "VILLE"]),
    ("📞 Contact",         ["DATE", "TELEPHONE", "EMAIL"]),
    ("🌱 Terrain",         ["CULTURES 2", "agric", "STATUT", "STATUT DET",
                            "CONSIGNES", "CONSIGNE_1", "REMARQUES"]),
]

if page == "📡 Recherche station":
    st.title("📡 Recherche Station – EDF 3D 2026")
    df = load_station_data()
    st.caption(f"{len(df):,} stations indexées · {df['Cadagri_26'].ne('').sum():,} avec parcelle")

    query = st.text_input(
        "Numéro de station", placeholder="ex: 10005090", max_chars=20,
    ).strip()

    if not query:
        st.info("Saisissez un numéro de station pour afficher les informations de la parcelle.")
        st.stop()

    mask_exact = df["station"].str.strip() == query
    matches = df[mask_exact]
    if matches.empty:
        matches = df[df["station"].str.contains(query, na=False)]

    if matches.empty:
        st.error(f"Aucune station trouvée pour « {query} »")
        st.stop()

    if len(matches) > 1:
        options = matches["station"].tolist()
        selected = st.selectbox(f"{len(matches)} stations correspondent – choisissez :", options)
        row = matches[matches["station"] == selected].iloc[0]
    else:
        row = matches.iloc[0]

    statut = row.get("STATUT", "").strip().upper()
    bg = STATUT_COLORS.get(statut, "#f8f9fa")

    st.markdown(
        f"""
        <div style="background:{bg}; padding:14px 18px; border-radius:8px; margin-bottom:16px">
            <h2 style="margin:0">Station {row.get('station','—')}</h2>
            <p style="margin:4px 0 0; color:#555">
                {row.get('NOM_COM','')}&nbsp;&nbsp;·&nbsp;&nbsp;
                <strong>{row.get('STATUT DET', row.get('STATUT',''))}</strong>
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    for section_title, cols in SECTIONS:
        filled = {c: row.get(c, "").strip() for c in cols
                  if c in row.index and row.get(c, "").strip() and row.get(c, "").strip() != "nan"}
        if not filled:
            continue
        with st.expander(section_title, expanded=True):
            for col, val in filled.items():
                label = LABELS.get(col, col)
                c1, c2 = st.columns([1, 2])
                c1.markdown(f"**{label}**")
                if col == "TELEPHONE" and val:
                    c2.markdown(f"[{val}](tel:{val.replace(' ','')})")
                elif col == "EMAIL" and val:
                    c2.markdown(f"[{val}](mailto:{val})")
                else:
                    c2.write(val)


# ═════════════════════════════════════════════════════════════════════════════
# PAGE 2 – EXPLOITANTS AGRICOLES (carte)
# ═════════════════════════════════════════════════════════════════════════════

elif page == "🌾 Exploitants agricoles":
    import folium
    from streamlit_folium import st_folium
    import simplekml

    st.title("🌾 Exploitants agricoles – Carte")

    try:
        rcv_df, src_df, exp_df, parcelles_json = load_agri_data()
    except FileNotFoundError:
        st.error(
            "Données non disponibles. Lancer **prepare_agri_data.py** "
            "en local puis pousser le dossier `data/` sur GitHub."
        )
        st.stop()

    # ── Sélection exploitant ─────────────────────────────────────────────────
    # Liste depuis l'Excel (inclut les exploitants sans RP)
    exp_df_sorted = exp_df.sort_values("agri_display")
    key_to_display = dict(zip(exp_df_sorted["agri_key"], exp_df_sorted["agri_display"]))
    display_to_key = {v: k for k, v in key_to_display.items()}
    display_list   = exp_df_sorted["agri_display"].tolist()

    col_sel1, col_sel2 = st.columns([3, 2])
    with col_sel1:
        chosen_display = st.selectbox("Exploitant", display_list)
    with col_sel2:
        rp_query = st.text_input("Ou N° RP / station", placeholder="ex: 10005076").strip()

    # Recherche par N° station → déterminer l'exploitant
    if rp_query:
        match_rp = rcv_df[rcv_df["station"].str.strip() == rp_query]
        if match_rp.empty:
            match_rp = rcv_df[rcv_df["station"].str.contains(rp_query, na=False)]
        if not match_rp.empty:
            found_key = match_rp.iloc[0]["agri_key"]
            if found_key in key_to_display:
                chosen_display = key_to_display[found_key]
                st.success(f"RP {rp_query} → **{chosen_display}**")
        else:
            st.warning(f"Station « {rp_query} » non trouvée dans les exploitants agricoles.")

    selected_key = display_to_key.get(chosen_display, "")

    # ── Filtrage des données ──────────────────────────────────────────────────
    rcv_sel = rcv_df[rcv_df["agri_key"] == selected_key].copy()
    src_sel = src_df[src_df["agri_key"] == selected_key].copy()
    feats_sel = [f for f in parcelles_json.get("features", [])
                 if f.get("properties", {}).get("agri_key") == selected_key]
    parcelles_sel = {"type": "FeatureCollection", "features": feats_sel}

    # ── Mise en page ──────────────────────────────────────────────────────────
    col_map, col_info = st.columns([3, 1])

    # ── Panneau info (depuis l'Excel) ─────────────────────────────────────────
    exp_row = exp_df[exp_df["agri_key"] == selected_key]
    has_exp = not exp_row.empty
    e = exp_row.iloc[0] if has_exp else None

    STATUT_BG = {
        "OK": "#d4edda", "PVG": "#fff3cd", "PG": "#fff3cd",
        "NC": "#f8d7da", "PC": "#d1ecf1", "R": "#e8d5f5", "NR": "#e2e3e5",
    }
    STATUT_IC = {
        "OK": "🟢", "PVG": "🟡", "PG": "🟡",
        "NC": "🔴", "PC": "🔵", "R": "🟣", "NR": "⚪",
    }

    with col_info:
        statut_val = str(e["STATUT"]).strip().upper() if has_exp else ""
        bg_col = STATUT_BG.get(statut_val, "#f8f9fa")
        ic_col = STATUT_IC.get(statut_val, "⚫")
        statut_det = str(e["STATUT DET"]).strip() if has_exp and "STATUT DET" in e.index else statut_val

        st.markdown(
            f"<div style='background:{bg_col};padding:10px 12px;border-radius:8px;margin-bottom:8px'>"
            f"<b style='font-size:1.1em'>{chosen_display}</b><br>"
            f"<span style='color:#555'>{ic_col} {statut_det or statut_val}</span>"
            "</div>",
            unsafe_allow_html=True,
        )

        if has_exp:
            tel = str(e.get("TEL_FORMAT", "")).strip()
            if tel and tel != "nan":
                st.markdown(f"📞 [{tel}](tel:{tel.replace(' ','')})")
            culture = str(e.get("CULTURES 2", "") or "").strip()
            if not culture or culture == "nan":
                culture = str(e.get("agric", "") or "").strip()
            if culture and culture != "nan":
                st.write(f"🌱 {culture}")
            surf = str(e.get("surface_2", "")).strip()
            if surf and surf != "nan":
                try:
                    st.write(f"📐 {int(float(surf)):,} m²".replace(",", " "))
                except ValueError:
                    pass
            cadagri = str(e.get("Cadagri_26", "")).strip()
            if cadagri and cadagri != "nan":
                st.caption(f"Cadagri : {cadagri}")
            consignes = str(e.get("CONSIGNES", "")).strip()
            if consignes and consignes != "nan":
                st.warning(f"⚠️ {consignes}")

        st.divider()

        m1, m2, m3 = st.columns(3)
        m1.metric("RP", len(rcv_sel))
        m2.metric("SP", len(src_sel))
        m3.metric("Parcelles", len(feats_sel))

        st.divider()

        # ── Export KML ────────────────────────────────────────────────────────
        def build_kml():
            kml = simplekml.Kml()
            kml.document.name = chosen_display

            # Styles
            sty_rp = simplekml.Style()
            sty_rp.iconstyle.color = simplekml.Color.blue
            sty_rp.iconstyle.scale = 0.8
            sty_rp.iconstyle.icon.href = (
                "http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png"
            )

            sty_sp = simplekml.Style()
            sty_sp.iconstyle.color = simplekml.Color.red
            sty_sp.iconstyle.scale = 0.7
            sty_sp.iconstyle.icon.href = (
                "http://maps.google.com/mapfiles/kml/shapes/triangle.png"
            )

            sty_pol = simplekml.Style()
            sty_pol.polystyle.color = simplekml.Color.changealpha("80", simplekml.Color.yellow)
            sty_pol.linestyle.color = simplekml.Color.orange
            sty_pol.linestyle.width = 2

            # Dossier Parcelles
            fld_parc = kml.newfolder(name="Parcelles")
            for feat in feats_sel:
                geom = feat.get("geometry", {})
                props = feat.get("properties", {})
                gtype = geom.get("type", "")
                coords_raw = geom.get("coordinates", [])
                if gtype in ("Polygon", "MultiPolygon"):
                    polys = [coords_raw] if gtype == "Polygon" else coords_raw
                    for poly_coords in polys:
                        outer = poly_coords[0] if poly_coords else []
                        pol = fld_parc.newpolygon(
                            name=props.get("Cadagri_26", ""),
                            outerboundaryis=[(c[0], c[1]) for c in outer],
                        )
                        pol.style = sty_pol
                        desc_parts = [
                            f"Commune: {props.get('NOM_COM','')}",
                            f"Culture: {props.get('CULTURES 2','')}",
                            f"Statut: {props.get('STATUT','')} – {props.get('STATUT DET','')}",
                            f"Consignes: {props.get('CONSIGNES','')}",
                        ]
                        pol.description = "\n".join(p for p in desc_parts if p.split(": ")[1])

            # Dossier RP
            fld_rp = kml.newfolder(name="Receivers RP")
            for _, row in rcv_sel.iterrows():
                try:
                    lat, lon = float(row["lat"]), float(row["lon"])
                except (ValueError, KeyError):
                    continue
                pt = fld_rp.newpoint(name=str(row.get("station", "")), coords=[(lon, lat)])
                pt.style = sty_rp
                desc_parts = [
                    f"Ligne: {row.get('line','')}",
                    f"Point: {row.get('point','')}",
                    f"Statut: {row.get('STATUT','')} – {row.get('STATUT DET','')}",
                    f"Culture: {row.get('CULTURES 2','')}",
                    f"Consignes: {row.get('CONSIGNES','')}",
                ]
                pt.description = "\n".join(p for p in desc_parts if p.split(": ")[1])

            # Dossier SP
            fld_sp = kml.newfolder(name="Sources SP")
            for _, row in src_sel.iterrows():
                try:
                    lat, lon = float(row["lat"]), float(row["lon"])
                except (ValueError, KeyError):
                    continue
                pt = fld_sp.newpoint(
                    name=f"SP {row.get('id','')}", coords=[(lon, lat)]
                )
                pt.style = sty_sp
                pt.description = f"Type: {row.get('type','')}\nDensité: {row.get('Densite','')}"

            buf = io.BytesIO()
            buf.write(kml.kml().encode("utf-8"))
            buf.seek(0)
            return buf

        if not rcv_sel.empty or feats_sel:
            kml_buf = build_kml()
            fname = chosen_display.replace(" ", "_").replace("/", "-") + ".kml"
            st.download_button(
                "⬇️ Télécharger KML (MyMaps)",
                data=kml_buf,
                file_name=fname,
                mime="application/vnd.google-earth.kml+xml",
            )

    # ── Carte Folium ──────────────────────────────────────────────────────────
    with col_map:
        # Calcul du centre
        lats, lons = [], []
        for _, row in rcv_sel.iterrows():
            try:
                lats.append(float(row["lat"]))
                lons.append(float(row["lon"]))
            except (ValueError, KeyError):
                pass
        for feat in feats_sel:
            geom = feat.get("geometry", {})
            if geom.get("type") == "Polygon":
                for c in geom["coordinates"][0]:
                    lons.append(c[0]); lats.append(c[1])
            elif geom.get("type") == "MultiPolygon":
                for poly in geom["coordinates"]:
                    for c in poly[0]:
                        lons.append(c[0]); lats.append(c[1])

        if lats:
            center_lat = sum(lats) / len(lats)
            center_lon = sum(lons) / len(lons)
        else:
            center_lat, center_lon = 44.65, 4.78

        m = folium.Map(location=[center_lat, center_lon], zoom_start=14, control_scale=True)

        # Fond satellite Esri (sans clé API)
        folium.TileLayer(
            tiles=(
                "https://server.arcgisonline.com/ArcGIS/rest/services/"
                "World_Imagery/MapServer/tile/{z}/{y}/{x}"
            ),
            attr="Esri",
            name="Satellite",
            overlay=False,
            control=True,
        ).add_to(m)

        # Étiquettes OSM par-dessus (optionnel)
        folium.TileLayer(
            tiles=(
                "https://services.arcgisonline.com/ArcGIS/rest/services/"
                "Reference/World_Boundaries_and_Places/MapServer/tile/{z}/{y}/{x}"
            ),
            attr="Esri",
            name="Étiquettes",
            overlay=True,
            control=True,
            opacity=0.6,
        ).add_to(m)

        # Parcelles
        if feats_sel:
            folium.GeoJson(
                parcelles_sel,
                name="Parcelles",
                style_function=lambda _: {
                    "fillColor": "#FFD700",
                    "color": "#FF8C00",
                    "weight": 2,
                    "fillOpacity": 0.35,
                },
                tooltip=folium.GeoJsonTooltip(
                    fields=["Cadagri_26", "NOM_COM", "CULTURES 2", "STATUT"],
                    aliases=["Cadagri", "Commune", "Culture", "Statut"],
                    sticky=False,
                ),
            ).add_to(m)

        # Points RP (bleu)
        rp_group = folium.FeatureGroup(name="RP (receivers)", show=True)
        statut_color_map = {
            "OK": "green", "PVG": "orange", "PG": "orange",
            "NC": "red", "PV": "blue", "NR": "gray",
        }
        for _, row in rcv_sel.iterrows():
            try:
                lat, lon = float(row["lat"]), float(row["lon"])
            except (ValueError, KeyError):
                continue
            statut = str(row.get("STATUT", "")).strip().upper()
            color  = statut_color_map.get(statut, "blue")
            popup_html = (
                f"<b>RP {row.get('station','')}</b><br>"
                f"Ligne {row.get('line','')} · Point {row.get('point','')}<br>"
                f"Statut : <b>{row.get('STATUT DET', row.get('STATUT',''))}</b><br>"
                f"Culture : {row.get('CULTURES 2','')}<br>"
                f"Consignes : {row.get('CONSIGNES','')}"
            )
            folium.CircleMarker(
                location=[lat, lon],
                radius=6,
                color=color,
                fill=True,
                fill_color=color,
                fill_opacity=0.85,
                popup=folium.Popup(popup_html, max_width=280),
                tooltip=f"RP {row.get('station','')}",
            ).add_to(rp_group)
        rp_group.add_to(m)

        # Points SP (rouge/orange)
        sp_group = folium.FeatureGroup(name="SP (sources)", show=True)
        for _, row in src_sel.iterrows():
            try:
                lat, lon = float(row["lat"]), float(row["lon"])
            except (ValueError, KeyError):
                continue
            popup_html = (
                f"<b>SP {row.get('id','')}</b><br>"
                f"Type : {row.get('type','')}<br>"
                f"Densité : {row.get('Densite','')}"
            )
            folium.RegularPolygonMarker(
                location=[lat, lon],
                number_of_sides=3,
                radius=7,
                color="#cc4400",
                fill=True,
                fill_color="#ff6600",
                fill_opacity=0.85,
                popup=folium.Popup(popup_html, max_width=240),
                tooltip=f"SP {row.get('id','')}",
            ).add_to(sp_group)
        sp_group.add_to(m)

        folium.LayerControl(collapsed=False).add_to(m)

        # Auto-fit sur les données
        if lats:
            m.fit_bounds([[min(lats), min(lons)], [max(lats), max(lons)]])

        st_folium(m, use_container_width=True, height=620)
