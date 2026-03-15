"""
Optimiseur de Tournées – Assainissement / WC Chimiques
Dépôt fixe : Impasse Gaston Phoebus, Saint Sulpice la Pointe, 81370
"""

import streamlit as st
import pandas as pd
import folium
from streamlit_folium import st_folium
import requests
from geopy.geocoders import Nominatim
import time
from io import BytesIO
import datetime
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# reportlab – gestion UTF-8 native, aucun problème d'encodage
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                 Paragraph, Spacer, HRFlowable)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTES
# ─────────────────────────────────────────────────────────────────────────────

DEPOT_ADDRESS = "Imp. Gaston Phoebus, 81370 Saint-Sulpice-la-Pointe"
DEPOT_COORDS  = (43.7746, 1.9028)  # GPS fixe – jamais géocodé
OSRM_URL      = "http://router.project-osrm.org"

ACTION_COLORS = {
    "Nettoyer": "#AEC6E8",
    "Déposer":  "#B7E5B4",
    "Retirer":  "#F4B8C1",
}
ACTION_BORDER_COLORS = {
    "Nettoyer": "#1f6aa5",
    "Déposer":  "#28a745",
    "Retirer":  "#dc3545",
}
ACTION_MAP_COLORS = {"Nettoyer": "blue", "Déposer": "green", "Retirer": "red"}
ACTION_MAP_ICONS  = {"Nettoyer": "tint", "Déposer": "arrow-down", "Retirer": "arrow-up"}

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG PAGE
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(page_title="Optimiseur de Tournées", page_icon="🚛",
                   layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
    [data-testid="stMetric"] { background:#f8f9fa; border-radius:10px; padding:12px; }
    [data-testid="stMetricValue"] { font-size:1.6rem !important; color:#1f4e79; }
    .stop-card { border-radius:8px; padding:8px 12px; margin:4px 0; }
    .depot-card { background:#fff3cd; border-left:4px solid #ffc107; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────────────────────

def _init_df():
    return pd.DataFrame({
        "Action":        ["Nettoyer", "Nettoyer", "Nettoyer"],
        "Quantité":      ["1 WC",     "1 WC",     "1 WC"],
        "Nom du client": ["",         "",         ""],
        "Adresse":       ["",         "",         ""],
    })

if "df_stops"  not in st.session_state: st.session_state.df_stops  = _init_df()
if "result"    not in st.session_state: st.session_state.result    = None
if "tour_date" not in st.session_state: st.session_state.tour_date = datetime.date.today()
if "driver"    not in st.session_state: st.session_state.driver    = ""

# ─────────────────────────────────────────────────────────────────────────────
# GEOCODAGE – API Adresse gouv.fr + fallback Nominatim
# ─────────────────────────────────────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def geocode(address: str):
    """Retourne (lat, lon) ou None."""
    try:
        r = requests.get("https://api-adresse.data.gouv.fr/search/",
                         params={"q": address, "limit": 1}, timeout=10)
        features = r.json().get("features", [])
        if features:
            lon, lat = features[0]["geometry"]["coordinates"]
            return (lat, lon)
    except Exception:
        pass
    try:
        geo = Nominatim(user_agent="tournee_optimizer_v4")
        loc = geo.geocode(address + ", France", timeout=10)
        if loc:
            return (loc.latitude, loc.longitude)
        time.sleep(1.1)
    except Exception:
        pass
    return None

# ─────────────────────────────────────────────────────────────────────────────
# ROUTAGE OSRM
# ─────────────────────────────────────────────────────────────────────────────

def _tour_cost(tour, matrix):
    """Cout total d'un tour (duree) + retour au depot."""
    cost = sum(matrix[tour[i]][tour[i + 1]] for i in range(len(tour) - 1))
    cost += matrix[tour[-1]][tour[0]]
    return cost


def _two_opt(tour, matrix):
    """Amelioration 2-opt : echange de segments jusqu'a convergence."""
    best = tour[:]
    improved = True
    while improved:
        improved = False
        for i in range(1, len(best) - 1):
            for j in range(i + 1, len(best)):
                candidate = best[:i] + best[i:j + 1][::-1] + best[j + 1:]
                if _tour_cost(candidate, matrix) < _tour_cost(best, matrix):
                    best = candidate
                    improved = True
    return best


def osrm_trip(coords_latlon):
    """
    Optimisation TSP depot fixe :
    1. Matrice de durees OSRM (/table)
    2. Nearest-neighbor depuis le depot (index 0)
    3. Amelioration 2-opt
    4. Route finale (/route) pour geometrie et distances reelles
    """
    n         = len(coords_latlon)
    coord_str = ";".join(f"{lon},{lat}" for lat, lon in coords_latlon)

    # Etape 1 : matrice de durees
    try:
        r    = requests.get(f"{OSRM_URL}/table/v1/driving/{coord_str}",
                            params={"annotations": "duration"}, timeout=30)
        data = r.json()
    except Exception as e:
        st.error(f"Erreur reseau OSRM table : {e}")
        return None
    if data.get("code") != "Ok":
        st.error(f"OSRM table error : {data.get('code')}")
        return None

    matrix = data["durations"]   # liste n x n de durees en secondes

    # Etape 2 : nearest-neighbor depuis le depot
    unvisited = list(range(1, n))
    tour      = [0]
    current   = 0
    while unvisited:
        nearest = min(unvisited, key=lambda j: matrix[current][j] or float("inf"))
        tour.append(nearest)
        unvisited.remove(nearest)
        current = nearest

    # Etape 3 : amelioration 2-opt
    tour = _two_opt(tour, matrix)

    # Etape 4 : route reelle pour geometrie et distances
    tour_with_return = tour + [tour[0]]
    route_coords     = [coords_latlon[i] for i in tour_with_return]
    route_coord_str  = ";".join(f"{lon},{lat}" for lat, lon in route_coords)

    try:
        r2    = requests.get(f"{OSRM_URL}/route/v1/driving/{route_coord_str}",
                             params={"overview": "full", "geometries": "geojson"},
                             timeout=20)
        rdata = r2.json()
    except Exception as e:
        st.error(f"Erreur reseau OSRM route : {e}")
        return None
    if rdata.get("code") != "Ok":
        st.error(f"OSRM route error : {rdata.get('code')}")
        return None

    route = rdata["routes"][0]
    geom  = [[lat, lon] for lon, lat in route["geometry"]["coordinates"]]

    return {
        "order":        tour,
        "distance_km":  route["distance"] / 1000,
        "duration_min": route["duration"] / 60,
        "geometry":     geom,
    }



def osrm_route_distance(coords_latlon):
    coords    = coords_latlon + [coords_latlon[0]]
    coord_str = ";".join(f"{lon},{lat}" for lat, lon in coords)
    try:
        r    = requests.get(f"{OSRM_URL}/route/v1/driving/{coord_str}",
                            params=dict(overview="false"), timeout=20)
        data = r.json()
    except Exception:
        return None
    if data.get("code") != "Ok":
        return None
    route = data["routes"][0]
    return {"distance_km":  route["distance"] / 1000,
            "duration_min": route["duration"] / 60}

# ─────────────────────────────────────────────────────────────────────────────
# CARTE FOLIUM
# ─────────────────────────────────────────────────────────────────────────────

def build_map(depot_coords, stops_ordered, geometry):
    m = folium.Map(location=depot_coords, zoom_start=11, tiles="CartoDB positron")
    if geometry:
        folium.PolyLine(geometry, color="#1f4e79", weight=4, opacity=0.85).add_to(m)
    folium.Marker(depot_coords,
                  popup=folium.Popup(f"<b>Dépôt</b><br>{DEPOT_ADDRESS}", max_width=260),
                  tooltip="Dépôt – Départ & Retour",
                  icon=folium.Icon(color="orange", icon="home", prefix="fa")).add_to(m)
    for stop in stops_ordered:
        if stop["lat"] is None:
            continue
        color = ACTION_MAP_COLORS.get(stop["action"], "gray")
        icon  = ACTION_MAP_ICONS.get(stop["action"], "map-marker")
        popup_html = (f"<b>Arrêt {stop['order_num']}</b><br>"
                      f"<b>Action :</b> {stop['action']}<br>"
                      f"<b>Qté :</b> {stop['quantity']}<br>"
                      + (f"<b>Client :</b> {stop['client']}<br>" if stop['client'] else "")
                      + f"<b>Adresse :</b> {stop['address']}")
        folium.Marker([stop["lat"], stop["lon"]],
                      popup=folium.Popup(popup_html, max_width=260),
                      tooltip=f"{stop['order_num']}. {stop['action']} – {stop['address']}",
                      icon=folium.Icon(color=color, icon=icon, prefix="fa")).add_to(m)
        bg = {"blue": "#337ab7", "green": "#28a745", "red": "#dc3545"}.get(color, "#555")
        folium.Marker([stop["lat"], stop["lon"]],
                      icon=folium.DivIcon(
                          html=(f'<div style="font-size:11px;font-weight:bold;color:#fff;'
                                f'background:{bg};border-radius:50%;width:20px;height:20px;'
                                f'display:flex;align-items:center;justify-content:center;'
                                f'margin-top:-42px;margin-left:14px;">{stop["order_num"]}</div>'),
                          icon_size=(20, 20), icon_anchor=(0, 0))).add_to(m)
    return m

# ─────────────────────────────────────────────────────────────────────────────
# EXPORT EXCEL
# ─────────────────────────────────────────────────────────────────────────────

def export_excel(result, tour_date, driver_name, fuel_price_per_l):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Feuille de tournée"

    hdr_fill = PatternFill("solid", fgColor="1F4E79")
    hdr_font = Font(bold=True, color="FFFFFF", size=11)
    bold_f   = Font(bold=True)
    center   = Alignment(horizontal="center", vertical="center")
    left     = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin     = Side(style="thin", color="CCCCCC")
    brd      = Border(left=thin, right=thin, top=thin, bottom=thin)
    action_fills = {
        "Nettoyer": PatternFill("solid", fgColor="AEC6E8"),
        "Déposer":  PatternFill("solid", fgColor="B7E5B4"),
        "Retirer":  PatternFill("solid", fgColor="F4B8C1"),
    }

    ws.merge_cells("A1:G1")
    c = ws["A1"]
    c.value     = f"FEUILLE DE TOURNÉE – {tour_date.strftime('%d/%m/%Y')}"
    c.font      = Font(bold=True, size=15, color="1F4E79")
    c.alignment = center
    ws.row_dimensions[1].height = 28

    for label, val in [
        ("Chauffeur",           driver_name or "—"),
        ("Dépôt départ/retour", DEPOT_ADDRESS),
        ("Distance totale",     f"{result['distance_km']:.1f} km"),
        ("Durée estimée",       f"{int(result['duration_min']//60)}h{int(result['duration_min']%60):02d}"),
        ("Carburant estimé",    f"{result['fuel_liters']:.1f} L  ({result['fuel_cost']:.2f} €)"),
        ("Gain optimisation",   f"{result['km_saved']:.1f} km – {int(result['time_saved_min'])} min"),
    ]:
        r = ws.max_row + 1
        ws.cell(r, 1).value = label + " :"
        ws.cell(r, 1).font  = bold_f
        ws.cell(r, 2).value = val
        ws.merge_cells(f"B{r}:G{r}")
        ws.cell(r, 2).alignment = left
    ws.append([])

    headers = ["Ordre", "Action", "Quantité", "Nom du client", "Adresse", "Heure passage", "✓ Fait"]
    ws.append(headers)
    hr = ws.max_row
    for col, h in enumerate(headers, 1):
        c = ws.cell(hr, col)
        c.value = h; c.fill = hdr_fill; c.font = hdr_font
        c.alignment = center; c.border = brd
    ws.row_dimensions[hr].height = 20

    for stop in result["stops_ordered"]:
        ws.append([stop["order_num"], stop["action"], stop["quantity"],
                   stop["client"] or "", stop["address"], "", ""])
        r = ws.max_row
        for col in range(1, 8):
            c = ws.cell(r, col)
            c.border = brd
            c.alignment = center if col != 5 else left
        ws.cell(r, 2).fill = action_fills.get(stop["action"],
                                               PatternFill("solid", fgColor="F0F0F0"))
        ws.row_dimensions[r].height = 18

    ws.append(["↩", "Retour dépôt", "", "", DEPOT_ADDRESS, "", ""])
    r = ws.max_row
    for col in range(1, 8):
        c = ws.cell(r, col)
        c.fill = PatternFill("solid", fgColor="FFF3CD")
        c.font = bold_f; c.border = brd
        c.alignment = center if col != 5 else left

    for col, width in zip(range(1, 8), [8, 14, 14, 22, 44, 16, 8]):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = width

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ─────────────────────────────────────────────────────────────────────────────
# EXPORT PDF – reportlab (UTF-8 natif, zéro problème d'encodage)
# ─────────────────────────────────────────────────────────────────────────────

def export_pdf(result, tour_date, driver_name):
    buf    = BytesIO()
    doc    = SimpleDocTemplate(buf, pagesize=A4,
                               leftMargin=15*mm, rightMargin=15*mm,
                               topMargin=15*mm, bottomMargin=15*mm)
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle("title", parent=styles["Heading1"],
                                 fontSize=16, textColor=colors.HexColor("#1F4E79"),
                                 alignment=TA_CENTER, spaceAfter=4)
    small_style = ParagraphStyle("small", parent=styles["Normal"],
                                 fontSize=8, textColor=colors.grey,
                                 alignment=TA_CENTER)

    story = []

    # Titre
    story.append(Paragraph(
        f"FEUILLE DE TOURNÉE – {tour_date.strftime('%d/%m/%Y')}", title_style))
    story.append(HRFlowable(width="100%", thickness=1,
                             color=colors.HexColor("#1F4E79"), spaceAfter=6))

    # Récapitulatif
    recap_data = [
        ["Chauffeur :",             driver_name or "—"],
        ["Dépôt départ / retour :", DEPOT_ADDRESS],
        ["Distance totale :",       f"{result['distance_km']:.1f} km"],
        ["Durée de trajet :",
         f"{int(result['duration_min']//60)}h{int(result['duration_min']%60):02d}"],
        ["Carburant estimé :",
         f"{result['fuel_liters']:.1f} L  ({result['fuel_cost']:.2f} \u20ac)"],
        ["Économie optimisation :",
         f"\u2212{result['km_saved']:.1f} km  /  \u2212{int(result['time_saved_min'])} min"],
    ]
    recap_table = Table(recap_data, colWidths=[55*mm, 120*mm])
    recap_table.setStyle(TableStyle([
        ("FONTNAME",  (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTNAME",  (1, 0), (1, -1), "Helvetica"),
        ("FONTSIZE",  (0, 0), (-1, -1), 10),
        ("VALIGN",    (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 0), (-1, -1),
         [colors.white, colors.HexColor("#F5F5F5")]),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING",    (0, 0), (-1, -1), 4),
    ]))
    story.append(recap_table)
    story.append(Spacer(1, 8*mm))

    # Tableau des arrêts
    action_bg = {
        "Nettoyer": colors.HexColor("#AEC6E8"),
        "Déposer":  colors.HexColor("#B7E5B4"),
        "Retirer":  colors.HexColor("#F4B8C1"),
    }

    table_data = [["N°", "Action", "Quantité", "Client", "Adresse"]]
    row_styles = []

    # Ligne dépôt départ
    table_data.append(["", "Dépôt – Départ", "", "", DEPOT_ADDRESS])
    row_styles += [("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#FFF3CD")),
                   ("FONTNAME",   (0, 1), (-1, 1), "Helvetica-Bold")]

    for i, stop in enumerate(result["stops_ordered"]):
        table_data.append([str(stop["order_num"]), stop["action"],
                           stop["quantity"], stop["client"] or "", stop["address"]])
        ri = i + 2
        bg = action_bg.get(stop["action"], colors.HexColor("#F0F0F0"))
        row_styles.append(("BACKGROUND", (1, ri), (1, ri), bg))

    # Ligne dépôt retour
    table_data.append(["", "Dépôt – Retour", "", "", DEPOT_ADDRESS])
    last = len(table_data) - 1
    row_styles += [("BACKGROUND", (0, last), (-1, last), colors.HexColor("#FFF3CD")),
                   ("FONTNAME",   (0, last), (-1, last), "Helvetica-Bold")]

    stops_table = Table(table_data, colWidths=[10*mm, 28*mm, 24*mm, 32*mm, 83*mm],
                        repeatRows=1)
    base_style = [
        ("BACKGROUND",  (0, 0), (-1, 0), colors.HexColor("#1F4E79")),
        ("TEXTCOLOR",   (0, 0), (-1, 0), colors.white),
        ("FONTNAME",    (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE",    (0, 0), (-1, 0), 10),
        ("ALIGN",       (0, 0), (-1, 0), "CENTER"),
        ("FONTNAME",    (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE",    (0, 1), (-1, -1), 9),
        ("ALIGN",       (0, 1), (-1, -1), "CENTER"),
        ("ALIGN",       (4, 1), (4, -1), "LEFT"),
        ("VALIGN",      (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -2),
         [colors.white, colors.HexColor("#F9F9F9")]),
        ("GRID",        (0, 0), (-1, -1), 0.5, colors.HexColor("#CCCCCC")),
        ("TOPPADDING",  (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]
    stops_table.setStyle(TableStyle(base_style + row_styles))
    story.append(stops_table)
    story.append(Spacer(1, 8*mm))

    # Footer
    story.append(HRFlowable(width="100%", thickness=0.5,
                             color=colors.grey, spaceBefore=4))
    story.append(Paragraph(
        f"Généré le {datetime.datetime.now().strftime('%d/%m/%Y à %H:%M')}"
        " — Optimiseur de Tournées", small_style))

    doc.build(story)
    buf.seek(0)
    return buf

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("## 🚛 Optimiseur de Tournées")
    st.caption("Assainissement · WC Chimiques")
    st.divider()

    st.subheader("📅 Tournée")
    st.session_state.tour_date = st.date_input(
        "Date", value=st.session_state.tour_date, format="DD/MM/YYYY")
    st.session_state.driver = st.text_input(
        "👤 Chauffeur", value=st.session_state.driver,
        placeholder="ex : Jean Dupont")

    st.divider()
    st.subheader("🚛 Véhicule")
    fuel_conso = st.number_input(
        "Consommation (L/100 km)", min_value=5.0, max_value=30.0,
        value=15.0, step=0.5,
        help="Petit camion avec remorque : environ 14–16 L/100 km")
    fuel_price = st.number_input(
        "Prix carburant (€/L)", min_value=1.0, max_value=3.5,
        value=1.85, step=0.01)

    st.divider()
    st.caption(f"**Dépôt :**\n{DEPOT_ADDRESS}")

    if st.session_state.result:
        st.divider()
        r = st.session_state.result
        st.success(
            f"✅ Dernière optimisation\n\n"
            f"**{r['distance_km']:.1f} km** · "
            f"**{int(r['duration_min']//60)}h{int(r['duration_min']%60):02d}**\n\n"
            f"⛽ {r['fuel_liters']:.1f} L ({r['fuel_cost']:.2f} €)\n\n"
            f"⏱ Gain : {int(r['time_saved_min'])} min / {r['km_saved']:.1f} km")

# ─────────────────────────────────────────────────────────────────────────────
# ONGLETS
# ─────────────────────────────────────────────────────────────────────────────

st.title("🚛 Optimiseur de Tournées — Assainissement")
st.caption(f"🏭 Dépôt fixe : **{DEPOT_ADDRESS}**")

tab_saisie, tab_optim, tab_export = st.tabs([
    "📋  Saisie des arrêts",
    "🗺️  Tournée optimisée",
    "📥  Export",
])

# ══════════════════════════════════════════════════════════════════════════════
# ONGLET 1 — SAISIE
# ══════════════════════════════════════════════════════════════════════════════

with tab_saisie:
    st.subheader("Saisie des arrêts de la tournée")
    st.info(
        f"💡 Saisissez vos arrêts ci-dessous. "
        f"Départ et retour au dépôt (**{DEPOT_ADDRESS}**) sont automatiques. "
        f"L'ordre de saisie n'a pas d'importance.")

    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("➕ Ajouter un arrêt", use_container_width=True):
            new = pd.DataFrame({"Action": ["Nettoyer"], "Quantité": ["1 WC"],
                                "Nom du client": [""], "Adresse": [""]})
            st.session_state.df_stops = pd.concat(
                [st.session_state.df_stops, new], ignore_index=True)
            st.rerun()
    with c2:
        if st.button("➖ Supprimer le dernier", use_container_width=True):
            if len(st.session_state.df_stops) > 1:
                st.session_state.df_stops = (
                    st.session_state.df_stops.iloc[:-1].reset_index(drop=True))
                st.rerun()
    with c3:
        if st.button("🗑️ Tout vider", use_container_width=True, type="secondary"):
            st.session_state.df_stops = _init_df()
            st.session_state.result   = None
            st.rerun()

    st.markdown("---")
    leg_cols = st.columns(3)
    for col, (action, color) in zip(leg_cols, ACTION_COLORS.items()):
        col.markdown(
            f'<span style="background:{color};padding:3px 10px;'
            f'border-radius:4px;font-size:0.85em">■ {action}</span>',
            unsafe_allow_html=True)
    st.markdown("")

    edited = st.data_editor(
        st.session_state.df_stops,
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "Action": st.column_config.SelectboxColumn(
                "Action", required=True, width="small",
                options=["Nettoyer", "Déposer", "Retirer"]),
            "Quantité": st.column_config.SelectboxColumn(
                "Quantité", required=True, width="small",
                options=["1 WC", "2 WC", "1 Lave Main", "2 Lave Main", "1 WC + 1 LM"]),
            "Nom du client": st.column_config.TextColumn(
                "Nom du client", width="medium"),
            "Adresse": st.column_config.TextColumn(
                "Adresse complète (rue, ville, CP)", width="large",
                help="Ex : Place d'Hautpoul 81600 Gaillac"),
        },
        hide_index=False,
        key="editor_stops",
    )
    st.session_state.df_stops = edited

    valid_rows = edited[edited["Adresse"].str.strip() != ""]
    n_valid    = len(valid_rows)
    st.caption(f"📍 **{n_valid}** arrêt(s) avec adresse renseignée")
    st.markdown("---")

    if st.button("🚀 Optimiser la tournée", type="primary",
                 use_container_width=True, disabled=(n_valid < 1)):
        valid_stops = valid_rows.reset_index(drop=True)

        with st.spinner("🔍 Géocodage des adresses en cours…"):
            depot_coords = DEPOT_COORDS
            pb           = st.progress(0, text="Géocodage…")
            geo          = {}
            all_addrs    = valid_stops["Adresse"].tolist()
            for i, addr in enumerate(all_addrs):
                geo[addr] = geocode(addr)
                pb.progress((i + 1) / len(all_addrs),
                             text=f"Géocodage : {addr[:50]}…")
            pb.empty()

        failed = [a for a, c in geo.items() if c is None]
        if failed:
            st.warning("⚠️ Ces adresses n'ont pas pu être géocodées :\n"
                       + "\n".join(f"- {a}" for a in failed))

        coords_list = [depot_coords]
        for i, row in valid_stops.iterrows():
            c = geo.get(row["Adresse"])
            if c:
                coords_list.append(c)

        if len(coords_list) < 2:
            st.error("❌ Aucune adresse valide après géocodage.")
            st.stop()

        with st.spinner("🗺️ Calcul de l'itinéraire optimisé…"):
            trip = osrm_trip(coords_list)
            if not trip:
                st.stop()
            orig = osrm_route_distance(coords_list)

        order         = trip["order"]
        stops_ordered = []
        rank          = 1
        for orig_idx in order:
            if orig_idx == 0:
                continue
            sri = orig_idx - 1
            if sri >= len(valid_stops):
                continue
            row    = valid_stops.iloc[sri]
            coords = geo.get(row["Adresse"])
            stops_ordered.append({
                "order_num": rank,
                "action":    row["Action"],
                "quantity":  row["Quantité"],
                "client":    row["Nom du client"],
                "address":   row["Adresse"],
                "lat":       coords[0] if coords else None,
                "lon":       coords[1] if coords else None,
            })
            rank += 1

        dist_km   = trip["distance_km"]
        dur_min   = trip["duration_min"]
        fuel_l    = dist_km * fuel_conso / 100
        fuel_cost = fuel_l * fuel_price
        km_saved  = max(0, (orig["distance_km"]  - dist_km))  if orig else 0
        min_saved = max(0, (orig["duration_min"] - dur_min)) if orig else 0

        st.session_state.result = {
            "stops_ordered": stops_ordered,
            "distance_km":   dist_km,
            "duration_min":  dur_min,
            "fuel_liters":   fuel_l,
            "fuel_cost":     fuel_cost,
            "km_saved":      km_saved,
            "time_saved_min":min_saved,
            "geometry":      trip["geometry"],
            "depot_coords":  depot_coords,
        }
        st.success("✅ Tournée optimisée ! Consultez l'onglet **Tournée optimisée**.")
        st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# ONGLET 2 — RÉSULTATS
# ══════════════════════════════════════════════════════════════════════════════

with tab_optim:
    if st.session_state.result is None:
        st.info("👈 Saisissez vos arrêts puis cliquez sur **Optimiser la tournée**.")
    else:
        r     = st.session_state.result
        hours = int(r["duration_min"] // 60)
        mins  = int(r["duration_min"] % 60)

        st.subheader("📊 Récapitulatif de la tournée optimisée")
        c1, c2, c3, c4, c5, c6 = st.columns(6)
        c1.metric("📍 Arrêts",        len(r["stops_ordered"]))
        c2.metric("🛣️ Distance",       f"{r['distance_km']:.1f} km")
        c3.metric("⏱️ Durée trajet",   f"{hours}h{mins:02d}")
        c4.metric("⛽ Carburant",      f"{r['fuel_liters']:.1f} L")
        c5.metric("💶 Coût carburant", f"{r['fuel_cost']:.2f} €")
        c6.metric("⏳ Temps gagné",    f"{int(r['time_saved_min'])} min",
                  delta=f"-{r['km_saved']:.1f} km", delta_color="inverse")

        st.markdown("---")
        col_map, col_list = st.columns([3, 2])

        with col_map:
            st.subheader("🗺️ Carte de la tournée")
            m = build_map(r["depot_coords"], r["stops_ordered"], r["geometry"])
            st_folium(m, use_container_width=True, height=520, returned_objects=[])

        with col_list:
            st.subheader("📋 Ordre des arrêts")
            st.markdown(
                f'<div class="stop-card depot-card"><b>🏭 Dépôt — Départ</b><br>'
                f'<small>{DEPOT_ADDRESS}</small></div>', unsafe_allow_html=True)
            for stop in r["stops_ordered"]:
                bg  = ACTION_COLORS.get(stop["action"], "#f0f0f0")
                brd = ACTION_BORDER_COLORS.get(stop["action"], "#999")
                cli = f" · {stop['client']}" if stop['client'] else ""
                st.markdown(
                    f'<div class="stop-card" style="background:{bg};'
                    f'border-left:4px solid {brd};">'
                    f'<b>#{stop["order_num"]} {stop["action"]}</b>{cli}<br>'
                    f'<small>📍 {stop["address"]}</small><br>'
                    f'<small>📦 {stop["quantity"]}</small></div>',
                    unsafe_allow_html=True)
            st.markdown(
                f'<div class="stop-card depot-card"><b>🏭 Dépôt — Retour</b><br>'
                f'<small>{DEPOT_ADDRESS}</small></div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# ONGLET 3 — EXPORT
# ══════════════════════════════════════════════════════════════════════════════

with tab_export:
    if st.session_state.result is None:
        st.info("👈 Optimisez d'abord une tournée pour pouvoir l'exporter.")
    else:
        r = st.session_state.result
        st.subheader("📥 Exporter la feuille de tournée")
        col_xl, col_pdf = st.columns(2)

        with col_xl:
            st.markdown("### 📊 Excel (.xlsx)")
            st.write("Feuille de route mise en forme avec codes couleur, "
                     "colonne heure de passage et case à cocher ✓ Fait.")
            xl_buf = export_excel(r, st.session_state.tour_date,
                                  st.session_state.driver, fuel_price)
            st.download_button(
                "⬇️ Télécharger Excel", data=xl_buf,
                file_name=f"tournee_{st.session_state.tour_date.strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, type="primary")

        with col_pdf:
            st.markdown("### 📄 PDF")
            st.write("Document imprimable prêt pour le chauffeur, "
                     "avec tableau coloré et récapitulatif des indicateurs.")
            pdf_buf = export_pdf(r, st.session_state.tour_date,
                                 st.session_state.driver)
            st.download_button(
                "⬇️ Télécharger PDF", data=pdf_buf,
                file_name=f"tournee_{st.session_state.tour_date.strftime('%Y%m%d')}.pdf",
                mime="application/pdf",
                use_container_width=True, type="primary")

        st.markdown("---")
        st.caption("💡 Le fichier Excel contient une colonne **Heure de passage** "
                   "à renseigner manuellement et une colonne **✓ Fait** pour validation terrain.")