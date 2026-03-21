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
                                 Paragraph, Spacer, HRFlowable, Image as RLImage)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTES
# ─────────────────────────────────────────────────────────────────────────────

DEPOTS = {
    "Saint-Sulpice":   "Imp. Gaston Phoebus, 81370 Saint-Sulpice-la-Pointe",
    "Villemur-sur-Tarn": "1 Avenue du Président Roosevelt 31340 Villemur-sur-Tarn",
}
DEPOT_DEPART_DEFAULT = "Saint-Sulpice"
DEPOT_RETOUR_DEFAULT = "Saint-Sulpice"
OSRM_URL      = "http://router.project-osrm.org"

ACTION_COLORS = {
    "Nettoyer":      "#AEC6E8",
    "Déposer":       "#B7E5B4",
    "Retirer":       "#F4B8C1",
    "Chargement":    "#D4B8E0",
    "Déchargement":  "#FFE0B2",
}
ACTION_BORDER_COLORS = {
    "Nettoyer":      "#1f6aa5",
    "Déposer":       "#28a745",
    "Retirer":       "#dc3545",
    "Chargement":    "#7B1FA2",
    "Déchargement":  "#E65100",
}
ACTION_MAP_COLORS = {
    "Nettoyer":     "blue",
    "Déposer":      "green",
    "Retirer":      "red",
    "Chargement":   "purple",
    "Déchargement": "orange",
}
ACTION_MAP_ICONS = {
    "Nettoyer":     "tint",
    "Déposer":      "arrow-down",
    "Retirer":      "arrow-up",
    "Chargement":   "upload",
    "Déchargement": "download",
}


# Consignes par type d'action (affichées dans le PDF)
ACTION_CONSIGNES = {
    "Nettoyer": (
        "Vidanger complètement la cuve. Nettoyer l'intérieur avec les produits homologués. "
        "Vérifier et réapprovisionner les consommables (papier, gel désinfectant, savon). "
        "Inspecter l'état général (porte, serrure, sol). Signaler tout dysfonctionnement "
        "ou dégradation sur la fiche de tournée."
    ),
    "Déposer": (
        "Positionner l'équipement sur la zone désignée par le client, hors obstacle et "
        "zone de passage. Vérifier la stabilité et l'aplomb. Vérifier la propreté avant "
        "remise au client. Informer le client de la mise en service et lui remettre les "
        "consignes d'utilisation si première installation."
    ),
    "Retirer": (
        "Vidanger la cuve avant enlèvement, même si partiellement remplie. "
        "Vérifier que la zone est propre et sans trace après retrait. "
        "Noter l'état de l'équipement au chargement (dégradation, pièce manquante). "
        "Obtenir la signature du bon de retrait si présence du client."
    ),
    "Chargement": (
        "Vérifier l'état de l'équipement avant chargement (noter les dommages existants). "
        "Arrimer correctement le chargement selon le plan de chargement. "
        "S'assurer que le poids total ne dépasse pas la charge utile du véhicule avec remorque. "
        "Vérifier les feux de la remorque avant départ."
    ),
    "Déchargement": (
        "Décharger avec précaution en utilisant les équipements de manutention adaptés. "
        "Vérifier l'état de l'équipement après déchargement. "
        "Positionner sur l'aire de stockage désignée. "
        "Renseigner le bon de livraison et faire signer le destinataire."
    ),
}

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG PAGE
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(page_title="Optimiseur de Tournées", page_icon="",
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
        "Action":        ["Nettoyer",    "Nettoyer",    "Nettoyer"],
        "Produit":       ["WC chimique", "WC chimique", "WC chimique"],
        "Option":        ["Lave-main",   "Lave-main",   "Lave-main"],
        "Quantité":      [1,             1,             1],
        "Nom du client": ["",            "",            ""],
        "Adresse":       ["",            "",            ""],
        "Durée (min)":   [30,            30,            30],
        "Pas avant":     ["",            "",            ""],
        "Pas après":     ["",            "",            ""],
        "Observations":  ["",            "",            ""],
    })

if "df_stops"          not in st.session_state: st.session_state.df_stops          = _init_df()
if "heure_min_depart"  not in st.session_state: st.session_state.heure_min_depart  = datetime.time(7, 0)
if "result"            not in st.session_state: st.session_state.result            = None
if "tour_date"         not in st.session_state: st.session_state.tour_date         = datetime.date.today()
if "driver"            not in st.session_state: st.session_state.driver            = ""
if "depot_depart_key" not in st.session_state: st.session_state.depot_depart_key = DEPOT_DEPART_DEFAULT
if "depot_retour_key" not in st.session_state: st.session_state.depot_retour_key = DEPOT_RETOUR_DEFAULT

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

def _parse_hhmm(s):
    """Convertit 'HH:MM' en minutes depuis minuit, ou None si vide/invalide."""
    if not s or not str(s).strip():
        return None
    s = str(s).strip()
    try:
        h, m = s.split(":")
        return int(h) * 60 + int(m)
    except Exception:
        return None


def _fmt_min(minutes):
    """Convertit des minutes depuis minuit en chaîne HH:MM."""
    if minutes is None:
        return ""
    h = int(minutes) // 60
    m = int(minutes) % 60
    return f"{h:02d}:{m:02d}"



def _compute_optimal_departure(tour, matrix, time_windows):
    """
    Calcule l'heure de départ optimale du dépôt.

    Logique :
    - On calcule pour chaque arrêt le temps cumulé depuis le dépôt
      (trajet + durées d'intervention des arrêts précédents).
    - Pour les contraintes "Pas après" (latest) : le départ ne peut pas
      dépasser latest_i - cumul_i  → on prend le minimum de ces valeurs.
    - Pour les contraintes "Pas avant" (earliest) : on peut partir plus tôt
      pour éviter l'attente → on prend le minimum entre le départ calculé
      depuis "Pas après" et earliest_i - cumul_i.
    - Si aucune contrainte : retourne None (pas de contrainte de départ).
    """
    n = len(tour)
    # Calcul des temps cumulés (trajet pur + interventions) depuis le dépôt
    cumul = 0.0
    cumuls = []          # cumuls[k] = temps cumulé pour atteindre tour[k] (k>=1)
    for k in range(1, n):
        prev = tour[k - 1]
        curr = tour[k]
        cumul += (matrix[prev][curr] or 0) / 60
        cumuls.append(cumul)
        # Ajouter la durée d'intervention de l'arrêt courant (sauf le dernier retour)
        dur = time_windows[curr].get("duration", 0) or 0
        cumul += dur

    candidates = []
    for k, idx in enumerate(tour[1:]):
        tw = time_windows[idx]
        cumul_k = cumuls[k]
        if tw.get("latest") is not None:
            # Dernier départ pour arriver à temps : latest - cumul
            candidates.append(tw["latest"] - cumul_k)
        if tw.get("earliest") is not None:
            # Départ idéal pour arriver exactement à l'ouverture (sans attente)
            candidates.append(tw["earliest"] - cumul_k)

    if not candidates:
        return None  # Aucune contrainte

    # Départ le plus tardif qui respecte toutes les contraintes
    optimal = min(candidates)
    # Arrondir à la minute inférieure
    return max(0, int(optimal))


def _qty_label(row):
    """
    Construit la description lisible de la commande à partir des colonnes
    Produit / Option / Quantité.
    Ex : "3 × WC chimique + 3 × Urinoir"
         "2 × Lave-main"
    """
    produit = str(row.get("Produit", "") or "").strip()
    option  = str(row.get("Option",  "") or "").strip()
    try:
        qty = int(row.get("Quantité", 1) or 1)
    except (ValueError, TypeError):
        qty = 1

    if not produit:
        return f"{qty} × ?"

    label = f"{qty} × {produit}"
    # Option uniquement pour WC chimique et si renseignée
    if produit == "WC chimique" and option:
        label += f" + {qty} × {option}"
    return label

def _compute_arrivals(tour, matrix, depart_min, time_windows):
    """
    Calcule l'heure d'arrivée réelle à chaque arrêt (en minutes depuis minuit),
    en tenant compte des fenêtres temporelles et des durées d'intervention.
    time_windows[i] contient aussi "duration" (durée intervention en minutes).
    Retourne une liste de dicts par arrêt (hors dépôt) :
      arrival_min, wait_min, departure_min, tw_early, tw_late, violated, duration_min
    """
    results      = []
    current_time = depart_min   # minutes depuis minuit
    current_idx  = tour[0]      # = 0 (dépôt)

    for step, next_idx in enumerate(tour[1:], 1):
        travel_min    = (matrix[current_idx][next_idx] or 0) / 60
        arrival_min   = current_time + travel_min
        tw            = time_windows[next_idx]
        earliest      = tw["earliest"]
        latest        = tw["latest"]
        duration_min  = tw.get("duration", 0) or 0
        wait_min      = 0

        # Si on arrive trop tôt → on attend jusqu'à l'ouverture
        if earliest is not None and arrival_min < earliest:
            wait_min    = earliest - arrival_min
            arrival_min = earliest

        # Violation : arrive après la limite latest
        violated = (latest is not None and arrival_min > latest)

        # Départ de l'arrêt = arrivée + durée intervention
        departure_min = arrival_min + duration_min

        results.append({
            "arrival_min":   arrival_min,
            "wait_min":      wait_min,
            "departure_min": departure_min,
            "tw_early":      earliest,
            "tw_late":       latest,
            "violated":      violated,
            "duration_min":  duration_min,
        })
        current_time = departure_min   # on repart après l'intervention
        current_idx  = next_idx

    return results


def _tour_cost_tw(tour, matrix, depart_min, time_windows):
    """
    Coût total tenant compte des fenêtres temporelles.
    Temps de trajet + attentes + pénalité forte pour violations.
    """
    PENALTY = 1e6   # pénalité par minute de dépassement
    arrivals = _compute_arrivals(tour, matrix, depart_min, time_windows)
    # Durée totale jusqu'au dernier arrêt + retour dépôt
    if not arrivals:
        return 0
    last_departure = arrivals[-1]["departure_min"]
    back_min = (matrix[tour[-1]][tour[0]] or 0) / 60
    total_time = (last_departure + back_min) - depart_min

    # Pénalités pour violations
    penalty = 0
    for a in arrivals:
        if a["violated"] and a["tw_late"] is not None:
            penalty += (a["arrival_min"] - a["tw_late"]) * PENALTY

    return total_time + penalty


def _two_opt(tour, matrix, depart_min=None, time_windows=None):
    """Amelioration 2-opt avec prise en compte optionnelle des fenêtres temporelles."""
    use_tw = depart_min is not None and time_windows is not None
    def cost(t):
        if use_tw:
            return _tour_cost_tw(t, matrix, depart_min, time_windows)
        total = sum(matrix[t[i]][t[i+1]] for i in range(len(t)-1))
        return total + matrix[t[-1]][t[0]]

    best     = tour[:]
    improved = True
    while improved:
        improved = False
        for i in range(1, len(best) - 1):
            for j in range(i + 1, len(best)):
                candidate = best[:i] + best[i:j+1][::-1] + best[j+1:]
                if cost(candidate) < cost(best):
                    best     = candidate
                    improved = True
    return best


def osrm_trip(coords_latlon, time_windows=None, depart_min=None):
    """
    Optimisation TSP depot fixe avec fenêtres temporelles optionnelles.
    1. Matrice de durees OSRM (/table)
    2. Nearest-neighbor depuis le depot (index 0)
    3. Amelioration 2-opt (avec TW si fournis)
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

    # Etape 3 : amelioration 2-opt (avec TW si fournis)
    tour = _two_opt(tour, matrix, depart_min=depart_min, time_windows=time_windows)

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
        "matrix":       matrix,
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

def build_map(depot_coords, stops_ordered, geometry, depot_depart_addr="Dépôt départ", depot_retour_addr=None, depot_retour_coords=None):
    m = folium.Map(location=depot_coords, zoom_start=11, tiles="CartoDB positron")
    if geometry:
        folium.PolyLine(geometry, color="#1f4e79", weight=4, opacity=0.85).add_to(m)
    folium.Marker(depot_coords,
                  popup=folium.Popup(f"<b>Dépôt départ</b><br>{r["depot_depart_addr"] if "r" in dir() else ""}", max_width=260),
                  tooltip="Dépôt – Départ",
                  icon=folium.Icon(color="orange", icon="home", prefix="fa")).add_to(m)
    # Marqueur retour dépôt si différent du départ
    if depot_retour_coords and depot_retour_coords != depot_coords:
        folium.Marker(depot_retour_coords,
                      popup=folium.Popup(f"<b>Dépôt retour</b><br>{depot_retour_addr}", max_width=260),
                      tooltip="Dépôt – Retour",
                      icon=folium.Icon(color="darkred", icon="flag", prefix="fa")).add_to(m)

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


def generate_map_image(depot_coords, stops_ordered, geometry, depot_retour_coords=None):
    """
    Génère une image PNG de la carte via Pillow (PIL).
    Disponible partout, aucune dépendance système supplémentaire.
    """
    from PIL import Image, ImageDraw

    W, H   = 900, 600
    MARGIN = 60
    BG     = (240, 239, 231)   # fond beige carte
    GRID   = (200, 200, 200)
    ROUTE  = (31, 78, 121)

    ACTION_RGB = {
        "Nettoyer":     (31,  106, 165),
        "Déposer":      (40,  167,  69),
        "Retirer":      (220,  53,  69),
        "Chargement":   (123,  31, 162),
        "Déchargement": (230,  81,   0),
    }

    img  = Image.new("RGB", (W, H), BG)
    draw = ImageDraw.Draw(img)

    # Points pour normalisation
    all_lats = [depot_coords[0]] + [s["lat"] for s in stops_ordered if s["lat"]]
    all_lons = [depot_coords[1]] + [s["lon"] for s in stops_ordered if s["lon"]]
    min_lat, max_lat = min(all_lats), max(all_lats)
    min_lon, max_lon = min(all_lons), max(all_lons)
    span_lat = max_lat - min_lat or 0.01
    span_lon = max_lon - min_lon or 0.01
    pad_lat, pad_lon = span_lat * 0.14, span_lon * 0.14
    min_lat -= pad_lat; max_lat += pad_lat
    min_lon -= pad_lon; max_lon += pad_lon
    span_lat = max_lat - min_lat
    span_lon = max_lon - min_lon

    def proj(lat, lon):
        x = int(MARGIN + (lon - min_lon) / span_lon * (W - 2 * MARGIN))
        y = int(H - MARGIN - (lat - min_lat) / span_lat * (H - 2 * MARGIN))
        return (x, y)

    # Grille
    for i in range(6):
        xi = MARGIN + int(i * (W - 2*MARGIN) / 5)
        yi = MARGIN + int(i * (H - 2*MARGIN) / 5)
        draw.line([(xi, MARGIN), (xi, H-MARGIN)], fill=GRID, width=1)
        draw.line([(MARGIN, yi), (W-MARGIN, yi)], fill=GRID, width=1)

    # Bordure
    draw.rectangle([MARGIN-1, MARGIN-1, W-MARGIN+1, H-MARGIN+1],
                   outline=(150, 150, 150), width=1)

    # Tracé route
    if geometry and len(geometry) > 1:
        pts = [proj(lat, lon) for lat, lon in geometry]
        for i in range(len(pts) - 1):
            # Ombre épaisse
            draw.line([pts[i], pts[i+1]], fill=(180, 200, 220), width=6)
        for i in range(len(pts) - 1):
            draw.line([pts[i], pts[i+1]], fill=ROUTE, width=3)

    # Dépôt
    dx, dy = proj(depot_coords[0], depot_coords[1])
    r = 12
    draw.ellipse([dx-r, dy-r, dx+r, dy+r], fill=(255, 193, 7), outline=(51, 51, 51), width=2)
    draw.text((dx, dy), "D", fill=(51, 51, 51), anchor="mm")
    draw.text((dx + r + 4, dy - 5), "Départ", fill=(51, 51, 51))

    # Dépôt retour (si différent du départ)
    if depot_retour_coords and depot_retour_coords != depot_coords:
        rx2, ry2 = proj(depot_retour_coords[0], depot_retour_coords[1])
        draw.ellipse([rx2-r, ry2-r, rx2+r, ry2+r], fill=(220, 53, 69), outline=(51, 51, 51), width=2)
        draw.text((rx2, ry2), "R", fill=(255, 255, 255), anchor="mm")
        draw.text((rx2 + r + 4, ry2 - 5), "Retour", fill=(51, 51, 51))

    # Arrêts
    for stop in stops_ordered:
        if stop["lat"] is None:
            continue
        sx, sy  = proj(stop["lat"], stop["lon"])
        rgb     = ACTION_RGB.get(stop["action"], (85, 85, 85))
        r_stop  = 11
        # Ombre
        draw.ellipse([sx-r_stop+1, sy-r_stop+1, sx+r_stop+1, sy+r_stop+1],
                     fill=(180, 180, 180))
        draw.ellipse([sx-r_stop, sy-r_stop, sx+r_stop, sy+r_stop],
                     fill=rgb, outline=(255, 255, 255), width=2)
        draw.text((sx, sy), str(stop["order_num"]), fill=(255, 255, 255), anchor="mm")
        # Étiquette
        label = stop["address"][:32] + ("…" if len(stop["address"]) > 32 else "")
        lx, ly = sx + r_stop + 4, sy - 7
        # Fond blanc semi-transparent
        bbox = draw.textbbox((lx, ly), label)
        draw.rectangle([bbox[0]-2, bbox[1]-1, bbox[2]+2, bbox[3]+1],
                       fill=(255, 255, 255, 200))
        draw.text((lx, ly), label, fill=(30, 30, 30))

    # Légende en bas
    lx, ly = MARGIN, H - MARGIN + 14
    draw.text((lx, ly), "●  Dépôt", fill=(51, 51, 51))
    lx += 80
    seen = []
    for stop in stops_ordered:
        a = stop["action"]
        if a not in seen:
            seen.append(a)
            rgb = ACTION_RGB.get(a, (85, 85, 85))
            draw.ellipse([lx, ly+2, lx+10, ly+12], fill=rgb)
            draw.text((lx+14, ly), a, fill=(51, 51, 51))
            lx += len(a) * 7 + 24

    # Titre
    title = "Tournée optimisée"
    tw = draw.textlength(title)
    draw.text(((W - tw) // 2, 14), title, fill=(31, 78, 121))

    buf = BytesIO()
    img.save(buf, format="PNG", optimize=True)
    buf.seek(0)
    return buf.read()


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
        "Nettoyer":     PatternFill("solid", fgColor="AEC6E8"),
        "Déposer":      PatternFill("solid", fgColor="B7E5B4"),
        "Retirer":      PatternFill("solid", fgColor="F4B8C1"),
        "Chargement":   PatternFill("solid", fgColor="D4B8E0"),
        "Déchargement": PatternFill("solid", fgColor="FFE0B2"),
    }

    ws.merge_cells("A1:G1")
    c = ws["A1"]
    c.value     = f"FEUILLE DE TOURNÉE – {tour_date.strftime('%d/%m/%Y')}"
    c.font      = Font(bold=True, size=15, color="1F4E79")
    c.alignment = center
    ws.row_dimensions[1].height = 28

    for label, val in [
        ("Chauffeur",           driver_name or "—"),
        ("Dépôt de départ",  result.get("depot_depart_addr", "")),
        ("Dépôt de retour",  result.get("depot_retour_addr", "")),
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

    headers = ["Ordre", "Action", "Produit", "Option", "Qté", "Nom du client", "Adresse", "Durée (min)", "Pas avant", "Pas après", "Arrivée", "Départ", "✓ Fait"]
    ws.append(headers)
    hr = ws.max_row
    for col, h in enumerate(headers, 1):
        c = ws.cell(hr, col)
        c.value = h; c.fill = hdr_fill; c.font = hdr_font
        c.alignment = center; c.border = brd
    ws.row_dimensions[hr].height = 20

    for stop in result["stops_ordered"]:
        ws.append([stop["order_num"], stop["action"],
                   stop.get("produit", ""), stop.get("option", ""),
                   stop.get("qty_num", ""),
                   stop["client"] or "", stop["address"],
                   stop.get("duration_min", ""),
                   _fmt_min(stop.get("tw_early")),
                   _fmt_min(stop.get("tw_late")),
                   _fmt_min(stop.get("arrival_min")),
                   _fmt_min(stop.get("departure_min")),
                   ""])
        r = ws.max_row
        for col in range(1, 14):
            c = ws.cell(r, col)
            c.border = brd
            c.alignment = center if col != 7 else left
            if col == 11 and stop.get("violated"):
                c.font = Font(bold=True, color="DC3545")
        ws.cell(r, 2).fill = action_fills.get(stop["action"],
                                               PatternFill("solid", fgColor="F0F0F0"))
        ws.row_dimensions[r].height = 18

    ws.append(["↩", "Retour dépôt", "", "", "", "", result.get("depot_retour_addr",""), "", "", "", "", "", ""])
    r = ws.max_row
    for col in range(1, 14):
        c = ws.cell(r, col)
        c.fill = PatternFill("solid", fgColor="FFF3CD")
        c.font = bold_f; c.border = brd
        c.alignment = center if col != 7 else left

    for col, width in zip(range(1, 14), [8, 14, 16, 14, 6, 22, 44, 12, 12, 12, 12, 12, 8]):
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
                               leftMargin=8*mm, rightMargin=8*mm,
                               topMargin=10*mm, bottomMargin=10*mm)
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle("title", parent=styles["Heading1"],
                                 fontSize=16, textColor=colors.HexColor("#1F4E79"),
                                 alignment=TA_CENTER, spaceAfter=4)
    small_style = ParagraphStyle("small", parent=styles["Normal"],
                                 fontSize=8, textColor=colors.grey,
                                 alignment=TA_CENTER)
    cell_style  = ParagraphStyle("cell", parent=styles["Normal"],
                                 fontSize=8, leading=10, wordWrap="CJK")
    cell_bold   = ParagraphStyle("cell_bold", parent=styles["Normal"],
                                 fontSize=8, leading=10, fontName="Helvetica-Bold")
    cell_white  = ParagraphStyle("cell_white", parent=styles["Normal"],
                                 fontSize=8, leading=10, fontName="Helvetica-Bold",
                                 textColor=colors.white)
    check_style = ParagraphStyle("check", parent=styles["Normal"],
                                 fontSize=9, leading=13, leftIndent=4)
    obs_style          = ParagraphStyle("obs_title", parent=styles["Normal"],
                                 fontSize=11, fontName="Helvetica-Bold",
                                 textColor=colors.HexColor("#1F4E79"), spaceAfter=2)
    consigne_title_style = ParagraphStyle("consigne_title", parent=styles["Heading2"],
                                          fontSize=12, textColor=colors.HexColor("#1F4E79"),
                                          spaceAfter=4)
    action_label_style   = ParagraphStyle("action_label", parent=styles["Normal"],
                                          fontSize=10, fontName="Helvetica-Bold",
                                          textColor=colors.white, spaceAfter=2)
    consigne_text_style  = ParagraphStyle("consigne_text", parent=styles["Normal"],
                                          fontSize=9, leftIndent=4, spaceAfter=8)

    story = []

    # Titre
    story.append(Paragraph(
        f"FEUILLE DE TOURNÉE – {tour_date.strftime('%d/%m/%Y')}", title_style))
    story.append(HRFlowable(width="100%", thickness=1,
                             color=colors.HexColor("#1F4E79"), spaceAfter=6))

    # Récapitulatif
    recap_data = [
        ["Chauffeur :",             driver_name or "—"],
        ["Dépôt de départ :", result.get("depot_depart_addr", "")],
        ["Dépôt de retour :",  result.get("depot_retour_addr", "")],
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
        "Nettoyer":     colors.HexColor("#AEC6E8"),
        "Déposer":      colors.HexColor("#B7E5B4"),
        "Retirer":      colors.HexColor("#F4B8C1"),
        "Chargement":   colors.HexColor("#D4B8E0"),
        "Déchargement": colors.HexColor("#FFE0B2"),
    }

    def P(txt, style=None):
        """Wrapper Paragraph pour word-wrap dans les cellules."""
        return Paragraph(str(txt) if txt is not None else "", style or cell_style)

    # En-têtes avec Paragraph pour alignement uniforme
    hdr_style = ParagraphStyle("hdr", parent=styles["Normal"],
                                fontSize=8, fontName="Helvetica-Bold",
                                textColor=colors.white, alignment=1)
    table_data = [[P(h, hdr_style) for h in
                   ["N°", "Action", "Produit", "Option", "Qté",
                    "Client", "Adresse", "Durée", "Pav.", "Pap.", "Arr.", "Dép.", "Obs."]]]
    row_styles = []

    depot_row_style = ParagraphStyle("dep_row", parent=styles["Normal"],
                                      fontSize=8, fontName="Helvetica-Bold", leading=10)

    # Ligne dépôt départ
    table_data.append([
        P(""), P("🏭 Dépôt départ", depot_row_style), P(""), P(""), P(""),
        P(""), P(result.get("depot_depart_addr",""), depot_row_style),
        P(""), P(""), P(""),
        P(_fmt_min(result.get("depart_min")) or "", depot_row_style), P(""), P("")
    ])
    row_styles += [("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#FFF3CD"))]

    for i, stop in enumerate(result["stops_ordered"]):
        arr_str = _fmt_min(stop.get("arrival_min")) or ""
        arr_para_style = cell_style
        if stop.get("violated"):
            arr_para_style = ParagraphStyle("viol", parent=styles["Normal"],
                                             fontSize=8, leading=10,
                                             textColor=colors.HexColor("#DC3545"),
                                             fontName="Helvetica-Bold")
        action_para_style = ParagraphStyle(f"act_{i}", parent=styles["Normal"],
                                            fontSize=8, leading=10,
                                            fontName="Helvetica-Bold",
                                            textColor=colors.HexColor("#1F4E79"))
        table_data.append([
            P(str(stop["order_num"])),
            P(stop["action"], action_para_style),
            P(stop.get("produit", "")),
            P(stop.get("option", "")),
            P(str(stop.get("qty_num", ""))),
            P(stop["client"] or ""),
            P(stop["address"]),
            P(str(stop.get("duration_min", "") or "")),
            P(_fmt_min(stop.get("tw_early")) or ""),
            P(_fmt_min(stop.get("tw_late")) or ""),
            P(arr_str, arr_para_style),
            P(_fmt_min(stop.get("departure_min")) or ""),
            P(stop.get("observations", "") or ""),
        ])
        ri = i + 2
        bg = action_bg.get(stop["action"], colors.HexColor("#F0F0F0"))
        row_styles.append(("BACKGROUND", (1, ri), (1, ri), bg))

    # Ligne dépôt retour
    table_data.append([
        P(""), P("🏁 Dépôt retour", depot_row_style), P(""), P(""), P(""),
        P(""), P(result.get("depot_retour_addr",""), depot_row_style),
        P(""), P(""), P(""), P(""),
        P(_fmt_min(result.get("return_min")) or "", depot_row_style), P("")
    ])
    last = len(table_data) - 1
    row_styles += [("BACKGROUND", (0, last), (-1, last), colors.HexColor("#FFF3CD"))]

    # Largeur utile A4 avec marges 8mm : 194mm
    # N°=7, Action=25, Produit=23, Option=17, Qté=7, Client=16, Adresse=30,
    # Durée=9, Pav.=10, Pap.=10, Arr.=10, Dép.=10, Obs.=20  → total=194mm
    CW = [7*mm, 25*mm, 23*mm, 17*mm, 7*mm, 16*mm, 30*mm, 9*mm, 10*mm, 10*mm, 10*mm, 10*mm, 20*mm]
    stops_table = Table(table_data, colWidths=CW, repeatRows=1)
    base_style = [
        ("BACKGROUND",    (0, 0), (-1, 0), colors.HexColor("#1F4E79")),
        ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
        ("ALIGN",         (6, 1), (6, -1), "LEFT"),
        ("ALIGN",         (1, 1), (3, -1), "LEFT"),
        ("ALIGN",         (12, 1), (12, -1), "LEFT"),
        ("NOSPLIT",        (1, 1), (3, -1)),
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
        ("ROWBACKGROUNDS",(0, 1), (-1, -2), [colors.white, colors.HexColor("#F9F9F9")]),
        ("GRID",          (0, 0), (-1, -1), 0.5, colors.HexColor("#CCCCCC")),
        ("TOPPADDING",    (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING",   (0, 0), (-1, -1), 3),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 3),
    ]
    stops_table.setStyle(TableStyle(base_style + row_styles))
    story.append(stops_table)
    story.append(Spacer(1, 6*mm))

    # ── Carte de la tournée ──
    story.append(Spacer(1, 6*mm))
    story.append(HRFlowable(width="100%", thickness=1,
                             color=colors.HexColor("#1F4E79"), spaceBefore=4, spaceAfter=6))
    map_title_style = ParagraphStyle("map_title", parent=styles["Heading2"],
                                      fontSize=12, textColor=colors.HexColor("#1F4E79"),
                                      spaceAfter=4)
    story.append(Paragraph("Carte de la tournée", map_title_style))
    try:
        map_bytes = generate_map_image(
            result["depot_coords"], result["stops_ordered"], result["geometry"]
        )
        img_buf = BytesIO(map_bytes)
        rl_img  = RLImage(img_buf, width=175*mm, height=120*mm)
        story.append(rl_img)
    except Exception as e:
        story.append(Paragraph(f"(Carte non disponible : {e})", styles["Normal"]))

    # ── Consignes par action ──
    story.append(Spacer(1, 6*mm))
    story.append(HRFlowable(width="100%", thickness=1,
                             color=colors.HexColor("#1F4E79"), spaceBefore=4, spaceAfter=6))
    story.append(Paragraph("Consignes par type d'intervention", consigne_title_style))

    action_header_colors = {
        "Nettoyer":     "#1f6aa5",
        "Déposer":      "#28a745",
        "Retirer":      "#dc3545",
        "Chargement":   "#7B1FA2",
        "Déchargement": "#E65100",
    }
    seen = []
    for stop in result["stops_ordered"]:
        if stop["action"] not in seen:
            seen.append(stop["action"])
    for action, consigne_text in ACTION_CONSIGNES.items():
        if action not in seen:
            continue
        hcol = colors.HexColor(action_header_colors.get(action, "#555555"))
        consigne_data = [[Paragraph(f"■  {action}", action_label_style)]]
        consigne_table = Table(consigne_data, colWidths=[194*mm])
        consigne_table.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (-1, -1), hcol),
            ("TOPPADDING",    (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING",   (0, 0), (-1, -1), 8),
        ]))
        story.append(consigne_table)
        story.append(Paragraph(consigne_text, consigne_text_style))

    # ── Checklist par action ──
    story.append(Spacer(1, 6*mm))
    story.append(HRFlowable(width="100%", thickness=1,
                             color=colors.HexColor("#1F4E79"), spaceBefore=4, spaceAfter=6))
    story.append(Paragraph("Checklist par intervention", consigne_title_style))

    ACTION_CHECKLIST = {
        "Nettoyer": [
            "Vidanger complètement la cuve",
            "Nettoyage intérieur avec produit homologué",
            "Réapprovisionnement papier / gel / savon",
            "Vérification porte et serrure",
            "Vérification état général (sol, parois)",
            "Photos / vidéos avant et après intervention",
            "Signalement des dégradations sur la fiche",
        ],
        "Déposer": [
            "Vérification propreté avant remise au client",
            "Positionnement sur zone désignée",
            "Vérification stabilité et aplomb",
            "Photos de l'installation en place",
            "Remise des consignes d'utilisation (1ère installation)",
            "Bon de livraison signé par le client",
        ],
        "Retirer": [
            "Vidange de la cuve avant enlèvement",
            "Nettoyage de la zone après retrait",
            "Photos / vidéos de l'état de l'équipement au retrait",
            "Contrôle état de l'équipement (noter dommages)",
            "Bon de retrait signé par le client",
        ],
        "Chargement": [
            "Photos de l'équipement avant chargement (état existant)",
            "Contrôle état de l'équipement avant chargement",
            "Arrimage correct du chargement",
            "Vérification charge utile respectée",
            "Vérification feux de la remorque",
        ],
        "Déchargement": [
            "Déchargement avec équipements adaptés",
            "Photos de l'équipement après déchargement",
            "Contrôle état après déchargement",
            "Positionnement sur aire de stockage désignée",
            "Bon de livraison signé",
        ],
    }
    action_header_colors_cl = {
        "Nettoyer":     "#1f6aa5",
        "Déposer":      "#28a745",
        "Retirer":      "#dc3545",
        "Chargement":   "#7B1FA2",
        "Déchargement": "#E65100",
    }
    seen_cl = []
    for stop in result["stops_ordered"]:
        if stop["action"] not in seen_cl:
            seen_cl.append(stop["action"])
    for action in seen_cl:
        checks = ACTION_CHECKLIST.get(action, [])
        if not checks:
            continue
        hcol = colors.HexColor(action_header_colors_cl.get(action, "#555555"))
        cl_title_data = [[Paragraph(f"&#9632;  {action}", cell_white)]]
        cl_title_table = Table(cl_title_data, colWidths=[194*mm])
        cl_title_table.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (-1, -1), hcol),
            ("TOPPADDING",    (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ("LEFTPADDING",   (0, 0), (-1, -1), 8),
        ]))
        story.append(cl_title_table)
        check_rows = [[
            Paragraph("☐", check_style),
            Paragraph(item, check_style)
        ] for item in checks]
        cl_table = Table(check_rows, colWidths=[8*mm, 186*mm])
        cl_table.setStyle(TableStyle([
            ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING",    (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("LEFTPADDING",   (0, 0), (0, -1), 4),
            ("ROWBACKGROUNDS",(0, 0), (-1, -1), [colors.white, colors.HexColor("#FAFAFA")]),
            ("GRID",          (0, 0), (-1, -1), 0.3, colors.HexColor("#DDDDDD")),
        ]))
        story.append(cl_table)
        story.append(Spacer(1, 3*mm))

    # ── Annotations / Observations ──
    story.append(Spacer(1, 4*mm))
    story.append(HRFlowable(width="100%", thickness=1,
                             color=colors.HexColor("#1F4E79"), spaceBefore=4, spaceAfter=6))
    story.append(Paragraph("Annotations / Observations", obs_style))
    story.append(Paragraph(
        "À remplir par le chauffeur durant ou à l'issue de la tournée :",
        ParagraphStyle("obs_sub", parent=styles["Normal"], fontSize=8,
                       textColor=colors.grey, spaceAfter=4)
    ))
    obs_rows = [[""] for _ in range(8)]
    obs_table = Table(obs_rows, colWidths=[194*mm], rowHeights=[10*mm]*8)
    obs_table.setStyle(TableStyle([
        ("GRID",          (0, 0), (-1, -1), 0.5, colors.HexColor("#CCCCCC")),
        ("BACKGROUND",    (0, 0), (-1, -1), colors.white),
        ("TOPPADDING",    (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    story.append(obs_table)
    story.append(Spacer(1, 4*mm))

    # ── Footer ──
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
    st.subheader("⚖️ Règles de travail")
    st.session_state.heure_min_depart = st.time_input(
        "🕖 Départ au plus tôt",
        value=st.session_state.heure_min_depart,
        step=300,
        help=(
            "Heure minimale de départ du dépôt. "
            "En France, le Code du travail impose un repos quotidien de 11h consécutives "
            "(art. L3131-1) et le transport routier interdit le travail avant 5h00 "
            "dans la plupart des conventions collectives. "
            "Valeur recommandée : 07h00."
        )
    )
    st.caption(
        "ℹ️ *Code du travail — art. L3131-1 :* "
        "repos quotidien minimum de **11h consécutives**. "
        "Convention collective transport : départ rarement avant **06h00**."
    )

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
    st.subheader("🏭 Dépôts")
    depot_options = list(DEPOTS.keys())

    st.session_state.depot_depart_key = st.selectbox(
        "📍 Dépôt de départ",
        options=depot_options,
        index=depot_options.index(st.session_state.depot_depart_key)
              if st.session_state.depot_depart_key in depot_options else 0,
    )
    st.caption(f"📌 {DEPOTS[st.session_state.depot_depart_key]}")

    st.session_state.depot_retour_key = st.selectbox(
        "🏁 Dépôt de retour",
        options=depot_options,
        index=depot_options.index(st.session_state.depot_retour_key)
              if st.session_state.depot_retour_key in depot_options else 0,
    )
    st.caption(f"📌 {DEPOTS[st.session_state.depot_retour_key]}")

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

st.title("Optimiseur automatique des tournées - Deldossi Assainissement")
st.caption(f"🏭 Départ : **{st.session_state.depot_depart_key}** · Retour : **{st.session_state.depot_retour_key}**")

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
        f"Départ : **{st.session_state.depot_depart_key}** · "
        f"Retour : **{st.session_state.depot_retour_key}**. "
        f"L'ordre de saisie n'a pas d'importance.")

    # ── Fusion des éditions en cours dans df_stops avant tout bouton ──
    # On récupère les modifications du data_editor depuis son état interne
    # (st.session_state["editor_stops"]) pour ne pas perdre les saisies.
    def _flush_editor():
        """Applique les édits en cours du tableau dans df_stops."""
        key = "editor_stops"
        if key not in st.session_state:
            return
        state = st.session_state[key]
        df    = st.session_state.df_stops.copy()
        # Éditions de cellules
        for row_idx, cols in (state.get("edited_rows") or {}).items():
            for col, val in cols.items():
                df.at[int(row_idx), col] = val
        # Lignes ajoutées via le "+" natif du data_editor
        for row in (state.get("added_rows") or []):
            df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
        # Lignes supprimées via la corbeille native
        deleted = sorted(state.get("deleted_rows") or [], reverse=True)
        for idx in deleted:
            df = df.drop(index=idx).reset_index(drop=True)
        st.session_state.df_stops = df

    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("➕ Ajouter un arrêt", use_container_width=True):
            _flush_editor()
            new_row = pd.DataFrame({
                "Action": ["Nettoyer"], "Produit": ["WC chimique"],
                "Option": ["Lave-main"], "Quantité": [1],
                "Nom du client": [""], "Adresse": [""],
                "Durée (min)": [30], "Pas avant": [""], "Pas après": [""],
                "Observations": [""],
            })
            st.session_state.df_stops = pd.concat(
                [st.session_state.df_stops, new_row], ignore_index=True)
            # Reset l'état interne du data_editor pour repartir proprement
            if "editor_stops" in st.session_state:
                del st.session_state["editor_stops"]
            st.rerun()
    with c2:
        if st.button("➖ Supprimer le dernier", use_container_width=True):
            _flush_editor()
            if len(st.session_state.df_stops) > 1:
                st.session_state.df_stops = (
                    st.session_state.df_stops.iloc[:-1].reset_index(drop=True))
            if "editor_stops" in st.session_state:
                del st.session_state["editor_stops"]
            st.rerun()
    with c3:
        if st.button("🗑️ Tout vider", use_container_width=True, type="secondary"):
            st.session_state.df_stops = _init_df()
            st.session_state.result   = None
            if "editor_stops" in st.session_state:
                del st.session_state["editor_stops"]
            st.rerun()

    st.markdown("---")
    leg_cols = st.columns(3)
    for col, (action, color) in zip(leg_cols, ACTION_COLORS.items()):
        col.markdown(
            f'<span style="background:{color};padding:3px 10px;'
            f'border-radius:4px;font-size:0.85em">■ {action}</span>',
            unsafe_allow_html=True)
    st.markdown("")

    st.data_editor(
        st.session_state.df_stops,
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "Action": st.column_config.SelectboxColumn(
                "Action", required=True, width="small",
                options=["Nettoyer", "Déposer", "Retirer", "Chargement", "Déchargement"]),
            "Produit": st.column_config.SelectboxColumn(
                "Produit", required=True, width="medium",
                options=["WC chimique", "Lave-main", "Urinoir", "WC handicapé"]),
            "Option": st.column_config.SelectboxColumn(
                "Option (WC chim. uniquement)", width="medium",
                options=["", "Lave-main", "Urinoir"],
                help="Obligatoire pour WC chimique. Laissez vide pour les autres produits."),
            "Quantité": st.column_config.NumberColumn(
                "Qté", required=True, width="small",
                min_value=1, max_value=20, step=1, default=1,
                help="Nombre d'unités. Pour WC chimique : 1 option par WC."),
            "Nom du client": st.column_config.TextColumn(
                "Nom du client", width="medium"),
            "Adresse": st.column_config.TextColumn(
                "Adresse complète (rue, ville, CP)", width="large",
                help="Ex : Place d'Hautpoul 81600 Gaillac"),
            "Durée (min)": st.column_config.NumberColumn(
                "⏱ Durée (min)", width="small", min_value=1, max_value=480,
                step=5, default=30,
                help="Durée de l'intervention sur place en minutes (ex: 30)"),
            "Pas avant": st.column_config.TextColumn(
                "⏰ Pas avant", width="small",
                help="Arriver au plus tôt à cette heure (format HH:MM). Ex : 09:00"),
            "Pas après": st.column_config.TextColumn(
                "⏰ Pas après", width="small",
                help="Arriver au plus tard à cette heure (format HH:MM). Ex : 11:30"),
            "Observations": st.column_config.TextColumn(
                "📝 Observations", width="large",
                help="Notes ou remarques particulières pour cet arrêt (ex : code portail, contact sur place…)"),
        },
        hide_index=False,
        key="editor_stops",
    )
    # Ne pas réécrire df_stops ici — c'est _flush_editor() qui le fait
    # uniquement quand une action (bouton, optimisation) est déclenchée.

    # Comptage des adresses valides en lisant l'état en temps réel
    def _current_df():
        """Retourne le df avec les éditions non encore flushées."""
        key   = "editor_stops"
        df    = st.session_state.df_stops.copy()
        if key not in st.session_state:
            return df
        state = st.session_state[key]
        for row_idx, cols in (state.get("edited_rows") or {}).items():
            for col, val in cols.items():
                df.at[int(row_idx), col] = val
        for row in (state.get("added_rows") or []):
            df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
        deleted = sorted(state.get("deleted_rows") or [], reverse=True)
        for idx in deleted:
            df = df.drop(index=idx).reset_index(drop=True)
        return df

    live_df    = _current_df()
    valid_rows = live_df[live_df["Adresse"].fillna("").str.strip() != ""]
    n_valid    = len(valid_rows)
    st.caption(f"📍 **{n_valid}** arrêt(s) avec adresse renseignée")

    # Validation Produit / Option
    if "Produit" in live_df.columns and "Option" in live_df.columns:
        wc_rows_no_opt = valid_rows[
            (valid_rows["Produit"] == "WC chimique") &
            (valid_rows["Option"].fillna("").str.strip() == "")
        ]
        if len(wc_rows_no_opt) > 0:
            st.warning(
                f"\U0001f6bd **{len(wc_rows_no_opt)} arr\u00eat(s) avec WC chimique sans option.** "
                "Chaque WC chimique doit \u00eatre accompagn\u00e9 d'un **Urinoir** ou d'un "
                "**Lave-main** (Code du travail, art. R4228-7). "
                "S\u00e9lectionnez une option dans la colonne *Option*."
            )
        non_wc_with_opt = valid_rows[
            (valid_rows["Produit"] != "WC chimique") &
            (valid_rows["Option"].fillna("").str.strip() != "")
        ]
        if len(non_wc_with_opt) > 0:
            st.info(
                f"\u2139\ufe0f **{len(non_wc_with_opt)} arr\u00eat(s)** ont une option renseign\u00e9e "
                "alors que le produit n'est pas un WC chimique — "
                "la colonne *Option* sera ignor\u00e9e pour ces lignes."
            )

    st.markdown("---")

    if st.button("🚀 Optimiser la tournée", type="primary",
                 use_container_width=True, disabled=(n_valid < 1)):
        _flush_editor()
        valid_stops = st.session_state.df_stops[
            st.session_state.df_stops["Adresse"].fillna("").str.strip() != ""
        ].reset_index(drop=True)

        with st.spinner("🔍 Géocodage des adresses en cours…"):
            # Géocodage des dépôts
            depot_depart_addr = DEPOTS[st.session_state.depot_depart_key]
            depot_retour_addr = DEPOTS[st.session_state.depot_retour_key]
            if not depot_depart_addr:
                st.error("❌ Veuillez renseigner le **Dépôt de départ** dans la barre latérale.")
                st.stop()
            depot_coords = geocode(depot_depart_addr)
            if not depot_coords:
                st.error(f"❌ Impossible de géocoder le dépôt de départ : {depot_depart_addr}")
                st.stop()
            if depot_retour_addr and depot_retour_addr != depot_depart_addr:
                depot_retour_coords = geocode(depot_retour_addr)
                if not depot_retour_coords:
                    st.warning(f"⚠️ Dépôt de retour introuvable, le dépôt de départ sera utilisé.")
                    depot_retour_coords = depot_coords
            else:
                depot_retour_coords = depot_coords

            pb           = st.progress(0, text="Géocodage des arrêts…")
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

        # Fenêtres temporelles (index 0 = dépôt, pas de contrainte)
        time_windows = [{"earliest": None, "latest": None, "duration": 0}]  # index 0 = dépôt
        for _, row in valid_stops.iterrows():
            if geo.get(row["Adresse"]):
                try:
                    dur = int(float(row.get("Durée (min)", 30) or 30))
                except (ValueError, TypeError):
                    dur = 30
                time_windows.append({
                    "earliest": _parse_hhmm(row.get("Pas avant", "")),
                    "latest":   _parse_hhmm(row.get("Pas après", "")),
                    "duration": dur,
                })

        with st.spinner("🗺️ Calcul de l'itinéraire optimisé…"):
            # Première passe sans heure de départ (ordre pur)
            trip = osrm_trip(coords_list, time_windows=time_windows, depart_min=None)
            if not trip:
                st.stop()
            orig = osrm_route_distance(coords_list)

        # Borne légale / réglementaire configurée dans la sidebar
        legal_min = (st.session_state.heure_min_depart.hour * 60
                     + st.session_state.heure_min_depart.minute)

        # Calcul automatique de l'heure de départ optimale (purement mathématique)
        depart_min_opt = _compute_optimal_departure(trip["order"], trip["matrix"], time_windows)

        if depart_min_opt is not None:
            depart_min_raw = depart_min_opt   # résultat pur sans contrainte légale

            if depart_min_raw < legal_min:
                # Le calcul pur tomberait avant l'heure légale → on clamp
                depart_min = legal_min
                st.warning(
                    f"⚠️ Le calcul théorique suggérait un départ à **{_fmt_min(depart_min_raw)}**, "
                    f"ce qui ne respecte pas l'heure de départ au plus tôt fixée à "
                    f"**{_fmt_min(legal_min)}**. "
                    f"Le départ est donc fixé à **{_fmt_min(depart_min)}**. "
                    f"Certaines contraintes client marquées 'Pas après' risquent de ne pas "
                    f"pouvoir être respectées — vérifiez les arrêts en rouge ci-dessous."
                )
            else:
                depart_min = depart_min_raw
                st.success(
                    f"🕖 **Heure de départ recommandée : {_fmt_min(depart_min)}** "
                    f"— calculée automatiquement pour respecter toutes les contraintes "
                    f"avec le minimum d'attente (départ légal : {_fmt_min(legal_min)})."
                )
        else:
            # Aucune contrainte client : on part à l'heure légale minimale
            depart_min = legal_min
            st.info(
                f"🕖 **Aucune contrainte horaire client renseignée.** "
                f"Heure de départ : **{_fmt_min(depart_min)}** "
                f"(heure légale minimale configurée)."
            )

        # Seconde passe 2-opt avec l'heure de départ calculée (affine si TW présentes)
        has_tw = any(
            tw.get("earliest") is not None or tw.get("latest") is not None
            for tw in time_windows[1:]
        )
        if has_tw:
            with st.spinner("⚙️ Affinage de l'ordre selon les contraintes horaires…"):
                trip2 = osrm_trip(coords_list, time_windows=time_windows, depart_min=depart_min)
                if trip2:
                    trip = trip2

        # Heures d'arrivée réelles
        arrivals = _compute_arrivals(trip["order"], trip["matrix"], depart_min, time_windows)

        order         = trip["order"]
        stops_ordered = []
        rank          = 1
        arr_idx       = 0
        for orig_idx in order:
            if orig_idx == 0:
                continue
            sri = orig_idx - 1
            if sri >= len(valid_stops):
                continue
            row    = valid_stops.iloc[sri]
            coords = geo.get(row["Adresse"])
            arr    = arrivals[arr_idx] if arr_idx < len(arrivals) else {}
            try:
                dur_stop = int(float(row.get("Durée (min)", 30) or 30))
            except (ValueError, TypeError):
                dur_stop = 30
            stops_ordered.append({
                "order_num":    rank,
                "action":       row["Action"],
                "quantity":     _qty_label(row),
                "produit":      row.get("Produit", ""),
                "option":       row.get("Option",  ""),
                "qty_num":      int(row.get("Quantité", 1) or 1),
                "client":       row["Nom du client"],
                "address":      row["Adresse"],
                "lat":          coords[0] if coords else None,
                "lon":          coords[1] if coords else None,
                "tw_early":     _parse_hhmm(row.get("Pas avant", "")),
                "tw_late":      _parse_hhmm(row.get("Pas après", "")),
                "duration_min": dur_stop,
                "arrival_min":  arr.get("arrival_min"),
                "departure_min":arr.get("departure_min"),
                "wait_min":     arr.get("wait_min", 0),
                "violated":     arr.get("violated", False),
                "observations": str(row.get("Observations", "") or ""),
            })
            rank    += 1
            arr_idx += 1

        dist_km   = trip["distance_km"]
        dur_min   = trip["duration_min"]
        fuel_l    = dist_km * fuel_conso / 100
        fuel_cost = fuel_l * fuel_price
        km_saved  = max(0, (orig["distance_km"]  - dist_km))  if orig else 0
        min_saved = max(0, (orig["duration_min"] - dur_min)) if orig else 0

        # Avertissements fenêtres non respectables
        violated_stops = [s for s in stops_ordered if s["violated"]]
        if violated_stops:
            lines = ["⚠️ Certaines contraintes horaires ne peuvent pas être respectées :"]
            for s in violated_stops:
                lines.append(
                    f"- **Arrêt {s['order_num']}** {s['address']} : "
                    f"arrivée estimée {_fmt_min(s['arrival_min'])} "
                    f"(limite : {_fmt_min(s['tw_late'])})"
                )
            st.warning("  \n".join(lines))

        # Temps de retour au dépôt de retour via OSRM /route
        if stops_ordered:
            last_stop_coords = [s for s in [
                (stops_ordered[-1]["lat"], stops_ordered[-1]["lon"])
            ] if s[0] is not None]
            if last_stop_coords and depot_retour_coords:
                try:
                    ret_coord_str = f"{last_stop_coords[0][1]},{last_stop_coords[0][0]};{depot_retour_coords[1]},{depot_retour_coords[0]}"
                    ret_r = requests.get(f"{OSRM_URL}/route/v1/driving/{ret_coord_str}",
                                         params={"overview": "false"}, timeout=10)
                    ret_data = ret_r.json()
                    travel_back = ret_data["routes"][0]["duration"] / 60 if ret_data.get("code") == "Ok" else 0
                except Exception:
                    travel_back = (trip["matrix"][order[-1]][order[0]] or 0) / 60
            else:
                travel_back = 0
            last_dep = stops_ordered[-1]["departure_min"] or stops_ordered[-1]["arrival_min"] or depart_min
            return_min = last_dep + travel_back
        else:
            return_min = depart_min

        st.session_state.result = {
            "stops_ordered":      stops_ordered,
            "distance_km":        dist_km,
            "duration_min":       dur_min,
            "fuel_liters":        fuel_l,
            "fuel_cost":          fuel_cost,
            "km_saved":           km_saved,
            "time_saved_min":     min_saved,
            "geometry":           trip["geometry"],
            "depot_coords":       depot_coords,
            "depot_retour_coords":depot_retour_coords,
            "depot_depart_addr":  depot_depart_addr,
            "depot_retour_addr":  depot_retour_addr,
            "depart_min":         depart_min,
            "return_min":         return_min,
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
        c1, c2, c3, c4, c5, c6, c7, c8 = st.columns(8)
        c1.metric("📍 Arrêts",        len(r["stops_ordered"]))
        c2.metric("🛣️ Distance",       f"{r['distance_km']:.1f} km")
        c3.metric("⏱️ Durée trajet",   f"{hours}h{mins:02d}")
        c4.metric("⛽ Carburant",      f"{r['fuel_liters']:.1f} L")
        c5.metric("💶 Coût carburant", f"{r['fuel_cost']:.2f} €")
        c6.metric("⏳ Temps gagné",    f"{int(r['time_saved_min'])} min",
                  delta=f"-{r['km_saved']:.1f} km", delta_color="inverse")
        c7.metric("🕖 Départ dépôt",   _fmt_min(r.get("depart_min")))
        c8.metric("🏁 Retour dépôt",   _fmt_min(r.get("return_min")))

        # Durée totale (trajet + toutes les interventions)
        if r.get("depart_min") is not None and r.get("return_min") is not None:
            total_tour_min = r["return_min"] - r["depart_min"]
            h_tot = int(total_tour_min // 60)
            m_tot = int(total_tour_min % 60)
            interv_total = sum(s.get("duration_min", 0) for s in r["stops_ordered"])
            st.info(
                f"🗓️ **Durée totale de la tournée** (trajet + interventions) : "
                f"**{h_tot}h{m_tot:02d}** "
                f"— dont {int(r['duration_min'])} min de trajet "
                f"et {interv_total} min d'interventions"
            )

        st.markdown("---")
        col_map, col_list = st.columns([3, 2])

        with col_map:
            st.subheader("🗺️ Carte de la tournée")
            m = build_map(r["depot_coords"], r["stops_ordered"], r["geometry"],
                          depot_depart_addr=r.get("depot_depart_addr","Dépôt départ"),
                          depot_retour_addr=r.get("depot_retour_addr"),
                          depot_retour_coords=r.get("depot_retour_coords"))
            st_folium(m, use_container_width=True, height=520, returned_objects=[])

        with col_list:
            st.subheader("📋 Ordre des arrêts")
            st.markdown(
                f'<div class="stop-card depot-card"><b>🏭 Dépôt — Départ</b><br>'
                f'<small>{r.get("depot_depart_addr","")}</small></div>', unsafe_allow_html=True)
            for stop in r["stops_ordered"]:
                bg  = ACTION_COLORS.get(stop["action"], "#f0f0f0")
                brd = ACTION_BORDER_COLORS.get(stop["action"], "#999")
                cli = f" · {stop['client']}" if stop['client'] else ""
                # Contrainte horaire
                tw_parts = []
                if stop.get("tw_early"):
                    tw_parts.append(f"⏰ Pas avant {_fmt_min(stop['tw_early'])}")
                if stop.get("tw_late"):
                    tw_parts.append(f"⏰ Pas après {_fmt_min(stop['tw_late'])}")
                tw_html = (f"<br><small style='color:#666'>"
                           + " · ".join(tw_parts) + "</small>") if tw_parts else ""
                # Heure d'arrivée estimée
                arr_str = _fmt_min(stop.get("arrival_min"))
                wait    = stop.get("wait_min", 0)
                wait_html = (f" <small style='color:#999'>(attente {int(wait)} min)</small>"
                             if wait and wait > 0.5 else "")
                violated = stop.get("violated", False)
                arr_color = "#dc3545" if violated else "#28a745"
                arr_icon  = "⚠️" if violated else "🕐"
                dep_str  = _fmt_min(stop.get("departure_min"))
                dur_str  = stop.get("duration_min", 0)
                dur_html = (f" <small style='color:#555'>(intervention : {dur_str} min → départ {dep_str})</small>"
                            if dur_str and dep_str else "")
                arr_html  = (f"<br><small style='color:{arr_color};font-weight:600'>"
                             f"{arr_icon} Arrivée estimée : {arr_str}</small>{wait_html}{dur_html}"
                             if arr_str else "")
                st.markdown(
                    f'<div class="stop-card" style="background:{bg};'
                    f'border-left:4px solid {brd};">'
                    f'<b>#{stop["order_num"]} {stop["action"]}</b>{cli}<br>'
                    f'<small>📍 {stop["address"]}</small><br>'
                    f'<small>📦 {stop["quantity"]}</small>'
                    f'{tw_html}{arr_html}</div>',
                    unsafe_allow_html=True)
            retour_str = _fmt_min(r.get("return_min"))
            retour_html = (f"<br><small style='color:#1f4e79;font-weight:600'>🏁 Retour estimé : {retour_str}</small>"
                           if retour_str else "")
            st.markdown(
                f'<div class="stop-card depot-card"><b>🏁 Dépôt — Retour</b><br>'
                f'<small>{r.get("depot_retour_addr","")}</small>{retour_html}</div>',
                unsafe_allow_html=True)

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
                file_name=f"Tournee_WC_{st.session_state.tour_date.strftime('%d-%m-%Y')}.xlsx",
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
                file_name=f"Tournee_WC_{st.session_state.tour_date.strftime('%d-%m-%Y')}.pdf",
                mime="application/pdf",
                use_container_width=True, type="primary")

        st.markdown("---")
        st.caption("💡 Le fichier Excel contient une colonne **Heure de passage** "
                   "à renseigner manuellement et une colonne **✓ Fait** pour validation terrain.")
