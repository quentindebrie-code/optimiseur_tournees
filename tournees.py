"""Test de non-régression : l'export doit refléter l'ordre manuel.

Scénario reproduit :
  1. une tournée optimisée à 4 arrêts (A, B, C, D) ;
  2. l'utilisateur réordonne manuellement en D, C, B, A ;
  3. on génère Excel et PDF à partir du résultat manuel ;
  4. on vérifie que l'ordre écrit dans les fichiers est bien D, C, B, A.
"""
import datetime, io, re
import openpyxl
import core_test as T

ADDR = ["A - Rue Alpha", "B - Rue Bravo", "C - Rue Charlie", "D - Rue Delta"]

def stop(i):
    return {
        "order_num": i + 1, "action": "Nettoyer", "produit": "WC chimique",
        "option": "", "couleur": "", "quantity": "1 u", "qty_num": 1,
        "client": f"Client {ADDR[i][0]}", "address": ADDR[i],
        "duration_min": 30, "tw_early": None, "tw_late": None,
        "arrival_min": 480 + i * 60, "departure_min": 510 + i * 60,
        "wait_min": 0, "violated": False, "observations": "",
        "lat": 43.6 + i * 0.01, "lon": 1.44 + i * 0.01,
    }

n = 4
# Matrice de trajets fictive (secondes), indice 0 = dépôt
matrix = [[0 if i == j else 900 for j in range(n + 1)] for i in range(n + 1)]

result = {
    "stops_ordered": [stop(i) for i in range(n)],
    "distance_km": 100.0, "duration_min": 240.0,
    "fuel_liters": 12.2, "fuel_cost": 22.5,
    "km_saved": 10.0, "time_saved_min": 25,
    "geometry": [], "depot_coords": (43.60, 1.44),
    "depot_retour_coords": (43.60, 1.44),
    "depot_depart_addr": "Dépôt Toulouse", "depot_retour_addr": "Dépôt Toulouse",
    "depart_min": 420, "return_min": 780,
    "pause_dejeuner": True, "etat_lieux_par_arret": False,
    "peri_affectes_nb": 0, "stop_matrix": matrix,
    "fuel_conso": 12.2, "fuel_price": 1.85, "besoins_records": [],
}

# ── 1. Réorganisation manuelle : ordre inversé ──
new_order = [3, 2, 1, 0]
manual = T._recalc_manual_route(result, new_order)

assert manual["is_manual"] is True
assert manual["manual_order"] == new_order
ordre_manuel = [s["address"] for s in manual["stops_ordered"]]
attendu = [ADDR[i] for i in new_order]
assert ordre_manuel == attendu, ordre_manuel
assert [s["order_num"] for s in manual["stops_ordered"]] == [1, 2, 3, 4]
print("1. _recalc_manual_route  → ordre =", ordre_manuel)
print("   recalc_ok =", manual["recalc_ok"], "(False attendu : OSRM injoignable ici)")

d = datetime.date(2026, 8, 13)

def ordre_dans_excel(res):
    wb = openpyxl.load_workbook(io.BytesIO(T.export_excel(res, d, "Chauffeur", 1.85).getvalue()))
    ws = wb["Feuille de tournée"]
    out = []
    for row in ws.iter_rows(values_only=True):
        if row and isinstance(row[0], int) and isinstance(row[7], str) and row[7] in ADDR:
            out.append(row[7])
    return out, ws

# ── 2. Export Excel du résultat MANUEL ──
xl_manuel, ws = ordre_dans_excel(manual)
assert xl_manuel == attendu, xl_manuel
entete = [c[1].value for c in ws.iter_rows(min_row=1, max_row=14) if c[0].value == "Ordre des arrêts :"]
print("2. export_excel(manuel)  → ordre =", xl_manuel)
print("   entête XLSX          →", entete[0][:55], "…")

# ── 3. Export Excel du résultat OPTIMISÉ (contrôle : inchangé) ──
xl_optim, _ = ordre_dans_excel(result)
assert xl_optim == ADDR, xl_optim
print("3. export_excel(optim)   → ordre =", xl_optim, "(non régressé)")

# ── 4. Export PDF du résultat MANUEL ──
pdf_bytes = T.export_pdf(manual, d, "Chauffeur").getvalue()
assert pdf_bytes[:4] == b"%PDF"
try:
    from pypdf import PdfReader
    txt = "".join(p.extract_text() for p in PdfReader(io.BytesIO(pdf_bytes)).pages)
    pos = [txt.find(a.split(" - ")[1]) for a in attendu]
    assert all(p >= 0 for p in pos), pos
    assert pos == sorted(pos), pos
    print("4. export_pdf(manuel)    → ordre respecté dans le texte du PDF", pos)
except ImportError:
    print("4. export_pdf(manuel)    → PDF généré (%d octets), pypdf absent" % len(pdf_bytes))

print("\n✅ TOUS LES TESTS PASSENT — l'export suit l'ordre manuel.")
