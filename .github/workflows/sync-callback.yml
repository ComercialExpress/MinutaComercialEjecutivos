#!/usr/bin/env python3
"""
Sync RegistrosCallback (SharePoint List) -> JSON para el panel de Callback JT.

Genera:
  data/callback_sucursal.json   -> Buena/Mala/Neutra por SucursalPDV
  data/origen_problema.json     -> Q y % por OrigenProblema

Auth: app-only (client credentials) contra Microsoft Graph, misma app
      'MinutaComercialEjecutivos' ya usada en el flujo de Minuta de Ejecutivos.

Variables de entorno requeridas (se configuran como GitHub Secrets):
  TENANT_ID
  CLIENT_ID
  CLIENT_SECRET
"""

import os
import sys
import json
import requests

TENANT_ID = os.environ["TENANT_ID"]          # a2c91f5e-f3c4-4e1a-a85e-d114d2650b65
CLIENT_ID = os.environ["CLIENT_ID"]           # 9abfdc51-4b4c-40f8-86be-4ceb3bd1306f
CLIENT_SECRET = os.environ["CLIENT_SECRET"]   # NUNCA hardcodear, siempre desde secret

SITE_ID = "clcomercialexpress.sharepoint.com,d331a44f-8002-4cab-8db4-061cb13eb493,1d32dc11-bde8-40af-8d21-d070452a7b2b"
LIST_ID = "e6f62de4-a0ba-404d-9885-d75c0df7ba29"

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
# Los JSON viven en la RAÍZ del repo, igual que ejecutivos.json / kpi.json / tienda.json
OUT_DIR = "."

# Regla de agregación Calif_Global (texto "1".."5") -> categoría.
# Ajustar acá si el corte real de negocio es distinto.
def calif_to_categoria(valor_texto):
    try:
        n = int(str(valor_texto).strip())
    except (TypeError, ValueError):
        return None  # vacío / no numérico -> se excluye del conteo
    if n <= 2:
        return "mala"
    if n == 3:
        return "neutra"
    if n >= 4:
        return "buena"
    return None


def get_access_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }
    resp = requests.post(url, data=data, timeout=30)
    resp.raise_for_status()
    return resp.json()["access_token"]


def fetch_all_items(token):
    """Trae todos los items de la lista, paginando con @odata.nextLink."""
    headers = {"Authorization": f"Bearer {token}"}
    url = (
        f"{GRAPH_BASE}/sites/{SITE_ID}/lists/{LIST_ID}/items"
        f"?expand=fields&$top=200"
    )
    items = []
    while url:
        resp = requests.get(url, headers=headers, timeout=30)
        resp.raise_for_status()
        payload = resp.json()
        items.extend(payload.get("value", []))
        url = payload.get("@odata.nextLink")
    return items


def main():
    token = get_access_token()
    raw_items = fetch_all_items(token)
    print(f"Items obtenidos de RegistrosCallback: {len(raw_items)}", file=sys.stderr)

    sucursal_agg = {}   # { sucursal: {"buena":0,"mala":0,"neutra":0} }
    origen_agg = {}      # { origen: count }
    excluidos = 0

    for item in raw_items:
        f = item.get("fields", {})
        sucursal = (f.get("SucursalPDV") or "").strip()
        origen = (f.get("OrigenProblema") or "").strip()
        categoria = calif_to_categoria(f.get("Calif_Global"))

        if sucursal and categoria:
            sucursal_agg.setdefault(sucursal, {"buena": 0, "mala": 0, "neutra": 0})
            sucursal_agg[sucursal][categoria] += 1
        else:
            excluidos += 1

        if origen:
            origen_agg[origen] = origen_agg.get(origen, 0) + 1

    if excluidos:
        print(f"Registros excluidos del conteo por sucursal (sin SucursalPDV o Calif_Global inválido): {excluidos}", file=sys.stderr)

    # ---- Shape: callback_sucursal.json ----
    callback_sucursal = [
        {
            "sucursal": nombre,
            "buena": v["buena"],
            "mala": v["mala"],
            "neutra": v["neutra"],
            "total": v["buena"] + v["mala"] + v["neutra"],
        }
        for nombre, v in sorted(sucursal_agg.items())
    ]

    # ---- Shape: origen_problema.json ----
    total_origen = sum(origen_agg.values()) or 1
    origen_problema = [
        {
            "origen": nombre,
            "q": q,
            "pct": round(q / total_origen * 100, 2),
        }
        for nombre, q in sorted(origen_agg.items(), key=lambda kv: -kv[1])
    ]

    with open(os.path.join(OUT_DIR, "callback_sucursal.json"), "w", encoding="utf-8") as fh:
        json.dump(callback_sucursal, fh, ensure_ascii=False, indent=2)
    with open(os.path.join(OUT_DIR, "origen_problema.json"), "w", encoding="utf-8") as fh:
        json.dump(origen_problema, fh, ensure_ascii=False, indent=2)

    print("OK: callback_sucursal.json y origen_problema.json generados en la raíz del repo.", file=sys.stderr)


if __name__ == "__main__":
    main()
