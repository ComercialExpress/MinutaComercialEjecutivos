#!/usr/bin/env python3
"""
Sync RegistrosCallback (SharePoint List) -> JSON para el panel de Callback JT.

Genera:
  data/callback_registros.json  -> un registro por callback calificado, con
                                    fecha, mes (YYYY-MM), sucursal, categoria
                                    (buena/mala/neutra) y origen del problema.

El HTML (PanelCallbackEpa.html) consume este único archivo y hace toda la
segmentación por mes/sucursal en el cliente, así que este script NO agrega
por sucursal ni por mes — solo aplana y clasifica cada registro. Esto evita
tener que regenerar agregaciones distintas cada vez que se agrega un nuevo
corte (mes, sucursal, ambos a la vez, etc.).

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
import datetime
import requests

TENANT_ID = os.environ["TENANT_ID"]          # a2c91f5e-f3c4-4e1a-a85e-d114d2650b65
CLIENT_ID = os.environ["CLIENT_ID"]           # 9abfdc51-4b4c-40f8-86be-4ceb3bd1306f
CLIENT_SECRET = os.environ["CLIENT_SECRET"]   # NUNCA hardcodear, siempre desde secret

SITE_ID = "clcomercialexpress.sharepoint.com,d331a44f-8002-4cab-8db4-061cb13eb493,1d32dc11-bde8-40af-8d21-d070452a7b2b"
LIST_ID = "e6f62de4-a0ba-404d-9885-d75c0df7ba29"

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
# El JSON vive en la RAÍZ del repo, igual que ejecutivos.json / kpi.json / tienda.json
OUT_DIR = "."

# Regla de agregación Calif_Global (texto/entero "1".."5") -> categoría.
# 1-2 = mala, 3 = neutra, 4-5 = buena. Ajustar acá si el corte real de
# negocio cambia — es el único lugar donde vive esta regla.
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


def parse_fecha(valor):
    """SharePoint suele devolver fechas ISO (YYYY-MM-DDTHH:MM:SSZ) via Graph,
    pero se deja tolerancia a dd/mm/aaaa por si la columna viaja como texto."""
    if not valor:
        return None
    valor = str(valor).strip()
    for fmt in ("%Y-%m-%dT%H:%M:%SZ", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.datetime.strptime(valor, fmt).date()
        except ValueError:
            continue
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

    # --- DEBUG TEMPORAL: ver TODAS las claves reales que trae 'fields' ---
    if raw_items:
        f0 = raw_items[0].get("fields", {})
        print(f"[DEBUG] claves disponibles en fields: {sorted(f0.keys())}", file=sys.stderr)
        print(f"[DEBUG] fields completo item 0: {f0}", file=sys.stderr)
    # --- FIN DEBUG TEMPORAL ---

    registros = []
    excluidos = 0

    for item in raw_items:
        f = item.get("fields", {})
        sucursal = (f.get("SucursalPDV") or "").strip()
        origen = (f.get("OrigenProblema") or "").strip() or None
        categoria = calif_to_categoria(f.get("Calif_Global"))
        fecha = parse_fecha(f.get("FechaAtencion"))

        if not sucursal or categoria is None or fecha is None:
            excluidos += 1
            continue

        registros.append({
            "fecha": fecha.strftime("%Y-%m-%d"),
            "mes": fecha.strftime("%Y-%m"),
            "sucursal": sucursal,
            "categoria": categoria,
            "origen": origen,
        })

    if excluidos:
        print(
            f"Registros excluidos (sin SucursalPDV, Calif_Global inválido o "
            f"FechaAtencion no parseable): {excluidos}",
            file=sys.stderr,
        )

    # Orden estable por fecha para que el diff del commit sea legible.
    registros.sort(key=lambda r: (r["fecha"], r["sucursal"]))

    with open(os.path.join(OUT_DIR, "callback_registros.json"), "w", encoding="utf-8") as fh:
        json.dump(registros, fh, ensure_ascii=False, indent=2)

    print(f"OK: callback_registros.json generado en la raíz del repo ({len(registros)} registros).", file=sys.stderr)


if __name__ == "__main__":
    main()
