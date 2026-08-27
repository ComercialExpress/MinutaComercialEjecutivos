#!/usr/bin/env python3
"""
Sync DepositosSucursalesBCI (SharePoint List) -> JSON para Cuadratura de Caja.

Genera:
  depositos.json  -> un registro por depósito bancario informado desde el
                      formulario de sucursales, tal cual viene de la Lista
                      (sin agregar por sucursal ni por fecha).

El HTML (cuadratura_caja.html) consume este único archivo y hace toda la
agrupación por sucursal + FechaRecaudacion (sumando MontoDepositado si hay
más de un registro el mismo día) en el cliente — mismo criterio que
sync_callback.py con callback_registros.json: este script NO agrega, solo
aplana cada item de la Lista.

Auth: app-only (client credentials) contra Microsoft Graph, misma app
      'MinutaComercialEjecutivos' ya usada en el resto del ecosistema.

Variables de entorno requeridas (ya existen como GitHub Secrets en este
repo, reusadas de sync_callback.py / sync-data.yml — no hace falta crear
ninguna nueva):
  TENANT_ID
  CLIENT_ID
  CLIENT_SECRET
"""

import os
import sys
import json
import requests

TENANT_ID = os.environ["TENANT_ID"]
CLIENT_ID = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]  # NUNCA hardcodear, siempre desde secret

# ⚠️ Asumido igual al SITE_ID de sync_callback.py, porque DepositosSucursalesBCI
# vive en el mismo sitio (`clcomercialexpress.sharepoint.com/sites/BIBLIOTECA_CEX`)
# que RegistrosCallback. Si el primer run tira 404, confirmar en Graph Explorer:
#   GET https://graph.microsoft.com/v1.0/sites/clcomercialexpress.sharepoint.com:/sites/BIBLIOTECA_CEX
SITE_ID = "clcomercialexpress.sharepoint.com,d331a44f-8002-4cab-8db4-061cb13eb493,1d32dc11-bde8-40af-8d21-d070452a7b2b"

# TODO: reemplazar por el list-id real de "DepositosSucursalesBCI".
# Forma más rápida de conseguirlo: entrar a la lista en SharePoint ->
# Configuración de la lista (la pantalla que ya capturaste) -> mirar la URL
# de esa página: .../listedit.aspx?List=%7B<GUID>%7D  — ese <GUID> (sin las
# llaves ni el %7B/%7D) ES el list-id.
LIST_ID = "a06c0e6a-b2d8-4d7d-a8a2-21d661ccfa03"

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
OUT_PATH = "depositos.json"  # raíz del repo, mismo patrón que ejecutivos.json / callback_registros.json

# ⚠️ MAPEO DE COLUMNAS — las claves de la IZQUIERDA son las que espera
# cuadratura_caja.html. Los valores de la DERECHA son los "internal name"
# asumidos en Graph (idénticos a los nombres visibles en SharePoint, mismo
# patrón que SucursalPDV/OrigenProblema/Calif_Global en sync_callback.py).
# Verificar contra DEPOSITOS_RAW_SAMPLE (ver más abajo) en el primer run real
# y ajustar acá si no calzan.
FIELD_MAP = {
    "Titulo": "Title",
    "NombreSucursal": "NombreSucursal",
    "OrigenForm": "OrigenForm",
    "FechaRecaudacion": "FechaRecaudacion",
    "FechaDeposito": "FechaDeposito",
    "MontoDepositado": "MontoDepositado",
    "NumeroDeposito": "NumeroDeposito",
    "NombreUsuario": "NombreUsuario",
    "CuentaDestino": "CuentaDestino",
    "HojaTransporte": "HojaTransporte",
    "BolsaTransporte": "BolsaTransporte",
    "ComentarioRevision": "ComentarioRevision",
}


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
    url = f"{GRAPH_BASE}/sites/{SITE_ID}/lists/{LIST_ID}/items?expand=fields&$top=200"
    items = []
    while url:
        resp = requests.get(url, headers=headers, timeout=30)
        resp.raise_for_status()
        payload = resp.json()
        items.extend(payload.get("value", []))
        url = payload.get("@odata.nextLink")
    return items


def map_item(raw_item):
    fields = raw_item.get("fields", {})
    return {out_key: fields.get(in_key) for out_key, in_key in FIELD_MAP.items()}


def main():
    token = get_access_token()
    raw_items = fetch_all_items(token)
    print(f"Items obtenidos de DepositosSucursalesBCI: {len(raw_items)}", file=sys.stderr)

    # Muestra cruda para verificar nombres reales de columna (ver FIELD_MAP).
    with open("depositos-raw-sample.json", "w", encoding="utf-8") as fh:
        json.dump([item.get("fields", {}) for item in raw_items[:3]], fh, ensure_ascii=False, indent=2)

    registros = []
    excluidos = 0
    for item in raw_items:
        mapped = map_item(item)
        if not mapped.get("NombreSucursal") or not mapped.get("FechaRecaudacion"):
            excluidos += 1
            continue
        registros.append(mapped)

    if excluidos:
        print(
            f"Registros excluidos (sin NombreSucursal o sin FechaRecaudacion): {excluidos}",
            file=sys.stderr,
        )

    with open(OUT_PATH, "w", encoding="utf-8") as fh:
        json.dump(registros, fh, ensure_ascii=False, indent=2)

    print(f"OK: {OUT_PATH} generado en la raíz del repo ({len(registros)} registros).", file=sys.stderr)


if __name__ == "__main__":
    main()
