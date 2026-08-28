#!/usr/bin/env python3
"""
Sync recaudación XStore -> recaudacion.json para Cuadratura de Caja.

REEMPLAZA la versión anterior (Lista SharePoint `RecaudacionXStore`).
Motivo del cambio (confirmado en pruebas, 27-08-2026): la Lista tocó el
"list view threshold" de SharePoint (5.000 ítems por vista) — solo el
efectivo ya promedia ~5.000 registros/mes, y el total con todos los medios
de pago (efectivo + tarjetas + TBD) ronda 12.000-15.000/mes. Una Lista de
SharePoint no es un buen lugar para ese volumen a grano transaccional.

NUEVO ORIGEN: los libros Excel mensuales que Guillermo ya mantiene hoy
(descarga diaria de XStore/Oracle, pegada a mano bajo los mismos
encabezados) — ya viven en OneDrive/SharePoint. En vez de migrar ese hábito
a una Lista, este script lee esos libros DIRECTO por la Graph Workbook API
(`/workbook/tables/{tabla}/rows`), exactamente el camino que
`metodologia_pipelines_html.md` §2.3 ya recomendaba para tablas de Excel.

Ventajas de este cambio, además de resolver el problema de volumen:
  - Cero cambio en el hábito diario de Guillermo — sigue pegando en el
    mismo libro de siempre. No hace falta ningún script de carga
    (`load_xstore_excel.py` / `carga_diaria.py` quedan obsoletos para este
    flujo, ver infraestructura_automatizacion.md sesión 7).
  - Solo requiere permiso de LECTURA (`Files.Read.All` / `Sites.Read.All`),
    que la app `MinutaComercialEjecutivos` ya tiene — no hace falta ampliar
    permisos ni pedirle nada a IT.
  - Sin límite de ítems por vista: un libro Excel aguanta muchísimo más
    volumen que una Lista sin degradarse.

REQUISITO en el libro Excel (una sola vez por archivo mensual, ver
README_XSTORE.md): el rango de datos tiene que estar formateado como Tabla
de Excel (Ctrl+T), no un rango suelto — así la Graph Workbook API puede
direccionarlo por nombre (`TABLE_NAME` abajo) y crece solo cuando Guillermo
pega filas nuevas justo debajo de la última fila de la tabla.

Genera:
  recaudacion.json -> [{ "codigo": <AGENCIA>, "dias": [{"fecha","recaudado"}, ...] }, ...]
  Misma forma que la versión anterior (Lista) y que stores_data.js — el
  HTML no necesita ningún cambio por este pivote.

Auth: app-only (client credentials) contra Microsoft Graph, misma app
      'MinutaComercialEjecutivos' ya usada en el resto del ecosistema.

Variables de entorno requeridas:
  TENANT_ID, CLIENT_ID, CLIENT_SECRET

Variable opcional:
  XSTORE_LOOKBACK_MONTHS (default: 2) -> cuántos libros mensuales leer
  (mes actual + N-1 anteriores), para no depender de un único archivo y
  cubrir correcciones tardías de fin de mes.
"""

import os
import sys
import json
import calendar
import requests
from datetime import date

TENANT_ID = os.environ["TENANT_ID"]
CLIENT_ID = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
OUT_PATH = "recaudacion.json"

# Confirmado 27-08-2026: el libro vive en la biblioteca de documentos
# ("Documentos compartidos") del mismo sitio BIBLIOTECA_CEX que ya usan
# DepositosSucursalesBCI/RegistrosCallback — mismo SITE_ID de siempre.
# Ruta compartida por Guillermo:
#   https://clcomercialexpress.sharepoint.com/sites/BIBLIOTECA_CEX/Documentos%20compartidos/BASES%20ENTEL/REPORTE_XSTORE
# "Documentos compartidos" ES la biblioteca por defecto (ya representada por
# la raíz del drive del sitio en Graph) — no es una carpeta más dentro de
# CARPETA, por eso no aparece en el valor de abajo.
SITE_ID = "clcomercialexpress.sharepoint.com,d331a44f-8002-4cab-8db4-061cb13eb493,1d32dc11-bde8-40af-8d21-d070452a7b2b"
CARPETA = "BASES ENTEL/REPORTE_XSTORE"

# ⚠️ TODO — confirmar el patrón real de nombre de archivo. Se asume igual
# al ya usado en el proyecto (REPORTE_XSTORE_2026_08.xlsx: prefijo fijo +
# año + mes con cero a la izquierda). Ajustar aquí si el nombre real difiere.
def nombre_archivo_mes(year, month):
    return f"REPORTE_XSTORE_{year:04d}_{month:02d}.xlsx"

# Confirmado 27-08-2026: nombre real de la Tabla de Excel en el libro
# mensual. Usar el mismo nombre exacto en cada libro nuevo (Ctrl+T ->
# Diseño de tabla -> Nombre de la tabla) para no tener que tocar esta
# constante cada mes.
TABLE_NAME = "TB_XSTORE"

# Mismo mapeo que la versión anterior — nombres de columna EXACTOS de la
# fila de encabezado en el Excel (no "internal name" de Graph: acá no
# aplica ese concepto, es texto de celda tal cual).
COL_AGENCIA = "AGENCIA"
COL_FECHA = "FECHA"
COL_HORA = "HORA"
COL_MONTO = "MONTO"
COL_MEDIO_PAGO = "MEDIO DE PAGO"

MEDIO_PAGO_EFECTIVO = "EFE"


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


def meses_a_leer(lookback_months):
    """[(year, month), ...] desde hace (lookback_months - 1) meses hasta el
    mes actual, inclusive. Con el default de 2: mes actual + el anterior."""
    hoy = date.today()
    meses = []
    for i in range(lookback_months - 1, -1, -1):
        m0 = hoy.month - 1 - i
        year = hoy.year + (m0 // 12)
        month = (m0 % 12) + 1
        meses.append((year, month))
    return meses


def workbook_path(nombre_archivo):
    """Arma el path de Graph al workbook, dentro de la biblioteca de
    documentos del sitio BIBLIOTECA_CEX (CARPETA = "BASES ENTEL/REPORTE_XSTORE").
    Se URL-encodea cada segmento (CARPETA trae un espacio: "BASES ENTEL")
    para que el path-based addressing de Graph no lo interprete mal."""
    from urllib.parse import quote
    ruta = quote(f"{CARPETA}/{nombre_archivo}", safe="/")
    return f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/{ruta}:/workbook"


def fetch_table_rows(token, nombre_archivo):
    """Trae todas las filas de la Tabla TABLE_NAME del libro `nombre_archivo`.

    BUG CORREGIDO (28-08-2026, sesión 11): la versión anterior paginaba
    confiando en que Graph devolviera "@odata.nextLink" en la respuesta del
    endpoint /workbook/tables/{tabla}/rows — ese endpoint NO lo entrega de
    forma confiable, a diferencia de otras colecciones de Graph. Con
    $top=200 y sin nextLink, el script se quedaba solo con la PRIMERA
    página (200 filas) y nunca pedía el resto. Confirmado con datos reales:
    las filas quedan en el archivo en el mismo orden en que Guillermo las
    pega cada día (el primer día pegado queda primero), y un solo día
    (01-08-2026) ya tenía 272 filas — más que el $top=200 — por eso
    recaudacion.json terminaba con un único registro por mes (el del
    primer día), en vez de uno por cada día real.

    Fix: paginar con $skip explícito, parando cuando una página vuelve con
    MENOS filas de las pedidas (señal estándar de "última página") — no
    depende de que el servidor arme nextLink."""
    headers = {"Authorization": f"Bearer {token}"}
    base_url = f"{workbook_path(nombre_archivo)}/tables/{TABLE_NAME}/rows"
    top = 200
    skip = 0
    filas = []
    while True:
        url = f"{base_url}?$top={top}&$skip={skip}"
        resp = requests.get(url, headers=headers, timeout=60)
        if resp.status_code == 404:
            print(f"'{nombre_archivo}' o la tabla '{TABLE_NAME}' no existe todavía — se omite.", file=sys.stderr)
            return []
        if not resp.ok:
            print(f"Graph respondió {resp.status_code} leyendo {nombre_archivo}: {resp.text[:500]}", file=sys.stderr)
            resp.raise_for_status()
        payload = resp.json()
        pagina = payload.get("value", [])
        filas.extend(pagina)
        if len(pagina) < top:
            break  # última página: vino incompleta, no hace falta pedir más
        skip += top
    return filas


def fetch_headers(token, nombre_archivo):
    """Encabezados reales de la tabla (fila 1), para mapear cada fila
    (que Graph devuelve como lista de valores posicionales, no como dict)
    a las columnas por nombre en vez de por índice fijo."""
    headers = {"Authorization": f"Bearer {token}"}
    url = f"{workbook_path(nombre_archivo)}/tables/{TABLE_NAME}/headerRowRange"
    resp = requests.get(url, headers=headers, timeout=30)
    resp.raise_for_status()
    return resp.json()["values"][0]


def solo_fecha(valor_fecha):
    """La columna FECHA puede volver como serial de Excel, como texto ISO,
    o como texto dd-mm-aaaa según cómo Graph serialice ese tipo de celda —
    no confirmado todavía contra una respuesta real. Este helper normaliza
    los casos más comunes; ajustar si el formato real es otro (revisar
    xstore-raw-sample.json, ver más abajo)."""
    if valor_fecha is None:
        return None
    s = str(valor_fecha)
    if len(s) >= 10 and s[4] == "-" and s[7] == "-":
        return s[:10]  # ya viene ISO (YYYY-MM-DD...)
    try:
        # Serial de fecha de Excel (días desde 1899-12-30)
        from datetime import datetime, timedelta
        base = datetime(1899, 12, 30)
        return (base + timedelta(days=float(s))).strftime("%Y-%m-%d")
    except (ValueError, TypeError):
        print(f"⚠️  No se pudo interpretar FECHA={s!r} — revisar formato real.", file=sys.stderr)
        return None


def aggregate_mes(token, nombre_archivo, muestra_cruda):
    encabezados = fetch_headers(token, nombre_archivo)
    try:
        idx_agencia = encabezados.index(COL_AGENCIA)
        idx_fecha = encabezados.index(COL_FECHA)
        idx_monto = encabezados.index(COL_MONTO)
        idx_medio = encabezados.index(COL_MEDIO_PAGO)
    except ValueError:
        print(f"⚠️  Encabezados de '{nombre_archivo}' no coinciden con lo esperado: {encabezados}", file=sys.stderr)
        return {}

    filas = fetch_table_rows(token, nombre_archivo)
    print(f"{nombre_archivo}: {len(filas)} filas leídas", file=sys.stderr)
    muestra_cruda.extend([f["values"][0] for f in filas[:5]])

    acumulado = {}
    agencias_vistas = set()
    fechas_vistas = set()  # todas las fechas que cubre el archivo, sin importar agencia/medio de pago

    for fila in filas:
        valores = fila["values"][0]
        agencia_raw = valores[idx_agencia]
        fecha = solo_fecha(valores[idx_fecha])
        if agencia_raw is None or fecha is None:
            continue
        agencia = int(agencia_raw)
        agencias_vistas.add(agencia)
        fechas_vistas.add(fecha)
        acumulado.setdefault(agencia, {})
        if (valores[idx_medio] or "").strip().upper() != MEDIO_PAGO_EFECTIVO:
            continue  # cuenta para "el archivo cubre este día", pero no suma al monto EFE
        monto = float(valores[idx_monto] or 0)
        acumulado[agencia].setdefault(fecha, 0)
        acumulado[agencia][fecha] += monto

    # Recaudación real $0 en efectivo (28-08-2026, sesión 12): si el archivo
    # cubre una fecha (hubo transacciones de CUALQUIER medio de pago, de
    # CUALQUIER sucursal) pero esta agencia no tuvo ninguna transacción en
    # EFE ese día, se completa igual con $0 explícito — en vez de dejar la
    # fecha ausente. cuadratura_caja.html distingue "recaudado === 0" (dato
    # real, no aplica depósito, se muestra "Cuadra") de "sin dato" (todavía
    # no llegó información, se muestra "Sin datos") — sin este completado,
    # cualquier día sin efectivo quedaba indistinguible de un día que el
    # pipeline simplemente no había cargado todavía.
    for agencia in agencias_vistas:
        for fecha in fechas_vistas:
            acumulado[agencia].setdefault(fecha, 0)

    return acumulado


def main():
    lookback_months = int(os.environ.get("XSTORE_LOOKBACK_MONTHS", "2"))
    token = get_access_token()

    total = {}
    muestra_cruda = []
    for year, month in meses_a_leer(lookback_months):
        nombre_archivo = nombre_archivo_mes(year, month)
        parcial = aggregate_mes(token, nombre_archivo, muestra_cruda)
        for agencia, dias in parcial.items():
            total.setdefault(agencia, {})
            for fecha, monto in dias.items():
                total[agencia][fecha] = total[agencia].get(fecha, 0) + monto

    with open("xstore-raw-sample.json", "w", encoding="utf-8") as fh:
        json.dump(muestra_cruda[:5], fh, ensure_ascii=False, indent=2)

    salida = [
        {
            "codigo": agencia,
            "dias": [{"fecha": fecha, "recaudado": round(monto)} for fecha, monto in sorted(dias.items())],
        }
        for agencia, dias in sorted(total.items())
    ]

    with open(OUT_PATH, "w", encoding="utf-8") as fh:
        json.dump(salida, fh, ensure_ascii=False, indent=2)

    print(f"OK: {OUT_PATH} generado ({len(salida)} sucursales).", file=sys.stderr)


if __name__ == "__main__":
    main()
