"""Lee Bsale, metas y FCST de SharePoint; publica solo JSON validado.

Credenciales: TENANT_ID, CLIENT_ID, CLIENT_SECRET (mismos secrets de Azure).
Las rutas se refieren a la raíz de la biblioteca, no al disco C: del usuario.
"""
import datetime as dt
import io
import json
import math
import os
from pathlib import Path
import sys
from urllib.parse import quote

import openpyxl
import requests
from preparar_metas_avance import convert

SITE_ID = 'clcomercialexpress.sharepoint.com,d331a44f-8002-4cab-8db4-061cb13eb493,1d32dc11-bde8-40af-8d21-d070452a7b2b'
PIVOT_PATH = 'INFORMES COMERCIALES/B) INFORME DE VENTAS AVANCE/PIVOT_EXPORT.xlsx'
METAS_PATH = '0. METAS/{year}/METAS_POR_SUCURSAL_{year}_{month}.xlsx'
FCST_PATH = 'INFORMES COMERCIALES/D) FCST_EXPRESS/FCST_PUBLICADO/FCST_EXPRESS_{year}_{month}.xlsx'
COLUMNS = {'f':'Fecha Documento','v':'Vendedor','s':'Sucursal','sg':'SEGMENTO',
           't':'Tipo de Producto / Servicio','p':'Producto / Servicio + Variante',
           'm':'Suma de Subtotal Bruto','q':'Cantidad'}
HEADER_ALIASES = {key: (name,) for key, name in COLUMNS.items()}
HEADER_ALIASES['m'] = ('Suma de Subtotal Bruto', 'Subtotal Bruto')


def parse_pivot(source):
    workbook=openpyxl.load_workbook(source,read_only=True,data_only=True)
    candidates=[]
    try:
        for sheet in workbook:
            rows=iter(sheet.values)
            for raw in rows:
                headers=[str(c or '').strip() for c in raw]
                matches = {key: [i for i, h in enumerate(headers) if h in aliases]
                           for key, aliases in HEADER_ALIASES.items()}
                if all(matches.values()):
                    if any(len(indices)!=1 for indices in matches.values()):
                        raise ValueError('El export contiene columnas duplicadas o alias ambiguos.')
                    candidates.append(({key: indices[0] for key, indices in matches.items()},list(rows)))
                    break
        if len(candidates)!=1: raise ValueError('Debe existir una sola hoja con las columnas verificadas de Bsale.')
        indices,records=candidates[0]
        out=[]; dates=set()
        for index,raw in enumerate(records,1):
            if all(v is None or v=='' for v in raw): continue
            row={key:raw[index] for key,index in indices.items()}
            value=row['f']
            if isinstance(value,(dt.datetime,dt.date)): date=value
            elif isinstance(value,str):
                try: date=dt.datetime.strptime(value.strip(),'%d/%m/%Y')
                except ValueError: raise ValueError(f'Registro {index}: fecha inválida.') from None
            else: raise ValueError(f'Registro {index}: fecha no reconocida.')
            row['f']=date.strftime('%d/%m/%Y');dates.add(row['f'])
            for key in ['v','s','sg','t','p']:
                if not isinstance(row[key],str) or row[key].strip() in ('','(blanco)'):
                    raise ValueError(f'Registro {index}: falta {COLUMNS[key]}.')
                row[key]=row[key].strip()
            for key in ['m','q']:
                value=row[key]
                if isinstance(value,bool) or not isinstance(value,(int,float)) or not math.isfinite(value) or value<0:
                    raise ValueError(f'Registro {index}: {COLUMNS[key]} no es un valor válido.')
            if row['q']!=int(row['q']): raise ValueError(f'Registro {index}: Cantidad debe ser entera.')
            # Montos brutos originales. El HTML los divide una sola vez por 1,19.
            out.append(row)
        if not out: raise ValueError('El export de Bsale está vacío.')
        if len(dates)!=1: raise ValueError('El export mezcla días. Exportar un único día de ventas.')
        return {'schemaVersion':1,'generatedAt':dt.datetime.now(dt.timezone.utc).isoformat(),
                'asOf':next(iter(dates)),'goal':33.5,'rows':out}
    finally: workbook.close()


def get_token():
    names=['TENANT_ID','CLIENT_ID','CLIENT_SECRET']
    if any(not os.environ.get(k) for k in names):
        raise ValueError('Faltan secrets de Azure: AZURE_TENANT_ID, AZURE_CLIENT_ID o AZURE_CLIENT_SECRET.')
    response=requests.post(f'https://login.microsoftonline.com/{quote(os.environ["TENANT_ID"],safe="")}/oauth2/v2.0/token',
        data={'client_id':os.environ['CLIENT_ID'],'client_secret':os.environ['CLIENT_SECRET'],
              'scope':'https://graph.microsoft.com/.default','grant_type':'client_credentials'},timeout=40)
    if response.status_code!=200: raise ValueError(f'No se pudo autenticar con Azure (HTTP {response.status_code}).')
    return response.json()['access_token']


def download(token, relative_path):
    site=os.environ.get('AVANCE_SITE_ID') or SITE_ID
    url=f'https://graph.microsoft.com/v1.0/sites/{site}/drive/root:/{quote(relative_path,safe="/")}:/content'
    # requests quita Authorization en redirecciones a otro host. No registrar
    # la URL temporal de descarga ni respuestas con datos del archivo original.
    try:
        response=requests.get(url,headers={'Authorization':'Bearer '+token},timeout=(20,90),stream=True)
        with response:
            if response.status_code!=200: raise ValueError(f'No se pudo leer {relative_path} (HTTP {response.status_code}).')
            data=bytearray()
            for chunk in response.iter_content(65536):
                data.extend(chunk)
                if len(data)>25*1024*1024: raise ValueError('El archivo de origen supera 25 MB.')
        return io.BytesIO(data)
    except requests.RequestException:
        raise ValueError('Error de red leyendo SharePoint; no se publican cambios.') from None


def stable_data(data):
    return {k:v for k,v in data.items() if k!='generatedAt'}


def save_if_changed(path, data):
    path=Path(path)
    if path.exists():
        previous=json.loads(path.read_text(encoding='utf-8'))
        if stable_data(previous)==stable_data(data): return False
    temporary=path.with_suffix('.json.tmp')
    temporary.write_text(json.dumps(data,ensure_ascii=False,separators=(',',':'))+'\n',encoding='utf-8')
    temporary.replace(path)
    return True


def main():
    token=get_token()
    sales=parse_pivot(download(token,os.environ.get('AVANCE_PIVOT_PATH') or PIVOT_PATH))
    day=dt.datetime.strptime(sales['asOf'],'%d/%m/%Y')
    period={'year':f'{day.year:04d}','month':f'{day.month:02d}'}
    metas=(os.environ.get('AVANCE_METAS_PATH') or METAS_PATH).format(**period)
    fcst=(os.environ.get('AVANCE_FCST_PATH') or FCST_PATH).format(**period)
    goals=convert(download(token,metas),download(token,fcst),int(os.environ.get('AVANCE_METAS_HEADER_ROW') or '6'))
    if goals['period']!=day.strftime('%Y-%m'): raise ValueError('Metas y ventas pertenecen a períodos diferentes.')
    # Ningún archivo se guarda hasta validar las tres fuentes.
    changed_sales=save_if_changed('avance_comercial.json',sales)
    changed_goals=save_if_changed('avance_metas.json',goals)
    print(f'Corte {sales["asOf"]}: {len(sales["rows"])} registros. Metas {goals["period"]}: {len(goals["stores"])} sucursales.')
    print(f'Cambios en ventas: {changed_sales}. Cambios en metas: {changed_goals}.')


if __name__=='__main__':
    try: main()
    except Exception as exc:
        # No imprimir tracebacks de librerías HTTP que pudieran contener URLs firmadas.
        message=str(exc) if isinstance(exc,ValueError) else f'Fallo de lectura/validación ({type(exc).__name__}).'
        print('ERROR: '+message,file=sys.stderr)
        sys.exit(1)
