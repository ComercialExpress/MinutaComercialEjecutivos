"""Convierte metas mensuales y FCST a JSON. Requiere openpyxl.

python preparar_metas_avance.py --metas METAS.xlsx --fcst FCST.xlsx --salida avance_metas.json
El encabezado predeterminado es la fila 6 (primer bloque confirmado).
"""
import argparse
import calendar
import datetime as dt
import json
import math
from pathlib import Path
import unicodedata
import openpyxl

FIELDS = {'movilPersona':'MOVIL_PERSONA','fibra':'FIBRA_SOLICITUD',
          'vozPortado':'VOZ_PORTADO','vozSS':'VOZ_SS',
          'movilEmpresa':'MOVIL_EMPRESA','fijoEmpresa':'FIJO_EMPRESA'}

def normalized(value):
    return ''.join(c for c in unicodedata.normalize('NFD',str(value).strip().upper()) if unicodedata.category(c)!='Mn')

def number(value, field):
    if isinstance(value,bool) or not isinstance(value,(int,float)) or not math.isfinite(value) or value<0:
        raise ValueError(f'{field}: se requiere un número no negativo, no una celda vacía o fórmula sin valor guardado.')
    return value

def convert(metas, fcst, header_row=6):
    workbook=openpyxl.load_workbook(metas,read_only=True,data_only=True)
    values=list(workbook['Export'].values)
    headers=values[header_row-1]
    if len([h for h in headers if h=='SUCURSAL'])!=1:
        raise ValueError('Encabezado de metas incorrecto.')
    stores={}; period=None; source_total=None
    for raw in values[header_row:]:
        row=dict(zip(headers,raw))
        if not row.get('SUCURSAL'): break
        if row['SUCURSAL']=='SUCURSAL': raise ValueError('Bloques de metas sin separación.')
        ym=str(row['YM'])
        if not ym.isdigit() or len(ym)!=6: raise ValueError('YM inválido.')
        row_period=ym[:4]+'-'+ym[4:]
        if period is not None and row_period!=period: raise ValueError('El bloque mezcla períodos.')
        period=row_period
        out={key:number(row.get(field),field) for key,field in FIELDS.items()}
        for key,field in [('equiposCLP','MONTO_EQUIPOS'),('accCLP','MONTO_ACCESORIOS')]:
            # Confirmado: persona + empresa, factor 1000; las metas ya son netas.
            out[key]=(number(row.get(field),field)+number(row.get(field+'_EMPRESA'),field+'_EMPRESA'))*1000
        out['mixPorta']=number(row.get('PCT_MIX_PORTA'),'PCT_MIX_PORTA')*100
        if out['mixPorta']>100: raise ValueError('PCT_MIX_PORTA debe ser fracción 0..1.')
        name=normalized(row['SUCURSAL'])
        if name=='CEX':
            source_total=out
            break
        if name in stores: raise ValueError(f'Sucursal duplicada: {name}')
        stores[name]=out
    if not stores or source_total is None: raise ValueError('Faltan sucursales o la fila CEX del bloque.')
    differences={field:sum(s[field] for s in stores.values())-source_total[field]
                 for field in [*FIELDS,'equiposCLP','accCLP']}
    workbook.close()
    forecast=openpyxl.load_workbook(fcst,read_only=True,data_only=True)
    values=list(forecast['Export'].values)
    headers=[str(h or '').replace('\n',' ').strip() for h in values[1]]
    date_index=headers.index('FECHA'); weight_index=headers.index('PESO DIARIO')
    weights={}
    for raw in values[2:]:
        date=raw[date_index]
        if date is None: continue
        if not isinstance(date,(dt.datetime,dt.date)): raise ValueError('FECHA del FCST debe ser una fecha Excel.')
        day=date.strftime('%Y-%m-%d')
        if day[:7]!=period: raise ValueError('El FCST y las metas no corresponden al mismo mes.')
        if day in weights: raise ValueError('Fecha duplicada en FCST.')
        weights[day]=number(raw[weight_index],'PESO DIARIO')
    forecast.close()
    y,m=map(int,period.split('-'))
    if len(weights)!=calendar.monthrange(y,m)[1] or not math.isclose(sum(weights.values()),1,abs_tol=1e-8):
        raise ValueError('FCST incompleto o pesos que no suman 100%.')
    return {'schemaVersion':1,'period':period,'sourceHeaderRow':header_row,
            'moneyScale':1000,'moneyBasis':'net','moneySegments':'PERSONA+EMPRESA',
            'generatedAt':dt.datetime.now(dt.timezone.utc).isoformat(), 'stores':stores,'weights':weights,
            'sourceTotalDifference':differences}

if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--metas',required=True); parser.add_argument('--fcst',required=True)
    parser.add_argument('--fila-encabezado',type=int,default=6)
    parser.add_argument('--salida',default='avance_metas.json')
    args=parser.parse_args()
    data=convert(args.metas,args.fcst,args.fila_encabezado)
    target=Path(args.salida)
    target.write_text(json.dumps(data,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(f'{data["period"]}: {len(data["stores"])} sucursales, {len(data["weights"])} días; pesos validados. El total global suma las sucursales.')
