# Avance Comercial Diario — Bsale

Este informe sigue la estructura existente de MinutaComercialEjecutivos: HTML y JSON en la raíz, configuración de URLs en el HTML y lectura desde `raw.githubusercontent.com` en la rama `main`.

## Archivos y primera publicación

- `avance_comercial_cex.html`: informe y herramienta de carga manual en `?admin=1`.
- `avance_comercial.json`: último corte validado, utilizado por todos los lectores.
- `avance_metas.json`: metas mensuales por sucursal y pesos diarios del FCST.
- `preparar_metas_avance.py`: conversión local de ambos Excel a JSON para futuros períodos (requiere Python y openpyxl).
- `AVANCE_COMERCIAL.md`: instrucciones y mapeo de datos.

Agregar los archivos de esta integración a la raíz de `main`. No reemplazar `index.html`, los otros informes, sus JSON ni los workflows existentes. Esta integración no necesita Azure, SharePoint, credenciales nuevas ni un workflow de sincronización.

Si Pages ya publica la raíz de `main`, el enlace esperado es:
`https://comercialexpress.github.io/MinutaComercialEjecutivos/avance_comercial_cex.html`

Verificar la configuración actual de Pages en Settings → Pages antes de modificarla. Si hay un dominio personalizado o un proceso de despliegue distinto, usar la misma base de los informes existentes. Los enlaces anteriores son esperados, no una confirmación de despliegue.

## Actualizar cada corte

1. Exportar un único día desde Bsale en formato `.xlsx`, con las columnas detalladas abajo. No editar el Excel original.
2. Abrir el informe con `?admin=1` al final de la dirección.
3. Pulsar **Cargar Excel Bsale**, seleccionar el archivo y revisar la vista previa local. La fecha de ventas selecciona automáticamente el peso del FCST y las metas de ese período.
4. Pulsar **Descargar JSON para publicar**. Las metas se administran en `avance_metas.json`, no en el archivo de ventas.
5. En la raíz del repositorio, usar **Add file → Upload files** y cargar el archivo descargado con el nombre exacto `avance_comercial.json`. Si el navegador agrega `(1)` al nombre, renombrarlo antes. Reemplazar el archivo existente y confirmar con **Commit changes**, según los permisos y reglas del repositorio.
6. Volver al informe y pulsar **Ver datos publicados**. Verificar el corte y la hora de preparación. Compartir el enlace sin `?admin=1` con los lectores.

Cargar un Excel y descargar el JSON no lo publica. Solo el cambio confirmado en GitHub comparte la carga. Los lectores consultan nuevos datos cada 60 segundos mientras la pestaña está visible, al regresar a la pestaña o mediante **Actualizar**. La caché de GitHub puede añadir demora, por lo que un minuto no es una garantía de propagación. La vista previa local suspende las actualizaciones para que no se sobrescriba durante la revisión.

`?admin=1` muestra herramientas, no autentica ni concede permisos de escritura. La publicación requiere iniciar sesión en GitHub con permisos sobre el repositorio. No se guardan tokens en el HTML.

## Datos y reglas

Las sucursales activas se definen en `CONFIG.storeOrder`. Siempre aparecen en la matriz, tarjetas y buscador, incluso sin registros en el corte, con indicadores en cero y detalle vacío. Si el archivo incluye una sucursal adicional, también se conserva para no omitir sus ventas. Actualizar el catálogo cuando una sucursal deje de estar activa o se incorpore una nueva.

Fuente comprobada: `PIVOT_EXPORT.xlsx`, hoja `Flexmonster Pivot Table`, encabezados en la fila 1, 557 registros del 10/09/2026. No requiere una tabla formal de Excel: se busca una única hoja con todos los encabezados requeridos.

| Clave JSON | Columna de Bsale | Uso |
|---|---|---|
| f | Fecha Documento | Fecha del corte |
| v | Vendedor | Detalle por vendedor |
| s | Sucursal | Agrupación por sucursal |
| sg | SEGMENTO | PERSONA / EMPRESA |
| t | Tipo de Producto / Servicio | Clasificación del indicador |
| p | Producto / Servicio + Variante | Voz, portabilidad y desglose |
| q | Cantidad | Unidades y desglose por producto |
| m | Suma de Subtotal Bruto | Venta bruta de equipos y accesorios |

Se suma **Cantidad**, según confirmación del usuario, para Móvil, Fibra, Voz, unidades de Equipos y Accesorios y sus desgloses. Los montos convierten **Suma de Subtotal Bruto ÷ 1,19** a neto sin IVA y no se multiplican de nuevo por Cantidad. Se conservan los decimales durante la agregación y se redondea a pesos enteros solo al mostrar. El JSON y el Excel mantienen el monto bruto original; la conversión se aplica una sola vez en el HTML, tanto en la vista publicada como en la carga local. Las columnas MONTO EQUIPO SIN RECARGO y MONTO ACCESORIO SIN RECARGO no se usan: se usa el subtotal bruto de origen para obtener el neto sin IVA solicitado.

Reglas comerciales conservadas del modelo original:

- MOVIL + PERSONA → Móvil Persona; MOVIL + EMPRESA → Móvil Empresa.
- HOGAR FIBRA + PERSONA → Fibra Solicitud; HOGAR FIBRA + EMPRESA → Solicitud Fijo Empresa.
- Productos MOVIL/PERSONA que empiezan por VOZ → Voz SS; los que además contienen PORTADO → Voz Portado.
- EQUIPO y ACCESORIO → montos netos sin IVA y unidades, para ambos segmentos.
- %Mix Porta = Voz Portado / Voz SS × 100. %Mix Voz SS = Voz SS / Móvil Persona × 100. Se calculan sobre las cantidades agregadas, no promediando porcentajes de vendedores o sucursales. Con denominador cero se conserva el criterio original de mostrar 0%.
- Las otras categorías permanecen en el conteo de registros, pero no se asignan a indicadores nuevos. Registros no equivale a documentos únicos: un documento puede tener varias líneas.

Se rechazan archivos vacíos, columnas faltantes/duplicadas, múltiples hojas compatibles, fechas inválidas o mezcladas y cantidades/montos no numéricos. Cantidades fraccionarias o negativas y montos negativos requieren acordar el tratamiento de devoluciones antes de incorporarlos. No se descartan silenciosamente filas. Límite de carga: 25 MB.

El JSON solo conserva las ocho columnas indicadas. No incluye Codigo cliente, Numero del documento, Tipo de Documento, PLATAFORMA ni el Excel original. Como el repositorio es público, los datos incluidos —también nombres de vendedores y cifras comerciales— son públicos.

Contrato: `{schemaVersion:1, generatedAt, asOf, goal, rows}`. `generatedAt` identifica la preparación del archivo, no la hora de venta ni el momento exacto del commit. La fecha visible del corte se verifica contra las filas. Al fallar la descarga, se conserva el último corte recibido con aviso. Si no hay datos anteriores, se muestra un estado de error sin cifras de ejemplo.

## Conciliación del corte inicial

| Concepto | Resultado |
|---|---:|
| Registros | 557 |
| Sucursales | 9 |
| Vendedores | 73 |
| Cantidad de todas las categorías | 600 |
| Móvil Persona | 142 |
| Móvil Empresa | 23 |
| Voz SS Persona | 115 |
| Voz Portado Persona | 33 |
| Fibra total (persona + empresa) | 14 |
| Unidades Equipos | 158 |
| Unidades Accesorios | 158 |
| Venta neta Equipos (bruto $83.322.453 ÷ 1,19) | $70.018.868 |
| Venta neta Accesorios (bruto $11.389.462 ÷ 1,19) | $9.570.976 |

El lector Excel usa SheetJS, como el HTML de origen. Requiere conexión para cargar esa biblioteca desde CDN. El informe conserva la identidad visual CEX y las fuentes Sora e Inter.

## Metas diarias y FCST

Para septiembre 2026 se usa el **primer bloque**, encabezados en fila 6, sucursales en filas 7–15 y total CEX en fila 16 de `METAS_POR_SUCURSAL_2026_09.xlsx`, hoja Export. El segundo bloque queda excluido por confirmación del usuario.

Se toma FECHA y PESO DIARIO del archivo `FCST_EXPRESS_2026_09.xlsx`, hoja Export, encabezados fila 2. Sus fechas son septiembre aunque el título dice agosto. Los 30 pesos suman 100%. No se utiliza el peso acumulado, `%AVANCE`, ni `% RESTANTE`.

**Meta diaria por tienda e indicador = meta mensual × peso de la fecha de ventas.** Cumplimiento = venta del día / meta diaria × 100. No es un objetivo ajustado a la hora del día: una carga parcial se compara con la meta completa del día.

| Indicador HTML | Meta mensual de origen |
|---|---|
| Móvil Persona | MOVIL_PERSONA |
| Fibra Solicitud | FIBRA_SOLICITUD |
| Voz Portado | VOZ_PORTADO |
| Voz SS | VOZ_SS |
| Equipos netos | (MONTO_EQUIPOS + MONTO_EQUIPOS_EMPRESA) × 1.000 |
| Accesorios netos | (MONTO_ACCESORIOS + MONTO_ACCESORIOS_EMPRESA) × 1.000 |
| Móvil Empresa | MOVIL_EMPRESA |
| Solicitud Fijo Empresa | FIJO_EMPRESA |
| %Mix Porta | PCT_MIX_PORTA × 100, objetivo constante del período |
| %Mix Voz SS | VOZ_SS / MOVIL_PERSONA × 100, objetivo derivado constante del período |

El usuario confirmó que las metas monetarias incluyen Persona + Empresa, el factor es exactamente 1.000 y los valores ya son netos: **no dividir las metas por 1,19**. Las ventas continúan convirtiéndose desde bruto a neto una sola vez. Cantidades objetivo con hasta dos decimales y pesos con hasta cuatro decimales porcentuales en pantalla; el cálculo conserva toda la precisión disponible. No redondear cantidades al entero antes de evaluar cumplimiento.

La meta global suma las metas de sucursales incluidas. Para %Mix Porta global se ponderan los objetivos porcentuales de las tiendas por sus metas mensuales VOZ_SS; para %Mix Voz SS se divide la suma de metas VOZ_SS por la suma de metas MOVIL_PERSONA. No se asignan metas de sucursal a vendedores: su detalle sigue siendo referencial.

La fila CEX del Excel difiere ligeramente de la suma de sucursales: Móvil Persona 4.868 frente a 4.867, Voz Portado 1.454 frente a 1.453, Voz SS 4.381 frente a 4.382 y Fibra Solicitud 445 frente a 444. Se conserva cada meta de sucursal y el total visible suma esas metas. Las diferencias se registran en `sourceTotalDifference` del JSON.

Verde indica cumplimiento desde 100%, rojo avance inferior y gris meta ausente o cero. Si la meta diaria es cero (incluidos los días 18 y 19 con peso cero), se muestra “Sin meta para hoy” y no se divide por cero. Si no corresponde el período o falta una sucursal, se muestra “Meta no disponible”. Si falla la descarga de metas, se mantienen las ventas y se suspende la evaluación hasta recuperar las metas. Las sucursales activas sin ventas permanecen visibles con venta cero y su meta correspondiente.

Las metas se consultan al abrir el informe, al actualizar, al volver a la pestaña y cada minuto. La carga local del Excel de ventas utiliza las mismas metas publicadas; su fecha de corte puede diferir de la fecha de publicación.

### Actualizar el período de metas

Ejecutar localmente con Python y openpyxl instalado:

```text
python preparar_metas_avance.py --metas "METAS_POR_SUCURSAL_2026_09.xlsx" --fcst "FCST_EXPRESS_2026_09.xlsx" --fila-encabezado 6 --salida avance_metas.json
```

Revisar la salida y reemplazar solo `avance_metas.json` en la raíz del repositorio. El script valida fechas, pesos, duplicados y valores numéricos. El encabezado y la hoja deben corresponder al formato documentado; si cambia la estructura, revisarla antes de importar. El JSON contiene un período: al publicar el siguiente mes, las ventas de otro mes siguen visibles pero sus metas se indican como no disponibles. No se necesitan cambios en los workflows existentes ni subir los Excel originales.
