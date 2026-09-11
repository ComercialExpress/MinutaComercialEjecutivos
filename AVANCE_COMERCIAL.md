# Avance Comercial Diario — Bsale

Este informe sigue la estructura existente de MinutaComercialEjecutivos: HTML y JSON en la raíz, configuración de URLs en el HTML y lectura desde `raw.githubusercontent.com` en la rama `main`.

## Archivos y primera publicación

- `avance_comercial_cex.html`: informe y herramienta de carga manual en `?admin=1`.
- `avance_comercial.json`: último corte validado, utilizado por todos los lectores.
- `AVANCE_COMERCIAL.md`: instrucciones y mapeo de datos.

Agregar estos tres archivos a la raíz de `main`. No reemplazar `index.html`, los otros informes, sus JSON ni los workflows existentes. Esta integración no necesita Azure, SharePoint, credenciales nuevas ni un workflow de sincronización.

Si Pages ya publica la raíz de `main`, el enlace esperado es:
`https://comercialexpress.github.io/MinutaComercialEjecutivos/avance_comercial_cex.html`

Verificar la configuración actual de Pages en Settings → Pages antes de modificarla. Si hay un dominio personalizado o un proceso de despliegue distinto, usar la misma base de los informes existentes. Los enlaces anteriores son esperados, no una confirmación de despliegue.

## Actualizar cada corte

1. Exportar un único día desde Bsale en formato `.xlsx`, con las columnas detalladas abajo. No editar el Excel original.
2. Abrir el informe con `?admin=1` al final de la dirección.
3. Pulsar **Cargar Excel Bsale**, seleccionar el archivo y revisar la vista previa local.
4. Ajustar la meta %Mix Porta si corresponde y pulsar **Descargar JSON para publicar**.
5. En la raíz del repositorio, usar **Add file → Upload files** y cargar el archivo descargado con el nombre exacto `avance_comercial.json`. Si el navegador agrega `(1)` al nombre, renombrarlo antes. Reemplazar el archivo existente y confirmar con **Commit changes**, según los permisos y reglas del repositorio.
6. Volver al informe y pulsar **Ver datos publicados**. Verificar el corte y la hora de preparación. Compartir el enlace sin `?admin=1` con los lectores.

Cargar un Excel y descargar el JSON no lo publica. Solo el cambio confirmado en GitHub comparte la carga. Los lectores consultan nuevos datos cada 60 segundos mientras la pestaña está visible, al regresar a la pestaña o mediante **Actualizar**. La caché de GitHub puede añadir demora, por lo que un minuto no es una garantía de propagación. La vista previa local suspende las actualizaciones para que no se sobrescriba durante la revisión.

`?admin=1` muestra herramientas, no autentica ni concede permisos de escritura. La publicación requiere iniciar sesión en GitHub con permisos sobre el repositorio. No se guardan tokens en el HTML.

## Datos y reglas

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

Se suma **Cantidad**, según confirmación del usuario, para Móvil, Fibra, Voz, unidades de Equipos y Accesorios y sus desgloses. Los montos suman **Suma de Subtotal Bruto** y no se multiplican de nuevo por Cantidad. Las columnas MONTO EQUIPO SIN RECARGO y MONTO ACCESORIO SIN RECARGO no se usan: se conserva el criterio de venta bruta del modelo original.

Reglas comerciales conservadas del modelo original:

- MOVIL + PERSONA → Móvil Persona; MOVIL + EMPRESA → Móvil Empresa.
- HOGAR FIBRA + PERSONA → Fibra Solicitud; HOGAR FIBRA + EMPRESA → Solicitud Fijo Empresa.
- Productos MOVIL/PERSONA que empiezan por VOZ → Voz SS; los que además contienen PORTADO → Voz Portado.
- EQUIPO y ACCESORIO → montos brutos y unidades, para ambos segmentos.
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
| Venta bruta Equipos | $83.322.453 |
| Venta bruta Accesorios | $11.389.462 |

El lector Excel usa SheetJS, como el HTML de origen. Requiere conexión para cargar esa biblioteca desde CDN. El informe conserva la identidad visual CEX y las fuentes Sora e Inter.
