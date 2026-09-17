# Detalle de encuestas EPA

Abrir `PanelCallbackEpa.html` y seleccionar **Detalle de encuestas EPA**.
Mes y sucursal controlan todas las vistas. La búsqueda consulta los 69 campos.
«Ver detalle» muestra la ficha, la grabación y el archivo/hoja/fila de origen.

Las encuestas están embebidas en el HTML para uso local. Son una copia de las
bases al importar, no una conexión en vivo a Excel.

En GitHub, el flujo **Sync Callback y encuestas EPA** actualiza esa copia cada
30 minutos, junto con los callbacks. También permite ejecución manual desde
Actions y se ejecuta al modificar los scripts o el workflow.

Lee por Microsoft Graph la carpeta `BASES ENTEL/EPA_CALIDAD` de la biblioteca
predeterminada del mismo sitio SharePoint del flujo comercial. Reutiliza
`AZURE_TENANT_ID`, `AZURE_CLIENT_ID` y `AZURE_CLIENT_SECRET`. Las variables
opcionales `EPA_SITE_ID` y `EPA_FOLDER_PATH` permiten cambiar el origen.
No guarda Excel ni credenciales en el repositorio. Una descarga o validación
fallida detiene la publicación y conserva el HTML anterior. Si los datos no
cambian, conserva la fecha de actualización y evita commits innecesarios.

El usuario autorizó publicar el detalle completo en este repositorio público,
incluidos teléfonos, RUT, comentarios y enlaces a grabaciones.

## Actualizar

Con Python y openpyxl instalados, ejecutar `python sync_epa.py` en esta carpeta.
Lee `../../BASES ENTEL/EPA_CALIDAD`; admite `--source` y `--html` para otras rutas.
Actualiza el bloque EPA sin modificar los Excel ni el resto del panel.

## Criterios

- Período según mes/año de la encuesta.
- Hoja principal Export/BaseEpaPersona. Las hojas auxiliares de retiros,
  cambios y mapeos no se agregan como encuestas adicionales.
- Se conservan los registros sin deduplicarlos y los 69 campos originales.
- Equivalencias PDV 4628 → MALL CENTER CURICÓ y 4172 → TALCA MALL confirmadas.
- Algunos textos tienen caracteres dañados en la fuente. Los valores con
  formato de fecha inválido se recuperan como texto original.
- Relación con callback por período/sucursal; no es un cruce individual,
  porque el JSON de callbacks no contiene el identificador de encuesta.
- Copia inicial: 9.674 encuestas, junio 2025 a septiembre 2026.

Verificado en Edge sin red: período, sucursal, búsqueda, nota, paginación,
ficha de 69 campos y ancho móvil.
