# Informe gerencial de eventos

Landing estática para presentar aperturas, reinauguraciones y otros eventos de tienda. La información se actualiza mediante una plantilla de Excel y permanece editable en el navegador.

## Uso

1. Abre `index.html` en Chrome o Edge.
2. Selecciona **Editar informe**.
3. Descarga y completa `plantilla-informe-evento.xlsx`.
4. Carga el Excel desde el panel de edición.
5. Ajusta textos o reemplaza fotografías directamente en la landing.
6. Guarda los cambios o utiliza **Imprimir / PDF**.

Los cambios se almacenan localmente en el navegador. La opción **Exportar respaldo** permite generar un archivo JSON para trasladarlos a otro equipo.

## Estructura del Excel

- `Evento`: información general, periodo y análisis.
- `Ventas`: facturas utilizadas para calcular venta, ATV y evolución diaria.
- `Trafico`: tráfico exterior, ingreso a tienda y conversión.
- `Pauta Digital`: impresiones, clics, CTR e inversión.
- `Momentos`: lectura previa, día del evento y periodo posterior.
- `Gastos`: proveedor, concepto, gasto y observaciones.
- `Acciones`: actividades de comunicación y ejecución.
- `Dinamicas`: mecánicas comerciales y resultados.
- `Resultados`: hallazgos gerenciales.
- `Fotos`: evidencias con título, categoría, descripción y enlace.

## Fotografías

Los enlaces de Google Drive deben estar compartidos con la opción **Cualquier persona con el enlace**. También es posible reemplazar las imágenes manualmente desde el modo edición.

## Publicación en GitHub Pages

1. Carga todos los archivos y la carpeta `assets` en la raíz del repositorio.
2. En GitHub abre **Settings > Pages**.
3. En **Build and deployment**, selecciona **Deploy from a branch**.
4. Elige la rama `main` y la carpeta `/ (root)`.
5. Guarda la configuración.

GitHub mostrará la dirección pública cuando termine la publicación.
