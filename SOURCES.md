# Trazabilidad de fuentes

Todos los montos y pesos del dashboard provienen de microdatos oficiales. Las coordenadas
son referencia geográfica pública (no son dato de comercio).

## 1. Microdatos de comercio — Servicio Nacional de Aduanas (Chile)

| Dato | Archivo local | Registros | Campos usados |
|---|---|---|---|
| Exportaciones (DUS) 2025 | `Antecedentes/Estadisticas COMEX/Exportaciones/salidas2025/Salidas2025.csv` | 328.443 | COD_TIPO_OPERACION, COD_PUERTO_EMBARQUE, COD_PAIS_DESTINO, COD_REGION_ORIGEN, COD_VIA_TRANSPORTE, ITEM_SA, FOB_US_DUSLEG, PESO_BRUTO_KG, MES |
| Importaciones (DIN) 2025 | `Antecedentes/Estadisticas COMEX/Importaciones/Por lugar e ingreso/ingresos_2025/ingresos_2025.csv` | 1.754.023 | COD_TIPO_OPERACION, COD_PUERTO_DESEMBARQUE, COD_PAIS_ORIGEN, COD_ADUANA_TRAMITACION, COD_VIA_TRANSPORTE, ITEM_SA, CIF_US, CANTIDAD_MERCANCIA, COD_UNIDAD_MEDIDA, MES |

Portal de origen (descarga pública): Servicio Nacional de Aduanas — Estadísticas de Comercio
Exterior (datos abiertos / microdatos DUS y DIN). https://www.aduana.cl

## 2. Diccionarios y clasificadores — Servicio Nacional de Aduanas

| Recurso | Archivo local | Uso |
|---|---|---|
| Tablas de códigos (Anexo 51) | `Exportaciones/tablas_de_codigos.xlsx` | Puertos (cod→nombre/tipo/país, chilenos = país 997), Países (cod→nombre/continente), Regiones, Aduanas, Vías de Transporte |
| Clasificador Arancelario 2022 v2.0 | `Exportaciones/clasificador2022_v2_0.xlsx` | ITEM_SA (8 díg) → NIVEL 1 / NIVEL 2 (export e import) y CAP |

## 3. Cartografía base (geometrías, no datos de comercio)

| Recurso | Origen | Licencia | Uso |
|---|---|---|---|
| Países del mundo (GeoJSON, id ISO3) | github.com/johan/world.geo.json (`countries.geo.json`) | dominio público / MIT | choropleth mundial |
| Regiones de Chile (GeoJSON, `codregion`) | github.com/caracena/chile-geojson (`regiones.json`) | uso abierto | choropleth y centroides de influencia nacional |
| Mapa base de teselas | CARTO Light + © OpenStreetMap contributors | ODbL / CC | fondo del mapa |

Ambos GeoJSON se descargaron, se re-codificaron a UTF-8, se simplificó la precisión de
coordenadas (2 decimales) y se embebieron en `data_bundle.js`. Copias originales en `assets/`.

## 4. Coordenadas de los puntos (referencia geográfica)

- Base: tabla de coordenadas del documento de especificación del estudio
  (`PROMPT_ClaudeCode_Mapa_Puertos_COMEX.md`).
- Extendida con ubicaciones geográficas públicas y conocidas para 11 puntos que quedaban sin
  geolocalizar (Caleta Coloso, Patache, Los Vilos, Taltal, Corral, Constitución, Juan
  Fernández, Isla de Pascua, Paso San Francisco, Mamuil Malal, Paso Palena). Marcadas como
  `coord_src = "referencia_anadida"`.
- Los 5 puntos restantes se completaron con fuente verificada — **Servicio Nacional de
  Aduanas, "Ubicación geográfica de los pasos fronterizos"**
  (https://www.aduana.cl/ubicacion-geografica-de-los-pasos-fronterizos/aduana/2007-02-28/125018.html):
  Abra de Napa (20°30'S 68°35'W) y Lago Verde (44°15'S 71°48'W). Panguipulli usa la localidad
  (no figura en la tabla) y Territorio Antártico la Base Frei (Isla Rey Jorge) como punto
  representativo. **Baker:** la longitud oficial (79°51'W) es errónea —cae en el océano—; se
  conservó la latitud oficial (47°09'S) y se ubicó en la zona del Río Baker / Cochrane (Aysén).
- Resultado: **0 puntos sin coordenada**; los 74 con comercio quedan geolocalizados. Ninguna
  coordenada fue inventada.

## Nota de método

- Filtros de operación de export/import según el spec del estudio (Anexo 51 de Aduanas).
- Cuadre de control verificado: ver README (94,7 % export, 96,7 % import de cobertura).
- La región de importación es la **aduana de nacionalización**, no el destino final
  (limitación del microdato DIN).
