# Comercio Exterior de Chile 2025 — Puntos de transferencia de carga

Sistema geográfico interactivo que ubica en un mapa **todos los puntos de transferencia de
carga de Chile** (puertos marítimos, aeropuertos, avanzadas fronterizas y ferrocarriles /
puertos secos) y muestra, para cada uno, su **perfil de comercio exterior con el mundo**:
valor y peso importado/exportado, países socios, tipos de carga, área de influencia nacional
y un mapa mundial (choropleth) por participación de países.

**Año de referencia: 2025.** Construido a partir de los microdatos oficiales del Servicio
Nacional de Aduanas. Regla de oro del proyecto: **no se inventan datos**.

## Cómo usarlo

Abre `index.html` en un navegador (funciona offline: todos los datos van embebidos en
`data_bundle.js`). También se publica como sitio estático (GitHub Pages).

- **Vista Puntos:** un marcador por punto; tamaño ∝ comercio total (FOB+CIF); color por tipo
  (azul marítimo, naranja aéreo, verde terrestre/fronterizo). Click → perfil en el panel.
- **Vista Influencia nacional:** muestra el *área de influencia* del punto seleccionado.
  En **verde**, las regiones de Chile **de donde proviene la carga exportada** (dato real,
  `COD_REGION_ORIGEN`); en **naranja**, la región donde se **nacionaliza la importación**
  (aduana de tramitación). Ambas capas se ven simultáneamente.
- **Vista Mundo:** choropleth de países del mundo según su % en las exportaciones (o
  importaciones) del punto seleccionado o del total nacional. Tooltip con país, % y monto.
- **Panel lateral:** pestañas *Perfil* (KPIs, destinos/orígenes, tipos de producto, 80/20),
  *Tendencias* (evolución mensual 2025, peso vs valor) y *Comparar* (hasta 5 puntos).
- Buscador de puntos (insensible a acentos) y vista *Total nacional* por defecto.

## Metodología

### Fuentes (microdatos en `Antecedentes/Estadisticas COMEX/`)
- **Exportaciones (DUS):** `Exportaciones/salidas2025/Salidas2025.csv` — 328.443 registros.
- **Importaciones (DIN):** `Importaciones/Por lugar e ingreso/ingresos_2025/ingresos_2025.csv` — 1.754.023 registros.
- **Tablas de códigos:** `Exportaciones/tablas_de_codigos.xlsx` (Puertos, Países, Regiones, Aduanas, Vías).
- **Clasificador:** `Exportaciones/clasificador2022_v2_0.xlsx` (ITEM_SA → NIVEL 1 / NIVEL 2 / CAP).

Formato CSV: delimitador `;`, encoding `latin-1`, decimales con coma.

### Procesamiento (`procesar_comex.py`)
1. **Filtros de operación.** Exportación: `COD_TIPO_OPERACION ∈ {200…216}`. Importación:
   `{101…180}` (lista completa en el script).
2. **Punto de transferencia.** Exportación = `COD_PUERTO_EMBARQUE`; Importación =
   `COD_PUERTO_DESEMBARQUE`. Se consideran sólo los puntos **chilenos** (`COD_PAIS = 997`
   en la tabla de Puertos): 105 puntos posibles, 74 con comercio registrado en 2025.
3. **Agregaciones por punto:** FOB/CIF total, peso, top países (FOB/CIF y %), tipo de
   producto (NIVEL 1), top productos (NIVEL 2 con % acumulado, principio 80/20), serie
   mensual y área de influencia regional.
4. **Valor / peso.** Export: `FOB_US_DUSLEG` y `PESO_BRUTO_KG`. Import: `CIF_US`; el DIN no
   trae peso bruto, así que el peso usa `CANTIDAD_MERCANCIA` sólo cuando `COD_UNIDAD_MEDIDA = 6` (kg).
5. **Geolocalización.** Match por nombre normalizado (mayúsculas, sin acentos, por substring)
   contra la tabla de coordenadas de referencia (no es dato COMEX).

### Cuadre de control (validación)
| Flujo | Total nacional (todas las ops) | Suma por puntos chilenos | Cobertura |
|---|---:|---:|---:|
| Exportación (FOB) | US$ 107.491 millones | US$ 101.798 millones | **94,7 %** |
| Importación (CIF) | US$ 79.445 millones | US$ 76.787 millones | **96,7 %** |

La diferencia (<100 %) corresponde a operaciones con puerto no chileno o código de puerto
0/desconocido en el microdato. Las cifras del dashboard cuadran con estos totales.

## Limitaciones (importante)

- **Área de influencia de importación.** El microdato DIN **no registra la región de destino
  o consumo final**. Como proxy se usa la **aduana de tramitación** (donde se nacionaliza la
  carga), que tiende a coincidir con la región del propio puerto. **No interpretar como
  destino final.** La influencia de *exportación* sí es real (región de origen de la carga).
- **Choropleth mundial por punto.** Usa los 8 principales socios del punto (≈ >90 % del valor).
  El choropleth nacional usa el universo completo de países.
- **Cobertura ISO3.** El mapeo nombre-país → ISO3 cubre ~99 % del FOB y ~99 % del CIF; el
  resto son códigos especiales (p. ej. rancho de naves) y una cola marginal de países.

## Puntos sin coordenada (5)

Geolocalización pendiente — montos sí registrados, marcadores no dibujados. **No se inventan
coordenadas.** Para incluirlos, agrega su `(lat, lon)` al diccionario `COORDS` de
`procesar_comex.py` y reejecuta.

| Cód | Nombre | Tipo | Total US$ |
|---|---|---|---:|
| 944 | Territorio Antártico Chileno | Puerto marítimo | 40.343 |
| 960 | Abra de Napa | Avanzada fronteriza | 571.230 |
| 971 | Panguipulli | Avanzada fronteriza | 800 |
| 973 | Lago Verde | Avanzada fronteriza | 22 |
| 978 | Baker | Avanzada fronteriza | 70.569 |

Coordenadas **añadidas** a la tabla de referencia del spec (ubicaciones geográficas públicas,
marcadas con `coord_src = "referencia_anadida"` en los datos): Caleta Coloso, Patache,
Los Vilos, Taltal, Corral, Constitución, Juan Fernández, Isla de Pascua, Paso San Francisco,
Mamuil Malal y Paso Palena.

## Archivos

```
procesar_comex.py      Pipeline de procesamiento (pandas/openpyxl, streaming)
index.html             Dashboard interactivo (Leaflet + Chart.js, autocontenido)
data_bundle.js         Datos embebidos para el HTML (puntos + países + geojson + meta)
data/puntos.json       Un registro por punto (perfil completo)
data/puntos_tidy.csv   Equivalente tabular
data/paises_comercio.json   Agregado país-socio nacional (export e import)
data/meta.json         Cuadre, fuentes, notas, puntos sin coordenada
assets/                GeoJSON mundo (ISO3) y regiones de Chile
SOURCES.md             Trazabilidad de fuentes
```

Para regenerar todo: `python procesar_comex.py`

## Fuente y cita

Servicio Nacional de Aduanas — microdatos DUS (exportaciones) y DIN (importaciones) 2025;
clasificador 2022 v2.0 y tablas de códigos (Anexo 51). Elaboración propia.
Coordenadas: referencia geográfica pública.
