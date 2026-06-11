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
  *Tendencias* (carga por año 2012–2025, mensual, por terminal), *Logística & estructura*,
  *Comparar* (hasta 5 puntos) y *Análisis del área de influencia*.
- **Logística & estructura** (por punto, desde el microdato): modo de transporte
  (`COD_VIA_TRANSPORTE`), tipo de carga (`COD_TIPO_CARGA`/`TPO_CARGA`), Incoterm
  (`CLAUSULA_VENTA`/`CL_COMPRA`), concentración HHI de socios y productos, precio implícito
  US$/kg por producto, arancel ad-valorem recaudado (`AD_VALOREM_US`), y —para avanzadas
  fronterizas— **tráfico terrestre** de camiones/buses/autos y carga 2012–2025.
- **Panorama nacional** (desde reportes agregados de Aduanas): balanza comercial 2002–2025,
  evolución de exportaciones e importaciones por tipo de producto 2012–2025, participación
  por continente, y país × tipo de producto (exportación 2025).
- Buscador de puntos (insensible a acentos) y vista *Total nacional* por defecto.
- **Análisis por región** (panel superior): perfil completo de cada región de Chile —
  cuánto exporta/importa, de qué tipo de producto, con qué socios, y **por qué puntos
  sale/entra su carga** (cruce región × punto desde el microdato). Exportación usa la
  región de origen real (`COD_REGION_ORIGEN`); importación, la región de la aduana de
  nacionalización (proxy). El cuadre región↔puntos es exacto. También se puede
  seleccionar una región haciendo click en su polígono en la vista de influencia.
- **Selección múltiple de puntos** (Ctrl/Cmd/Shift + click en el mapa o el listado):
  agrega los perfiles de varios terminales —útil para puertos vecinos como San Vicente +
  Lirquén + Coronel— y todas las pestañas muestran el agregado (HHI y precio por producto
  se omiten porque requieren el universo completo de operaciones por punto).

## Metodología

### Fuentes (microdatos en `Antecedentes/Estadisticas COMEX/`)
- **Exportaciones (DUS):** `Exportaciones/salidas2025/Salidas2025.csv` — 328.443 registros.
- **Importaciones (DIN):** `Importaciones/Por lugar e ingreso/ingresos_2025/ingresos_2025.csv` — 1.754.023 registros.
- **Tablas de códigos:** `Exportaciones/tablas_de_codigos.xlsx` (Puertos, Países, Regiones, Aduanas, Vías).
- **Clasificador:** `Exportaciones/clasificador2022_v2_0.xlsx` (ITEM_SA → NIVEL 1 / NIVEL 2 / CAP).

Formato CSV: delimitador `;`, encoding `latin-1`, decimales con coma.

### Series anuales (movimiento de carga 2012–2025)
Además del año base 2025 (microdato DUS/DIN), el dashboard incluye la **evolución anual
2012–2025** por punto (valor FOB/CIF y tonelaje), a partir de los reportes agregados
oficiales del Servicio Nacional de Aduanas por lugar de salida/ingreso
(`expo_pto_monto/peso_*`, `impo_pto_monto/peso_*`). El año 2025 de esta serie agregada
cuadra ~99% con el valor derivado del microdato; pequeñas diferencias provienen del filtro
de tipos de operación. La serie nacional es la suma de los puntos (72/74 con histórico;
sin serie: Isla de Pascua/Punta Delgada y "otros", marginales).

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

## Puntos sin coordenada

**Ninguno: los 74 puntos con comercio están geolocalizados.** Los 5 que faltaban se
completaron con fuente verificada (no se inventaron):

| Cód | Nombre | Coordenada | Fuente |
|---|---|---|---|
| 944 | Territorio Antártico Chileno | −62.196, −58.962 | Base Pdte. Frei / Villa Las Estrellas (Isla Rey Jorge), punto representativo |
| 960 | Abra de Napa | −20.500, −68.583 | Servicio Nacional de Aduanas — ubicación geográfica de pasos (20°30'S 68°35'W) |
| 971 | Panguipulli | −39.643, −72.334 | Localidad de Panguipulli (no figura en la tabla de pasos de Aduana) |
| 973 | Lago Verde | −44.250, −71.800 | Servicio Nacional de Aduanas — pasos (44°15'S 71°48'W) |
| 978 | Baker | −47.150, −72.550 | Aduana lat. 47°09'S; **long. oficial 79°51'W es errónea** (cae en el Pacífico): se ubicó en la zona del Río Baker / Cochrane (Aysén) conservando la latitud oficial |

> Fuente de pasos: Servicio Nacional de Aduanas — *Ubicación geográfica de los pasos
> fronterizos* (aduana.cl). Todas marcadas con `coord_src = "referencia_anadida"`.

Coordenadas **añadidas** previamente a la tabla de referencia del spec (ubicaciones geográficas públicas,
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
icons/                 Ícono SVG por tipo de terminal + catálogo (icons/index.html)
SOURCES.md             Trazabilidad de fuentes
```

### Íconos por tipo de terminal

Cada punto del mapa usa un marcador con **glifo según su tipo** y **color según su grupo**;
los marcadores escalan con el comercio total. Archivos sueltos en `icons/` (catálogo en
`icons/index.html`):

| Tipo | Ícono | Color |
|---|---|---|
| Puerto marítimo | ⚓ ancla | azul |
| Aeropuerto | ✈ avión | naranja |
| Avanzada fronteriza | 🚧 barrera | verde |
| Ferrocarril | 🚆 tren | teal |
| Antepuerto / terminal de carga | 📦 contenedores | violeta |

Ferrocarril y antepuerto no tienen comercio registrado en 2025 en el microdato, pero el
ícono queda disponible (el sistema es extensible: nuevos tipos caen en el ícono de
contenedores por defecto).

Para regenerar todo: `python procesar_comex.py`

## Fuente y cita

Servicio Nacional de Aduanas — microdatos DUS (exportaciones) y DIN (importaciones) 2025;
clasificador 2022 v2.0 y tablas de códigos (Anexo 51). Elaboración propia.
Coordenadas: referencia geográfica pública.
