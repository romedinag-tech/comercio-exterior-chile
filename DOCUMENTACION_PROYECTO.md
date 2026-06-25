# Documentación del proyecto — Dinámicas de Comercio Exterior de Chile

> **Propósito de este archivo:** handoff completo para retomar el trabajo en cualquier
> momento (yo mismo en otra sesión, otra persona, u otro tipo de análisis) sin re-descubrir
> nada. Aquí está: qué datos hay y dónde, en qué formato, qué calcula el pipeline, cómo está
> hecho el dashboard, las decisiones metodológicas y trampas, el flujo de despliegue, los bugs
> ya resueltos, y el material para el próximo paso (modelo camión/ferrocarril).
> Última actualización: 2026-06-23.

---

## 0. Resumen en una línea

Dashboard web estático (HTML + Leaflet + Chart.js, sin build step) que perfila el comercio
exterior 2025 de **los 74 puntos de transferencia de carga de Chile** (puertos, aeropuertos,
avanzadas fronterizas) y de **las 18 "regiones" de origen/nacionalización**, a partir de los
microdatos oficiales del Servicio Nacional de Aduanas. Publicado en GitHub Pages.

- **Repo:** https://github.com/romedinag-tech/comercio-exterior-chile
- **Web en vivo:** https://romedinag-tech.github.io/comercio-exterior-chile/
- **Push:** directo vía Git Credential Manager (`gh` NO está autenticado).
- **Autor commits:** `Rodrigo Medina González`. Co-Author: Claude.

---

## 1. Ubicaciones (rutas absolutas)

| Qué | Dónde |
|---|---|
| **Proyecto / output** | `C:\Users\Rodrigo\Análisis RMG\Puerto Calbuco\Estudio Demanda Futura Puerto Industrial\mapa_puertos_comex\` |
| **Microdatos COMEX** | `…\Estudio Demanda Futura Puerto Industrial\Antecedentes\Estadisticas COMEX\` |
| **Material ferroviario** (para modelo) | `…\Puerto Calbuco\Antecedentes\Tren\` ← *ojo: un nivel arriba del Estudio* |
| **Spec original** | `…\Estudio Demanda Futura Puerto Industrial\PROMPT_ClaudeCode_Mapa_Puertos_COMEX.md` |

> El proyecto **no** está en la carpeta de trabajo por defecto de Claude (GIS Gran Concepción);
> siempre operar con rutas absolutas a Puerto Calbuco.

### Estructura de `mapa_puertos_comex/` (el repo publicado)
```
procesar_comex.py        Pipeline único (DuckDB no; usa csv+openpyxl streaming). Genera todo.
index.html               Dashboard (autocontenido salvo CDNs de Leaflet/Chart.js/fonts).
mapa_puertos.html        Copia idéntica de index.html (entregable con nombre del spec).
data_bundle.js           window.DATA = {...} embebido (1.3 MB). Lo lee el HTML.
data/                    puntos.json, regiones_perfil.json, paises_comercio.json,
                         meta.json, puntos_tidy.csv  (deliverables tabulares)
assets/                  world_countries.geojson (ISO3), chile_regiones.geojson (codregion)
icons/                   1 SVG por tipo de terminal + index.html (catálogo)
README.md                Metodología orientada al usuario.
SOURCES.md               Trazabilidad de fuentes.
DOCUMENTACION_PROYECTO.md  ESTE archivo.
```

---

## 2. Inventario de datos disponibles

### 2.1 Microdatos (lo que alimenta el dashboard 2025) — en `Antecedentes/Estadisticas COMEX/`
| Archivo | Qué es | Formato |
|---|---|---|
| `Exportaciones/salidas2025/Salidas2025.csv` | DUS exportaciones 2025 | `;` latin-1, decimales con coma, 328.443 filas |
| `Importaciones/Por lugar e ingreso/ingresos_2025/ingresos_2025.csv` | DIN importaciones 2025 | idem, 1.754.023 filas, 131 MB |
| `Exportaciones/tablas_de_codigos.xlsx` | Diccionarios Anexo 51 | hojas: Puertos, Países, Regiones, Aduanas, Vías de Transporte, Tipos de Carga, Cláusulas… |
| `Exportaciones/clasificador2022_v2_0.xlsx` | ITEM_SA(8díg) → NIVEL 1/2 export·import, CAP | 9.373 aranceles |

**Columnas DUS (export), índice 0:** `0 PERIODO;1 MES;2 COD_ADUANA_TRAMITACION;3 COD_TIPO_OPERACION;4 COD_REGION_ORIGEN;5 COD_VIA_TRANSPORTE;6 COD_PUERTO_EMBARQUE;7 COD_PUERTO_DESEMBARQUE;8 COD_PAIS_DESTINO;9 COD_MODALIDAD_VENTA;10 MONEDA;11 CLAUSULA_VENTA;12 COD_TIPO_CARGA;13 ITEM_SA;14 FOB_US_DUSLEG;15 FOBUS_AJUSTADO_IVV;16 PESO_BRUTO_KG;17 CANTIDAD_MERCANCIA;18 COD_UNIDAD_MEDIDA`
- Punto = col 6 (embarque). Filtro export: `COD_TIPO_OPERACION ∈ {200,201,202,203,204,205,206,207,210,211,212,213,216}`.

**Columnas DIN (import), índice 0:** `0 PERIODO;1 MES;2 COD_ADUANA_TRAMITACION;3 COD_TIPO_OPERACION;4 COD_PAIS_ORIGEN;5 COD_PAIS_ADQUISICION;6 COD_REGIMEN_IMPORTACION;7 COD_PUERTO_EMBARQUE;8 COD_PUERTO_DESEMBARQUE;9 COD_VIA_TRANSPORTE;10 CL_COMPRA;11 ITEM_SA;12 CIF_US;13 AD_VALOREM_US;14 MONEDA;15 CANTIDAD_MERCANCIA;16 COD_UNIDAD_MEDIDA;17 TPO_CARGA`
- Punto = col 8 (desembarque). Filtro import: `{101,102,103,104,105,113,115,116,119,120,121,122,123,129,130,134,142,151,152,165,171,179,180}`.
- **El DIN NO trae peso bruto:** se usa `CANTIDAD_MERCANCIA` (col 15) solo cuando `COD_UNIDAD_MEDIDA=6` (kg).

### 2.2 Agregados oficiales por puerto/producto/país (series históricas)
- **Por puerto × año** (monto FOB/CIF y peso kg): `expo_pto_monto_2002_2021.xlsx`, `expo_pto_monto_2022_2025.xlsx`, `impo_pto_monto_*`, `*_peso_*`. → alimentan la serie anual 2012–2025 por punto.
- **Por tipo de producto NIVEL 1 × año** (nacional): `expo_prod_monto_*`, `impo_prod_monto_*`. → Panorama nacional.
- **País × producto × año**: `expo_paisprod_monto_*`. → Panorama (país×producto + continente export).
- **Continente/país × año** (import): `impo_contipais_monto_*`. → Panorama continente import.
- **Balanza nacional 2002–2025**: `resumen_intercambiocomercial_2002_2025.xlsx`.
- Hay 24 xlsx en Exportaciones y 26 en Importaciones; los relevantes están en `procesar_comex.py` (sección lote 2). Estructura de cada uno: fila de título + encabezado con años; **autodetección de encabezado** en el lector (`read_pto_year`, `_read_named`).

### 2.3 Microdatos históricos (NO procesados aún — para futuro)
Hay salidas DUS de otros años en la carpeta: `bdexportacionesano2018`, `salidas_0112_2021` (2021), `salidas_2024.csv`, `salidas_mayo2026`, etc. URLs oficiales de los faltantes (2012–2017, 2020, 2023) en aduana.cl/publicaciones-base-de-datos-salidas. **Con el microdato de cada año se puede tener región×año×tipo-de-producto** (lo que los agregados históricos no permiten).

### 2.4 Tráfico terrestre — en `Antecedentes/Estadisticas COMEX/`
- `2010_2021_trafico_terrestre_final_v2.xlsx`, `2022_2025_trafico_terrestre_final.xlsx`.
- 4 hojas c/u (autos, buses, camiones, carga), por **Flujo (Ingreso/Salida) × Región × Avanzada × año**. → pestaña Flujo de carga terrestre (camiones por sentido 2010–2025).

### 2.5 Material ferroviario (para el modelo camión/tren) — en `Puerto Calbuco\Antecedentes\Tren\`
| Archivo | Contenido |
|---|---|
| `Carga transportadas en ferrocarril mensualizado.csv` | Observatorio MTT: ton y ton-km mensuales por **tipo de carga** (2017–2026) |
| `Toneladas transferidas por puertos transportadas en ferrocarril.csv` | Ton ferroviarias por **puerto** (2022–2026), con lat/lon |
| `Partición Modal Comercio Exterior Importación y Exportación.csv` | Modo de la pata internacional (marítimo/carretero/aéreo/ducto), 2018–2025 |
| `SDG_2011_Costos_Competitividad_Modos_Transporte_Carga.pdf` | Estudio madre de costos modales (CLP 2011): tren 11,9–17,2 / camión 18,8–21,2 $/ton-km |
| `MTT_2011_Analisis_FFCC_Carga_InformeEjecutivo.pdf` | Diagnóstico ferroviario |
| `Memoria_Fepasa_2024.pdf` | Tarifa real implícita ~54 CLP/ton-km (62.729 MM$ / 1.158 M ton-km) |
| `COSTOS_TREN_CAMION_resumen.md` | **Resumen mío de costos** + dato clave de transbordo (~$40.000/contenedor 2011 ≈ 1.600–2.000 $/ton) y distancia de equilibrio |

---

## 3. El pipeline: `procesar_comex.py`

Un solo script. Corre con `python procesar_comex.py` desde la carpeta del proyecto. No usa DuckDB
(lee CSV en streaming con `csv` + Excel con `openpyxl`). Tarda ~1–2 min (el DIN son 1.75M filas).

**Flujo:**
1. `load_dicts()` — Puertos (chilenos = `COD_PAIS=997`), Países (+continente), Regiones, Aduanas, Vías, Tipos de Carga, Cláusulas, clasificador. Encabezado real de los xlsx en fila 5 (offset de 1 columna por columna A vacía).
2. **Streaming export + import**: acumula por punto (`ports`) Y por región (`regs`) en el mismo paso. Counters por país, producto (N1/N2), región, vía, tipo carga, incoterm, mes, kg por producto, ad-valorem.
3. **Series anuales** (`load_anual`): merge de agregados por puerto 2002-2021 + 2022-2025, con **alias multi-variante** por punto (aeropuertos cambian de nombre entre archivos).
4. **Construcción de registros** por punto y por región (mismas funciones `top_pais/top_tipo/top_prod/estructura/hhi/precio_prod`).
5. **Geolocalización** punto→coordenada (tabla `COORDS` del spec + `COORDS_EXTRA` añadidas con fuente; `coord_src` marca cuáles). Punto→región por **point-in-polygon** (`punto_region`) contra el geojson regional (para la serie anual regional).
6. **Lote 2 (macro)**: balanza, producto×año, país×producto, continente, tráfico terrestre.
7. **Escribe** `data/*.json/csv`, `data_bundle.js`, y **estampa el cache-busting** (`data_bundle.js?v=<hash md5>`) y la **insignia de build** (`build <fecha> · <hash>`) en index.html y mapa_puertos.html.

**Helpers reutilizables:** `norm()` (mayúsc+sin acentos), `nkey()` (alnum), `fnum()` (coma→punto), `fint()`, `hhi()`, `precio_prod()`, `_read_named()`/`read_pto_year()` (lectores con autodetección de encabezado y filas de años).

---

## 4. KPIs y estadísticas calculadas (estructura del bundle)

`window.DATA` (en `data_bundle.js`) tiene: `meta, nacional, puntos[74], paises, world, regiones (geojson), region_nombres, macro, regiones_perfil[18]`.

**Cada PUNTO (y cada REGIÓN) trae:**
`cod, nombre, tipo, grupo (maritimo/aereo/terrestre/region), zona, lat, lon, coord_src,
exp_fob, imp_cif, total, peso_exp, peso_imp, balance (=cif−fob), via_dom,
top_destinos / top_origenes [{cod,nombre,iso3,valor,pct}],
top_tipo_exp / top_tipo_imp (NIVEL 1),
top_prod_exp / top_prod_imp (NIVEL 2, con pct_acum → 80/20),
infl_exp / infl_imp [{cod,nombre,valor,pct}] (regiones; export=COD_REGION_ORIGEN real, import=aduana proxy),
mes_exp / mes_imp [12] (estacionalidad 2025),
anual {years:[2012..2025], exp_fob, imp_cif, exp_kg, imp_kg},
estructura {via_exp/imp, tcarga_exp/imp, incoterm_exp/imp, hhi{destinos,origenes,prod_exp,prod_imp},
            precio_exp/imp (US$/kg por producto), usd_kg_exp/imp, advalorem, advalorem_pct},
region_ubicacion (solo puntos: codregion donde cae geográficamente)`
- Las **regiones** además traen `top_puntos_exp / top_puntos_imp` (por qué terminales sale/entra su carga).

**`macro` (nacional):** `prod {years, exp[], imp[]}` (producto NIVEL 1 2012-2025), `continente {years, imp[], exp[]}`, `paisprod[]` (top países × tipo producto 2025), `balanza {years, exp_fob, imp_cif, saldo}` (2002-2025).

**`meta.cuadre` (validación):** export Σpuntos = 101.798 MM US$ = **94,7%** del total nacional (107.491 MM); import Σpuntos = 76.787 MM = **96,65%** de 79.445 MM. El <100% son operaciones con puerto no chileno o código 0.

**KPIs visibles:** Exportación FOB nacional $101.80 MM (miles de millones), Importación CIF $76.79 MM, 74 puntos. (En la UI "MM" = miles de millones.)

---

## 5. Arquitectura del dashboard (`index.html`)

**Layout escritorio:** header (marca + ☰ + ☀ + 3 stats) · barra de regiones (panel superior, scroll horizontal) · `.top` (sidebar 264px + mapa) · panel inferior redimensionable con pestañas · footer (fuente + cuadre + **insignia build**).

**Selección — dos ejes coherentes (todo `state.sel`):**
- **Por lugar** (sidebar izquierdo): click = 1 punto; **Ctrl/Cmd/Shift+click = multi-selección** (agrega registros vía `mergeRecs()`).
- **Por región** (panel superior): click chip = perfil de esa región; también se puede clickear el polígono en el mapa.

**3 vistas de mapa (`setView`):** `puntos` (marcadores por tipo, tamaño ∝ comercio) · `influencia` (regiones de Chile + flujos región↔punto) · `mundo` (choropleth países). Encuadre automático **sin animación** (`animate:false`, idempotente) — clave para que no se "clave".

**4 formas cartográficas (control "Forma cartográfica", solo en influencia/mundo):** `coro_col` (Coropleta + Columnas 3D, default) · `coropleta` · `simbolos` (círculos proporcionales) · `columnas` (columnas 3D canvas con % encima). Helpers `isChoro()/isColumns()`.

**Barra de zonas (derecha):** macrozonas N→S por latitud; en choropleth = colorbar de %; **sombrea la importancia relativa de la selección a nivel país** (★ X% del país).

**7 pestañas (panel inferior):** Perfil · Tendencias (carga/año 2012-2025 US$/Ton + mensual + peso vs valor) · Logística & estructura (modo transporte, tipo carga, incoterm, HHI, precio US$/kg, ad-valorem) · Comparar (grupos desplegables por tipo + gráficos) · Análisis del área de influencia (nacional/mundial × export/import) · Flujo de carga terrestre (camiones ingreso/salida/total 2010-2025 + ranking de pasos) · Panorama nacional (balanza, producto, continente, país×producto).

**Estética:** tema oscuro "instrumento marino" (navy + acento **cobre** = el cobre es la export #1 de Chile), Space Grotesk (display/números) + Inter (UI), mapa CARTO dark/light. **Toggle tema claro/oscuro** y **auto-ocultar paneles** (hover, solo escritorio), persistidos en localStorage.

**Responsive iPad/táctil:** `@media (max-width:1024px), (hover:none) and (pointer:coarse)` → layout en **flujo vertical scrolleable** (NO flex+height:100%, que rompe en WebKit), sidebar = cajón deslizante (hamburguesa ☰ + backdrop), controles compactos sin leyenda, header sticky.

---

## 6. Decisiones metodológicas y TRAMPAS (no tropezar)

1. **Influencia de IMPORTACIÓN = aduana de nacionalización, NO destino final.** El DIN no registra región de destino. La de EXPORTACIÓN sí es real (`COD_REGION_ORIGEN`). Etiquetado con su advertencia en la UI.
2. **Región × punto:** el perfil de región (cuánto exporta/importa, de qué tipo) sale del microdato fila a fila (export por COD_REGION_ORIGEN, import por aduana). La **serie anual regional 2012-2025** es distinta: suma las series de los **terminales ubicados** en la región (point-in-polygon). Por eso KPI ≠ serie (ej. Biobío exporta 4,4 MM pero sus terminales mueven 12,6 MM — también sacan carga de regiones vecinas). Esto está explicado en una nota en la propia tarjeta.
3. **KPI 2025 vs serie de tendencia:** el KPI sale del microdato filtrado; la serie del reporte agregado oficial ("operaciones ajustadas con documentos modificatorios") → ±1% export, ~9% import. Nota automática cuando la brecha >1,5%.
4. **Código región 20 = "Mercancías extranjeras nacionalizadas"** (reexportaciones), no es región geográfica → no se pinta en choropleth, se etiqueta aparte.
5. **Coordenadas = referencia geográfica pública**, no dato COMEX. 5 puntos sin coord originalmente; completados con fuente (tabla de pasos de Aduanas). Baker: la long oficial 79°51'W cae en el océano → se conservó la lat y se ubicó en el Río Baker.
6. **"AEROPUERTO CARRIEL SUR 945"** está mal clasificado como "Puerto marítimo" en la tabla SNA → se normaliza a Aeropuerto en el procesador.
7. **`catastro_2025_1.csv`** (otro proyecto) trae filas corruptas — irrelevante aquí pero ojo si se reusa la carpeta.
8. **"MM" en la UI = miles de millones** (no millones).

---

## 7. Despliegue y caché

1. Editar en `mapa_puertos_comex/` → `cp index.html mapa_puertos.html` → `python procesar_comex.py` (estampa build+hash) → `git add -A && git commit && git push`.
2. **GitHub Pages cachea ~10 min** y **Chrome/Safari iOS cachean agresivo.** El cache-busting `data_bundle.js?v=<hash>` evita la mezcla HTML-nuevo/datos-viejos. La **insignia de build** en el pie permite verificar qué versión corre el navegador. Para forzar recarga: añadir `?x=1` a la URL.
3. **Modo `?debug`** en la URL → panel verde que registra cada movimiento del mapa (setView/fitBounds/panTo con origen) y excepciones. Útil para diagnosticar en el dispositivo del usuario.

---

## 8. Bugs históricos resueltos (lecciones)

- **NEUT() recursiva:** un reemplazo global `'#0e1726'→NEUT()` también pisó el literal dentro de la propia función → recursión infinita SOLO en tema oscuro → influencia/mundo rotos. **Lección: verificar SIEMPRE en ambos temas; el preview guardaba tema claro en localStorage y ocultaba el bug.**
- **Encuadre que "no funcionaba":** los fitBounds eran animados y la animación se cancelaba (re-render, invalidateSize). **Solución: todos los movimientos de mapa con `animate:false`.**
- **Colisión de nombre `n2`:** los loops usan `n2 = clasif_get(...)` (producto) y pisaban una función helper `n2()` a nivel módulo → se renombró el helper a `nkey()`.
- **iPad "secciones encimadas":** iOS/WebKit colapsa la cadena `height:100%`+flex. **Solución: layout de flujo de documento scrolleable disparado por `(pointer:coarse)`** (no solo por ancho — el iPad horizontal mide >1024). **Chrome del iPad también es WebKit.**
- **Screenshots del preview:** la herramienta de captura del preview falló de forma intermitente toda la sesión → la verificación se hizo por **inspección del DOM** (posiciones, computed styles, conteos), que es confiable.

---

## 9. Receta para un análisis NUEVO

1. **¿El dato ya está?** Revisa §2 y la estructura del bundle (§4). Mucho ya está calculado por punto/región/nacional.
2. **Si es del microdato 2025:** agrega contadores en el loop de `procesar_comex.py` (export ~línea del bloque "EXPORTACIONES", import idem) y un campo en `estructura()` o en el registro; cablea una tarjeta en la pestaña que corresponda (patrón: `renderX(body,p)` + dispatch en `renderPanel`).
3. **Si es serie histórica por producto/región:** necesitas el **microdato de cada año** (§2.3) — los agregados no traen producto×región. Es descarga + procesamiento acotado.
4. **Front nuevo:** entrada en el catálogo correspondiente, render en pestaña (botón `data-t`, `renderX`, dispatch). Reusa `barsHTML()`, los charts con `gridY`, `fmtUSD/fmtTon`.
5. **Siempre:** correr el procesador (re-estampa build), copiar a mapa_puertos.html, verificar por DOM en ambos temas, commit+push.

---

## 10. Pendientes / ideas futuras

- **Modelo de elección discreta camión/ferrocarril** (el norte del proyecto): matriz **región×terminal×segmento** (toneladas) desde el microdato como demanda total; particiones modales observadas por puerto (Observatorio + memorias); logit incremental (pivot-point) calibrado con corredores reales; costo generalizado con transbordos. Material en `Antecedentes/Tren/`. Ver `COSTOS_TREN_CAMION_resumen.md`.
- Netear graneles cautivos (mineroductos/cintas: Coloso, Patache, Los Vilos, Quintero…) ≈29% del tonelaje no compite por modo.
- Tendencia regional **por tipo de carga** (requiere microdato histórico por año).
- Arcos a países en influencia internacional; símbolos graduados; vista 3D real (deck.gl) como pestaña aparte (rompería el offline).
- Banco Central: precio del cobre / tipo de cambio para leer series en términos reales.
- 5 puntos sin serie histórica y los marginales sin coordenada (ya listados).
