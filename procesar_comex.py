# -*- coding: utf-8 -*-
"""
procesar_comex.py
Procesa los microdatos oficiales del Servicio Nacional de Aduanas (DUS exportaciones /
DIN importaciones, anio 2025) y construye el perfil de comercio exterior de cada punto
de transferencia de carga de Chile (puertos maritimos, aeropuertos, avanzadas fronterizas,
ferrocarriles / puertos secos).

Salidas:
  data/puntos.json            -> un registro por punto (perfil completo)
  data/puntos_tidy.csv        -> equivalente tabular
  data/paises_comercio.json   -> agregado pais-socio (nacional)
  data/meta.json              -> cuadre de control, fechas, fuentes, puntos sin coordenada
  data_bundle.js              -> bundle embebible para el HTML (puntos + paises + geojson + meta)

Regla de oro: NO se inventan datos. Todo monto/peso sale del microdato. Las coordenadas son
referencia geografica (tabla del spec). Los puntos sin coordenada se listan, no se inventan.

Fuente: Servicio Nacional de Aduanas, microdatos DUS/DIN 2025; clasificador 2022 v2.0 y
tablas de codigos del Servicio. Elaboracion propia.
"""
import csv, json, unicodedata, sys, os
from collections import defaultdict, Counter
import openpyxl
import warnings
warnings.filterwarnings("ignore")

csv.field_size_limit(10_000_000)

# ---------------------------------------------------------------- rutas
HERE   = os.path.dirname(os.path.abspath(__file__))
STUDY  = os.path.dirname(HERE)
COMEX  = os.path.join(STUDY, "Antecedentes", "Estadisticas COMEX")
EXPO   = os.path.join(COMEX, "Exportaciones", "salidas2025", "Salidas2025.csv")
IMPO   = os.path.join(COMEX, "Importaciones", "Por lugar e ingreso", "ingresos_2025", "ingresos_2025.csv")
TABLAS = os.path.join(COMEX, "Exportaciones", "tablas_de_codigos.xlsx")
CLASIF = os.path.join(COMEX, "Exportaciones", "clasificador2022_v2_0.xlsx")
DATA   = os.path.join(HERE, "data");   os.makedirs(DATA, exist_ok=True)
ASSETS = os.path.join(HERE, "assets")

YEAR = 2025
CHILE_PAIS = 997

EXPORT_OPS = {200,201,202,203,204,205,206,207,210,211,212,213,216}
IMPORT_OPS = {101,102,103,104,105,113,115,116,119,120,121,122,123,129,130,134,142,151,152,165,171,179,180}

# aduana de tramitacion -> codigo de region (proxy de destino de importacion)
ADUANA_REGION = {
    3:15, 7:1, 10:2, 14:2, 17:3, 25:4, 33:5, 34:5, 39:5,
    48:13, 55:8, 67:10, 69:10, 83:11, 90:11, 92:12,
}

# Tabla de coordenadas = referencia geografica (del spec). Match por substring normalizado.
COORDS = {
 "ARICA": (-18.476,-70.323), "IQUIQUE": (-20.213,-70.152), "PATILLOS": (-20.74,-70.20),
 "TOCOPILLA": (-22.092,-70.197), "MEJILLONES": (-23.10,-70.45), "ANGAMOS": (-23.06,-70.43),
 "ANTOFAGASTA": (-23.65,-70.40), "CALDERA": (-27.067,-70.823), "BARQUITO": (-26.36,-70.66),
 "CHANARAL": (-26.35,-70.62), "HUASCO": (-28.46,-71.22), "GUACOLDA": (-28.46,-71.22),
 "COQUIMBO": (-29.95,-71.34), "GUAYACAN": (-29.97,-71.36), "VENTANAS": (-32.75,-71.48),
 "QUINTERO": (-32.78,-71.53), "VALPARAISO": (-33.04,-71.63), "SAN ANTONIO": (-33.59,-71.61),
 "LIRQUEN": (-36.71,-72.98), "PENCO": (-36.74,-72.99), "TALCAHUANO": (-36.72,-73.11),
 "SAN VICENTE": (-36.73,-73.13), "HUACHIPATO": (-36.74,-73.13), "CORONEL": (-37.03,-73.15),
 "LOTA": (-37.09,-73.16), "PUERTO MONTT": (-41.48,-72.94), "CALBUCO": (-41.77,-73.13),
 "CHACABUCO": (-45.46,-72.82), "PUERTO AYSEN": (-45.40,-72.69), "PUNTA ARENAS": (-53.16,-70.92),
 "NATALES": (-51.73,-72.51), "PUERTO WILLIAMS": (-54.93,-67.61), "CABO NEGRO": (-52.95,-70.80),
 "GREGORIO": (-52.62,-70.20), "TERMINAL GRANELES DEL NORTE": (-23.06,-70.43),
 "MERINO BENITEZ": (-33.39,-70.79), "ARTURO MERINO": (-33.39,-70.79),
 "CHACALLUTA": (-18.35,-70.34), "DIEGO ARACENA": (-20.54,-70.18), "CERRO MORENO": (-23.45,-70.44),
 "CARRIEL SUR": (-36.77,-73.06), "EL TEPUAL": (-41.44,-73.09), "IBANEZ DEL CAMPO": (-53.00,-70.85),
 "VISVIRI": (-17.60,-69.48), "CHUNGARA": (-18.22,-69.20), "COLCHANE": (-19.28,-68.64),
 "OLLAGUE": (-21.23,-68.25), "JAMA": (-23.23,-67.05), "SAN PEDRO DE ATACAMA": (-22.91,-67.99),
 "HITO CAJON": (-22.91,-67.77), "CONCORDIA": (-18.35,-70.32), "AGUA NEGRA": (-30.18,-69.82),
 "LOS LIBERTADORES": (-32.84,-70.08), "CRISTO REDENTOR": (-32.84,-70.08),
 "PEHUENCHE": (-35.98,-70.40), "PINO HACHADO": (-38.66,-70.89), "LIUCURA": (-38.66,-70.89),
 "CARDENAL SAMORE": (-40.72,-71.94), "FUTALEUFU": (-43.18,-71.85), "HUEMULES": (-45.62,-71.55),
 "COYHAIQUE ALTO": (-45.49,-71.30), "INTEGRACION AUSTRAL": (-52.16,-69.49),
 "MONTE AYMOND": (-52.16,-69.49), "DOROTEA": (-51.66,-72.36), "SAN SEBASTIAN": (-53.32,-68.62),
}

# Coordenadas de referencia ANADIDAS (ubicaciones geograficas publicas y conocidas, no dato COMEX).
# Extienden la tabla del spec para puntos relevantes que quedaban sin geolocalizar.
# Se marcan como "referencia anadida" en meta para trazabilidad.
COORDS_EXTRA = {
 "CALETA COLOSO": (-23.758,-70.470),   # terminal minero al sur de Antofagasta (Escondida)
 "COLOSO": (-23.758,-70.470),
 "PATACHE": (-20.810,-70.200),         # Punta Patache, al sur de Iquique
 "LOS VILOS": (-31.911,-71.510),
 "TALTAL": (-25.411,-70.484),
 "CORRAL": (-39.887,-73.430),
 "CONSTITUCION": (-35.333,-72.411),
 "JUAN FERNANDEZ": (-33.638,-78.830),
 "ISLA DE PASCUA": (-27.150,-109.433),
 "SAN FRANCISCO": (-26.900,-68.270),   # Paso San Francisco (Atacama)
 "MAHUIL MALAL": (-39.570,-71.500),    # Paso Mamuil Malal / Tromen
 "PALENA": (-43.620,-71.720),          # Paso Palena - Carrenleufu
 # --- completadas con fuente Servicio Nacional de Aduanas (ubicacion geografica de pasos) ---
 "ABRA DE NAPA": (-20.500,-68.583),    # Aduana: 20 30'S 68 35'W (altiplano, frontera con Bolivia)
 "LAGO VERDE": (-44.250,-71.800),      # Aduana: 44 15'S 71 48'W (Aysen)
 "PANGUIPULLI": (-39.643,-72.334),     # localidad (no figura en tabla de pasos de Aduana)
 "TERRITORIO ANTARTICO": (-62.196,-58.962),  # Base Pdte. Frei / Villa Las Estrellas, Isla Rey Jorge (punto representativo)
 "BAKER": (-47.150,-72.550),           # Aduana lat 47 09'S; long oficial 79 51'W es erronea (cae en el oceano); ubicado en zona Rio Baker / Cochrane (Aysen)
}
COORDS.update(COORDS_EXTRA)
COORDS_EXTRA_KEYS = set(COORDS_EXTRA)

# nombre Aduanas (normalizado, sin acentos, mayusculas) -> ISO3 del GeoJSON mundial
PAIS_ISO3 = {
 "CHINA":"CHN","ESTADOS UNIDOS DE AMERICA":"USA","JAPON":"JPN","BRASIL":"BRA",
 "COREA DEL SUR":"KOR","INDIA":"IND","HOLANDA":"NLD","ESPANA":"ESP","MEXICO":"MEX",
 "SUIZA":"CHE","PERU":"PER","ALEMANIA":"DEU","FRANCIA":"FRA","CANADA":"CAN",
 "ARGENTINA":"ARG","COLOMBIA":"COL","TAIWAN (FORMOSA)":"TWN","ITALIA":"ITA",
 "REINO UNIDO":"GBR","THAILANDIA":"THA","ECUADOR":"ECU","BELGICA":"BEL","RUSIA":"RUS",
 "SUECIA":"SWE","COSTA RICA":"CRI","VIETNAM":"VNM","BOLIVIA":"BOL","NORUEGA":"NOR",
 "AUSTRALIA":"AUS","POLONIA":"POL","FINLANDIA":"FIN","GUATEMALA":"GTM","BULGARIA":"BGR",
 "URUGUAY":"URY","MALASIA":"MYS","REPUBLICA DOMINICANA":"DOM","PARAGUAY":"PRY",
 "EMIRATOS ARABES UNIDOS":"ARE","BAHREIN":"BHR","DINAMARCA":"DNK","COSTA DE MARFIL":"CIV",
 "PUERTO RICO":"PRI","TURQUIA":"TUR","PANAMA":"PAN","NIGERIA":"NGA","ISRAEL":"ISR",
 "VENEZUELA":"VEN","FILIPINAS":"PHL","INDONESIA":"IDN","SINGAPUR":"SGP",
 "ARABIA SAUDITA":"SAU","SUDAFRICA":"ZAF","EGIPTO":"EGY","GRECIA":"GRC","PORTUGAL":"PRT",
 "IRLANDA":"IRL","AUSTRIA":"AUT","REPUBLICA CHECA":"CZE","HUNGRIA":"HUN","RUMANIA":"ROU",
 "UCRANIA":"UKR","HONDURAS":"HND","EL SALVADOR":"SLV","NICARAGUA":"NIC","CUBA":"CUB",
 "JAMAICA":"JAM","TRINIDAD Y TOBAGO":"TTO","NUEVA ZELANDIA":"NZL","NUEVA ZELANDA":"NZL",
 "PAKISTAN":"PAK","BANGLADESH":"BGD","SRI LANKA":"LKA","MARRUECOS":"MAR","ARGELIA":"DZA",
 "TUNEZ":"TUN","KENIA":"KEN","GHANA":"GHA","ANGOLA":"AGO","TANZANIA":"TZA",
 "ESLOVAQUIA":"SVK","ESLOVENIA":"SVN","CROACIA":"HRV","SERBIA":"SRB","LITUANIA":"LTU",
 "LETONIA":"LVA","ESTONIA":"EST","ISLANDIA":"ISL","LUXEMBURGO":"LUX","CHIPRE":"CYP",
 "QATAR":"QAT","KUWAIT":"KWT","OMAN":"OMN","JORDANIA":"JOR","LIBANO":"LBN","IRAN":"IRN",
 "IRAK":"IRQ","MYANMAR (BIRMANIA)":"MMR","CAMBOYA":"KHM","KAZAJSTAN":"KAZ",
 "HONG KONG":"HKG","MACAO":"HKG","BELICE":"BLZ","GUYANA":"GUY","SURINAM":"SUR",
 "MOZAMBIQUE":"MOZ","NAMIBIA":"NAM","ZAMBIA":"ZMB","ZIMBABWE":"ZWE","SENEGAL":"SEN",
 "CAMERUN":"CMR","GABON":"GAB","ETIOPIA":"ETH","MADAGASCAR":"MDG","MAURICIO":"MUS",
}

# ---------------------------------------------------------------- helpers
def norm(s):
    if s is None: return ""
    s = str(s).strip().upper()
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    return s

def fnum(x):
    if x is None: return 0.0
    x = str(x).strip().replace(",", ".")
    if x == "": return 0.0
    try: return float(x)
    except ValueError: return 0.0

def fint(x):
    try: return int(str(x).strip())
    except (ValueError, TypeError): return None

def tipo_grupo(tipo):
    t = norm(tipo)
    if "MARITIMO" in t: return "maritimo"
    if "AEROPUERTO" in t or "AEREO" in t: return "aereo"
    return "terrestre"   # avanzada fronteriza, ferrocarril, puerto seco

def match_coord(nombre):
    n = norm(nombre)
    for k,(la,lo) in COORDS.items():
        if k in n or n in k:
            return la, lo, k
    # match por palabra significativa
    for k,(la,lo) in COORDS.items():
        for w in k.split():
            if len(w) >= 5 and w in n:
                return la, lo, k
    return None

# ---------------------------------------------------------------- diccionarios
def load_dicts():
    wb = openpyxl.load_workbook(TABLAS, read_only=True, data_only=True)
    def sheet(pref):
        for n in wb.sheetnames:
            if norm(n).startswith(norm(pref)): return wb[n]
        raise KeyError(pref)

    puertos = {}
    for r in sheet("Puertos").iter_rows(min_row=6, values_only=True):
        cod = fint(r[1])
        if cod is None: continue
        puertos[cod] = dict(nombre=r[2], tipo=r[3], cod_pais=fint(r[4]),
                            pais=r[5], zona=r[6])
    paises = {}
    for r in sheet("Pais").iter_rows(min_row=6, values_only=True):
        cod = fint(r[1])
        if cod is None: continue
        paises[cod] = dict(nombre=r[2], continente=r[3])
    regiones = {}
    for r in sheet("Regiones").iter_rows(min_row=5, values_only=True):
        cod = fint(r[1])
        if cod is None: continue
        regiones[cod] = r[2]
    wb.close()

    # clasificador: arancel(8) -> niveles
    wc = openpyxl.load_workbook(CLASIF, read_only=True, data_only=True)
    ws = wc[wc.sheetnames[0]]
    clasif = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        ar = r[1]
        if ar is None: continue
        key = str(ar).strip().zfill(8)
        clasif[key] = dict(cap=str(r[0]).strip() if r[0] is not None else "",
                           n1_imp=r[3], n2_imp=r[4], n1_exp=r[6], n2_exp=r[7])
    wc.close()
    return puertos, paises, regiones, clasif

print("Cargando diccionarios...", flush=True)
PUERTOS, PAISES, REGIONES, CLASIF = load_dicts()
CHILE_PORTS = {c for c,p in PUERTOS.items() if p["cod_pais"] == CHILE_PAIS}
print(f"  puertos CL={len(CHILE_PORTS)}  paises={len(PAISES)}  regiones={len(REGIONES)}  aranceles={len(CLASIF)}", flush=True)

def clasif_get(item, field):
    key = str(item).strip().zfill(8)
    rec = CLASIF.get(key)
    if rec is None:
        # intenta por capitulo (2 primeros digitos) -> no; devuelve sin clasificar
        return None
    return rec.get(field)

# ---------------------------------------------------------------- acumuladores
def new_port():
    return dict(fob=0.0, peso_exp=0.0, cif=0.0, peso_imp=0.0,
                dest=Counter(), orig=Counter(),
                texp=Counter(), timp=Counter(),
                pexp=Counter(), pimp=Counter(),
                reg_exp=Counter(), reg_imp=Counter(),
                via_exp=Counter(), via_imp=Counter(),
                mes_exp=[0.0]*12, mes_imp=[0.0]*12)

ports = defaultdict(new_port)
nat = dict(fob=0.0, cif=0.0, peso_exp=0.0, peso_imp=0.0,
           dest=Counter(), orig=Counter(),
           mes_exp=[0.0]*12, mes_imp=[0.0]*12,
           reg_exp=Counter(), reg_imp=Counter())
nat_fob_all = 0.0   # exportacion nacional total (todas las ops export, cualquier puerto)
nat_cif_all = 0.0

# ---------------------------------------------------------------- EXPORTACIONES
print("Procesando exportaciones (Salidas2025.csv)...", flush=True)
n=0; kept=0
with open(EXPO, encoding="latin-1", newline="") as f:
    rd = csv.reader(f, delimiter=";")
    next(rd, None)
    for row in rd:
        n += 1
        if len(row) < 17: continue
        op = fint(row[3])
        if op not in EXPORT_OPS: continue
        fob = fnum(row[14]); peso = fnum(row[16])
        nat_fob_all += fob
        port = fint(row[6])
        if port not in CHILE_PORTS:   # solo puntos chilenos para el mapa
            continue
        kept += 1
        pais = fint(row[8]); region = fint(row[4]); via = fint(row[5])
        item = row[13]; mes = fint(row[1])
        d = ports[port]
        d["fob"] += fob; d["peso_exp"] += peso
        if mes and 1 <= mes <= 12: d["mes_exp"][mes-1] += fob
        if pais is not None: d["dest"][pais] += fob
        n1 = clasif_get(item, "n1_exp"); n2 = clasif_get(item, "n2_exp")
        if n1: d["texp"][n1] += fob
        if n2: d["pexp"][n2] += fob
        if region is not None: d["reg_exp"][region] += fob
        if via is not None: d["via_exp"][via] += fob
        nat["fob"] += fob; nat["peso_exp"] += peso
        if pais is not None: nat["dest"][pais] += fob
        if mes and 1 <= mes <= 12: nat["mes_exp"][mes-1] += fob
        if region is not None: nat["reg_exp"][region] += fob
print(f"  filas={n}  export-ops puntos CL={kept}  FOB nacional(all)={nat_fob_all:,.0f}", flush=True)

# ---------------------------------------------------------------- IMPORTACIONES
print("Procesando importaciones (ingresos_2025.csv)...", flush=True)
n=0; kept=0
with open(IMPO, encoding="latin-1", newline="") as f:
    rd = csv.reader(f, delimiter=";")
    next(rd, None)
    for row in rd:
        n += 1
        if len(row) < 17: continue
        op = fint(row[3])
        if op not in IMPORT_OPS: continue
        cif = fnum(row[12])
        nat_cif_all += cif
        port = fint(row[8])
        if port not in CHILE_PORTS:
            continue
        kept += 1
        pais = fint(row[4]); item = row[11]; via = fint(row[9]); aduana = fint(row[2])
        cant = fnum(row[15]); unidad = fint(row[16]); mes = fint(row[1])
        peso = cant if unidad == 6 else 0.0
        d = ports[port]
        d["cif"] += cif; d["peso_imp"] += peso
        if mes and 1 <= mes <= 12: d["mes_imp"][mes-1] += cif
        if pais is not None: d["orig"][pais] += cif
        n1 = clasif_get(item, "n1_imp"); n2 = clasif_get(item, "n2_imp")
        if n1: d["timp"][n1] += cif
        if n2: d["pimp"][n2] += cif
        reg = ADUANA_REGION.get(aduana)
        if reg is not None: d["reg_imp"][reg] += cif
        if via is not None: d["via_imp"][via] += cif
        nat["cif"] += cif; nat["peso_imp"] += peso
        if pais is not None: nat["orig"][pais] += cif
        if mes and 1 <= mes <= 12: nat["mes_imp"][mes-1] += cif
        reg = ADUANA_REGION.get(aduana)
        if reg is not None: nat["reg_imp"][reg] += cif
        if n % 300000 == 0: print(f"    ...{n} filas", flush=True)
print(f"  filas={n}  import-ops puntos CL={kept}  CIF nacional(all)={nat_cif_all:,.0f}", flush=True)

# ---------------------------------------------------------------- armado de registros
def pais_nombre(c):
    r = PAISES.get(c); return r["nombre"] if r else f"PAIS_{c}"
def pais_iso3(c):
    r = PAISES.get(c)
    if not r: return None
    return PAIS_ISO3.get(norm(r["nombre"]))
def region_nombre(c):
    return REGIONES.get(c, f"REGION_{c}")

def top_pais(counter, total, k=8):
    out=[]
    for c,v in counter.most_common(k):
        out.append(dict(cod=c, nombre=pais_nombre(c), iso3=pais_iso3(c), valor=round(v,1),
                        pct=round(100*v/total,2) if total else 0))
    return out

def top_tipo(counter, total, k=8):
    out=[]
    for name,v in counter.most_common(k):
        out.append(dict(nombre=name, valor=round(v,1),
                        pct=round(100*v/total,2) if total else 0))
    return out

def top_prod(counter, total, k=12):
    out=[]; acum=0.0
    for name,v in counter.most_common(k):
        acum += v
        out.append(dict(nombre=name, valor=round(v,1),
                        pct=round(100*v/total,2) if total else 0,
                        pct_acum=round(100*acum/total,2) if total else 0))
    return out

def top_region(counter, total, k=16):
    out=[]
    for c,v in counter.most_common(k):
        out.append(dict(cod=c, nombre=region_nombre(c), valor=round(v,1),
                        pct=round(100*v/total,2) if total else 0))
    return out

def via_dominante(d):
    tot = d["via_exp"] + d["via_imp"]
    if not tot: return None
    return tot.most_common(1)[0][0]

registros=[]; sin_coord=[]
for cod in sorted(ports):
    d = ports[cod]
    meta = PUERTOS[cod]
    total = d["fob"] + d["cif"]
    if total <= 0: continue
    coord = match_coord(meta["nombre"])
    if coord is None:
        lat, lon, coord_src = None, None, None
        sin_coord.append(dict(cod=cod, nombre=meta["nombre"], tipo=meta["tipo"],
                              total=round(total,0)))
    else:
        lat, lon, mkey = coord
        coord_src = "referencia_anadida" if mkey in COORDS_EXTRA_KEYS else "spec"
    reg = dict(
        cod=cod, nombre=meta["nombre"], tipo=meta["tipo"],
        grupo=tipo_grupo(meta["tipo"]), zona=meta["zona"],
        lat=lat, lon=lon, coord_src=coord_src,
        exp_fob=round(d["fob"],1), imp_cif=round(d["cif"],1),
        total=round(total,1),
        peso_exp=round(d["peso_exp"],1), peso_imp=round(d["peso_imp"],1),
        balance=round(d["cif"]-d["fob"],1),
        via_dom=via_dominante(d),
        top_destinos=top_pais(d["dest"], d["fob"]),
        top_origenes=top_pais(d["orig"], d["cif"]),
        top_tipo_exp=top_tipo(d["texp"], d["fob"]),
        top_tipo_imp=top_tipo(d["timp"], d["cif"]),
        top_prod_exp=top_prod(d["pexp"], d["fob"]),
        top_prod_imp=top_prod(d["pimp"], d["cif"]),
        infl_exp=top_region(d["reg_exp"], d["fob"]),   # regiones de origen de la carga exportada
        infl_imp=top_region(d["reg_imp"], d["cif"]),   # region de la aduana (proxy destino import)
        mes_exp=[round(x,1) for x in d["mes_exp"]],
        mes_imp=[round(x,1) for x in d["mes_imp"]],
    )
    registros.append(reg)

registros.sort(key=lambda r: r["total"], reverse=True)
sum_fob = sum(r["exp_fob"] for r in registros)
sum_cif = sum(r["imp_cif"] for r in registros)

# ---------------------------------------------------------------- paises_comercio (nacional)
paises_comercio = dict(
    exp=[dict(cod=c, nombre=pais_nombre(c), iso3=pais_iso3(c),
              continente=(PAISES.get(c) or {}).get("continente"),
              fob=round(v,1), pct=round(100*v/nat["fob"],3) if nat["fob"] else 0)
         for c,v in nat["dest"].most_common()],
    imp=[dict(cod=c, nombre=pais_nombre(c), iso3=pais_iso3(c),
              continente=(PAISES.get(c) or {}).get("continente"),
              cif=round(v,1), pct=round(100*v/nat["cif"],3) if nat["cif"] else 0)
         for c,v in nat["orig"].most_common()],
)

# ---------------------------------------------------------------- meta / cuadre
meta = dict(
    anio=YEAR,
    generado="procesar_comex.py",
    n_puntos=len(registros),
    cuadre=dict(
        exp_nacional_total=round(nat_fob_all,0),
        suma_fob_puntos=round(sum_fob,0),
        exp_cobertura_pct=round(100*sum_fob/nat_fob_all,2) if nat_fob_all else 0,
        imp_nacional_total=round(nat_cif_all,0),
        suma_cif_puntos=round(sum_cif,0),
        imp_cobertura_pct=round(100*sum_cif/nat_cif_all,2) if nat_cif_all else 0,
    ),
    sin_coordenada=sin_coord,
    fuentes=[
        "Servicio Nacional de Aduanas - microdato DUS (exportaciones) 2025",
        "Servicio Nacional de Aduanas - microdato DIN (importaciones) 2025",
        "Servicio Nacional de Aduanas - tablas de codigos (Anexo 51)",
        "Servicio Nacional de Aduanas - clasificador 2022 v2.0",
        "Coordenadas: referencia geografica publica (tabla del estudio, extendida)",
    ],
    notas=[
        "Influencia de EXPORTACION (infl_exp): region de origen real de la carga "
        "(COD_REGION_ORIGEN del DUS). Responde 'de que region del pais proviene lo exportado'.",
        "Influencia de IMPORTACION (infl_imp): region de la ADUANA DE NACIONALIZACION "
        "(COD_ADUANA_TRAMITACION). Indica donde se nacionaliza la carga, NO la region de "
        "destino/consumo final, que el microdato DIN no registra.",
        "Peso de importacion: el DIN no trae peso bruto; se usa CANTIDAD_MERCANCIA solo cuando "
        "COD_UNIDAD_MEDIDA = 6 (kilogramos).",
        "Cobertura del cuadre <100%: corresponde a operaciones con puerto no chileno o "
        "codigo de puerto 0/desconocido en el microdato.",
        "Coordenadas completadas (Abra de Napa, Lago Verde, Baker) desde la tabla de ubicacion "
        "geografica de pasos del Servicio Nacional de Aduanas. La longitud oficial de BAKER "
        "(79 51'W) es erronea (cae en el oceano Pacifico); se conservo la latitud oficial "
        "(47 09'S) y se ubico en la zona del Rio Baker / Cochrane (Aysen). Panguipulli usa la "
        "localidad y Territorio Antartico la Base Frei (Isla Rey Jorge) como punto representativo.",
    ],
)

# ---------------------------------------------------------------- escritura
with open(os.path.join(DATA,"puntos.json"),"w",encoding="utf-8") as f:
    json.dump(registros, f, ensure_ascii=False, indent=1)
with open(os.path.join(DATA,"paises_comercio.json"),"w",encoding="utf-8") as f:
    json.dump(paises_comercio, f, ensure_ascii=False, indent=1)
with open(os.path.join(DATA,"meta.json"),"w",encoding="utf-8") as f:
    json.dump(meta, f, ensure_ascii=False, indent=2)

# csv tidy
with open(os.path.join(DATA,"puntos_tidy.csv"),"w",encoding="utf-8-sig",newline="") as f:
    w = csv.writer(f, delimiter=";")
    w.writerow(["cod","nombre","tipo","grupo","zona","lat","lon",
                "exp_fob_usd","imp_cif_usd","total_usd","peso_exp_kg","peso_imp_kg",
                "balance_usd","top_destino","top_origen","top_tipo_exp","top_tipo_imp"])
    for r in registros:
        w.writerow([r["cod"],r["nombre"],r["tipo"],r["grupo"],r["zona"],r["lat"],r["lon"],
                    r["exp_fob"],r["imp_cif"],r["total"],r["peso_exp"],r["peso_imp"],r["balance"],
                    r["top_destinos"][0]["nombre"] if r["top_destinos"] else "",
                    r["top_origenes"][0]["nombre"] if r["top_origenes"] else "",
                    r["top_tipo_exp"][0]["nombre"] if r["top_tipo_exp"] else "",
                    r["top_tipo_imp"][0]["nombre"] if r["top_tipo_imp"] else ""])

# ---------------------------------------------------------------- geojson minificados + bundle
def round_geom(geom, nd=2):
    def rc(c):
        if isinstance(c[0], (int,float)):
            return [round(c[0],nd), round(c[1],nd)]
        return [rc(x) for x in c]
    geom["coordinates"] = rc(geom["coordinates"])
    return geom

# mundo: id ISO3 + name
world = json.load(open(os.path.join(ASSETS,"world_countries.geojson"),encoding="utf-8"))
for ft in world["features"]:
    ft["properties"] = {"name": ft["properties"].get("name")}
    round_geom(ft["geometry"], 2)

# chile regiones: codregion + nombre limpio (desde tabla SNA)
chreg = json.load(open(os.path.join(ASSETS,"chile_regiones.geojson"),encoding="latin-1"))
for ft in chreg["features"]:
    cr = ft["properties"].get("codregion")
    ft["properties"] = {"codregion": cr, "nombre": region_nombre(cr)}
    round_geom(ft["geometry"], 2)

# registro nacional (Total nacional) para la vista por defecto
nacional = dict(
    cod=0, nombre="TOTAL NACIONAL", tipo="Nacional", grupo="nacional", zona="Chile",
    lat=None, lon=None, coord_src=None,
    exp_fob=round(nat["fob"],1), imp_cif=round(nat["cif"],1),
    total=round(nat["fob"]+nat["cif"],1),
    peso_exp=round(nat["peso_exp"],1), peso_imp=round(nat["peso_imp"],1),
    balance=round(nat["cif"]-nat["fob"],1), via_dom=None,
    top_destinos=top_pais(nat["dest"], nat["fob"]),
    top_origenes=top_pais(nat["orig"], nat["cif"]),
    top_tipo_exp=[], top_tipo_imp=[], top_prod_exp=[], top_prod_imp=[],
    infl_exp=top_region(nat["reg_exp"], nat["fob"]),
    infl_imp=top_region(nat["reg_imp"], nat["cif"]),
    mes_exp=[round(x,1) for x in nat["mes_exp"]],
    mes_imp=[round(x,1) for x in nat["mes_imp"]],
)

bundle = dict(meta=meta, nacional=nacional, puntos=registros, paises=paises_comercio,
              world=world, regiones=chreg, region_nombres=REGIONES)
with open(os.path.join(HERE,"data_bundle.js"),"w",encoding="utf-8") as f:
    f.write("window.DATA = ")
    json.dump(bundle, f, ensure_ascii=False)
    f.write(";\n")

# ---------------------------------------------------------------- reporte
print("\n================ CUADRE DE CONTROL ================")
print(f"Exportacion nacional (todas ops):  US$ {nat_fob_all:,.0f}")
print(f"Suma FOB puntos chilenos:          US$ {sum_fob:,.0f}  ({meta['cuadre']['exp_cobertura_pct']}%)")
print(f"Importacion nacional (todas ops):  US$ {nat_cif_all:,.0f}")
print(f"Suma CIF puntos chilenos:          US$ {sum_cif:,.0f}  ({meta['cuadre']['imp_cobertura_pct']}%)")
print(f"\nPuntos con comercio: {len(registros)}")
print(f"Puntos SIN coordenada ({len(sin_coord)}):")
for s in sin_coord:
    print(f"   - [{s['cod']}] {s['nombre']} ({s['tipo']})  total US$ {s['total']:,.0f}")
print("\nArchivos escritos en:", DATA)
print("Bundle:", os.path.join(HERE,"data_bundle.js"))
