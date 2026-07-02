# ============================================================
# MAPA INVEX - ULTRA AMIGABLE PARA CELULAR
# Con botones de minimizar panel y cerrar tarjeta de info
# ============================================================

import os
import re
import sys
import json
import math
import time
import subprocess
import shutil
import unicodedata
import warnings
warnings.filterwarnings("ignore")

# ============================================================
# INSTALAR DEPENDENCIAS
# ============================================================

def instalar_si_falta(modulo, paquete_pip=None):
    if paquete_pip is None:
        paquete_pip = modulo
    try:
        __import__(modulo)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", paquete_pip])

instalar_si_falta("pandas")
instalar_si_falta("openpyxl")
instalar_si_falta("geopy")
instalar_si_falta("requests")
instalar_si_falta("bs4", "beautifulsoup4")

import pandas as pd
import requests
from bs4 import BeautifulSoup
from getpass import getpass
from datetime import datetime
from pathlib import Path

try:
    from zoneinfo import ZoneInfo
except Exception:
    ZoneInfo = None

from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter

# Este script está pensado para correr LOCALMENTE desde la carpeta del repo de GitHub.
EN_COLAB = False
files = None

try:
    from IPython.display import display
except Exception:
    def display(obj):
        print(obj)

# ============================================================
# CONFIGURACION
# ============================================================

SOLO_TIENDAS_OPERANDO = True
ESTATUS_FINALIZADOS = {"FINALIZADO"}

# Carpeta base: debe ser la carpeta del repo clonado
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
CARPETA_RECORDS_LOCAL = BASE_DIR / "records_local"
CARPETA_SALIDA_LOCAL = BASE_DIR / "salida_local"

DATA_DIR.mkdir(exist_ok=True)
CARPETA_RECORDS_LOCAL.mkdir(exist_ok=True)
CARPETA_SALIDA_LOCAL.mkdir(exist_ok=True)

# Archivos fijos que debes dejar dentro de /data
RUTA_ASISTENCIA_LOCAL = DATA_DIR / "ASISTENCIA.xlsx"
RUTA_BASE_LOCAL = DATA_DIR / "BASE_DE_TIENDAS.xlsx"

# Salidas
SALIDA_HTML_REPO = BASE_DIR / "index.html"
SALIDA_EXCEL_LOCAL = CARPETA_SALIDA_LOCAL / "resumen_mapa_invex.xlsx"

# Cache de coordenadas dentro del repo local, para que no recalcule siempre
ARCHIVO_CACHE_COORDS = str(BASE_DIR / "cache_coordenadas_tiendas_v3_validado.xlsx")

# GitHub
AUTO_GIT_PUSH = True
ARCHIVOS_A_SUBIR_GITHUB = ["index.html"]

# ============================================================
# DESCARGA AUTOMATICA DE RECORDS DEL DIA
# ============================================================

USAR_RECORDS_AUTOMATICO = True
DESCARGAR_RECORDS_A_PC = False  # Cambia a True si quieres descargar tambien el Records original

URL_BASE = "https://modulos.invex.maatai.com"
ADMIN_URL = URL_BASE + "/admin/"
REPORT_URL = URL_BASE + "/admin/records/recordreport/"


PALETA_REGIONES = [
    "#2563EB", "#7C3AED", "#DB2777", "#EA580C", "#059669",
    "#0891B2", "#4F46E5", "#BE123C", "#65A30D", "#9333EA",
    "#0284C7", "#C2410C", "#16A34A", "#A21CAF", "#0F766E",
    "#CA8A04", "#DC2626", "#1D4ED8", "#047857", "#6D28D9",
    "#B45309", "#0369A1", "#9F1239", "#15803D", "#4338CA"
]

# ============================================================
# FUNCIONES DE LIMPIEZA
# ============================================================

def limpiar_texto(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).strip().upper()
    texto = unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("utf-8")
    texto = re.sub(r"\s+", " ", texto)
    return texto.strip()

def limpiar_texto_mostrar(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).strip()
    texto = re.sub(r"\s+", " ", texto)
    return texto.strip()

def extraer_determinante(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).strip()
    match = re.search(r"\b(\d{3,5})\b", texto)
    if match:
        return match.group(1)
    return ""

def normalizar_columna(col):
    return limpiar_texto(col).replace("\n", " ").strip()

def mapa_columnas(df):
    return {normalizar_columna(c): c for c in df.columns}

def buscar_columna(df, opciones, obligatorio=True):
    cols = mapa_columnas(df)
    opciones_norm = [normalizar_columna(o) for o in opciones]
    for op in opciones_norm:
        if op in cols:
            return cols[op]
    for op in opciones_norm:
        for col_norm, col_original in cols.items():
            if op in col_norm or col_norm in op:
                return col_original
    if obligatorio:
        raise ValueError(f"No encontre columna. Busque alguna de estas opciones: {opciones}")
    return None

def leer_excel_auto(ruta, hoja_preferida=None):
    xls = pd.ExcelFile(ruta)
    if hoja_preferida and hoja_preferida in xls.sheet_names:
        print(f"Hoja usada en {os.path.basename(ruta)}: {hoja_preferida}")
        return pd.read_excel(ruta, sheet_name=hoja_preferida)
    print(f"Hoja usada en {os.path.basename(ruta)}: {xls.sheet_names[0]}")
    return pd.read_excel(ruta, sheet_name=xls.sheet_names[0])

# ============================================================
# SEMAFORO
# ============================================================

def color_semaforo(finalizados):
    if finalizados <= 0:
        return "#DC2626"
    elif finalizados == 1:
        return "#FACC15"
    else:
        return "#16A34A"

def texto_semaforo(finalizados):
    if finalizados <= 0:
        return "ROJO - Sin tramites listos"
    elif finalizados == 1:
        return "AMARILLO - 1 tramite listo"
    else:
        return "VERDE - 2 o mas tramites listos"

def mensaje_simple(finalizados):
    if finalizados <= 0:
        return "Necesita apoyo"
    elif finalizados == 1:
        return "Va empezando"
    else:
        return "Va bien"

# ============================================================
# DESCARGA AUTOMATICA DE RECORDS DE HOY
# ============================================================

def fecha_hoy_mexico():
    if ZoneInfo:
        return datetime.now(ZoneInfo("America/Mexico_City"))
    return datetime.now()

def descargar_records_hoy():
    """
    Descarga automaticamente el reporte Records del dia actual desde MAAT/Invex.
    Guarda el archivo localmente y devuelve la ruta para usarlo en el mapa.
    """
    print("\n" + "═" * 60)
    print("     📥 DESCARGA AUTOMATICA DE RECORDS DEL DIA")
    print("═" * 60 + "\n")

    # Para que sea automático con Programador de tareas, guarda estas variables en Windows:
    # setx INVEX_USER "tu_usuario"
    # setx INVEX_PASS "tu_contraseña"
    usuario = os.environ.get("INVEX_USER", "").strip()
    password = os.environ.get("INVEX_PASS", "")

    if not usuario:
        usuario = input("👤 Usuario Invex/MAAT: ").strip()
    if not password:
        password = getpass("🔑 Contraseña Invex/MAAT: ")

    if not usuario or not password:
        raise ValueError("❌ Ingresa usuario y contraseña o configura INVEX_USER / INVEX_PASS.")

    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0"})

    print("\n[1/3] 🔐 Iniciando sesión...")

    resp = session.get(ADMIN_URL, timeout=30)
    soup = BeautifulSoup(resp.text, "html.parser")

    csrf = soup.find("input", {"name": "csrfmiddlewaretoken"})
    nxt = soup.find("input", {"name": "next"})

    if not csrf:
        raise ValueError("❌ No se encontró CSRF token en login. Puede haber cambiado la página.")

    payload = {
        "username": usuario,
        "password": password,
        "csrfmiddlewaretoken": csrf.get("value"),
        "next": nxt.get("value") if nxt else "/admin/",
    }

    login_resp = session.post(
        resp.url,
        data=payload,
        headers={
            "Referer": resp.url,
            "User-Agent": "Mozilla/5.0",
        },
        allow_redirects=True,
        timeout=30,
    )

    texto_login = login_resp.text.lower()

    if 'name="username"' in texto_login and 'name="password"' in texto_login:
        raise ValueError("❌ Login fallido. Revisa usuario/contraseña.")

    print("✅ Login OK")

    print("\n[2/3] 📄 Solicitando Records de hoy...")

    hoy = fecha_hoy_mexico()
    fecha_form = hoy.strftime("%d/%m/%Y")
    fecha_archivo = hoy.strftime("%Y%m%d_%H%M%S")

    r = session.get(REPORT_URL, timeout=30)
    soup = BeautifulSoup(r.text, "html.parser")

    csrf_input = soup.find("input", {"name": "csrfmiddlewaretoken"})

    if not csrf_input:
        debug_path = os.path.join(CARPETA_RECORDS_LOCAL, f"debug_reporte_sin_csrf_{fecha_archivo}.html")
        with open(debug_path, "w", encoding="utf-8") as f:
            f.write(r.text)
        raise ValueError(f"❌ No se encontró CSRF token en reporte. Debug guardado en: {debug_path}")

    form_data = {
        "csrfmiddlewaretoken": csrf_input.get("value"),
        "start": fecha_form,
        "end": fecha_form,
        "format": "XLSX",
    }

    r = session.post(
        REPORT_URL,
        data=form_data,
        headers={
            "Referer": REPORT_URL,
            "User-Agent": "Mozilla/5.0",
        },
        allow_redirects=True,
        timeout=180,
    )

    es_excel = False
    ext = ".xlsx"

    if len(r.content) > 4:
        magic = r.content[:4]

        if magic == b"PK\x03\x04":
            es_excel = True
            ext = ".xlsx"
        elif magic == b"\xd0\xCF\x11\xE0":
            es_excel = True
            ext = ".xls"

    if not es_excel or len(r.content) <= 100:
        debug_path = os.path.join(CARPETA_RECORDS_LOCAL, f"debug_records_{fecha_archivo}.html")
        try:
            with open(debug_path, "w", encoding="utf-8") as f:
                f.write(r.text)
        except Exception:
            debug_path = "No se pudo guardar debug"

        raise ValueError(
            "❌ No se descargó un Excel válido.\n"
            f"Status: {r.status_code}\n"
            f"Content-Type: {r.headers.get('Content-Type', 'N/A')}\n"
            f"Tamaño: {len(r.content)} bytes\n"
            f"Debug: {debug_path}"
        )

    nombre_archivo = f"Records_HOY_{fecha_archivo}{ext}"
    ruta_records = os.path.join(CARPETA_RECORDS_LOCAL, nombre_archivo)

    with open(ruta_records, "wb") as f:
        f.write(r.content)

    print(f"✅ Records descargado: {nombre_archivo}")
    print(f"📁 Guardado en: {ruta_records}")
    print(f"📅 Fecha descargada: {fecha_form}")

    if DESCARGAR_RECORDS_A_PC and EN_COLAB and files is not None:
        try:
            files.download(ruta_records)
        except Exception:
            print("⚠️ No se pudo descargar el Records a tu PC, pero sí quedó guardado localmente.")

    print("\n[3/3] ✅ Records listo para actualizar el mapa\n")
    return ruta_records


# ============================================================
# ARCHIVOS LOCALES - ASISTENCIA Y BASE DE TIENDAS
# RECORDS SE DESCARGA AUTOMATICAMENTE
# ============================================================

def obtener_archivos_locales():
    """
    Lee archivos fijos desde la carpeta /data del repo.
    Estructura esperada:
      data/ASISTENCIA.xlsx
      data/BASE_DE_TIENDAS.xlsx
    """
    faltantes = []

    if not RUTA_ASISTENCIA_LOCAL.exists():
        faltantes.append(str(RUTA_ASISTENCIA_LOCAL))

    if not RUTA_BASE_LOCAL.exists():
        faltantes.append(str(RUTA_BASE_LOCAL))

    if faltantes:
        raise FileNotFoundError(
            "❌ Faltan archivos locales:\n" +
            "\n".join(f" - {x}" for x in faltantes) +
            "\n\nGuarda tus archivos con estos nombres exactos dentro de la carpeta data."
        )

    return RUTA_ASISTENCIA_LOCAL, RUTA_BASE_LOCAL

# ============================================================
# COORDENADAS
# ============================================================


def construir_cache_key(row):
    """
    Llave para el cache. Incluye direccion/tienda/ciudad/estado para que,
    si actualizas la BASE DE TIENDAS, se vuelvan a calcular las coordenadas.
    """
    partes = [
        str(row.get("DET", "")),
        limpiar_texto(row.get("DIRECCION_TIENDA_N", "")),
        limpiar_texto(row.get("TIENDA_N", "")),
        limpiar_texto(row.get("CIUDAD_N", "")),
        limpiar_texto(row.get("ESTADO_N", "")),
    ]
    return "|".join(partes)

def cargar_cache_coords():
    columnas = [
        "CACHE_KEY", "DET", "LAT", "LON", "GEOCODE_QUERY",
        "TIPO_UBICACION", "UBICACION_APROXIMADA"
    ]
    if os.path.exists(ARCHIVO_CACHE_COORDS):
        try:
            cache = pd.read_excel(ARCHIVO_CACHE_COORDS)
            for col in columnas:
                if col not in cache.columns:
                    cache[col] = pd.NA
            cache["DET"] = cache["DET"].astype(str)
            cache["CACHE_KEY"] = cache["CACHE_KEY"].astype(str)
            return cache[columnas]
        except:
            pass
    return pd.DataFrame(columns=columnas)

def guardar_cache_coords(cache):
    cache = cache.drop_duplicates(subset=["CACHE_KEY"], keep="last")
    cache.to_excel(ARCHIVO_CACHE_COORDS, index=False)

def limpiar_query_geo(texto):
    texto = "" if pd.isna(texto) else str(texto)
    texto = texto.replace("\n", " ")
    texto = re.sub(r"\s+", " ", texto)
    texto = texto.strip(" ,")
    return texto

def normalizar_para_geo(texto):
    """Normaliza texto para comparar estado/ciudad en resultados de geocodificacion."""
    texto = "" if pd.isna(texto) else str(texto)
    texto = unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("utf-8")
    texto = texto.upper()
    texto = re.sub(r"[^A-Z0-9 ]+", " ", texto)
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto

def variantes_estado_geo(estado):
    """Devuelve variantes comunes del estado para validar que Nominatim no mande la tienda a otro estado."""
    e = normalizar_para_geo(estado)
    variantes = {e}

    mapa = {
        "ESTADO DE MEXICO": {"ESTADO DE MEXICO", "MEXICO STATE", "STATE OF MEXICO", "EDOMEX", "EDO MEX"},
        "CIUDAD DE MEXICO": {"CIUDAD DE MEXICO", "CDMX", "MEXICO CITY"},
        "MICHOACAN": {"MICHOACAN", "MICHOACAN DE OCAMPO"},
        "VERACRUZ": {"VERACRUZ", "VERACRUZ DE IGNACIO DE LA LLAVE"},
        "COAHUILA": {"COAHUILA", "COAHUILA DE ZARAGOZA"},
        "NUEVO LEON": {"NUEVO LEON"},
        "SAN LUIS POTOSI": {"SAN LUIS POTOSI"},
        "QUERETARO": {"QUERETARO"},
        "YUCATAN": {"YUCATAN"},
    }

    # Quitar prefijos que a veces vienen en base
    e_sin_prefijo = e.replace("ESTADO DE ", "").strip()
    if e_sin_prefijo:
        variantes.add(e_sin_prefijo)

    for clave, vals in mapa.items():
        if e == clave or e in vals or clave in e:
            variantes.update(vals)

    return {v for v in variantes if v}

def resultado_coincide_estado(loc, estado_esperado):
    """
    Evita aceptar coordenadas de otro estado.
    Ejemplo: si la base dice ESTADO DE MEXICO, no acepta un resultado en CHIHUAHUA.
    """
    if not estado_esperado:
        return True

    texto_resultado = normalizar_para_geo(getattr(loc, "address", ""))
    try:
        raw = getattr(loc, "raw", {}) or {}
        address = raw.get("address", {}) or {}
        partes = " ".join(str(v) for v in address.values())
        texto_resultado = normalizar_para_geo(texto_resultado + " " + partes)
    except Exception:
        pass

    variantes = variantes_estado_geo(estado_esperado)
    return any(v in texto_resultado for v in variantes)

def geocodificar_tiendas(df_base):
    """
    Primero intenta ubicar por DIRECCION TIENDA.
    Si no encuentra resultado, usa aproximaciones:
    1) nombre de tienda + ciudad + estado
    2) ciudad + estado
    3) estado
    En el mapa/Excel queda marcado si fue aproximada.
    """
    df_base = df_base.copy()
    df_base["CACHE_KEY"] = df_base.apply(construir_cache_key, axis=1)

    columnas_cache = [
        "CACHE_KEY", "LAT", "LON", "GEOCODE_QUERY",
        "TIPO_UBICACION", "UBICACION_APROXIMADA"
    ]

    cache = cargar_cache_coords()

    df_base = df_base.drop(
        columns=["LAT", "LON", "GEOCODE_QUERY", "TIPO_UBICACION", "UBICACION_APROXIMADA"],
        errors="ignore"
    )

    if not cache.empty:
        df_base = df_base.merge(cache[columnas_cache], on="CACHE_KEY", how="left")
    else:
        df_base["LAT"] = pd.NA
        df_base["LON"] = pd.NA
        df_base["GEOCODE_QUERY"] = ""
        df_base["TIPO_UBICACION"] = ""
        df_base["UBICACION_APROXIMADA"] = True

    faltan = df_base[df_base["LAT"].isna() | df_base["LON"].isna()].copy()

    if len(faltan) == 0:
        print("Todas las tiendas ya tienen coordenadas en cache.")
        return df_base

    print(f"Buscando coordenadas de {len(faltan)} tiendas...")
    print("Prioridad: DIRECCION TIENDA. Si no encuentra, usa ubicacion aproximada por tienda/ciudad/estado.")
    print("La primera vez puede tardar unos minutos.")

    geolocator = Nominatim(user_agent="invex_mapa_celular_direccion_tienda")
    geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1.1)

    nuevos = []

    for _, row in faltan.iterrows():
        det = str(row["DET"])
        direccion = limpiar_query_geo(row.get("DIRECCION_TIENDA_N", ""))
        tienda = limpiar_query_geo(row.get("TIENDA_N", ""))
        ciudad = limpiar_query_geo(row.get("CIUDAD_N", ""))
        estado = limpiar_query_geo(row.get("ESTADO_N", ""))

        candidatos = []

        # 1) Intento principal: direccion exacta de la BASE DE TIENDAS
        if direccion:
            candidatos.append((
                f"{direccion}, {ciudad}, {estado}, Mexico",
                "DIRECCION TIENDA",
                False
            ))

        # 2) Respaldos aproximados
        if tienda and ciudad:
            candidatos.append((
                f"{tienda}, {ciudad}, {estado}, Mexico",
                "APROXIMADA POR TIENDA",
                True
            ))
        if ciudad and estado:
            candidatos.append((
                f"{ciudad}, {estado}, Mexico",
                "APROXIMADA POR CIUDAD",
                True
            ))
        if estado:
            candidatos.append((
                f"{estado}, Mexico",
                "APROXIMADA POR ESTADO",
                True
            ))

        lat = None
        lon = None
        query_usada = ""
        tipo_ubicacion = "SIN COORDENADAS"
        aproximada = True

        for q, tipo, es_aprox in candidatos:
            q = limpiar_query_geo(q)
            if not q:
                continue

            try:
                # country_codes="mx" evita resultados fuera de Mexico.
                # addressdetails=True nos ayuda a validar estado contra la BASE DE TIENDAS.
                loc = geocode(q, country_codes="mx", addressdetails=True, exactly_one=True)
                if loc and resultado_coincide_estado(loc, estado):
                    lat = loc.latitude
                    lon = loc.longitude
                    query_usada = q
                    tipo_ubicacion = tipo
                    aproximada = es_aprox
                    break
                elif loc:
                    # Si encontro algo, pero en otro estado, NO lo aceptamos.
                    # Asi evitamos casos tipo Estado de Mexico cayendo en Chihuahua.
                    continue
            except Exception:
                pass

        nuevos.append({
            "CACHE_KEY": row["CACHE_KEY"],
            "DET": det,
            "LAT": lat,
            "LON": lon,
            "GEOCODE_QUERY": query_usada,
            "TIPO_UBICACION": tipo_ubicacion,
            "UBICACION_APROXIMADA": aproximada
        })

    nuevos = pd.DataFrame(nuevos)

    cache = pd.concat([cache, nuevos], ignore_index=True)
    cache = cache.drop_duplicates(subset=["CACHE_KEY"], keep="last")
    guardar_cache_coords(cache)

    df_base = df_base.drop(
        columns=["LAT", "LON", "GEOCODE_QUERY", "TIPO_UBICACION", "UBICACION_APROXIMADA"],
        errors="ignore"
    )
    df_base = df_base.merge(cache[columnas_cache], on="CACHE_KEY", how="left")

    total_ok = df_base["LAT"].notna().sum()
    total_exactas = ((df_base["LAT"].notna()) & (df_base["UBICACION_APROXIMADA"] == False)).sum()
    total_aprox = ((df_base["LAT"].notna()) & (df_base["UBICACION_APROXIMADA"] == True)).sum()

    print(f"Coordenadas encontradas: {total_ok}")
    print(f" - Por DIRECCION TIENDA: {total_exactas}")
    print(f" - Aproximadas: {total_aprox}")

    return df_base


def aplicar_separacion_puntos(df):
    """
    Separa visualmente puntos que caen en la misma coordenada SIN revolver tiendas.

    IMPORTANTE:
    La version anterior hacia listas nuevos_lat/nuevos_lon por grupos y luego las pegaba
    al DataFrame completo. Como groupby cambia el orden, eso podia asignar la coordenada
    de una tienda a otra tienda. Esta version escribe por indice original.
    """
    df = df.copy()
    df["LAT"] = pd.to_numeric(df["LAT"], errors="coerce")
    df["LON"] = pd.to_numeric(df["LON"], errors="coerce")

    # Por default, cada tienda conserva SU coordenada real.
    df["LAT_MAPA"] = df["LAT"]
    df["LON_MAPA"] = df["LON"]

    df["GRUPO_COORD"] = df["LAT"].round(4).astype(str) + "|" + df["LON"].round(4).astype(str)

    for grupo, temp in df.groupby("GRUPO_COORD", sort=False):
        n = len(temp)
        if n <= 1:
            continue

        radio = 0.015
        for pos, idx in enumerate(temp.index):
            angulo = 2 * math.pi * pos / n
            df.loc[idx, "LAT_MAPA"] = df.loc[idx, "LAT"] + radio * math.sin(angulo)
            df.loc[idx, "LON_MAPA"] = df.loc[idx, "LON"] + radio * math.cos(angulo)

    return df.drop(columns=["GRUPO_COORD"], errors="ignore")

# ============================================================
# CREAR HTML ULTRA AMIGABLE - CON MINIMIZAR Y CERRAR
# ============================================================

def crear_html_mapa(data_tiendas, salida_html):
    tiendas_json = json.dumps(data_tiendas, ensure_ascii=False).replace("</", "<\\/")

    regiones_con_color = {}
    for t in data_tiendas:
        if t["region"] and t["region"] not in regiones_con_color:
            regiones_con_color[t["region"]] = t["colorRegion"]
    regiones_ordenadas = sorted(regiones_con_color.keys(), key=lambda x: str(x))

    opciones_region = ""
    for r in regiones_ordenadas:
        c = regiones_con_color[r]
        opciones_region += f'<option value="{r}" data-color="{c}">{r}</option>\n'

    total_rojas = sum(1 for t in data_tiendas if t["finalizados"] <= 0)
    total_amarillas = sum(1 for t in data_tiendas if t["finalizados"] == 1)
    total_verdes = sum(1 for t in data_tiendas if t["finalizados"] >= 2)

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no"/>
<meta http-equiv="refresh" content="7200"/>
<title>Mapa de Tiendas</title>
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>
<style>
    * {{ box-sizing: border-box; }}
    html, body {{ height: 100%; margin: 0; padding: 0; font-family: 'Segoe UI', Arial, sans-serif; background: #F3F4F6; }}

    #map {{ position: fixed; inset: 0; width: 100%; height: 100%; z-index: 1; }}

    /* ===== PANEL PRINCIPAL ===== */
    .panel {{
        position: fixed;
        top: 10px;
        left: 10px;
        width: 390px;
        max-height: 94vh;
        background: rgba(255,255,255,0.98);
        z-index: 999;
        border-radius: 20px;
        box-shadow: 0 12px 32px rgba(0,0,0,0.22);
        overflow: hidden;
        display: flex;
        flex-direction: column;
        border: 1px solid #E5E7EB;
        transition: transform 0.3s cubic-bezier(0.4, 0, 0.2, 1), opacity 0.3s;
    }}
    .panel.hidden-panel {{
        transform: translateX(calc(-100% - 20px));
        opacity: 0;
        pointer-events: none;
    }}
    .panel.collapsed-mobile {{
        transform: translateY(calc(100% - 50px));
    }}

    /* Boton minimizar */
    .btn-minimize {{
        position: absolute;
        top: 10px;
        right: 10px;
        width: 32px;
        height: 32px;
        border-radius: 10px;
        background: rgba(255,255,255,0.2);
        border: 1px solid rgba(255,255,255,0.3);
        color: white;
        font-size: 18px;
        display: flex;
        align-items: center;
        justify-content: center;
        cursor: pointer;
        z-index: 10;
        transition: background 0.15s;
        padding: 0;
    }}
    .btn-minimize:active {{ background: rgba(255,255,255,0.35); transform: scale(0.92); }}

    /* Barra arrastre */
    .drag-handle {{
        display: none;
        width: 44px;
        height: 5px;
        background: #CBD5E1;
        border-radius: 999px;
        margin: 8px auto 2px auto;
    }}

    .panel-header {{
        padding: 14px 16px 12px 16px;
        background: #1E40AF;
        color: white;
        flex-shrink: 0;
        position: relative;
        cursor: default;
    }}
    .panel-header h1 {{ margin: 0; font-size: 19px; line-height: 1.2; font-weight: 800; padding-right: 36px; }}
    .panel-header p {{ margin: 5px 0 0 0; font-size: 12px; color: #BFDBFE; font-weight: 500; }}

    .panel-body {{
        padding: 12px 14px;
        overflow-y: auto;
        flex: 1;
    }}

    /* Boton flotante para reabrir panel */
    .fab-reopen {{
        position: fixed;
        top: 14px;
        left: 14px;
        width: 50px;
        height: 50px;
        border-radius: 16px;
        background: #1E40AF;
        color: white;
        border: none;
        font-size: 22px;
        display: none;
        align-items: center;
        justify-content: center;
        cursor: pointer;
        z-index: 998;
        box-shadow: 0 6px 20px rgba(30,64,175,0.35);
        transition: transform 0.15s;
    }}
    .fab-reopen:active {{ transform: scale(0.9); }}
    .fab-reopen.visible {{ display: flex; }}

    /* ===== REGION ===== */
    .region-row {{
        display: flex;
        align-items: center;
        gap: 8px;
        margin-bottom: 10px;
    }}
    .region-dot {{
        width: 16px;
        height: 16px;
        border-radius: 999px;
        flex-shrink: 0;
        border: 2px solid white;
        box-shadow: 0 0 0 2px #D1D5DB;
    }}
    select {{
        flex: 1;
        border: 2px solid #D1D5DB;
        border-radius: 14px;
        padding: 12px 14px;
        font-size: 16px;
        font-weight: 700;
        background: white;
        outline: none;
        color: #1F2937;
        cursor: pointer;
    }}
    select:focus {{ border-color: #2563EB; }}

    label {{
        display: block;
        font-size: 12px;
        font-weight: 800;
        color: #374151;
        margin-bottom: 5px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }}
    input[type="text"] {{
        width: 100%;
        border: 2px solid #D1D5DB;
        border-radius: 14px;
        padding: 13px 16px;
        font-size: 17px;
        font-weight: 700;
        background: white;
        outline: none;
        color: #1F2937;
        margin-bottom: 8px;
    }}
    input[type="text"]:focus {{ border-color: #2563EB; box-shadow: 0 0 0 4px rgba(37,99,235,0.12); }}
    input::placeholder {{ color: #9CA3AF; font-weight: 500; }}

    .btn-row {{
        display: grid;
        grid-template-columns: 2fr 1fr;
        gap: 8px;
        margin-bottom: 10px;
    }}
    button {{
        border: none;
        border-radius: 14px;
        padding: 14px;
        font-weight: 800;
        font-size: 15px;
        cursor: pointer;
        transition: transform 0.05s;
    }}
    button:active {{ transform: scale(0.97); }}
    .btn-primary {{
        background: #2563EB;
        color: white;
        box-shadow: 0 4px 12px rgba(37,99,235,0.25);
    }}
    .btn-clean {{
        background: #E5E7EB;
        color: #374151;
    }}

    /* ===== RESUMEN ===== */
    .summary-grid {{
        display: grid;
        grid-template-columns: 1fr 1fr 1fr;
        gap: 6px;
        margin-bottom: 10px;
    }}
    .summary-card {{
        border-radius: 16px;
        padding: 10px 6px;
        text-align: center;
        color: white;
        font-weight: 800;
    }}
    .summary-card.rojo {{ background: #DC2626; }}
    .summary-card.amarillo {{ background: #F59E0B; color: #111827; }}
    .summary-card.verde {{ background: #16A34A; }}
    .summary-card .big {{
        font-size: 28px;
        display: block;
        line-height: 1;
        margin-bottom: 3px;
    }}
    .summary-card .small {{
        font-size: 10px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        opacity: 0.9;
    }}

    .help-box {{
        background: #EFF6FF;
        border: 1px solid #BFDBFE;
        border-radius: 14px;
        padding: 10px 12px;
        font-size: 12px;
        color: #1E40AF;
        line-height: 1.5;
        margin-bottom: 10px;
        font-weight: 600;
    }}
    .help-box b {{ color: #1E3A8A; }}

    .count-box {{
        display: flex;
        justify-content: space-between;
        align-items: center;
        background: #F3F4F6;
        border: 1px solid #E5E7EB;
        color: #374151;
        padding: 10px 12px;
        border-radius: 14px;
        font-size: 14px;
        font-weight: 700;
        margin-bottom: 8px;
    }}
    .store-list {{
        max-height: 220px;
        overflow-y: auto;
        padding-right: 2px;
    }}
    .store-card {{
        background: white;
        border: 2px solid #E5E7EB;
        border-left-width: 6px;
        border-radius: 16px;
        padding: 11px 12px;
        margin-bottom: 7px;
        cursor: pointer;
        transition: all 0.1s;
    }}
    .store-card:active {{ transform: scale(0.98); background: #F9FAFB; }}
    .store-title {{
        font-size: 15px;
        font-weight: 800;
        color: #111827;
        margin-bottom: 3px;
    }}
    .store-sub {{
        font-size: 12px;
        color: #6B7280;
        font-weight: 600;
        margin-bottom: 6px;
    }}
    .pill-row {{
        display: flex;
        flex-wrap: wrap;
        gap: 5px;
        align-items: center;
    }}
    .pill {{
        font-size: 12px;
        padding: 5px 10px;
        border-radius: 999px;
        background: #F3F4F6;
        color: #374151;
        font-weight: 700;
    }}
    .pill.semaforo {{
        color: white;
        font-weight: 800;
    }}
    .pill.semaforo.amarillo {{ background: #F59E0B; color: #111827; }}

    /* ===== TARJETA DE INFO ===== */
    .info-card {{
        position: fixed;
        right: 12px;
        bottom: 12px;
        width: 310px;
        background: rgba(255,255,255,0.98);
        z-index: 999;
        border-radius: 20px;
        box-shadow: 0 12px 32px rgba(0,0,0,0.22);
        border: 1px solid #E5E7EB;
        overflow: hidden;
        transition: transform 0.3s cubic-bezier(0.4, 0, 0.2, 1), opacity 0.3s;
    }}
    .info-card.hidden-card {{
        transform: translateY(calc(100% + 20px));
        opacity: 0;
        pointer-events: none;
    }}

    .btn-close-info {{
        position: absolute;
        top: 10px;
        right: 10px;
        width: 30px;
        height: 30px;
        border-radius: 10px;
        background: rgba(255,255,255,0.2);
        border: 1px solid rgba(255,255,255,0.3);
        color: white;
        font-size: 16px;
        display: flex;
        align-items: center;
        justify-content: center;
        cursor: pointer;
        z-index: 10;
        transition: background 0.15s;
        padding: 0;
    }}
    .btn-close-info:active {{ background: rgba(255,255,255,0.35); transform: scale(0.92); }}

    .info-top {{
        padding: 14px 16px;
        background: #1E40AF;
        color: white;
        position: relative;
    }}
    .info-top h2 {{ margin: 0; font-size: 16px; font-weight: 800; padding-right: 36px; }}
    .info-body {{ padding: 14px 16px; font-size: 14px; color: #111827; }}
    .big-number-row {{
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 10px;
        margin-top: 10px;
    }}
    .big-number {{
        background: #F9FAFB;
        border: 2px solid #E5E7EB;
        border-radius: 16px;
        padding: 12px 8px;
        text-align: center;
    }}
    .big-number b {{
        display: block;
        font-size: 30px;
        line-height: 1;
        margin-bottom: 4px;
        color: #1E40AF;
    }}
    .big-number span {{
        font-size: 11px;
        color: #6B7280;
        font-weight: 700;
        text-transform: uppercase;
    }}
    .status-box {{
        margin-top: 12px;
        padding: 14px;
        border-radius: 16px;
        font-weight: 800;
        text-align: center;
        font-size: 16px;
    }}
    .status-box.rojo {{ background: #FEE2E2; color: #991B1B; }}
    .status-box.amarillo {{ background: #FEF3C7; color: #92400E; }}
    .status-box.verde {{ background: #DCFCE7; color: #166534; }}

    /* Tooltip y controles */
    .leaflet-tooltip {{
        border-radius: 14px !important;
        padding: 10px !important;
        border: 1px solid #E5E7EB !important;
        box-shadow: 0 8px 20px rgba(0,0,0,0.15) !important;
        font-family: 'Segoe UI', Arial, sans-serif !important;
        font-size: 13px !important;
    }}
    .leaflet-control-zoom {{
        margin-top: 12px !important;
        margin-right: 12px !important;
    }}

    .store-list::-webkit-scrollbar {{ width: 5px; }}
    .store-list::-webkit-scrollbar-track {{ background: transparent; }}
    .store-list::-webkit-scrollbar-thumb {{ background: #D1D5DB; border-radius: 999px; }}

    /* ===== CELULAR ===== */
    @media (max-width: 768px) {{
        .panel {{
            top: auto;
            left: 0;
            right: 0;
            bottom: 0;
            width: 100%;
            max-height: 58vh;
            border-radius: 24px 24px 0 0;
        }}
        .panel.hidden-panel {{
            transform: translateY(120%);
        }}
        .drag-handle {{ display: block; }}
        .panel-header {{
            padding: 4px 16px 10px 16px;
            cursor: pointer;
        }}
        .panel-header h1 {{ font-size: 16px; }}
        .panel-body {{ padding: 10px 14px; }}
        select, input[type="text"] {{ font-size: 16px; padding: 12px 14px; }}
        button {{ padding: 13px; font-size: 14px; }}
        .summary-card .big {{ font-size: 24px; }}
        .store-list {{ max-height: 130px; }}
        .store-card {{ padding: 10px; }}
        .btn-minimize {{
            width: 28px;
            height: 28px;
            font-size: 16px;
            top: 8px;
            right: 8px;
        }}

        .info-card {{
            left: 10px;
            right: 10px;
            bottom: 10px;
            width: auto;
            border-radius: 18px;
        }}
        .info-top {{ padding: 12px 14px; }}
        .info-top h2 {{ font-size: 15px; }}
        .info-body {{ padding: 12px 14px; }}
        .big-number b {{ font-size: 26px; }}

        .fab-reopen {{
            top: auto;
            bottom: 20px;
            left: 50%;
            transform: translateX(-50%);
            width: 56px;
            height: 56px;
            border-radius: 18px;
            font-size: 24px;
        }}
        .fab-reopen.visible {{ display: flex; }}
    }}
</style>
</head>

<body>

<div id="map"></div>

<!-- Boton flotante para reabrir panel -->
<button class="fab-reopen" id="fabReopen" onclick="mostrarPanel()" title="Abrir menu">
    &#9776;
</button>

<!-- ===== PANEL PRINCIPAL ===== -->
<div class="panel" id="mainPanel">
    <div class="drag-handle"></div>
    <div class="panel-header" onclick="togglePanelMobile()">
        <button class="btn-minimize" onclick="event.stopPropagation(); ocultarPanel()" title="Minimizar">
            &#8211;
        </button>
        <h1>Mapa de Tiendas</h1>
        <p>Toca un punto para ver detalles</p>
    </div>

    <div class="panel-body">
        <label>Buscar tienda</label>
        <input id="searchInput" type="text" placeholder="Numero de tienda..."/>

        <div class="btn-row">
            <button class="btn-primary" onclick="buscarTienda()">Buscar</button>
            <button class="btn-clean" onclick="limpiarBusqueda()">Borrar</button>
        </div>

        <label>Ver por zona</label>
        <div class="region-row">
            <div class="region-dot" id="regionDot" style="background:#6B7280;"></div>
            <select id="regionSelect" onchange="cambiarColorRegion()">
                <option value="TODAS">TODAS LAS ZONAS</option>
                {opciones_region}
            </select>
        </div>

        <div class="summary-grid">
            <div class="summary-card rojo">
                <span class="big">{total_rojas}</span>
                <span class="small">Necesitan apoyo</span>
            </div>
            <div class="summary-card amarillo">
                <span class="big">{total_amarillas}</span>
                <span class="small">Empezando</span>
            </div>
            <div class="summary-card verde">
                <span class="big">{total_verdes}</span>
                <span class="small">Van bien</span>
            </div>
        </div>

        <div class="help-box">
            <b>Como leer el mapa:</b><br>
            El <b>circulo de colores</b> dice como va la tienda.<br>
            El <b>borde</b> dice a que zona pertenece.<br>
            Si no encuentra la direccion exacta, la ubicacion queda como <b>aproximada</b>.
        </div>

        <div class="count-box">
            <span>Tiendas encontradas</span>
            <span id="countText">0</span>
        </div>

        <div id="storeList" class="store-list"></div>
    </div>
</div>

<!-- ===== TARJETA DE INFO ===== -->
<div id="infoCard" class="info-card hidden-card">
    <div class="info-top">
        <button class="btn-close-info" onclick="cerrarInfo()" title="Cerrar">
            &#10005;
        </button>
        <h2 id="infoTitle">Toca un punto en el mapa</h2>
    </div>
    <div class="info-body">
        <div id="infoSubtitle" style="color:#6B7280; font-weight:600;">
            Aqui aparece la informacion de la tienda
        </div>

        <div class="big-number-row">
            <div class="big-number">
                <b id="infoActivos">-</b>
                <span>Personas activas</span>
            </div>
            <div class="big-number">
                <b id="infoFinalizados">-</b>
                <span>Tramites listos</span>
            </div>
        </div>

        <div id="infoStatus" class="status-box rojo">
            Toca una tienda para ver
        </div>
    </div>
</div>

<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<script>
const tiendas = {tiendas_json};

let map = L.map("map", {{ zoomControl: false }}).setView([23.6345, -102.5528], 5);
L.control.zoom({{ position: "topright" }}).addTo(map);
L.tileLayer("https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png", {{
    maxZoom: 19, attribution: "Mapa"
}}).addTo(map);

let markersLayer = L.layerGroup().addTo(map);
let markerByDet = {{}};

function esc(texto) {{
    if (texto === null || texto === undefined) return "";
    return String(texto).replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}}

/* ===== MOSTRAR / OCULTAR PANEL ===== */
let panelVisible = true;
let panelColapsadoMobile = false;

function ocultarPanel() {{
    document.getElementById("mainPanel").classList.add("hidden-panel");
    document.getElementById("fabReopen").classList.add("visible");
    panelVisible = false;
}}

function mostrarPanel() {{
    const panel = document.getElementById("mainPanel");
    panel.classList.remove("hidden-panel");
    if (window.innerWidth < 768) {{
        panel.classList.remove("collapsed-mobile");
        panelColapsadoMobile = false;
    }}
    document.getElementById("fabReopen").classList.remove("visible");
    panelVisible = true;
}}

function togglePanelMobile() {{
    if (window.innerWidth >= 768) return;
    const panel = document.getElementById("mainPanel");
    panelColapsadoMobile = !panelColapsadoMobile;
    panel.classList.toggle("collapsed-mobile", panelColapsadoMobile);
}}

/* ===== MOSTRAR / CERRAR INFO ===== */
let infoVisible = false;

function mostrarInfo(t) {{
    document.getElementById("infoTitle").textContent = t.det + " - " + t.tienda;

    const direccionTxt = (t.direccion && String(t.direccion).trim() !== "")
        ? esc(t.direccion)
        : "Sin direccion en base";

    const tipoTxt = t.ubicacionAproximada
        ? "Ubicacion aproximada"
        : "Ubicacion por DIRECCION TIENDA";

    document.getElementById("infoSubtitle").innerHTML =
        "<b>Zona:</b> " + esc(t.region) + "<br>" +
        "<b>Ciudad:</b> " + esc(t.ciudad) + "<br>" +
        "<b>Estado:</b> " + esc(t.estado) + "<br>" +
        "<b>Direccion:</b> " + direccionTxt + "<br>" +
        "<b>Ubicacion:</b> " + esc(tipoTxt);

    document.getElementById("infoActivos").textContent = t.activos;
    document.getElementById("infoFinalizados").textContent = t.finalizados;

    const st = document.getElementById("infoStatus");
    st.textContent = t.mensaje;
    st.className = "status-box";
    if (t.finalizados <= 0) st.classList.add("rojo");
    else if (t.finalizados === 1) st.classList.add("amarillo");
    else st.classList.add("verde");

    document.getElementById("infoCard").classList.remove("hidden-card");
    infoVisible = true;
}}

function cerrarInfo() {{
    document.getElementById("infoCard").classList.add("hidden-card");
    infoVisible = false;
}}

function cambiarColorRegion() {{
    const sel = document.getElementById("regionSelect");
    const dot = document.getElementById("regionDot");
    const opt = sel.options[sel.selectedIndex];
    const color = opt.getAttribute("data-color") || "#6B7280";
    dot.style.background = color;
    dot.style.boxShadow = "0 0 0 2px " + color + "40";
    aplicarFiltros();
}}

function crearTooltip(t) {{
    const direccionTxt = (t.direccion && String(t.direccion).trim() !== "")
        ? esc(t.direccion)
        : "Sin direccion en base";
    const tipoTxt = t.ubicacionAproximada
        ? "Ubicacion aproximada"
        : "Ubicacion por direccion";

    return `
        <div style="font-family:'Segoe UI',Arial; min-width:220px; font-size:13px;">
            <b style="font-size:15px;">${{esc(t.tienda)}}</b><br>
            <span style="color:#6B7280;">Tienda ${{esc(t.det)}} | ${{esc(t.region)}}</span><br>
            <span style="color:#6B7280;">${{direccionTxt}}</span><br>
            <span style="color:#2563EB; font-weight:700;">${{esc(tipoTxt)}}</span><br><br>
            <b>${{t.activos}}</b> personas activas<br>
            <b>${{t.finalizados}}</b> tramites listos<br>
            <b style="color:${{t.colorSemaforo}};">${{esc(t.semaforo)}}</b>
        </div>
    `;
}}

function agregarMarker(t) {{
    const esCelular = window.innerWidth < 768;
    const marker = L.circleMarker([t.lat, t.lon], {{
        radius: esCelular ? 11 : 9,
        color: t.colorRegion,
        weight: 4,
        fillColor: t.colorSemaforo,
        fillOpacity: 0.95,
        opacity: 1
    }});

    marker.bindTooltip(crearTooltip(t), {{ sticky: true, direction: "top" }});

    marker.on("click", () => {{
        mostrarInfo(t);
        map.setView([t.lat, t.lon], Math.max(map.getZoom(), 14), {{animate: true}});
        if (esCelular && !panelColapsadoMobile) {{
            togglePanelMobile();
        }}
    }});

    marker.addTo(markersLayer);
    markerByDet[t.det] = marker;
}}

function tiendaCoincide(t, region, texto) {{
    const txt = texto.trim().toUpperCase();
    let okRegion = region === "TODAS" || t.region === region;
    let okTexto = true;
    if (txt.length > 0) {{
        okTexto = String(t.det).includes(txt) ||
                  String(t.tienda).toUpperCase().includes(txt) ||
                  String(t.ciudad).toUpperCase().includes(txt) ||
                  String(t.estado).toUpperCase().includes(txt) ||
                  String(t.direccion).toUpperCase().includes(txt);
    }}
    return okRegion && okTexto;
}}

function aplicarFiltros() {{
    const region = document.getElementById("regionSelect").value;
    const texto = document.getElementById("searchInput").value;
    markersLayer.clearLayers();
    markerByDet = {{}};

    const visibles = tiendas.filter(t => tiendaCoincide(t, region, texto));
    visibles.forEach(t => agregarMarker(t));
    renderListaTiendas(visibles);
    document.getElementById("countText").textContent = visibles.length;

    if (visibles.length > 0) {{
        const bounds = L.latLngBounds(visibles.map(t => [t.lat, t.lon]));
        map.fitBounds(bounds.pad(0.15));
    }}
}}

function renderListaTiendas(lista) {{
    const cont = document.getElementById("storeList");
    cont.innerHTML = "";
    if (lista.length === 0) {{
        cont.innerHTML = '<div class="store-card"><div class="store-title">No se encontro ninguna tienda</div><div class="store-sub">Prueba con otro numero o zona</div></div>';
        return;
    }}

    const ordenadas = [...lista].sort((a, b) => {{
        if (a.region !== b.region) return a.region.localeCompare(b.region, "es", {{numeric:true}});
        return String(a.det).localeCompare(String(b.det), "es", {{numeric:true}});
    }});

    const limite = 100;
    const mostrar = ordenadas.slice(0, limite);

    mostrar.forEach(t => {{
        const div = document.createElement("div");
        div.className = "store-card";
        div.style.borderLeftColor = t.colorRegion;
        const pillClass = t.finalizados === 1 ? "amarillo" : "";

        div.innerHTML = `
            <div class="store-title">${{esc(t.det)}} - ${{esc(t.tienda)}}</div>
            <div class="store-sub">${{esc(t.region)}} | ${{esc(t.ciudad)}}</div>
            <div class="pill-row">
                <span class="pill">${{t.activos}} activos</span>
                <span class="pill">${{t.finalizados}} listos</span>
                <span class="pill semaforo ${{pillClass}}" style="background:${{t.colorSemaforo}};">
                    ${{esc(t.semaforoCorto)}}
                </span>
            </div>
        `;
        div.onclick = () => enfocarTienda(t.det);
        cont.appendChild(div);
    }});

    if (ordenadas.length > limite) {{
        const aviso = document.createElement("div");
        aviso.className = "store-card";
        aviso.innerHTML = `<div class="store-title">Hay mas tiendas</div><div class="store-sub">Usa el buscador para encontrar una</div>`;
        cont.appendChild(aviso);
    }}
}}

function enfocarTienda(det) {{
    const t = tiendas.find(x => String(x.det) === String(det));
    if (!t) return;
    mostrarInfo(t);
    map.setView([t.lat, t.lon], 14, {{animate: true}});
    setTimeout(() => {{ if (markerByDet[det]) markerByDet[det].openTooltip(); }}, 250);
    if (window.innerWidth < 768 && !panelColapsadoMobile) {{
        togglePanelMobile();
    }}
}}

function buscarTienda() {{
    const texto = document.getElementById("searchInput").value.trim();
    if (!texto) {{ aplicarFiltros(); return; }}
    const region = document.getElementById("regionSelect").value;

    const exacta = tiendas.find(t => String(t.det) === texto && (region === "TODAS" || t.region === region));
    if (exacta) {{ aplicarFiltros(); enfocarTienda(exacta.det); return; }}

    const parecida = tiendas.find(t => String(t.det).includes(texto) && (region === "TODAS" || t.region === region));
    if (parecida) {{ aplicarFiltros(); enfocarTienda(parecida.det); return; }}

    aplicarFiltros();
}}

function limpiarBusqueda() {{
    document.getElementById("searchInput").value = "";
    document.getElementById("regionSelect").value = "TODAS";
    cambiarColorRegion();
    aplicarFiltros();
    document.getElementById("infoTitle").textContent = "Toca un punto en el mapa";
    document.getElementById("infoSubtitle").textContent = "Aqui aparece la informacion de la tienda, direccion y tipo de ubicacion";
    document.getElementById("infoActivos").textContent = "-";
    document.getElementById("infoFinalizados").textContent = "-";
    const st = document.getElementById("infoStatus");
    st.textContent = "Toca una tienda para ver";
    st.className = "status-box rojo";
    cerrarInfo();
}}

document.getElementById("searchInput").addEventListener("keyup", function(e) {{
    if (e.key === "Enter") buscarTienda();
}});

/* ===== GESTOS TOUCH PARA PANEL EN CELULAR ===== */
let touchStartY = 0;
let touchStartX = 0;

const panelEl = document.getElementById("mainPanel");

panelEl.addEventListener("touchstart", e => {{
    if (window.innerWidth >= 768) return;
    touchStartY = e.touches[0].clientY;
    touchStartX = e.touches[0].clientX;
}}, {{ passive: true }});

panelEl.addEventListener("touchmove", e => {{
    if (window.innerWidth >= 768) return;
    const y = e.touches[0].clientY;
    const diffY = y - touchStartY;
    const diffX = Math.abs(e.touches[0].clientX - touchStartX);

    // Solo si el movimiento es vertical predominante
    if (diffX > Math.abs(diffY)) return;

    if (!panelColapsadoMobile && diffY > 60) {{
        panelColapsadoMobile = true;
        panelEl.classList.add("collapsed-mobile");
    }} else if (panelColapsadoMobile && diffY < -50) {{
        panelColapsadoMobile = false;
        panelEl.classList.remove("collapsed-mobile");
    }}
}}, {{ passive: true }});

/* ===== INICIALIZAR ===== */
cambiarColorRegion();
</script>

</body>
</html>"""

    with open(salida_html, "w", encoding="utf-8") as f:
        f.write(html)

# ============================================================
# SUBIR RESULTADO A GITHUB
# ============================================================

def subir_a_github():
    """Hace git add/commit/push del index.html generado."""
    if not AUTO_GIT_PUSH:
        print("ℹ️ AUTO_GIT_PUSH está desactivado. No se sube a GitHub.")
        return

    if not (BASE_DIR / ".git").exists():
        print("⚠️ Esta carpeta no parece ser un repo de Git. No se hizo git push.")
        return

    if shutil.which("git") is None:
        print("⚠️ Git no está disponible en PATH. No se hizo git push.")
        return

    print("\n" + "═" * 55)
    print("     🚀 SUBIENDO MAPA A GITHUB")
    print("═" * 55 + "\n")

    try:
        subprocess.run(["git", "add"] + ARCHIVOS_A_SUBIR_GITHUB, cwd=str(BASE_DIR), check=True)

        status = subprocess.run(
            ["git", "status", "--porcelain"] + ARCHIVOS_A_SUBIR_GITHUB,
            cwd=str(BASE_DIR),
            capture_output=True,
            text=True,
            check=True,
        )

        if not status.stdout.strip():
            print("✅ No hubo cambios en el mapa. No se creó commit.")
            return

        mensaje = "Actualización automática mapa " + datetime.now().strftime("%d/%m/%Y %H:%M")

        subprocess.run(["git", "commit", "-m", mensaje], cwd=str(BASE_DIR), check=True)
        subprocess.run(["git", "push"], cwd=str(BASE_DIR), check=True)

        print("✅ Mapa actualizado en GitHub correctamente.")

    except subprocess.CalledProcessError as e:
        print("❌ No se pudo subir a GitHub.")
        print("Revisa que hayas iniciado sesión en Git y que tengas permisos del repo.")
        print(e)

# ============================================================
# PROCESO PRINCIPAL
# ============================================================

def main():
    ruta_records = descargar_records_hoy() if USAR_RECORDS_AUTOMATICO else None
    ruta_asistencia, ruta_base = obtener_archivos_locales()

    if not ruta_records:
        raise ValueError("❌ USAR_RECORDS_AUTOMATICO está en False, pero este script local espera descargar Records automáticamente.")

    print("\nArchivos detectados:")
    print("Records:   ", ruta_records)
    print("Asistencia:", ruta_asistencia)
    print("Base:      ", ruta_base)

    records = leer_excel_auto(ruta_records)
    asistencia = leer_excel_auto(ruta_asistencia, hoja_preferida="Hoja1")
    base = leer_excel_auto(ruta_base, hoja_preferida="BASE DE TIENDAS INVEX")

    col_det_base = buscar_columna(base, ["NUMERO DE TIENDA/DETERMINANTE", "DETERMINANTE"])
    col_region_base = buscar_columna(base, ["REGION", "REGIÓN"])
    col_estado_base = buscar_columna(base, ["ESTADO"])
    col_ciudad_base = buscar_columna(base, ["CIUDAD"])
    col_tienda_base = buscar_columna(base, ["TIENDA"])
    col_direccion_base = buscar_columna(base, ["DIRECCION TIENDA", "DIRECCIÓN TIENDA", "DIRECCION", "DIRECCIÓN", "DOMICILIO", "DOMICILIO TIENDA"], obligatorio=False)
    col_estatus_base = buscar_columna(base, ["ESTATUS"], obligatorio=False)

    col_segundo_lugar = buscar_columna(records, ["Segundo lugar de venta", "Lugar de venta", "DETERMINANTE"])
    col_estatus_proceso = buscar_columna(records, ["Estatus del proceso"])
    col_folio_records = buscar_columna(records, ["Folio FICO", "FOLIO"], obligatorio=False)

    col_det_asis = buscar_columna(asistencia, ["DETERMINANTE"])
    col_nombre_asis = buscar_columna(asistencia, ["NOMBRE COMPLETO", "NOMBRE"])
    col_usuario_asis = buscar_columna(asistencia, ["USUARIO"], obligatorio=False)
    col_estatus_asis = buscar_columna(asistencia, ["ESTATUS"])

    print("\nColumnas detectadas correctamente.")

    base = base.copy()
    base["DET"] = base[col_det_base].apply(extraer_determinante)
    base["REGION_N"] = base[col_region_base].apply(limpiar_texto_mostrar)
    base["ESTADO_N"] = base[col_estado_base].apply(limpiar_texto_mostrar)
    base["CIUDAD_N"] = base[col_ciudad_base].apply(limpiar_texto_mostrar)
    base["TIENDA_N"] = base[col_tienda_base].apply(limpiar_texto_mostrar)
    if col_direccion_base:
        base["DIRECCION_TIENDA_N"] = base[col_direccion_base].apply(limpiar_texto_mostrar)
    else:
        base["DIRECCION_TIENDA_N"] = ""
    if col_estatus_base:
        base["ESTATUS_TIENDA_N"] = base[col_estatus_base].apply(limpiar_texto)
    else:
        base["ESTATUS_TIENDA_N"] = ""
    base = base[base["DET"] != ""].copy()
    base = base.drop_duplicates(subset=["DET"], keep="first")
    if SOLO_TIENDAS_OPERANDO and col_estatus_base:
        base = base[base["ESTATUS_TIENDA_N"] == "OPERANDO"].copy()

    records = records.copy()
    records["DET"] = records[col_segundo_lugar].apply(extraer_determinante)
    records["ESTATUS_PROCESO_N"] = records[col_estatus_proceso].apply(limpiar_texto)
    records_finalizados = records[
        (records["DET"] != "") & (records["ESTATUS_PROCESO_N"].isin(ESTATUS_FINALIZADOS))
    ].copy()
    if col_folio_records:
        finalizados_tienda = records_finalizados.groupby("DET")[col_folio_records].nunique().reset_index(name="FINALIZADOS")
    else:
        finalizados_tienda = records_finalizados.groupby("DET").size().reset_index(name="FINALIZADOS")

    asistencia = asistencia.copy()
    asistencia["DET"] = asistencia[col_det_asis].apply(extraer_determinante)
    asistencia["ESTATUS_PERSONA_N"] = asistencia[col_estatus_asis].apply(limpiar_texto)
    id_persona = col_usuario_asis if col_usuario_asis else col_nombre_asis
    asistencia_activos = asistencia[
        (asistencia["DET"] != "") & (asistencia["ESTATUS_PERSONA_N"] == "ACTIVO")
    ].copy()
    activos_tienda = asistencia_activos.groupby("DET")[id_persona].nunique().reset_index(name="ACTIVOS")

    resumen = base.merge(finalizados_tienda, on="DET", how="left")
    resumen = resumen.merge(activos_tienda, on="DET", how="left")
    resumen["FINALIZADOS"] = resumen["FINALIZADOS"].fillna(0).astype(int)
    resumen["ACTIVOS"] = resumen["ACTIVOS"].fillna(0).astype(int)
    resumen["COLOR_SEMAFORO"] = resumen["FINALIZADOS"].apply(color_semaforo)
    resumen["SEMAFORO"] = resumen["FINALIZADOS"].apply(texto_semaforo)
    resumen["SEMAFORO_CORTO"] = resumen["FINALIZADOS"].apply(lambda x: "ROJO" if x <= 0 else ("AMARILLO" if x == 1 else "VERDE"))
    resumen["MENSAJE"] = resumen["FINALIZADOS"].apply(mensaje_simple)

    regiones = sorted(resumen["REGION_N"].dropna().unique(), key=lambda x: str(x))
    color_region = {}
    for i, region in enumerate(regiones):
        color_region[region] = PALETA_REGIONES[i % len(PALETA_REGIONES)]
    resumen["COLOR_REGION"] = resumen["REGION_N"].map(color_region).fillna("#111827")

    resumen = geocodificar_tiendas(resumen)
    resumen = resumen[resumen["LAT"].notna() & resumen["LON"].notna()].copy()
    resumen = aplicar_separacion_puntos(resumen)

    if len(resumen) == 0:
        raise ValueError("No hay tiendas con coordenadas. Revisa direcciones o conexion a internet.")

    data_tiendas = []
    for _, row in resumen.iterrows():
        ubicacion_aprox = row.get("UBICACION_APROXIMADA", True)
        if pd.isna(ubicacion_aprox):
            ubicacion_aprox = True

        data_tiendas.append({
            "det": str(row["DET"]), "region": str(row["REGION_N"]),
            "estado": str(row["ESTADO_N"]), "ciudad": str(row["CIUDAD_N"]),
            "tienda": str(row["TIENDA_N"]),
            "direccion": str(row.get("DIRECCION_TIENDA_N", "")),
            "tipoUbicacion": str(row.get("TIPO_UBICACION", "")),
            "ubicacionAproximada": bool(ubicacion_aprox),
            "geocodeQuery": str(row.get("GEOCODE_QUERY", "")),
            "activos": int(row["ACTIVOS"]),
            "finalizados": int(row["FINALIZADOS"]),
            "semaforo": str(row["SEMAFORO"]), "semaforoCorto": str(row["SEMAFORO_CORTO"]),
            "mensaje": str(row["MENSAJE"]), "colorSemaforo": str(row["COLOR_SEMAFORO"]),
            "colorRegion": str(row["COLOR_REGION"]),
            "lat": float(row["LAT_MAPA"]), "lon": float(row["LON_MAPA"]),
        })

    resumen_region = resumen.groupby("REGION_N").agg(
        TIENDAS=("DET", "nunique"), ACTIVOS=("ACTIVOS", "sum"),
        FINALIZADOS=("FINALIZADOS", "sum"),
        ROJAS=("FINALIZADOS", lambda x: (x <= 0).sum()),
        AMARILLAS=("FINALIZADOS", lambda x: (x == 1).sum()),
        VERDES=("FINALIZADOS", lambda x: (x >= 2).sum()),
    ).reset_index().sort_values("REGION_N")

    print("\nResumen por region:")
    display(resumen_region)

    salida_html = str(SALIDA_HTML_REPO)
    salida_excel = str(SALIDA_EXCEL_LOCAL)

    crear_html_mapa(data_tiendas, salida_html)

    with pd.ExcelWriter(salida_excel, engine="openpyxl") as writer:
        resumen[[
            "REGION_N", "ESTADO_N", "CIUDAD_N", "DET", "TIENDA_N",
            "DIRECCION_TIENDA_N", "TIPO_UBICACION", "UBICACION_APROXIMADA",
            "GEOCODE_QUERY", "ACTIVOS", "FINALIZADOS", "SEMAFORO",
            "LAT", "LON", "LAT_MAPA", "LON_MAPA"
        ]].sort_values(["REGION_N", "DET"]).to_excel(writer, sheet_name="TIENDAS", index=False)
        resumen_region.to_excel(writer, sheet_name="REGIONES", index=False)

    print("\nListo! Archivos generados:")
    print("HTML :", salida_html)
    print("Excel:", salida_excel)

    if EN_COLAB and files is not None:
        files.download(salida_html)
        files.download(salida_excel)
    else:
        print("Archivos guardados en la carpeta actual.")

main()