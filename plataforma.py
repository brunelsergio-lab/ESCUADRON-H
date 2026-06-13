# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font
from datetime import datetime
import os
from io import BytesIO
from html import escape
import streamlit.components.v1 as components

LIVE_SEARCH_COMPONENT = components.declare_component(
    "live_search_input",
    path=os.path.join(os.path.dirname(__file__), "live_search_component")
)

def live_search_input(label, placeholder, key):
    value = LIVE_SEARCH_COMPONENT(
        label=label,
        placeholder=placeholder,
        value=st.session_state.get(key, ""),
        key=f"{key}_live",
        default=st.session_state.get(key, "")
    )
    value = value or ""
    st.session_state[key] = value
    return value

def excel_bytes(wb):
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

DIAS_SEMANA = {
    0: "Lunes",
    1: "Martes",
    2: "Miércoles",
    3: "Jueves",
    4: "Viernes",
    5: "Sábado",
    6: "Domingo",
}

ESTADOS_AUSENCIA = {"AUSENTE", "ART", "DAF", "LES", "SSD", "LAO", "AUTORIZADO", "DESCANSO DE GUARDIA"}
ESTADOS_GUARDIA_DIURNA = {"ENTRANTE GUARDIA DIURNA", "SERVICIO DE ARMAS DIURNA"}
ESTADOS_GUARDIA_NOCTURNA = {"ENTRANTE GUARDIA NOCTURNA", "SERVICIO DE ARMAS NOCTURNA"}
AMBITOS_NOVEDAD = {
    "AUSENTE": "Ausente",
    "INSTITUTO": "Presente en instituto",
    "ESCUADRON": "Presente en escuadrón",
}

def ambito_por_defecto(estado):
    if estado in ESTADOS_AUSENCIA:
        return "AUSENTE"
    if estado == "PRESENTE EN INSTITUTO":
        return "INSTITUTO"
    if estado == "PRESENTE EN ESCUADRÓN":
        return "ESCUADRON"
    if estado == "COMISIÓN":
        return "INSTITUTO"
    return "ESCUADRON"

def ambito_efectivo(novedad):
    estado = novedad.get('estado', '')
    ambito_guardado = novedad.get('ambito')
    ambito_estado = ambito_por_defecto(estado)
    if estado in {"AUSENTE", "PRESENTE EN INSTITUTO", "PRESENTE EN ESCUADRÓN"}:
        return ambito_estado
    return ambito_guardado or ambito_estado

def es_tercer_anio(valor):
    return "III" in str(valor).upper() or "TERCER" in str(valor).upper()

def es_aop(valor):
    texto = str(valor).upper()
    return "AOP" in texto or "CAO" in texto or "AUXILIAR" in texto

def numero_letras(n):
    mapa = {
        0: "CERO", 1: "UN", 2: "DOS", 3: "TRES", 4: "CUATRO", 5: "CINCO",
        6: "SEIS", 7: "SIETE", 8: "OCHO", 9: "NUEVE", 10: "DIEZ",
        11: "ONCE", 12: "DOCE", 13: "TRECE", 14: "CATORCE", 15: "QUINCE",
        16: "DIECISÉIS", 17: "DIECISIETE", 18: "DIECIOCHO", 19: "DIECINUEVE",
        20: "VEINTE"
    }
    return mapa.get(int(n), str(n))

def formatear_lista_novedades(novedades, estado, curso_fn):
    filtradas = [n for n in novedades if n['estado'] == estado and curso_fn(n.get('grado', ''))]
    if not filtradas:
        return ".-"
    lineas = []
    for idx, nov in enumerate(filtradas, 1):
        grado = "ASP III AÑO" if es_tercer_anio(nov.get('grado', '')) else "ASP I"
        detalle = f" {nov.get('detalle', '').strip()}" if nov.get('detalle') else ""
        lineas.append(f"{idx}. {grado} {nov['nombre']}{detalle} (D: {nov['fecha_ini']}, H: {nov['fecha_fin']})")
    return "\n".join(lineas)

def formatear_servicio(novedades, estados, curso_fn, titulo):
    filtradas = [n for n in novedades if n['estado'] in estados and curso_fn(n.get('grado', ''))]
    if not filtradas:
        return f"▫️ {titulo}:"
    plural = "ASPIRANTES" if len(filtradas) != 1 else "ASPIRANTE"
    lineas = [f"▫️ {titulo}: {numero_letras(len(filtradas))} ({len(filtradas)}) {plural}."]
    for idx, nov in enumerate(filtradas, 1):
        grado = "ASP III AÑO" if es_tercer_anio(nov.get('grado', '')) else "ASP I"
        detalle = f" {nov.get('detalle', '').strip()}" if nov.get('detalle') else ""
        lineas.append(f"{idx}. {grado} {nov['nombre']}{detalle}")
    return "\n".join(lineas)

def generar_minuta_informativa():
    fecha_minuta = st.session_state.fecha_reporte.strftime('%d%b%y').upper()
    novedades = st.session_state.novedades_lista

    df_tercero = df[df['GRADO'].map(es_tercer_anio)]
    df_aop = df[df['GRADO'].map(es_aop)]
    ausentes_tercero = set(df_tercero['ORDEN_LIMP']) & total_ausentes
    ausentes_aop = set(df_aop['ORDEN_LIMP']) & total_ausentes
    formados_tercero = len(df_presentes_primera[df_presentes_primera['GRADO'].map(es_tercer_anio)])
    formados_aop = len(df_presentes_primera[df_presentes_primera['GRADO'].map(es_aop)])

    lineas = [
        f'MINUTA INFORMATIVA DEL ESCUADRÓN H "CABO MARCELO GODOY" DEL DÍA {fecha_minuta}',
        "",
        f"FE: {TOTAL_ESCUADRON}",
        f"P: {disponibles}",
        f"A: {len(total_ausentes)}",
        f"FORMADOS A PRIMERA OBLIGACIÓN: {primera_total}",
        "",
        "✅ CURSO DE TERCER AÑO",
        "",
        f"FE: {len(df_tercero)}",
        f"P: {len(df_tercero) - len(ausentes_tercero)}",
        f"A: {len(ausentes_tercero)}",
        f"FORMADOS PRIMERA OBLIGACIÓN: {formados_tercero}",
        "",
        "OBS:",
        "",
        "▫️ INGRESO HORARIO DIFERENCIADO:",
        "",
        formatear_servicio(novedades, ESTADOS_GUARDIA_DIURNA, es_tercer_anio, "SERVICIO DE ARMAS DIURNA"),
        "",
        formatear_servicio(novedades, ESTADOS_GUARDIA_NOCTURNA | {"DESCANSO DE GUARDIA"}, es_tercer_anio, "DESCANSO DE SERVICIO DE ARMAS NOCTURNO"),
        "",
        "NOVEDADES SANITARIAS:",
        "",
        "▫️ SIN SERVICIO EN DOMICILIO:",
        formatear_lista_novedades(novedades, "SSD", es_tercer_anio),
        "",
        "▫️ ART:",
        formatear_lista_novedades(novedades, "ART", es_tercer_anio),
        "",
        "▫️ DAF:",
        formatear_lista_novedades(novedades, "DAF", es_tercer_anio),
        "",
        "▫️ AUTORIZADO:",
        formatear_lista_novedades(novedades, "AUTORIZADO", es_tercer_anio),
        "",
        "✅ CURSO AUXILIAR OPERATIVO",
        f"FE: {len(df_aop)}",
        f"P: {len(df_aop) - len(ausentes_aop)}",
        f"A: {len(ausentes_aop)}",
        f"FORMADOS PRIMERA OBLIGACIÓN: {formados_aop}",
        "",
        "OBS:",
        "",
        "▫️ INGRESO EN HORARIO DIFERENCIAL:.-",
        "",
        formatear_servicio(novedades, ESTADOS_GUARDIA_DIURNA, es_aop, "SERVICIO DE ARMAS DIURNO"),
        "",
        formatear_servicio(novedades, ESTADOS_GUARDIA_NOCTURNA | {"DESCANSO DE GUARDIA"}, es_aop, "DESCANSO DE SERVICIO DE ARMAS NOCTURNO"),
        "",
        "▫️ ART",
        formatear_lista_novedades(novedades, "ART", es_aop),
        "",
        "▫️ SIN SERVICIO EN DOMICILIO:",
        formatear_lista_novedades(novedades, "SSD", es_aop),
        "",
        "▫️ LAO: (A CUENTA DE LAO)",
        formatear_lista_novedades(novedades, "LAO", es_aop),
        "",
        "▫️ LES:",
        formatear_lista_novedades(novedades, "LES", es_aop),
        "",
        "▫️ DAF:",
        formatear_lista_novedades(novedades, "DAF", es_aop),
    ]
    return "\n".join(lineas)

def normalizar_aula(aula):
    return str(aula).strip().upper().replace(" ", "")

def cargar_horarios_txt(path):
    if not os.path.exists(path):
        return 0

    df_horarios = pd.read_csv(path, encoding="utf-8-sig")
    df_horarios.columns = df_horarios.columns.str.strip()
    df_horarios = df_horarios.rename(columns={
        "Aula": "aula",
        "Dia": "dia",
        "Entrada_Mañana": "ent_m",
        "Salida_Mañana": "sal_m",
        "Entrada_Tarde": "ent_t",
        "Salida_Tarde": "sal_t",
    })

    required = {"aula", "dia", "ent_m", "sal_m", "ent_t", "sal_t"}
    if not required.issubset(set(df_horarios.columns)):
        return 0

    total = 0
    for _, row in df_horarios.iterrows():
        guardar_horarios_dia(normalizar_aula(row["aula"]), str(row["dia"]).strip(), {
            "ent_m": str(row["ent_m"]).strip(),
            "sal_m": str(row["sal_m"]).strip(),
            "ent_t": str(row["ent_t"]).strip(),
            "sal_t": str(row["sal_t"]).strip(),
        })
        total += 1
    return total

# 🔹 IMPORTS CORREGIDOS (NO OMITIR NINGUNA FUNCIÓN)
from db_manager import (
    init_db,
    obtener_novedades, agregar_novedad, actualizar_novedad, eliminar_novedad, vaciar_novedades,
    obtener_estado_aulas, guardar_estado_aula,
    obtener_almuerzo, agregar_almuerzo, quitar_almuerzo,
    obtener_horarios, guardar_horarios, obtener_horarios_dia, guardar_horarios_dia,
    obtener_asistencia, actualizar_asistencia,
    obtener_contacto, obtener_todos_contactos, guardar_contacto,
    registrar_movimiento, obtener_movimientos
)
st.set_page_config(page_title="Gestión de Parte Diario - Escuadrón H", layout="wide")

# 🔹 1. CSS PARA ESPACIADO (PEGAR AQUI AL INICIO)
st.markdown("""
<style>
    .main .block-container { padding-top: 2rem !important; padding-bottom: 3rem !important; }
    [data-testid="stMetric"] { margin-bottom: 1rem !important; }
    .row-widget.stHorizontal { margin-bottom: 2rem !important; }
    .stDataFrame table th, .stDataFrame table td { padding: 0.7rem 0.9rem !important; font-size: 0.95rem !important; }
    .stButton { margin-top: 1rem !important; margin-bottom: 1rem !important; }
    hr { margin: 2rem 0 !important; border-color: #3a3f47 !important; }
</style>
""", unsafe_allow_html=True)
# 🔹 CSS PARA ESPACIADO Y LEGIBILIDAD AL 100%
st.markdown("""
<style>
    /* Espaciado general del contenedor principal */
    .main .block-container {
        padding-top: 2.5rem !important;
        padding-bottom: 3rem !important;
    }
    /* Espacio entre métricas y secciones */
    [data-testid="stMetric"] {
        margin-bottom: 1.2rem !important;
    }
    /* Separación entre filas horizontales (las dos filas de métricas) */
    .row-widget.stHorizontal {
        margin-bottom: 1.8rem !important;
    }
    /* Tablas más legibles y menos apretadas */
    .stDataFrame table th, .stDataFrame table td {
        padding: 0.6rem 0.9rem !important;
        min-width: 90px !important;
        font-size: 0.95rem !important;
    }
    /* Botones con aire alrededor */
    .stButton {
        margin-top: 1.2rem !important;
        margin-bottom: 1.2rem !important;
    }
    /* Líneas divisorias más visibles */
    hr {
        margin: 1.8rem 0 !important;
        border-color: #3a3f47 !important;
    }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 1. CARGA DE DATOS
# ==============================================================================
@st.cache_data(ttl=300)
def cargar_personal():
    nombre_archivo = "alumnos.csv"
    try:
        if os.path.exists(nombre_archivo):
            df = pd.read_csv(nombre_archivo, sep=";", encoding="utf-8")
            df.columns = df.columns.str.strip().str.upper()
            df['ORDEN_LIMP'] = pd.to_numeric(df['ORDEN'], errors='coerce')
            df['NOMBRE_COMPLETO'] = df['NOMBRE'].astype(str).str.strip().str.upper()
            df['DNI'] = df['DNI'].astype(str).str.strip()
            df['CE'] = df['CE'].astype(str).str.strip()
            df['GRADO'] = df['CURSO'].astype(str).str.strip().str.upper()
            df['AULA'] = df['AULA'].astype(str).str.strip()
            df = df.dropna(subset=['ORDEN_LIMP'])
            return df[['ORDEN_LIMP', 'AULA', 'GRADO', 'NOMBRE_COMPLETO', 'DNI', 'CE']].sort_values('ORDEN_LIMP')
        else:
            st.error("No se encontró 'alumnos.csv' en la carpeta del proyecto.")
            return pd.DataFrame()
    except Exception as e:
        st.error(f"Error al cargar datos: {e}")
        return pd.DataFrame()

# ==============================================================================
# 2. INICIALIZACIÓN & DB
# ==============================================================================

# Inicializar base de datos
if 'db_iniciada' not in st.session_state:
    init_db()
    st.session_state.db_iniciada = True

if 'horarios_txt_importado' not in st.session_state:
    st.session_state.horarios_txt_importado = cargar_horarios_txt(r"C:\Users\admin\Desktop\horarios.txt")

df = cargar_personal()
if df.empty:
    st.stop()

TOTAL_ESCUADRON = len(df)
AULAS_UNICAS = sorted(df['AULA'].unique())

# Fecha del reporte
if 'fecha_reporte' not in st.session_state:
    st.session_state.fecha_reporte = datetime.now().date()
FECHA_STR = st.session_state.fecha_reporte.isoformat()

# Novedades
if 'novedades_lista' not in st.session_state:
    st.session_state.novedades_lista = obtener_novedades(FECHA_STR)

# Estado de aulas
if 'estado_aulas' not in st.session_state:
    db_estado = obtener_estado_aulas(FECHA_STR)
    st.session_state.estado_aulas = {}
    for aula in AULAS_UNICAS:
        aula_data = db_estado.get(aula, {})
        st.session_state.estado_aulas[aula] = {
            "estado_m": aula_data.get("estado_m", "EN INSTITUTO"),
            "estado_t": aula_data.get("estado_t", "EN INSTITUTO"),
            "salida_m": aula_data.get("salida_m"),
            "salida_t": aula_data.get("salida_t"),
            "ubicacion_m": aula_data.get("ubicacion_m", "EN AULA"),  # 👈 NUEVO
            "ubicacion_t": aula_data.get("ubicacion_t", "EN AULA")   # 👈 NUEVO
        }

# Lista de almuerzo
if 'lista_almuerzo' not in st.session_state:
    st.session_state.lista_almuerzo = obtener_almuerzo(FECHA_STR)

# Horarios config
if 'horarios_config' not in st.session_state:
    dia_reporte = DIAS_SEMANA[st.session_state.fecha_reporte.weekday()]
    db_hor = obtener_horarios_dia(dia_reporte) or obtener_horarios()
    st.session_state.horarios_config = {}
    for aula in AULAS_UNICAS:
        st.session_state.horarios_config[aula] = db_hor.get(normalizar_aula(aula), db_hor.get(aula, {
            "ent_m": "06:00", "sal_m": "12:00", "ent_t": "13:00", "sal_t": "19:00"
        }))

# Asistencia diaria
if 'estado_asistencia' not in st.session_state:
    st.session_state.estado_asistencia = obtener_asistencia(FECHA_STR)

# Variables de control UI (¡ESTAS SON LAS QUE FALTABAN!)
if 'editando_idx' not in st.session_state:
    st.session_state.editando_idx = None

if 'sel_nov' not in st.session_state:
    st.session_state.sel_nov = None

def log_movimiento(modulo, accion, orden=None, nombre=None, aula=None, detalle=""):
    registrar_movimiento(FECHA_STR, modulo, accion, orden, nombre, aula, detalle)

def limpiar_form_novedad():
    for key in ("sel_estado", "sel_ambito", "txt_detalle"):
        if key in st.session_state:
            del st.session_state[key]

if st.session_state.pop("limpiar_form_novedad_pendiente", False):
    limpiar_form_novedad()

# ==============================================================================
# 3. MÉTRICAS EN TIEMPO REAL (CON SINCRONIZACIÓN AUTOMÁTICA)
# ==============================================================================

# # 🔹 1. RECARGAR DATOS DESDE DB (Prioridad a session_state)
st.session_state.novedades_lista = obtener_novedades(FECHA_STR)
st.session_state.lista_almuerzo = obtener_almuerzo(FECHA_STR)

# Asistencia: NO sobrescribir si ya hay datos (prioridad a cambios manuales)
if not st.session_state.estado_asistencia:
    st.session_state.estado_asistencia = obtener_asistencia(FECHA_STR)
else:
    # Sincronizar solo registros nuevos de DB
    db_asistencia = obtener_asistencia(FECHA_STR)
    for orden, estado in db_asistencia.items():
        if orden not in st.session_state.estado_asistencia:
            st.session_state.estado_asistencia[orden] = estado

# 🔹 2. CÁLCULO DE MÉTRICAS (Se recalcula en cada interacción)
ambito_por_orden = {
    n['orden']: ambito_efectivo(n)
    for n in st.session_state.novedades_lista
}
ausentes_fijos = {orden for orden, ambito in ambito_por_orden.items() if ambito == "AUSENTE"}
presentes_instituto_novedad = {orden for orden, ambito in ambito_por_orden.items() if ambito in {"INSTITUTO", "ESCUADRON"}}
presentes_escuadron_novedad = {orden for orden, ambito in ambito_por_orden.items() if ambito == "ESCUADRON"}
ausentes_manuales = {orden for orden, estado in st.session_state.estado_asistencia.items() if estado == "AUSENTE"}
presentes_instituto_manuales = {
    orden for orden, estado in st.session_state.estado_asistencia.items()
    if estado in {"PRESENTE", "PRESENTE EN INSTITUTO", "PRESENTE EN ESCUADRÓN"}
}
presentes_escuadron_manuales = {
    orden for orden, estado in st.session_state.estado_asistencia.items()
    if estado in {"PRESENTE", "PRESENTE EN ESCUADRÓN"}
}

total_ausentes = (ausentes_fijos | ausentes_manuales) - presentes_instituto_manuales

en_instituto = 0
fuera_por_aula = 0
presentes_escuadron = 0
for _, row in df.iterrows():
    orden = row['ORDEN_LIMP']
    aula = row['AULA']
    if orden in total_ausentes: continue
    en_aula_instituto = st.session_state.estado_aulas.get(aula, {}).get('estado_m', 'EN INSTITUTO') == 'EN INSTITUTO'
    if orden in presentes_instituto_novedad or orden in presentes_instituto_manuales or en_aula_instituto:
        en_instituto += 1
        if orden in presentes_escuadron_novedad or orden in presentes_escuadron_manuales or (orden not in ambito_por_orden and orden not in presentes_instituto_manuales and en_aula_instituto):
            presentes_escuadron += 1
    else:
        fuera_por_aula += 1

disponibles = TOTAL_ESCUADRON - len(total_ausentes)
total_fuera = fuera_por_aula + len(total_ausentes)

df_presentes_primera = df[
    (~df['ORDEN_LIMP'].isin(total_ausentes)) &
    (
        df['ORDEN_LIMP'].isin(presentes_escuadron_novedad) |
        df['ORDEN_LIMP'].isin(presentes_escuadron_manuales) |
        ((~df['ORDEN_LIMP'].isin(ambito_por_orden.keys())) &
         (~df['ORDEN_LIMP'].isin(presentes_instituto_manuales)) &
         (df['AULA'].map(lambda aula: st.session_state.estado_aulas.get(aula, {}).get('estado_m', 'EN INSTITUTO')) == 'EN INSTITUTO'))
    )
]
primera_total = len(df_presentes_primera)
primera_tercer_anio = len(df_presentes_primera[df_presentes_primera['GRADO'].map(es_tercer_anio)])
primera_aop = len(df_presentes_primera[df_presentes_primera['GRADO'].map(es_aop)])

ubicacion_dist = {"EN AULA": [], "URF": [], "EDUCACIÓN FÍSICA": [], "EN INSTITUTO": []}
for aula in AULAS_UNICAS:
    cfg = st.session_state.estado_aulas[aula]
    if cfg['estado_m'] == 'EN INSTITUTO':
        ubic = cfg.get('ubicacion_m', 'EN AULA')
        if ubic in ubicacion_dist:
            ubicacion_dist[ubic].append(len(df[df['AULA'] == aula]))

# 🔹 3. CSS Y RENDERIZADO (Métricas fijas arriba)
st.markdown("""
<style>
    .sticky-bar { position: sticky !important; top: 0 !important; z-index: 999 !important; 
                  background: #161B15 !important; padding: 10px 0 12px 0 !important; 
                  border-bottom: 2px solid #C4A000 !important; margin-bottom: 15px !important; 
                  box-shadow: 0 4px 8px rgba(0,0,0,0.6) !important; }
    .header-title { text-align: center !important; font-size: 1.15rem !important; font-weight: 900 !important; 
                    color: #FFFFFF !important; letter-spacing: 2.5px !important; margin-bottom: 8px !important; 
                    text-transform: uppercase !important; text-shadow: 1px 1px 3px #000000 !important; }
    [data-testid="stMetric"] { padding: 2px 4px !important; }
    [data-testid="stMetricValue"] { font-size: 1.05rem !important; font-weight: 700 !important; color: #FFFFFF !important; }
    [data-testid="stMetricLabel"] { font-size: 0.65rem !important; color: #A8B099 !important; text-transform: uppercase !important; letter-spacing: 1px !important; }
    [data-testid="stMetricDelta"] { font-size: 0.65rem !important; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="header-title">ESCUADRÓN H "CABO MARCELO GODOY"</div>', unsafe_allow_html=True)


en_aula_count = sum(ubicacion_dist.get('EN AULA', []))
urf_count = sum(ubicacion_dist.get('URF', []))
edfis_key = next((k for k in ubicacion_dist.keys() if 'FISICA' in k.upper() or 'FÍSICA' in k.upper()), 'EDUCACIÓN FÍSICA')
edfis_count = sum(ubicacion_dist.get(edfis_key, []))
activ_count = sum(ubicacion_dist.get('EN INSTITUTO', []))

kpis = [
    ("Total", TOTAL_ESCUADRON, "Dotacion registrada", "neutral"),
    ("Presentes", disponibles, f"Ausentes {len(total_ausentes)}", "ok" if disponibles == TOTAL_ESCUADRON else "warn"),
    ("En instituto", en_instituto, "Personal dentro", "ok"),
    ("En escuadrón", presentes_escuadron, "Sin comisión", "ok"),
    ("Fuera", total_fuera, "Ausentes o retirados", "alert" if total_fuera else "neutral"),
    ("1ra oblig.", primera_total, "Formados 06:00", "ok"),
    ("3er año", primera_tercer_anio, "En formación", "neutral"),
    ("AOP", primera_aop, "En formación", "neutral"),
    ("En aula", en_aula_count, f"{len(ubicacion_dist.get('EN AULA', []))} aulas", "neutral"),
    ("URF", urf_count, f"{len(ubicacion_dist.get('URF', []))} aulas", "neutral"),
    ("Ed. fisica", edfis_count, f"{len(ubicacion_dist.get(edfis_key, []))} aulas", "neutral"),
    ("Actividad", activ_count, f"{len(ubicacion_dist.get('EN INSTITUTO', []))} aulas", "neutral"),
]

kpi_html = "".join(
    f'<div class="rrhh-kpi rrhh-kpi-{status}"><span>{label}</span><strong>{value}</strong><small>{caption}</small></div>'
    for label, value, caption, status in kpis
)

def texto_obs_motivos(novedades):
    if not novedades:
        return "Sin novedades"
    motivos = pd.Series([n["estado"] for n in novedades]).value_counts().sort_index()
    partes = []
    for motivo, cantidad in motivos.items():
        label = str(motivo).lower()
        partes.append(f"{int(cantidad)} {label}")
    return ", ".join(partes)

def tarjeta_monitor_novedades(titulo, novedades):
    total = len(novedades)
    obs = texto_obs_motivos(novedades)
    return (
        '<div class="nov-card">'
        f'<div><span>{escape(titulo)}</span><strong>{total}</strong></div>'
        f'<p><b>OBS:</b> {escape(obs)}</p>'
        '</div>'
    )

novedades_ausentes = [n for n in st.session_state.novedades_lista if ambito_efectivo(n) == "AUSENTE"]
novedades_ausentes_tercero = [n for n in novedades_ausentes if es_tercer_anio(n.get("grado", ""))]
novedades_ausentes_aop = [n for n in novedades_ausentes if es_aop(n.get("grado", ""))]

monitor_items_html = (
    tarjeta_monitor_novedades("Ausentes de 3er año", novedades_ausentes_tercero)
    + tarjeta_monitor_novedades("Ausentes AOP", novedades_ausentes_aop)
)

st.markdown(f"""
<style>
    .header-title {{
        display: none !important;
    }}
    .rrhh-panel {{
        background: #111827;
        border: 1px solid rgba(148, 163, 184, 0.22);
        border-radius: 8px;
        padding: 1rem;
        margin: 0.25rem 0 1rem 0;
    }}
    .rrhh-head {{
        display: flex;
        justify-content: space-between;
        gap: 1rem;
        align-items: flex-start;
        margin-bottom: 0.9rem;
    }}
    .rrhh-eyebrow {{
        margin: 0 0 0.25rem 0;
        color: #9CA3AF;
        font-size: 0.76rem;
        letter-spacing: 0.08em;
        text-transform: uppercase;
    }}
    .rrhh-title {{
        margin: 0;
        color: #F9FAFB;
        font-size: 1.2rem;
        line-height: 1.25;
        font-weight: 800;
    }}
    .rrhh-subtitle {{
        margin: 0.25rem 0 0 0;
        color: #CBD5E1;
        font-size: 0.9rem;
    }}
    .rrhh-date {{
        border: 1px solid rgba(148, 163, 184, 0.28);
        border-radius: 999px;
        color: #E5E7EB;
        padding: 0.35rem 0.7rem;
        font-size: 0.82rem;
        white-space: nowrap;
        background: rgba(15, 23, 42, 0.7);
    }}
    .nov-monitor {{
        border: 1px solid rgba(148, 163, 184, 0.22);
        background: rgba(15, 23, 42, 0.72);
        border-radius: 8px;
        padding: 0.75rem;
        margin: 0 0 0.85rem 0;
    }}
    .nov-monitor-head {{
        display: flex;
        justify-content: space-between;
        gap: 0.75rem;
        align-items: center;
        margin-bottom: 0.55rem;
    }}
    .nov-monitor-title {{
        color: #F9FAFB;
        font-size: 0.84rem;
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: 0.06em;
    }}
    .nov-monitor-sub {{
        color: #94A3B8;
        font-size: 0.76rem;
        white-space: nowrap;
    }}
    .nov-monitor-grid {{
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(230px, 1fr));
        gap: 0.5rem;
    }}
    .nov-card {{
        border: 1px solid rgba(148, 163, 184, 0.18);
        background: rgba(255, 255, 255, 0.045);
        border-radius: 8px;
        padding: 0.65rem 0.75rem;
        min-height: 74px;
    }}
    .nov-card > div {{
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 0.75rem;
    }}
    .nov-card span {{
        color: #CBD5E1;
        font-size: 0.74rem;
        font-weight: 700;
        line-height: 1.15;
        text-transform: uppercase;
    }}
    .nov-card strong {{
        color: #FFFFFF;
        font-size: 1.5rem;
        line-height: 1;
    }}
    .nov-card p {{
        margin: 0.45rem 0 0 0;
        color: #94A3B8;
        font-size: 0.78rem;
        line-height: 1.25;
    }}
    .nov-card b {{
        color: #E5E7EB;
    }}
    .rrhh-kpi-grid {{
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(132px, 1fr));
        gap: 0.65rem;
    }}
    .rrhh-kpi {{
        border: 1px solid rgba(148, 163, 184, 0.18);
        background: rgba(255, 255, 255, 0.04);
        border-radius: 8px;
        padding: 0.75rem;
        min-height: 96px;
    }}
    .rrhh-kpi span {{
        display: block;
        color: #A7B0BE;
        font-size: 0.72rem;
        text-transform: uppercase;
        letter-spacing: 0.06em;
    }}
    .rrhh-kpi strong {{
        display: block;
        color: #F9FAFB;
        font-size: 1.7rem;
        line-height: 1.1;
        margin-top: 0.35rem;
    }}
    .rrhh-kpi small {{
        display: block;
        color: #94A3B8;
        font-size: 0.78rem;
        margin-top: 0.3rem;
    }}
    .rrhh-kpi-ok {{
        border-left: 3px solid #22C55E;
    }}
    .rrhh-kpi-warn {{
        border-left: 3px solid #F59E0B;
    }}
    .rrhh-kpi-alert {{
        border-left: 3px solid #EF4444;
    }}
    @media (max-width: 720px) {{
        .rrhh-panel {{
            padding: 0.85rem;
        }}
        .rrhh-head {{
            display: block;
        }}
        .rrhh-date {{
            display: inline-block;
            margin-top: 0.7rem;
        }}
        .nov-monitor {{
            padding: 0.65rem;
            position: sticky;
            top: 0.25rem;
            z-index: 10;
        }}
        .nov-monitor-head {{
            display: block;
            margin-bottom: 0.45rem;
        }}
        .nov-monitor-sub {{
            display: block;
            margin-top: 0.2rem;
        }}
        .nov-monitor-grid {{
            display: flex;
            overflow-x: auto;
            gap: 0.45rem;
            padding-bottom: 0.1rem;
        }}
        .nov-card {{
            min-width: 245px;
            flex: 0 0 auto;
        }}
        .rrhh-kpi-grid {{
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 0.5rem;
        }}
        .rrhh-kpi {{
            min-height: 86px;
            padding: 0.65rem;
        }}
        .rrhh-kpi strong {{
            font-size: 1.45rem;
        }}
    }}
</style>
<section class="rrhh-panel">
    <div class="rrhh-head">
        <div>
            <p class="rrhh-eyebrow">Gestor RRHH - control en instalaciones</p>
            <h1 class="rrhh-title">Escuadron H "Cabo Marcelo Godoy"</h1>
            <p class="rrhh-subtitle">Situacion diaria, presentismo, ubicacion y novedades activas.</p>
        </div>
        <div class="rrhh-date">Parte: {st.session_state.fecha_reporte.strftime('%d/%m/%Y')}</div>
    </div>
    <div class="nov-monitor">
        <div class="nov-monitor-head">
            <div class="nov-monitor-title">Monitor de novedades</div>
            <div class="nov-monitor-sub">{len(st.session_state.novedades_lista)} activa(s)</div>
        </div>
        <div class="nov-monitor-grid">{monitor_items_html}</div>
    </div>
    <div class="rrhh-kpi-grid">{kpi_html}</div>
</section>
""", unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

st.divider()
# 🔹 BOTÓN OCULTO PARA FORZAR SINCRONIZACIÓN (Por si las dudas)
if st.button("🔄 Sincronizar Datos", key="sync_btn", help="Fuerza la recarga desde base de datos"):
    st.session_state.novedades_lista = obtener_novedades(FECHA_STR)
    st.session_state.estado_asistencia = obtener_asistencia(FECHA_STR)
    st.session_state.lista_almuerzo = obtener_almuerzo(FECHA_STR)
    st.success("✅ Datos sincronizados correctamente")
    st.rerun()
# 🔹 BOTÓN DE EMERGENCIA: RESETEAR ASISTENCIA
if st.button("🚨 RESETEAR ASISTENCIA DEL DÍA", key="reset_asistencia", help="Pone a TODOS en PRESENTE"):
    import sqlite3
    conn = sqlite3.connect("parte_diario.db")
    conn.execute("DELETE FROM asistencia_diaria WHERE fecha=?", (FECHA_STR,))
    conn.commit()
    conn.close()
    
    # Limpiar session_state
    st.session_state.estado_asistencia = {}
    st.success("✅ Asistencia reiniciada. Todos en PRESENTE.")
    st.rerun()
# ==============================================================================
# 4. PESTAÑAS
# ==============================================================================
tab_config, tab_nov, tab_seg, tab_alm, tab_plan, tab_res = st.tabs([
    "Día y horarios", "Novedades", "Ubicación", "Racionamiento", "Legajos y contactos", "Reportes"
])

# --- TAB: CONFIGURACIÓN ---
with tab_config:
    st.subheader("Configuración del Día y Horarios")
    fecha_anterior = st.session_state.fecha_reporte
    st.session_state.fecha_reporte = st.date_input("Fecha del Reporte", st.session_state.fecha_reporte)
    dia_reporte = DIAS_SEMANA[st.session_state.fecha_reporte.weekday()]
    if fecha_anterior != st.session_state.fecha_reporte:
        db_hor = obtener_horarios_dia(dia_reporte) or obtener_horarios()
        st.session_state.horarios_config = {}
        for aula in AULAS_UNICAS:
            st.session_state.horarios_config[aula] = db_hor.get(normalizar_aula(aula), db_hor.get(aula, {
                "ent_m": "06:00", "sal_m": "12:00", "ent_t": "13:00", "sal_t": "19:00"
            }))
        st.rerun()
    st.caption(f"Horarios cargados para: {dia_reporte}. Podés editarlos y guardarlos para ese día.")
    st.divider()
    for aula in AULAS_UNICAS:
        cfg = st.session_state.horarios_config[aula]
        with st.expander(f"**{aula}**", expanded=False):
            c1, c2, c3, c4 = st.columns(4)
            cfg['ent_m'] = c1.text_input("Entrada Mañana", cfg['ent_m'], key=f"em_{aula}")
            cfg['sal_m'] = c2.text_input("Salida Mañana", cfg['sal_m'], key=f"sm_{aula}")
            cfg['ent_t'] = c3.text_input("Entrada Tarde", cfg['ent_t'], key=f"et_{aula}")
            cfg['sal_t'] = c4.text_input("Salida Tarde", cfg['sal_t'], key=f"st_{aula}")
    
    if st.button("💾 Guardar configuración de horarios", type="primary"):
        for aula in AULAS_UNICAS:
            guardar_horarios_dia(normalizar_aula(aula), dia_reporte, st.session_state.horarios_config[aula])
            guardar_horarios(normalizar_aula(aula), st.session_state.horarios_config[aula])
        st.success("✅ Horarios guardados en base de datos")
        st.rerun()

# --- TAB: NOVEDADES ---
with tab_nov:
    edit_idx = st.session_state.editando_idx
    es_edicion = edit_idx is not None

    # Validación de seguridad por si la lista cambió mientras editabas
    if es_edicion and not (0 <= edit_idx < len(st.session_state.novedades_lista)):
        st.session_state.editando_idx = None
        st.warning("⚠️ La novedad que editabas cambió. Se reinició el formulario.")
        st.rerun()

    st.subheader("✏️ Editando Novedad" if es_edicion else "➕ Registrar Novedad")

    data = None
    if es_edicion:
        nov = st.session_state.novedades_lista[edit_idx]
        st.info(f"Editando a: **{nov['nombre']}**")
        data = nov
        
        st.divider()
        st.markdown(f"### Control de presencia: **{data['nombre']}**")
        
        orden = data["orden"]
        nombre_asp = data["nombre"]

        col_btn_aus, col_btn_p, col_btn_a = st.columns(3)
        
        with col_btn_aus:
            if st.button("Ausente", type="secondary", use_container_width=True, key="btn_aus_edit"):
                st.session_state.sel_ambito = "AUSENTE"
                st.toast(f"{nombre_asp} preparado como ausente")
                st.rerun()

        with col_btn_p:
            if st.button("Presente en instituto", type="secondary", use_container_width=True, key="btn_pres_inst_edit"):
                st.session_state.sel_ambito = "INSTITUTO"
                st.toast(f"{nombre_asp} preparado como presente en instituto")
                st.rerun()
                
        with col_btn_a:
            if st.button("Presente en escuadrón", type="primary", use_container_width=True, key="btn_pres_esc_edit"):
                st.session_state.sel_ambito = "ESCUADRON"
                st.toast(f"{nombre_asp} preparado como presente en escuadrón")
                st.rerun()
        st.divider()

    else:
        # Lógica de búsqueda y selección (Modo Registro)
        search = live_search_input("Buscar aspirante:", "Nombre, DNI o CE", "search_nov")
        if search.strip() and not st.session_state.sel_nov:
            s = search.strip().upper()
            res = df[
                (df['NOMBRE_COMPLETO'].str.contains(s, na=False)) |
                (df['DNI'].str.contains(s, na=False)) |
                (df['CE'].str.contains(s, na=False))
            ]
            if not res.empty:
                st.markdown("### Resultados de búsqueda")
                for i, (_, r) in enumerate(res.head(5).iterrows()):
                    c1, c2 = st.columns([4, 1])
                    with c1: st.markdown(f"**{r['NOMBRE_COMPLETO']}** | DNI: {r['DNI']} | CE: {r['CE']}")
                    with c2:
                        if st.button("👆 Seleccionar", key=f"sel_{i}"):
                            limpiar_form_novedad()
                            st.session_state.sel_nov = r.to_dict()
                            st.session_state.search_nov = ""
                            st.rerun()

        if st.session_state.sel_nov:
            data = st.session_state.sel_nov
            orden = data["ORDEN_LIMP"]
            nombre_asp = data.get('NOMBRE_COMPLETO', data.get('nombre', 'Aspirante'))
            st.divider()
            st.markdown(f"### Control de presencia: **{nombre_asp}**")

            col_btn_aus, col_btn_p, col_btn_a, col_btn_c = st.columns([2, 2, 2, 1])

            with col_btn_aus:
                if st.button("Ausente", type="secondary", use_container_width=True, key="btn_aus"):
                    st.session_state.sel_ambito = "AUSENTE"
                    st.toast(f"{nombre_asp} preparado como ausente")
                    st.rerun()

            with col_btn_p:
                if st.button("Presente en instituto", type="secondary", use_container_width=True, key="btn_pres_inst"):
                    st.session_state.sel_ambito = "INSTITUTO"
                    st.toast(f"{nombre_asp} preparado como presente en instituto")
                    st.rerun()

            with col_btn_a:
                if st.button("Presente en escuadrón", type="primary", use_container_width=True, key="btn_pres_esc"):
                    st.session_state.sel_ambito = "ESCUADRON"
                    st.toast(f"{nombre_asp} preparado como presente en escuadrón")
                    st.rerun()

            with col_btn_c:
                if st.button("Cambiar", use_container_width=True, help="Cambiar aspirante", key="btn_clear_sel"):
                    limpiar_form_novedad()
                    st.session_state.sel_nov = None
                    st.session_state.search_nov = ""
                    st.rerun()
            st.divider()

    # 🔹 FORMULARIO DE NOVEDAD (Visible si hay data, tanto en edición como en registro)
    if data is not None:
        st.markdown("### ⚙️ Detalles de Novedad")
        c1, c2 = st.columns(2)
        with c1:
            opts = [
    "AUSENTE",
    "ART",
    "DAF",
    "LES",
    "LAO",
    "SSD",
    "COMISIÓN",
    "AUTORIZADO",
    "ENTRANTE GUARDIA DIURNA",
    "ENTRANTE GUARDIA NOCTURNA",
    "DESCANSO DE GUARDIA"
]
            current_estado = data.get('estado', "ART")
            if current_estado not in opts:
                current_estado = "ART"
            if st.session_state.get("sel_estado") not in opts:
                st.session_state.sel_estado = current_estado
            idx_opts = opts.index(current_estado) if current_estado in opts else 0
            est = st.selectbox("Situación:", opts, index=idx_opts, key="sel_estado")
        with c2:
            det = st.text_input("Detalle:", value=data.get('detalle', ''), key="txt_detalle")

        ambito_actual = st.session_state.get("sel_ambito") or data.get('ambito') or ambito_por_defecto(est)
        ambito_keys = list(AMBITOS_NOVEDAD.keys())
        ambito_idx = ambito_keys.index(ambito_actual) if ambito_actual in ambito_keys else 0
        ambito = st.radio(
            "Presencia real:",
            ambito_keys,
            index=ambito_idx,
            format_func=lambda x: AMBITOS_NOVEDAD[x],
            key="sel_ambito",
            horizontal=True
        )

        cf1, cf2 = st.columns(2)
        with cf1:
            if es_edicion:
                try: fi_val = datetime.strptime(data['fecha_ini'], '%d%b%y')
                except: fi_val = datetime.now()
                fi = st.date_input("Desde:", value=fi_val, key="date_ini").strftime('%d%b%y').upper()
            else:
                fi = st.date_input("Desde:", value=datetime.now(), key="date_ini2").strftime('%d%b%y').upper()
        with cf2:
            is_no = (data.get('fecha_fin') == "N/O") if es_edicion else False
            sin_fin = st.checkbox("Sin término", value=is_no, key="chk_sintermino")
            if sin_fin:
                ff = "N/O"
            else:
                if es_edicion and not is_no:
                    try: ff_val = datetime.strptime(data['fecha_fin'], '%d%b%y')
                    except: ff_val = datetime.now()
                    ff = st.date_input("Hasta:", value=ff_val, key="date_fin").strftime('%d%b%y').upper()
                else:
                    ff = st.date_input("Hasta:", value=datetime.now(), key="date_fin2").strftime('%d%b%y').upper()

        b1, b2 = st.columns([3, 1])
        with b1:
            if es_edicion:
                if st.button("💾 Guardar Cambios", type="primary", use_container_width=True, key="btn_save_edit"):
                    nov_id = st.session_state.novedades_lista[edit_idx]['id']
                    actualizar_novedad(nov_id, {"estado": est, "detalle": det.upper(), "fecha_ini": fi, "fecha_fin": ff, "ambito": ambito})
                    log_movimiento("Novedades", "EDITAR NOVEDAD", data.get("orden"), data.get("nombre"), data.get("aula"), f"{est} | {AMBITOS_NOVEDAD.get(ambito, ambito)} | {fi} a {ff} | {det.upper()}")
                    st.session_state.novedades_lista = obtener_novedades(FECHA_STR)
                    st.session_state.editando_idx = None
                    st.session_state.limpiar_form_novedad_pendiente = True
                    st.success("✅ Novedad y asistencia actualizadas")
                    st.rerun()
            else:
                if st.button("💾 Grabar Novedad", use_container_width=True, key="btn_save_new"):
                    nombre_asp = data.get('NOMBRE_COMPLETO', data.get('nombre', 'Aspirante'))
                    agregar_novedad({
                        "orden": int(data["ORDEN_LIMP"]), "grado": data["GRADO"],
                        "nombre": nombre_asp, "dni": data["DNI"], "ce": data["CE"],
                        "aula": data["AULA"], "estado": est, "detalle": det.upper(),
                        "fecha_ini": fi, "fecha_fin": ff, "ambito": ambito
                    })
                    log_movimiento("Novedades", "ALTA NOVEDAD", int(data["ORDEN_LIMP"]), nombre_asp, data["AULA"], f"{est} | {AMBITOS_NOVEDAD.get(ambito, ambito)} | {fi} a {ff} | {det.upper()}")
                    st.session_state.novedades_lista = obtener_novedades(FECHA_STR)
                    st.session_state.sel_nov = None
                    st.session_state.limpiar_form_novedad_pendiente = True
                    st.success(f"✅ Novedad grabada para {nombre_asp}")
                    st.rerun()
        with b2:
            if st.button("🚫 Cancelar", use_container_width=True, key="btn_cancel"):
                st.session_state.limpiar_form_novedad_pendiente = True
                st.session_state.editando_idx = None
                st.session_state.sel_nov = None
                st.rerun()
    else:
        if not es_edicion:
            st.info("🔍 Busca un aspirante para registrar novedad o asistencia.")

    # 📋 LISTA DE NOVEDADES REGISTRADAS
    if st.session_state.novedades_lista:
        st.subheader("📋 Novedades Registradas en la Guardia")
        st.markdown("---")
        for idx, nov in enumerate(st.session_state.novedades_lista):
            col_datos, col_edit, col_borrar = st.columns([6, 1, 1])
            with col_datos:
                badge_color = "red" if nov['estado'] in ['ART','DAF','LES'] else "orange"
                st.markdown(f"**{nov['nombre']}** <span style='color:{badge_color};font-weight:bold'>[{nov['estado']}]</span>  |  DNI: {nov['dni']}  |  CE: {nov['ce']}", unsafe_allow_html=True)
                ambito_lbl = AMBITOS_NOVEDAD.get(ambito_efectivo(nov), "Sin definir")
                st.caption(f"📅 {nov['fecha_ini']} → {nov['fecha_fin']} | {ambito_lbl} | 📝 {nov['detalle']}")
            with col_edit:
                if st.button("✏️", key=f"edit_{idx}", use_container_width=True, help="Editar"):
                    limpiar_form_novedad()
                    st.session_state.editando_idx = idx
                    st.rerun()
            with col_borrar:
                if st.button("🗑️", key=f"del_{idx}", use_container_width=True, help="Eliminar"):
                    log_movimiento("Novedades", "ELIMINAR NOVEDAD", nov.get("orden"), nov.get("nombre"), nov.get("aula"), f"{nov.get('estado')} | {nov.get('fecha_ini')} a {nov.get('fecha_fin')} | {nov.get('detalle')}")
                    # 1. Eliminar de la base de datos
                    eliminar_novedad(nov['id'])
                    
                    # 2. Sincronizar asistencia a PRESENTE
                    st.session_state.estado_asistencia[nov['orden']] = "PRESENTE"
                    actualizar_asistencia(FECHA_STR, nov['orden'], "PRESENTE")
                    
                    # 3. Actualizar la lista y refrescar
                    st.session_state.novedades_lista = obtener_novedades(FECHA_STR)
                    st.toast("Novedad eliminada y asistencia actualizada")
                    st.rerun()
            st.markdown("<hr style='margin: 5px 0px; border-color: #333;'>", unsafe_allow_html=True)
        
        if st.button("🗑️ Vaciar Todas las Novedades", type="secondary", key="btn_clear_all_nov"):
            log_movimiento("Novedades", "VACIAR NOVEDADES", detalle=f"Se eliminaron {len(st.session_state.novedades_lista)} novedades activas")
            vaciar_novedades()
            for n in st.session_state.novedades_lista:
                st.session_state.estado_asistencia[n['orden']] = "PRESENTE"
                actualizar_asistencia(FECHA_STR, n['orden'], "PRESENTE")
            st.session_state.novedades_lista = []
            st.toast("Todo limpio: Novedades y asistencia reiniciadas")
            st.rerun()

# --- TAB: SEGUIMIENTO ---
with tab_seg:
    st.subheader("Ubicación del personal por aula")
    
    # Selector de turno
    turno_act = st.radio("Seleccionar Turno:", ["🌅 MAÑANA", "🌆 TARDE"], horizontal=True, label_visibility="collapsed")
    prefijo = "m" if turno_act == "🌅 MAÑANA" else "t"

    st.caption("Resumen compacto por aula. Abrí solo el aula que necesites modificar.")
    locations = [("Aula", "EN AULA"), ("URF", "URF"), ("Ed. física", "EDUCACIÓN FÍSICA"), ("Actividad", "EN INSTITUTO")]
    grid_cols = st.columns(2)

    for idx, aula in enumerate(AULAS_UNICAS):
        with grid_cols[idx % 2]:
            cfg = st.session_state.estado_aulas[aula]
            estado_key = f"estado_{prefijo}"
            salida_key = f"salida_{prefijo}"
            ubic_key = f"ubicacion_{prefijo}"
            alumnos = df[df['AULA'] == aula]
            total = len(alumnos)
            ausentes = sum(1 for n in st.session_state.novedades_lista if n['aula'] == aula and ambito_efectivo(n) == 'AUSENTE')
            presentes = total - ausentes
            is_inside = cfg[estado_key] == 'EN INSTITUTO'
            ubicacion_actual = cfg.get(ubic_key, 'EN AULA')
            estado_label = ubicacion_actual if is_inside else "FUERA"
            estado_color = "#22C55E" if is_inside else "#EF4444"
            salida_txt = f" | Salida: {cfg[salida_key]}" if cfg[salida_key] else ""

            st.markdown(
                f"""
                <div style="border:1px solid rgba(148,163,184,.22); border-radius:8px; padding:.65rem .75rem; margin:.35rem 0; background:rgba(255,255,255,.035);">
                    <div style="display:flex; justify-content:space-between; gap:.5rem; align-items:center;">
                        <strong style="font-size:1rem;">{aula}</strong>
                        <span style="font-size:.78rem; color:{estado_color}; font-weight:700;">{estado_label}</span>
                    </div>
                    <div style="font-size:.8rem; color:#A7B0BE; margin-top:.25rem;">Total {total} | Presentes {presentes} | Ausentes {ausentes}{salida_txt}</div>
                </div>
                """,
                unsafe_allow_html=True
            )

            with st.expander(f"Gestionar {aula}", expanded=False):
                if is_inside:
                    loc_cols = st.columns(4)
                    for i, (label, value) in enumerate(locations):
                        with loc_cols[i]:
                            is_active = ubicacion_actual == value
                            btn_type = "primary" if is_active else "secondary"
                            if st.button(label, key=f"loc_{prefijo}_{aula}_{i}", type=btn_type, use_container_width=True):
                                st.session_state.estado_aulas[aula][ubic_key] = value
                                guardar_estado_aula(FECHA_STR, aula, cfg['estado_m'], cfg['estado_t'], cfg.get('salida_m'), cfg.get('salida_t'), cfg.get('ubicacion_m', 'EN AULA'), cfg.get('ubicacion_t', 'EN AULA'))
                                log_movimiento("Ubicación", "CAMBIAR UBICACIÓN", aula=aula, detalle=f"Turno {turno_act}: {ubicacion_actual} -> {value}")
                                st.rerun()

                    if st.button("Retirar aula", key=f"out_{prefijo}_{aula}", use_container_width=True):
                        st.session_state.estado_aulas[aula][estado_key] = 'FUERA'
                        st.session_state.estado_aulas[aula][salida_key] = datetime.now().strftime("%H:%M")
                        guardar_estado_aula(FECHA_STR, aula, cfg['estado_m'], cfg['estado_t'], cfg.get('salida_m'), cfg.get('salida_t'), cfg.get('ubicacion_m', 'EN AULA'), cfg.get('ubicacion_t', 'EN AULA'))
                        log_movimiento("Ubicación", "RETIRAR AULA", aula=aula, detalle=f"Turno {turno_act} | Salida {st.session_state.estado_aulas[aula][salida_key]}")
                        st.rerun()
                else:
                    st.warning("Aula fuera del instituto. Reingresar para habilitar ubicación.")
                    if st.button("Reingresar aula", key=f"in_{prefijo}_{aula}", use_container_width=True):
                        st.session_state.estado_aulas[aula][estado_key] = 'EN INSTITUTO'
                        st.session_state.estado_aulas[aula][salida_key] = None
                        guardar_estado_aula(FECHA_STR, aula, cfg['estado_m'], cfg['estado_t'], cfg.get('salida_m'), cfg.get('salida_t'), cfg.get('ubicacion_m', 'EN AULA'), cfg.get('ubicacion_t', 'EN AULA'))
                        log_movimiento("Ubicación", "REINGRESAR AULA", aula=aula, detalle=f"Turno {turno_act}")
                        st.rerun()

# --- TAB: ALMUERZO ---
with tab_alm:
    st.subheader("Control de Personal que Almuerza")
    search = live_search_input("Buscar aspirante:", "Nombre, DNI o CE", "search_alm")
    if search.strip():
        s = search.strip().upper()
        res = df[
            (df['NOMBRE_COMPLETO'].str.contains(s, na=False)) |
            (df['DNI'].str.contains(s, na=False)) |
            (df['CE'].str.contains(s, na=False))
        ]
        if not res.empty:
            st.markdown("### Resultados de búsqueda")
            for i, (_, r) in enumerate(res.head(10).iterrows()):
                c1, c2 = st.columns([4, 1])
                with c1: 
                    st.markdown(f"**{r['NOMBRE_COMPLETO']}** | {r['AULA']} | Orden: {r['ORDEN_LIMP']}")
                with c2:
                    if r['ORDEN_LIMP'] in st.session_state.lista_almuerzo:
                        st.button("✅ Marcado", key=f"marked_{i}", use_container_width=True, disabled=True)
                    else:
                        if st.button("➕ Marcar", key=f"m_{i}", use_container_width=True):
                            st.session_state.lista_almuerzo.add(r['ORDEN_LIMP'])
                            agregar_almuerzo(FECHA_STR, r['ORDEN_LIMP'])
                            log_movimiento("Racionamiento", "AGREGAR ALMUERZO", int(r['ORDEN_LIMP']), r['NOMBRE_COMPLETO'], r['AULA'])
                            st.rerun()
    
    st.divider()
    st.subheader("📋 Lista de Almuerzo")
    if st.session_state.lista_almuerzo:
        st.success(f"**Total que almuerzan:** {len(st.session_state.lista_almuerzo)}")
        df_lista = df[df['ORDEN_LIMP'].isin(st.session_state.lista_almuerzo)].sort_values('ORDEN_LIMP')
        for idx, row in df_lista.iterrows():
            col_info, col_btn = st.columns([4, 1])
            with col_info:
                st.markdown(f"**{row['ORDEN_LIMP']}** - {row['NOMBRE_COMPLETO']} | {row['AULA']} | DNI: {row['DNI']}")
            with col_btn:
                if st.button("❌ Quitar", key=f"quit_{row['ORDEN_LIMP']}", use_container_width=True):
                    st.session_state.lista_almuerzo.discard(row['ORDEN_LIMP'])
                    quitar_almuerzo(FECHA_STR, row['ORDEN_LIMP'])
                    log_movimiento("Racionamiento", "QUITAR ALMUERZO", int(row['ORDEN_LIMP']), row['NOMBRE_COMPLETO'], row['AULA'])
                    st.toast(f"✅ {row['NOMBRE_COMPLETO']} removido")
                    st.rerun()
            st.markdown("<hr style='margin: 3px 0; border-color: #444;'>", unsafe_allow_html=True)
        if st.button("🗑️ Vaciar lista completa", type="secondary", key="clear_all_lunch"):
            log_movimiento("Racionamiento", "VACIAR ALMUERZO", detalle=f"Se quitaron {len(st.session_state.lista_almuerzo)} registros")
            for orden in list(st.session_state.lista_almuerzo):
                quitar_almuerzo(FECHA_STR, orden)
            st.session_state.lista_almuerzo.clear()
            st.rerun()
    else:
        st.info("ℹ️ Aún no hay personal marcado para almorzar.")

    st.divider()
    if st.button("📥 GENERAR PARTE DE RACIONAMIENTO", type="primary", use_container_width=True):
        from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "RACIONAMIENTO"
        
        ws.merge_cells('A1:F1')
        ws['A1'] = "PARTE DE RACIONAMIENTO - ESCUADRÓN H"
        ws['A1'].font = Font(bold=True, size=15, color="FFFFFF")
        ws['A1'].fill = PatternFill(start_color="E65100", end_color="E65100", fill_type="solid")
        ws['A1'].alignment = Alignment(horizontal="center", vertical="center")
        ws['A2'] = f"Fecha: {st.session_state.fecha_reporte.strftime('%d/%m/%Y')} | Total: {len(st.session_state.lista_almuerzo)}"
        ws['A2'].font = Font(italic=True, size=10, color="555555")
        ws.row_dimensions[1].height = 25
        
        headers = ["Nro", "NOMBRE COMPLETO", "GRADO", "CE", "DNI", "AULA"]
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=4, column=col, value=h)
            cell.font = Font(bold=True, color="FFFFFF", size=9)
            cell.fill = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
            cell.alignment = Alignment(horizontal="center")
            cell.border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
            
        row = 5
        for nro, orden in enumerate(sorted(st.session_state.lista_almuerzo), 1):
            p = df[df['ORDEN_LIMP'] == orden].iloc[0]
            ws.cell(row=row, column=1, value=nro)
            ws.cell(row=row, column=2, value=p['NOMBRE_COMPLETO'])
            ws.cell(row=row, column=3, value=p['GRADO'])
            ws.cell(row=row, column=4, value=p['CE'])
            ws.cell(row=row, column=5, value=p['DNI'])
            ws.cell(row=row, column=6, value=p['AULA'])
            for c in range(1, 7):
                ws.cell(row=row, column=c).border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
                ws.cell(row=row, column=c).alignment = Alignment(horizontal="center" if c in [1,3,6] else "left")
            row += 1
            
        for col, w in zip("ABCDEF", [10, 35, 12, 12, 15, 12]): 
            ws.column_dimensions[col].width = w
        output = f"RACIONAMIENTO_{datetime.now().strftime('%d%m%Y_%H%M')}.xlsx"
        st.success("✅ Parte de racionamiento listo para descargar")
        st.download_button(
            "📥 Descargar parte de racionamiento",
            data=excel_bytes(wb),
            file_name=output,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="download_racionamiento_excel"
        )

# ==================== PLAN DE LLAMADA (CONTACTOS) ====================
# --- TAB: PLAN DE LLAMADA ---
with tab_plan:
    st.subheader("📞 Plan de Llamada - Base de Contactos")
    st.info("Registra domicilios y contactos de emergencia para cada personal del escuadrón.")
    
    search = live_search_input("Buscar personal:", "Nombre, DNI, CE o Orden", "search_plan")
    
    if search.strip():
        s = search.strip().upper()
        res = df[
            (df['NOMBRE_COMPLETO'].str.contains(s, na=False)) |
            (df['DNI'].str.contains(s, na=False)) |
            (df['CE'].str.contains(s, na=False)) |
            (df['ORDEN_LIMP'].astype(str).str.contains(s, na=False))
        ]
        
        if not res.empty:
            st.markdown(f"### Resultados ({len(res)} encontrados)")
            for _, row in res.head(10).iterrows():
                orden = row['ORDEN_LIMP']
                contacto = obtener_contacto(orden)
                
                with st.expander(f"**{orden} - {row['NOMBRE_COMPLETO']}** | {row['AULA']} | DNI: {row['DNI']}", expanded=False):
                    c1, c2 = st.columns(2)
                    
                    with c1:
                        dom = st.text_area("📍 Domicilio", value=contacto.get('domicilio','') if contacto else '', 
                                          key=f"dom_{orden}", placeholder="Calle, número, barrio...")
                        tel_pers = st.text_input("📱 Teléfono Personal", value=contacto.get('telefono_personal','') if contacto else '',
                                                key=f"telp_{orden}", placeholder="2964-XXXXXX")
                        tel_urg = st.text_input("🚨 Teléfono Emergencia", value=contacto.get('telefono_emergencia','') if contacto else '',
                                               key=f"telu_{orden}", placeholder="2964-XXXXXX")
                    
                    with c2:
                        nom_urg = st.text_input("👤 Nombre Contacto Emergencia", 
                                               value=contacto.get('nombre_emergencia','') if contacto else '',
                                               key=f"nomu_{orden}", placeholder="Nombre completo")
                        parent = st.text_input("🔗 Parentesco", value=contacto.get('parentesco_emergencia','') if contacto else '',
                                              key=f"par_{orden}", placeholder="Esposa, Madre, Hermano...")
                        obs = st.text_area("📝 Observaciones", value=contacto.get('observaciones','') if contacto else '',
                                          key=f"obs_{orden}", placeholder="Alergias, grupo sanguíneo, etc.")
                    
                    if st.button("💾 Guardar Contacto", key=f"save_{orden}", type="primary"):
                        guardar_contacto({
                            'orden': int(orden), 'domicilio': dom, 'telefono_personal': tel_pers,
                            'telefono_emergencia': tel_urg, 'nombre_emergencia': nom_urg,
                            'parentesco_emergencia': parent, 'observaciones': obs
                        })
                        st.success(f"✅ Datos guardados para {row['NOMBRE_COMPLETO']}")
                        st.rerun()
    
    st.divider()
    st.subheader("📋 Resumen del Plan de Llamada")
    
    todos = obtener_todos_contactos()
    if todos:
        col1, col2, col3 = st.columns(3)
        with col1: st.metric("Total Registrados", len(todos))
        with col2: 
            con_tel = sum(1 for t in todos if t.get('telefono_personal'))
            st.metric("Con Teléfono Personal", con_tel)
        with col3:
            con_urg = sum(1 for t in todos if t.get('telefono_emergencia'))
            st.metric("Con Contacto Emergencia", con_urg)
        
        if st.button("📥 EXPORTAR PLAN DE LLAMADA (EXCEL)", type="primary", use_container_width=True):
            from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "PLAN DE LLAMADA"
            
            ws.merge_cells('A1:H1')
            ws['A1'] = "PLAN DE LLAMADA - ESCUADRÓN H"
            ws['A1'].font = Font(bold=True, size=16, color="FFFFFF")
            ws['A1'].fill = PatternFill(start_color="C62828", end_color="C62828", fill_type="solid")
            ws['A1'].alignment = Alignment(horizontal="center")
            
            headers = ["Nro", "NOMBRE", "AULA", "DOMICILIO", "TEL. PERSONAL", "TEL. EMERGENCIA", "CONTACTO EMERG.", "OBSERV."]
            for col, h in enumerate(headers, 1):
                cell = ws.cell(row=3, column=col, value=h)
                cell.font = Font(bold=True, color="FFFFFF", size=9)
                cell.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                cell.alignment = Alignment(horizontal="center")
                cell.border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
            
            row = 4
            for nro, cont in enumerate(sorted(todos, key=lambda x: x['orden']), 1):
                pers = df[df['ORDEN_LIMP'] == cont['orden']]
                nombre = pers.iloc[0]['NOMBRE_COMPLETO'] if not pers.empty else ""
                aula = pers.iloc[0]['AULA'] if not pers.empty else ""
                
                ws.cell(row=row, column=1, value=nro)
                ws.cell(row=row, column=2, value=nombre)
                ws.cell(row=row, column=3, value=aula)
                ws.cell(row=row, column=4, value=cont.get('domicilio',''))
                ws.cell(row=row, column=5, value=cont.get('telefono_personal',''))
                ws.cell(row=row, column=6, value=cont.get('telefono_emergencia',''))
                ws.cell(row=row, column=7, value=f"{cont.get('nombre_emergencia','')} ({cont.get('parentesco_emergencia','')})")
                ws.cell(row=row, column=8, value=cont.get('observaciones',''))
                
                for c in range(1, 9):
                    ws.cell(row=row, column=c).border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
                    ws.cell(row=row, column=c).alignment = Alignment(horizontal="center" if c in [1,3] else "left")
                row += 1
            
            for col, w in zip("ABCDEFGH", [8, 30, 10, 35, 15, 15, 25, 30]):
                ws.column_dimensions[col].width = w
            
            output = f"PLAN_LLAMADA_{datetime.now().strftime('%d%m%Y')}.xlsx"
            st.success("✅ Plan de llamada listo para descargar")
            st.download_button(
                "📥 Descargar plan de llamada",
                data=excel_bytes(wb),
                file_name=output,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="download_plan_llamada_excel"
            )
    else:
        st.warning("⚠️ Aún no hay contactos registrados. Usa el buscador para cargar datos.")
  

# --- TAB: RESUMEN ---
with tab_res:
    st.subheader("Resumen General y Novedades")

    # Recarga de datos
    st.session_state.novedades_lista = obtener_novedades(FECHA_STR)
    st.session_state.estado_asistencia = obtener_asistencia(FECHA_STR)
    st.session_state.lista_almuerzo = obtener_almuerzo(FECHA_STR)

    ambito_resumen = {
        n['orden']: ambito_efectivo(n)
        for n in st.session_state.novedades_lista
    }
    ausentes_resumen = {orden for orden, ambito in ambito_resumen.items() if ambito == "AUSENTE"}
    ausentes_manuales_resumen = {
        orden for orden, estado in st.session_state.estado_asistencia.items()
        if estado == "AUSENTE"
    }
    presentes_manuales_resumen = {
        orden for orden, estado in st.session_state.estado_asistencia.items()
        if estado in {"PRESENTE", "PRESENTE EN INSTITUTO", "PRESENTE EN ESCUADRÓN"}
    }
    total_ausentes_resumen = (ausentes_resumen | ausentes_manuales_resumen) - presentes_manuales_resumen

    # ==========================================================
    # RESUMEN POR AULA
    # ==========================================================

    data_aulas = []

    for aula in AULAS_UNICAS:

        cfg = st.session_state.estado_aulas[aula]
        alumnos = df[df['AULA'] == aula]

        total_aula = len(alumnos)

        ausentes_aula = len(
            {
                row['ORDEN_LIMP']
                for _, row in alumnos.iterrows()
                if row['ORDEN_LIMP'] in total_ausentes_resumen
            }
        )

        presentes_aula = total_aula - ausentes_aula

        almuerzan = sum(
            1
            for _, row in alumnos.iterrows()
            if row['ORDEN_LIMP']
            in st.session_state.lista_almuerzo
        )

        data_aulas.append({
            "Aula": aula,
            "Total": total_aula,
            "Presentes": presentes_aula,
            "Ausentes": ausentes_aula,
            "Almuerzan": almuerzan,
            "Ubicación": (
                cfg.get("ubicacion_m", "-")
                if cfg["estado_m"] == "EN INSTITUTO"
                else "FUERA"
            ),
            "Estado": cfg["estado_m"]
        })

    st.dataframe(
        pd.DataFrame(data_aulas),
        use_container_width=True,
        hide_index=True
    )

    st.divider()

# =====================================================
# TABLA PERSONAL AUSENTE O FUERA DE INSTALACIONES
# =====================================================

data_ausentes = []

for nov in st.session_state.novedades_lista:

    alumno_df = df[df['ORDEN_LIMP'] == nov['orden']]

    if not alumno_df.empty:

        alumno = alumno_df.iloc[0]

        data_ausentes.append({
            "Nombre": alumno["NOMBRE_COMPLETO"],
            "Motivo": nov["estado"],
            "Desde": nov["fecha_ini"],
            "Hasta": nov["fecha_fin"],
        })

# AUSENTES MANUALES
for orden, estado in st.session_state.estado_asistencia.items():

    if estado == "AUSENTE":

        # Evitar duplicados si ya tiene novedad cargada
        ya_existe = any(
            nov["orden"] == orden
            for nov in st.session_state.novedades_lista
        )

        if not ya_existe:

            alumno_df = df[df['ORDEN_LIMP'] == orden]

            if not alumno_df.empty:

                alumno = alumno_df.iloc[0]

                data_ausentes.append({
                    "Nombre": alumno["NOMBRE_COMPLETO"],
                    "Motivo": "AUSENTE",
                    "Desde": FECHA_STR,
                    "Hasta": FECHA_STR
                })

df_ausentes = pd.DataFrame(data_ausentes)

if not df_ausentes.empty:
    df_ausentes = df_ausentes.sort_values(["Motivo", "Nombre"]).reset_index(drop=True)
    df_ausentes.insert(0, "Nro", range(1, len(df_ausentes) + 1))
    st.dataframe(
        df_ausentes,
        use_container_width=True,
        hide_index=True
    )
else:
    st.success("Sin personal ausente.")

with tab_res:
    st.divider()
    st.subheader("Minuta informativa")
    minuta_texto = generar_minuta_informativa()
    st.text_area("Minuta generada automáticamente desde Novedades", value=minuta_texto, height=520, key="minuta_generada")
    st.download_button(
        "📄 Descargar minuta (.txt)",
        data=minuta_texto.encode("utf-8"),
        file_name=f"MINUTA_ESCUADRON_H_{st.session_state.fecha_reporte.strftime('%d%m%Y')}.txt",
        mime="text/plain",
        use_container_width=True
    )

    st.divider()
    st.subheader("Historial de movimientos")
    fecha_historial = st.date_input("Fecha a consultar", st.session_state.fecha_reporte, key="fecha_historial_mov")
    movimientos = obtener_movimientos(fecha_historial.isoformat())
    if movimientos:
        df_mov = pd.DataFrame(movimientos)
        df_mov = df_mov.rename(columns={
            "fecha_hora": "Fecha/hora",
            "fecha_parte": "Fecha parte",
            "modulo": "Módulo",
            "accion": "Acción",
            "orden": "Orden interno",
            "nombre": "Nombre",
            "aula": "Aula",
            "detalle": "Detalle",
        })
        columnas = ["Fecha/hora", "Módulo", "Acción", "Nombre", "Aula", "Detalle"]
        st.dataframe(df_mov[columnas], use_container_width=True, hide_index=True)
        st.download_button(
            "📥 Descargar historial (.csv)",
            data=df_mov[columnas].to_csv(index=False).encode("utf-8-sig"),
            file_name=f"HISTORIAL_MOVIMIENTOS_{fecha_historial.strftime('%d%m%Y')}.csv",
            mime="text/csv",
            use_container_width=True
        )
    else:
        st.info("No hay movimientos registrados para esa fecha.")

if st.button("📥 GENERAR PARTE DIARIO (EXCEL)", type="primary", use_container_width=True, key="btn_parte_diario_formal"):
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PARTE DIARIO"

    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    sub_fill = PatternFill(start_color="D9EAF7", end_color="D9EAF7", fill_type="solid")
    fecha_titulo = st.session_state.fecha_reporte.strftime('%d%b%y').upper()
    dia_reporte = DIAS_SEMANA[st.session_state.fecha_reporte.weekday()]

    ws.merge_cells('A1:J1')
    ws['A1'] = f"PARTE DIARIO DEL ESCUADRÓN H - {fecha_titulo}"
    ws['A1'].font = Font(bold=True, size=16, color="FFFFFF")
    ws['A1'].fill = header_fill
    ws['A1'].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    ws.merge_cells('A2:J2')
    ws['A2'] = f"Día: {dia_reporte} | Primera obligación: 06:00 hs | Generado: {datetime.now().strftime('%H:%M')}"
    ws['A2'].font = Font(italic=True, size=11, color="555555")
    ws['A2'].alignment = Alignment(horizontal="center")

    metric_headers = [
        "TOTAL", "EN INSTITUTO", "EN ESCUADRÓN", "AUSENTES",
        "GUARDIA D.", "GUARDIA N.", "COMISIÓN",
        "1RA OBLIG. 06:00", "3ER AÑO", "AOP"
    ]
    metric_values = [
        TOTAL_ESCUADRON, en_instituto, presentes_escuadron, len(total_ausentes),
        sum(1 for n in st.session_state.novedades_lista if n['estado'] == 'ENTRANTE GUARDIA DIURNA'),
        sum(1 for n in st.session_state.novedades_lista if n['estado'] == 'ENTRANTE GUARDIA NOCTURNA'),
        sum(1 for n in st.session_state.novedades_lista if n['estado'] == 'COMISIÓN'),
        primera_total, primera_tercer_anio, primera_aop
    ]
    for col, label in enumerate(metric_headers, 1):
        cell = ws.cell(row=4, column=col, value=label)
        cell.font = Font(bold=True, size=9, color="1F2937")
        cell.fill = sub_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
        val = ws.cell(row=5, column=col, value=metric_values[col - 1])
        val.font = Font(bold=True, size=12)
        val.alignment = Alignment(horizontal="center")
        val.border = border

    ws.merge_cells('A8:J8')
    ws['A8'] = "NOVEDADES DEL PERSONAL (AUSENTES JUSTIFICADOS)"
    ws['A8'].font = Font(bold=True, size=12, color="1F4E78")

    nov_headers = ["Nro", "GRADO", "APELLIDO Y NOMBRE", "DNI", "CE", "NOVEDAD", "DETALLE", "DESDE", "HASTA", "AULA"]
    for col, h in enumerate(nov_headers, 1):
        cell = ws.cell(row=9, column=col, value=h)
        cell.font = Font(bold=True, color="FFFFFF", size=9)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = border

    current_row = 10
    if st.session_state.novedades_lista:
        for i, nov in enumerate(st.session_state.novedades_lista, 1):
            row = current_row + i - 1
            values = [
                i, nov.get('grado', ''), nov['nombre'], nov.get('dni', ''), nov.get('ce', ''),
                nov['estado'], nov['detalle'], nov['fecha_ini'], nov['fecha_fin'], nov.get('aula', '-')
            ]
            for col, value in enumerate(values, 1):
                cell = ws.cell(row=row, column=col, value=value)
                cell.border = border
                cell.alignment = Alignment(horizontal="center" if col in [1, 2, 4, 5, 6, 8, 9, 10] else "left")
        current_row += len(st.session_state.novedades_lista)
    else:
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=10)
        ws.cell(row=current_row, column=1, value="Sin novedades registradas en la guardia").font = Font(italic=True, color="888888")
        ws.cell(row=current_row, column=1).alignment = Alignment(horizontal="center")
        current_row += 1

    current_row += 2
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=10)
    ws.cell(row=current_row, column=1, value=f"HORARIOS DE INGRESO - {dia_reporte.upper()}").font = Font(bold=True, size=12, color="1F4E78")
    current_row += 1

    aulas_0600 = []
    for aula in AULAS_UNICAS:
        hor = st.session_state.horarios_config[aula]
        if hor.get('ent_m') == '06:00':
            aulas_0600.append(aula)

    texto_0600 = (
        f"• Ingreso 06:00 hs (Primera obligación): Aula(s) {', '.join(aulas_0600) if aulas_0600 else 'N/A'} "
        f"— Forman {primera_total} aspirante(s): {primera_tercer_anio} de 3er año y {primera_aop} de AOP."
    )
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=10)
    ws.cell(row=current_row, column=1, value=texto_0600)
    ws.cell(row=current_row, column=1).alignment = Alignment(wrap_text=True)
    current_row += 1

    otros_ingresos = []
    for aula in AULAS_UNICAS:
        hor = st.session_state.horarios_config[aula]
        if hor.get('ent_m') != '06:00':
            cant = len(df_presentes_primera[df_presentes_primera['AULA'] == aula])
            otros_ingresos.append(f"{aula}: {hor.get('ent_m')} hs ({cant})")
    if otros_ingresos:
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=10)
        ws.cell(row=current_row, column=1, value=f"• Otros ingresos mañana: {', '.join(otros_ingresos)}")
        ws.cell(row=current_row, column=1).alignment = Alignment(wrap_text=True)
        current_row += 1

    current_row += 2
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=10)
    ws.cell(row=current_row, column=1, value="OBSERVACIONES").font = Font(bold=True, size=12, color="1F4E78")

    for col, width in enumerate([10, 14, 34, 12, 10, 18, 24, 12, 12, 12], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = width

    output = f"PARTE_DIARIO_ESCUADRON_H_{datetime.now().strftime('%d%m%Y_%H%M')}.xlsx"
    st.success("✅ Parte diario listo para descargar")
    st.download_button(
        "📥 Descargar parte diario",
        data=excel_bytes(wb),
        file_name=output,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key="download_parte_diario_excel"
    )

# ==============================================================================
# 5. EXPORTAR EXCEL (PARTE DIARIO DETALLADO)
# ==============================================================================
if st.button("📥 GENERAR PARTE DIARIO DETALLADO (EXCEL)", type="secondary", use_container_width=True):
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PARTE DIARIO"

    # ️ ENCABEZADO PRINCIPAL
    ws.merge_cells('A1:F1')
    ws['A1'] = "PARTE DIARIO - ESCUADRÓN H"
    ws['A1'].font = Font(bold=True, size=18, color="FFFFFF")
    ws['A1'].fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    ws['A1'].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells('A2:F2')
    ws['A2'] = f"Fecha: {st.session_state.fecha_reporte.strftime('%d/%m/%Y')} | Generado: {datetime.now().strftime('%H:%M')}"
    ws['A2'].font = Font(italic=True, size=11, color="555555")
    ws['A2'].alignment = Alignment(horizontal="center")
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 20

    # 🔹 SECCIÓN 1: MÉTRICAS GENERALES
    ws['A4'] = "📊 RESUMEN DE EFECTIVOS"
    ws['A4'].font = Font(bold=True, size=13, color="1F4E78")
    
    metrics = [
        ("TOTAL ESCUADRÓN", TOTAL_ESCUADRON),
        ("DISPONIBLES", disponibles),
        ("EN INSTITUTO", en_instituto),
        ("FUERA DEL INSTITUTO", total_fuera)
    ]
    for i, (label, value) in enumerate(metrics):
        row = 5 + i
        ws.cell(row=row, column=1, value=label).font = Font(bold=True)
        cell_val = ws.cell(row=row, column=2, value=value)
        cell_val.font = Font(bold=True, size=12, color="1F4E78")
        cell_val.alignment = Alignment(horizontal="center")
        for c in range(1, 3):
            ws.cell(row=row, column=c).border = Border(left=Side(style="thin"), right=Side(style="thin"),
                                                       top=Side(style="thin"), bottom=Side(style="thin"))

    # 🔹 SECCIÓN 2: NOVEDADES REGISTRADAS (CON COLUMNA AULA)
    ws['A10'] = "📝 NOVEDADES REGISTRADAS"
    ws['A10'].font = Font(bold=True, size=13, color="1F4E78")

    # ✅ AGREGAMOS "AULA" EN LA LISTA DE ENCABEZADOS
    nov_headers = ["Nro", "NOMBRE", "AULA", "ESTADO", "DETALLE", "PERÍODO"]
    for col, h in enumerate(nov_headers, 1):
        cell = ws.cell(row=11, column=col, value=h)
        cell.font = Font(bold=True, color="FFFFFF", size=10)
        cell.fill = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")
        cell.border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

    if st.session_state.novedades_lista:
        for i, nov in enumerate(st.session_state.novedades_lista, 1):
            row = 11 + i
            ws.cell(row=row, column=1, value=i)
            ws.cell(row=row, column=2, value=nov['nombre'])
            ws.cell(row=row, column=3, value=nov.get('aula', '-')) # ✅ AQUÍ SE AGREGA EL AULA
            ws.cell(row=row, column=4, value=nov['estado'])
            ws.cell(row=row, column=5, value=nov['detalle'])
            ws.cell(row=row, column=6, value=f"{nov['fecha_ini']} a {nov['fecha_fin']}")
            # Aplicar bordes y alineación a las 6 columnas
            for c in range(1, 7):
                ws.cell(row=row, column=c).border = Border(left=Side(style="thin"), right=Side(style="thin"),
                                                           top=Side(style="thin"), bottom=Side(style="thin"))
                # Alineación: Centro para números y estados, Izquierda para textos
                ws.cell(row=row, column=c).alignment = Alignment(horizontal="center" if c in [1, 4, 6] else "left")
    else:
        ws.merge_cells('A12:F12')
        ws.cell(row=12, column=1, value="Sin novedades registradas en la guardia").font = Font(italic=True, color="888888")
        ws.cell(row=12, column=1).alignment = Alignment(horizontal="center")

    # 🔹 SECCIÓN 3: PERSONAL QUE ALMUERZA
    ws['A18'] = "🍽️ PERSONAL QUE ALMUERZA"
    ws['A18'].font = Font(bold=True, size=13, color="1F4E78")

    alm_headers = ["Nro", "NOMBRE COMPLETO", "AULA"]
    for col, h in enumerate(alm_headers, 1):
        cell = ws.cell(row=19, column=col, value=h)
        cell.font = Font(bold=True, color="FFFFFF", size=10)
        cell.fill = PatternFill(start_color="E65100", end_color="E65100", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")
        cell.border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

    if st.session_state.lista_almuerzo:
        for i, orden in enumerate(sorted(st.session_state.lista_almuerzo), 1):
            row = 19 + i
            asp = df[df['ORDEN_LIMP'] == orden]
            if not asp.empty:
                ws.cell(row=row, column=1, value=i)
                ws.cell(row=row, column=2, value=asp.iloc[0]['NOMBRE_COMPLETO'])
                ws.cell(row=row, column=3, value=asp.iloc[0]['AULA'])
            for c in range(1, 4):
                ws.cell(row=row, column=c).border = Border(left=Side(style="thin"), right=Side(style="thin"),
                                                           top=Side(style="thin"), bottom=Side(style="thin"))
                ws.cell(row=row, column=c).alignment = Alignment(horizontal="center" if c == 1 else "left")
    else:
        ws.merge_cells('A20:C20')
        ws.cell(row=20, column=1, value="N/A").font = Font(italic=True, color="888888")
        ws.cell(row=20, column=1).alignment = Alignment(horizontal="center")

    # 🔹 SECCIÓN 4: CONTROL DE AULAS Y HORARIOS
    ws['A26'] = "🏫 CONTROL DE AULAS Y HORARIOS"
    ws['A26'].font = Font(bold=True, size=13, color="1F4E78")

    aula_headers = ["AULA", "HORARIO MAÑANA", "ESTADO M", "HORARIO TARDE", "ESTADO T"]
    for col, h in enumerate(aula_headers, 1):
        cell = ws.cell(row=27, column=col, value=h)
        cell.font = Font(bold=True, color="FFFFFF", size=10)
        cell.fill = PatternFill(start_color="4527A0", end_color="4527A0", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")
        cell.border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

    for i, aula in enumerate(AULAS_UNICAS):
        row = 28 + i
        cfg = st.session_state.estado_aulas[aula]
        hor = st.session_state.horarios_config[aula]
        ws.cell(row=row, column=1, value=aula)
        ws.cell(row=row, column=2, value=f"{hor['ent_m']} - {hor['sal_m']}")
        ws.cell(row=row, column=3, value=cfg['estado_m'])
        ws.cell(row=row, column=4, value=f"{hor['ent_t']} - {hor['sal_t']}")
        ws.cell(row=row, column=5, value=cfg['estado_t'])
        for c in range(1, 6):
            ws.cell(row=row, column=c).border = Border(left=Side(style="thin"), right=Side(style="thin"),
                                                       top=Side(style="thin"), bottom=Side(style="thin"))
            ws.cell(row=row, column=c).alignment = Alignment(horizontal="center" if c > 1 else "left")

    # 📐 Ajuste de columnas para acomodar la nueva columna "AULA"
    ws.column_dimensions['A'].width = 10  # Orden
    ws.column_dimensions['B'].width = 25  # Nombre
    ws.column_dimensions['C'].width = 12  # Aula (Nueva)
    ws.column_dimensions['D'].width = 12  # Estado
    ws.column_dimensions['E'].width = 25  # Detalle
    ws.column_dimensions['F'].width = 20  # Período

    # 💾 Guardar archivo
    output = f"PARTE_DIARIO_DETALLADO_{datetime.now().strftime('%d%m%Y_%H%M')}.xlsx"
    st.success("✅ Parte diario detallado listo para descargar")
    st.download_button(
        "📥 Descargar parte diario detallado",
        data=excel_bytes(wb),
        file_name=output,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key="download_parte_diario_detallado_excel"
    )
