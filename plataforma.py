# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime
import os
import io

# 🔹 IMPORTS DESDE BASE DE DATOS
from db_manager import (
    init_db,
    obtener_novedades, agregar_novedad, actualizar_novedad, eliminar_novedad, vaciar_novedades,
    obtener_estado_aulas, guardar_estado_aula,
    obtener_almuerzo, agregar_almuerzo, quitar_almuerzo,
    obtener_horarios, guardar_horarios,
    obtener_asistencia, actualizar_asistencia,
    obtener_todos_contactos, guardar_contacto
)

st.set_page_config(page_title="Gestión de Parte Diario - Escuadrón H", layout="wide")

# ==============================================================================
# 🎖️ ESTILOS E INTERFAZ INSTITUTIONAL
# ==============================================================================
st.markdown("""
<style>
/* Estilos generales */
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="header-box">
</div>
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
            df = df[df['ORDEN_LIMP'].between(1, 139)].dropna(subset=['ORDEN_LIMP'])
            return df[['ORDEN_LIMP', 'AULA', 'GRADO', 'NOMBRE_COMPLETO', 'DNI', 'CE']].sort_values('ORDEN_LIMP')
        else:
            st.error("No se encontró 'alumnos.csv' en la carpeta del proyecto.")
            return pd.DataFrame()
    except Exception as e:
        st.error(f"Error al cargar datos: {e}")
        return pd.DataFrame()

# Inicializar base de datos
if 'db_iniciada' not in st.session_state:
    init_db()
    st.session_state.db_iniciada = True

df = cargar_personal()
if df.empty:
    st.stop()

TOTAL_ESCUADRON = 139
AULAS_UNICAS = sorted(df['AULA'].unique())

# Fecha del reporte
if 'fecha_reporte' not in st.session_state:
    st.session_state.fecha_reporte = datetime.now().date()
FECHA_STR = st.session_state.fecha_reporte.isoformat()

# --- COPIÁ Y PEGÁ ESTO ACÁ ---
if "config_aulas" not in st.session_state:
    st.session_state.config_aulas = {
        "18TM": {"horario": "0600 a 0620 hs", "tipo_ingreso": "Normal"},
        "24TM": {"horario": "0600 a 0620 hs", "tipo_ingreso": "Normal"},
        "23TT": {"horario": "0600 a 0620 hs", "tipo_ingreso": "Normal"},
        "23TM": {"horario": "0810 hs", "tipo_ingreso": "Diferencial"},
        "26TM": {"horario": "0900 a 0915 hs", "tipo_ingreso": "Diferencial"},
        "28TM": {"horario": "0810 hs", "tipo_ingreso": "Diferencial"},
        "7TT": {"horario": "Normal", "tipo_ingreso": "Normal"},
        "8TM": {"horario": "Normal", "tipo_ingreso": "Normal"}
    }
# -----------------------------

# Novedades y configuraciones
if 'novedades_lista' not in st.session_state:
    st.session_state.novedades_lista = obtener_novedades()
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
            "ubicacion_m": aula_data.get("ubicacion_m", "EN AULA"),
            "ubicacion_t": aula_data.get("ubicacion_t", "EN AULA")
        }

if 'lista_almuerzo' not in st.session_state:
    st.session_state.lista_almuerzo = obtener_almuerzo(FECHA_STR)

if 'horarios_config' not in st.session_state:
    db_hor = obtener_horarios()
    st.session_state.horarios_config = {}
    for aula in AULAS_UNICAS:
        st.session_state.horarios_config[aula] = db_hor.get(aula, {
            "ent_m": "06:00", "sal_m": "12:00", "ent_t": "13:00", "sal_t": "19:00"
        })

if 'estado_asistencia' not in st.session_state:
    st.session_state.estado_asistencia = obtener_asistencia(FECHA_STR)

if 'editando_idx' not in st.session_state:
    st.session_state.editando_idx = None

if 'sel_nov' not in st.session_state:
    st.session_state.sel_nov = None

# ==============================================================================
# 3. CÁLCULO DE MÉTRICAS EN TIEMPO REAL
# ==============================================================================
st.session_state.novedades_lista = obtener_novedades()
st.session_state.lista_almuerzo = obtener_almuerzo(FECHA_STR)

if not st.session_state.estado_asistencia:
    st.session_state.estado_asistencia = obtener_asistencia(FECHA_STR)
else:
    db_asistencia = obtener_asistencia(FECHA_STR)
    for orden, estado in db_asistencia.items():
        if orden not in st.session_state.estado_asistencia:
            st.session_state.estado_asistencia[orden] = estado

ausentes_fijos = {n['orden'] for n in st.session_state.novedades_lista if n['estado'] in ['ART', 'DAF', 'LES']}
ausentes_manuales = {orden for orden, estado in st.session_state.estado_asistencia.items() if estado == "AUSENTE"}
presentes_manuales = {orden for orden, estado in st.session_state.estado_asistencia.items() if estado == "PRESENTE"}

total_ausentes = (ausentes_fijos | ausentes_manuales) - presentes_manuales

en_instituto = 0
fuera_por_aula = 0
for _, row in df.iterrows():
    orden = row['ORDEN_LIMP']
    aula = row['AULA']
    if orden in total_ausentes: continue
    if st.session_state.estado_aulas.get(aula, {}).get('estado_m', 'EN INSTITUTO') == 'EN INSTITUTO':
        en_instituto += 1
    else:
        fuera_por_aula += 1

disponibles = TOTAL_ESCUADRON - len(total_ausentes)
total_fuera = fuera_por_aula + len(total_ausentes)

ubicacion_dist = {"EN AULA": [], "URF": [], "EDUCACIÓN FÍSICA": [], "EN INSTITUTO": []}
for aula in AULAS_UNICAS:
    cfg = st.session_state.estado_aulas[aula]
    if cfg['estado_m'] == 'EN INSTITUTO':
        ubic = cfg.get('ubicacion_m', 'EN AULA')
        if ubic in ubicacion_dist:
            ubicacion_dist[ubic].append(len(df[df['AULA'] == aula]))

st.markdown("""
<style>
    .sticky-bar { position: sticky !important; top: 0 !important; z-index: 999 !important; 
                  background: #161B15 !important; padding: 10px 0 12px 0 !important; 
                  border-bottom: 2px solid #C4A000 !important; margin-bottom: 15px !important; 
                  box-shadow: 0 4px 8px rgba(0,0,0,0.6) !important; }
    .header-title { text-align: center !important; font-size: 1.15rem !important; font-weight: 900 !important; 
                    color: #FFFFFF !important; letter-spacing: 2.5px !important; margin-bottom: 8px !important; 
                    text-transform: uppercase !important; text-shadow: 1px 1px 3px #000000 !important; }
    [data-testid="stMetricValue"] { font-size: 1.05rem !important; font-weight: 700 !important; color: #FFFFFF !important; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="sticky-bar">', unsafe_allow_html=True)
st.markdown('<div class="header-title">ESCUADRÓN H "CABO MARCELO GODOY"</div>', unsafe_allow_html=True)

c1, c2, c3, c4, c5, c6, c7, c8 = st.columns(8)
with c1: st.metric("TOTAL", TOTAL_ESCUADRON)
with c2: st.metric("DISP.", disponibles, delta=f"-{len(total_ausentes)}" if total_ausentes else None)
with c3: st.metric("EN INST.", en_instituto)
with c4: st.metric("FUERA", total_fuera, delta_color="inverse" if total_fuera > 0 else "normal")
with c5: 
    v = sum(ubicacion_dist['EN AULA'])
    st.metric("AULA", v, f"{len(ubicacion_dist['EN AULA'])} a." if v>0 else None)
with c6: 
    v = sum(ubicacion_dist['URF'])
    st.metric("URF", v, f"{len(ubicacion_dist['URF'])} a." if v>0 else None)
with c7: 
    v = sum(ubicacion_dist['EDUCACIÓN FÍSICA'])
    st.metric("ED.FÍS", v, f"{len(ubicacion_dist['EDUCACIÓN FÍSICA'])} a." if v>0 else None)
with c8: 
    v = sum(ubicacion_dist['EN INSTITUTO'])
    st.metric("ACTIV.", v, f"{len(ubicacion_dist['EN INSTITUTO'])} a." if v>0 else None)
st.markdown('</div>', unsafe_allow_html=True)

st.divider()

# Botones de Control Sincro
col_sync1, col_sync2 = st.columns(2)
with col_sync1:
    if st.button("🔄 Sincronizar Datos", key="sync_btn", use_container_width=True):
        st.session_state.novedades_lista = obtener_novedades()
        st.session_state.estado_asistencia = obtener_asistencia(FECHA_STR)
        st.session_state.lista_almuerzo = obtener_almuerzo(FECHA_STR)
        st.success("✅ Datos sincronizados")
        st.rerun()
with col_sync2:
    if st.button("🚨 RESETEAR ASISTENCIA DEL DÍA", key="reset_asistencia", use_container_width=True):
        import sqlite3
        conn = sqlite3.connect("parte_diario.db")
        conn.execute(f"DELETE FROM asistencia_diaria WHERE fecha='{FECHA_STR}'")
        conn.commit()
        conn.close()
        st.session_state.estado_asistencia = {}
        st.success("✅ Asistencia reiniciada.")
        st.rerun()

# ==============================================================================
# 4. ESTRUCTURA DE PESTAÑAS (AL RAS DEL MARGEN)
# ==============================================================================
tab_config, tab_nov, tab_seg, tab_alm, tab_plan, tab_res = st.tabs([
    "⚙️ Configuración", "📝 Novedades", "🔍 Seguimiento", "🍴 Almuerzo", "📞 Plan de Llamada", "📊 Resumen"
])

# --- TAB: CONFIGURACIÓN ---
with tab_config:
    st.subheader("Configuración del Día y Horarios")
    st.session_state.fecha_reporte = st.date_input("Fecha del Reporte", st.session_state.fecha_reporte)
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
            guardar_horarios(aula, st.session_state.horarios_config[aula])
        st.success("✅ Horarios guardados")
        st.rerun()

# --- TAB: NOVEDADES ---
with tab_nov:
    edit_idx = st.session_state.editando_idx
    es_edicion = edit_idx is not None

    if es_edicion and not (0 <= edit_idx < len(st.session_state.novedades_lista)):
        st.session_state.editando_idx = None
        st.rerun()

    st.subheader("✏️ Editando Novedad" if es_edicion else "➕ Registrar Novedad")
    data = None
    if es_edicion:
        nov = st.session_state.novedades_lista[edit_idx]
        st.info(f"Editando a: **{nov['nombre']}**")
        data = nov
        orden = int(data["orden"]) # Forzamos entero para el presentismo
        nombre_asp = data["nombre"]
        estado_act = st.session_state.estado_asistencia.get(orden, "PRESENTE")

        col_btn_p, col_btn_a = st.columns(2)
        with col_btn_p:
            if st.button("✅ MARCAR PRESENTE", type="primary", use_container_width=True, key="btn_pres_edit"):
                # SOLUCIÓN: Guarda el presentismo de forma independiente como ENTERO sin borrar la novedad
                actualizar_asistencia(FECHA_STR, orden, "PRESENTE")
                st.session_state.estado_asistencia[orden] = "PRESENTE"
                st.success(f"{nombre_asp} marcado como PRESENTE en el Parte Diario")
                st.rerun()
        with col_btn_a:
            if st.button("❌ MARCAR AUSENTE", type="secondary", use_container_width=True, key="btn_aus_edit"):
                # SOLUCIÓN: Guarda el presentismo de forma independiente como ENTERO sin borrar la novedad
                actualizar_asistencia(FECHA_STR, orden, "AUSENTE")
                st.session_state.estado_asistencia[orden] = "AUSENTE"
                st.error(f"{nombre_asp} marcado como AUSENTE en el Parte Diario")
                st.rerun()
    else:
        search = st.text_input("🔍 Buscar aspirante:", placeholder="Nombre, DNI o CE", key="search_nov")
        if search.strip():
            s = search.strip().upper()
            res = df[(df['NOMBRE_COMPLETO'].str.contains(s, na=False)) | (df['DNI'].str.contains(s, na=False)) | (df['CE'].str.contains(s, na=False))]
            if not res.empty:
                for i, (_, r) in enumerate(res.head(5).iterrows()):
                    c1, c2 = st.columns([4, 1])
                    with c1: st.markdown(f"**{r['NOMBRE_COMPLETO']}** | DNI: {r['DNI']} | CE: {r['CE']}")
                    with c2:
                        if st.button("👆 Seleccionar", key=f"sel_{i}"):
                            st.session_state.sel_nov = r.to_dict()
                            st.rerun()

        if st.session_state.sel_nov:
            data = st.session_state.sel_nov
            orden = int(data["ORDEN_LIMP"]) # Forzamos entero para el presentismo
            nombre_asp = data.get('NOMBRE_COMPLETO', data.get('nombre', 'Aspirante'))
            estado_act = st.session_state.estado_asistencia.get(orden, "PRESENTE")

            col_btn_p, col_btn_a, col_btn_c = st.columns([2, 2, 1])
            with col_btn_p:
                if st.button("✅ PRESENTE", type="primary", use_container_width=True, key="btn_pres_reg"):
                    # SOLUCIÓN: Guarda el presentismo de forma independiente como ENTERO sin tocar la base de novedades
                    actualizar_asistencia(FECHA_STR, orden, "PRESENTE")
                    st.session_state.estado_asistencia[orden] = "PRESENTE"
                    st.success(f"{nombre_asp} marcado como PRESENTE")
                    st.rerun()
            with col_btn_a:
                if st.button("❌ AUSENTE", type="secondary", use_container_width=True, key="btn_aus_reg"):
                    # SOLUCIÓN: Guarda el presentismo de forma independiente como ENTERO sin tocar la base de novedades
                    actualizar_asistencia(FECHA_STR, orden, "AUSENTE")
                    st.session_state.estado_asistencia[orden] = "AUSENTE"
                    st.error(f"{nombre_asp} marcado como AUSENTE")
                    st.rerun()
            with col_btn_c:
                if st.button("🔄", use_container_width=True, key="btn_clear_sel"):
                    st.session_state.sel_nov = None
                    st.rerun()

    if data is not None:
        st.markdown("### ⚙️ Detalles de Novedad")
        c1, c2 = st.columns(2)
        with c1:
            opts = ["ART", "DAF", "LES", "SSD", "COMISIÓN", "AUTORIZADO", "ENTRANTE GUARDIA DIURNA", "ENTRANTE GUARDIA NOCTURNA", "DESCANSO DE GUARDIA"]
            current_estado = data.get('estado', "ART")
            idx_opts = opts.index(current_estado) if current_estado in opts else 0
            est = st.selectbox("Situación:", opts, index=idx_opts, key="sel_estado")
        with c2:
            det = st.text_input("Detalle:", value=data.get('detalle', ''), key="txt_detalle")

        cf1, cf2 = st.columns(2)
        with cf1:
            fi = st.date_input("Desde:", value=datetime.now(), key="date_ini_global").strftime('%d%b%y').upper()
        with cf2:
            sin_fin = st.checkbox("Sin término", value=(data.get('fecha_fin') == "N/O"), key="chk_sintermino")
            ff = "N/O" if sin_fin else st.date_input("Hasta:", value=datetime.now(), key="date_fin_global").strftime('%d%b%y').upper()

        b1, b2 = st.columns([3, 1])
        with b1:
            if es_edicion:
                if st.button("💾 Guardar Cambios", type="primary", use_container_width=True, key="btn_save_edit"):
                    actualizar_novedad(data['id'], {"estado": est, "detalle": det.upper(), "fecha_ini": fi, "fecha_fin": ff})
                    st.session_state.novedades_lista = obtener_novedades()
                    st.session_state.editando_idx = None
                    st.rerun()
            else:
                if st.button("💾 Grabar Novedad", use_container_width=True, key="btn_save_new"):
                    agregar_novedad({
                        "orden": int(data["ORDEN_LIMP"]), "grado": data["GRADO"], "nombre": nombre_asp,
                        "dni": data["DNI"], "ce": data["CE"], "aula": data["AULA"], "estado": est,
                        "detalle": det.upper(), "fecha_ini": fi, "fecha_fin": ff
                    })
                    st.session_state.novedades_lista = obtener_novedades()
                    st.session_state.sel_nov = None
                    st.rerun()
        with b2:
            if st.button("🚫 Cancelar", use_container_width=True, key="btn_cancel"):
                st.session_state.editando_idx = None
                st.session_state.sel_nov = None
                st.rerun()

    if st.session_state.novedades_lista:
        st.subheader("📋 Novedades Registradas")
        for idx, nov in enumerate(st.session_state.novedades_lista):
            col_datos, col_edit, col_borrar = st.columns([6, 1, 1])
            with col_datos:
                est_limpio = str(nov['estado']).replace("<span>", "").replace("</span>", "")
                
                # Buscamos el estado de presentismo real para mostrarlo al lado del nombre
                asis_actual = st.session_state.estado_asistencia.get(int(nov['orden']), "AUSENTE")
                color_asis = "🟢 PRESENTE" if asis_actual == "PRESENTE" else "🔴 AUSENTE"
                
                st.markdown(f"**{nov['nombre']}** | **[{est_limpio}]** | {color_asis} | Aula: {nov['aula']}")
                st.caption(f"📅 {nov['fecha_ini']} → {nov['fecha_fin']} | {nov['detalle']} (DNI: {nov['dni']})")
            with col_edit:
                if st.button("✏️", key=f"edit_{idx}"):
                    st.session_state.editando_idx = idx
                    st.rerun()
            with col_borrar:
                if st.button("🗑️", key=f"del_{idx}"):
                    # Aquí el usuario SÍ quiere borrar explícitamente la novedad con el tacho de basura
                    eliminar_novedad(nov['id'])
                    st.session_state.novedades_lista = obtener_novedades()
                    st.rerun()
# --- TAB: SEGUIMIENTO ---
with tab_seg:
    st.subheader("Control de Ingreso/Egreso - Seguimiento Diario")
    turno_act = st.radio("Seleccionar Turno:", ["🌅 MAÑANA", "🌆 TARDE"], horizontal=True, label_visibility="collapsed")
    prefijo = "m" if turno_act == "🌅 MAÑANA" else "t"
    
    for aula in AULAS_UNICAS:
        cfg = st.session_state.estado_aulas[aula]
        estado_key = f"estado_{prefijo}"
        salida_key = f"salida_{prefijo}"
        ubic_key = f"ubicacion_{prefijo}"
        
        alumnos = df[df['AULA'] == aula]
        total = len(alumnos)
        ausentes = sum(1 for n in st.session_state.novedades_lista if n['aula'] == aula and n['estado'] in ['ART', 'DAF', 'LES'])
        
        c1, c2, c3, c4 = st.columns([2, 2, 4, 1])
        with c1:
            st.markdown(f"**{aula}** (Tot: {total} | Aus: {ausentes})")
        with c2:
            is_inside = cfg[estado_key] == 'EN INSTITUTO'
            color = "green" if is_inside else "red"
            st.markdown(f"<div style='color:{color}; font-weight:bold;'>{cfg[estado_key]}</div>", unsafe_allow_html=True)
        with c3:
            if is_inside:
                ubicacion_actual = cfg.get(ubic_key, 'EN AULA')
                loc_cols = st.columns(4)
                locations = [("🏫 AULA", "EN AULA"), ("🏃 URF", "URF"), ("🏋️ ED. FÍS", "EDUCACIÓN FÍSICA"), ("🏛️ ACTIV.", "EN INSTITUTO")]
                for i, (label, value) in enumerate(locations):
                    with loc_cols[i]:
                        if st.button(label, key=f"loc_{prefijo}_{aula}_{i}", type="primary" if ubicacion_actual == value else "secondary"):
                            st.session_state.estado_aulas[aula][ubic_key] = value
                            guardar_estado_aula(FECHA_STR, aula, cfg['estado_m'], cfg['estado_t'], cfg.get('salida_m'), cfg.get('salida_t'), cfg.get('ubicacion_m', 'EN AULA'), cfg.get('ubicacion_t', 'EN AULA'))
                            st.rerun()
        with c4:
            if is_inside:
                if st.button("🚪 Retirar", key=f"out_{prefijo}_{aula}"):
                    st.session_state.estado_aulas[aula][estado_key] = 'FUERA'
                    st.session_state.estado_aulas[aula][salida_key] = datetime.now().strftime("%H:%M")
                    guardar_estado_aula(FECHA_STR, aula, cfg['estado_m'], cfg['estado_t'], cfg.get('salida_m'), cfg.get('salida_t'), cfg.get('ubicacion_m', 'EN AULA'), cfg.get('ubicacion_t', 'EN AULA'))
                    st.rerun()
            else:
                if st.button("🔙 Reingresar", key=f"in_{prefijo}_{aula}"):
                    st.session_state.estado_aulas[aula][estado_key] = 'EN INSTITUTO'
                    st.session_state.estado_aulas[aula][salida_key] = None
                    guardar_estado_aula(FECHA_STR, aula, cfg['estado_m'], cfg['estado_t'], cfg.get('salida_m'), cfg.get('salida_t'), cfg.get('ubicacion_m', 'EN AULA'), cfg.get('ubicacion_t', 'EN AULA'))
                    st.rerun()

# --- TAB: ALMUERZO ---
with tab_alm:
    st.subheader("Control de Personal que Almuerza")
    search_a = st.text_input("🔍 Buscar aspirante para almuerzo:", placeholder="Nombre o DNI", key="search_alm")
    if search_a.strip():
        s = search_a.strip().upper()
        res = df[(df['NOMBRE_COMPLETO'].str.contains(s, na=False)) | (df['DNI'].str.contains(s, na=False))]
        if not res.empty:
            for i, (_, r) in enumerate(res.head(5).iterrows()):
                c1, c2 = st.columns([4, 1])
                with c1: st.markdown(f"**{r['NOMBRE_COMPLETO']}** | {r['AULA']}")
                with c2:
                    if r['ORDEN_LIMP'] not in st.session_state.lista_almuerzo:
                        if st.button("➕ Marcar", key=f"malm_{i}"):
                            st.session_state.lista_almuerzo.add(r['ORDEN_LIMP'])
                            agregar_almuerzo(FECHA_STR, r['ORDEN_LIMP'])
                            st.rerun()

    st.divider()
    if st.session_state.lista_almuerzo:
        st.success(f"**Total Almuerzan:** {len(st.session_state.lista_almuerzo)}")
        df_lista = df[df['ORDEN_LIMP'].isin(st.session_state.lista_almuerzo)].sort_values('ORDEN_LIMP')
        for idx, row in df_lista.iterrows():
            col_info, col_btn = st.columns([4, 1])
            with col_info: st.markdown(f"**{row['ORDEN_LIMP']}** - {row['NOMBRE_COMPLETO']} | {row['AULA']}")
            with col_btn:
                if st.button("❌ Quitar", key=f"quit_{row['ORDEN_LIMP']}"):
                    st.session_state.lista_almuerzo.discard(row['ORDEN_LIMP'])
                    quitar_almuerzo(FECHA_STR, row['ORDEN_LIMP'])
                    st.rerun()
        
        # 📥 SECCIÓN EXCEL RACIONAMIENTO (Fijo adentro de Almuerzo)
        st.divider()
        wb_rac = openpyxl.Workbook()
        ws_rac = wb_rac.active
        ws_rac.title = "RACIONAMIENTO"
        ws_rac.merge_cells('A1:F1')
        ws_rac['A1'] = "PARTE DE RACIONAMIENTO - ESCUADRÓN H"
        ws_rac['A1'].font = Font(bold=True, size=14, color="FFFFFF")
        ws_rac['A1'].fill = PatternFill(start_color="E65100", end_color="E65100", fill_type="solid")
        ws_rac['A1'].alignment = Alignment(horizontal="center", vertical="center")
        ws_rac['A2'] = f"Fecha: {st.session_state.fecha_reporte.strftime('%d/%m/%Y')} | Total: {len(st.session_state.lista_almuerzo)}"
        ws_rac['A2'].font = Font(italic=True)
        
        headers_r = ["ORDEN", "NOMBRE COMPLETO", "GRADO", "CE", "DNI", "AULA"]
        for col_idx, h in enumerate(headers_r, 1):
            cell = ws_rac.cell(row=4, column=col_idx, value=h)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
        
        r_row = 5
        for orden in sorted(st.session_state.lista_almuerzo):
            p = df[df['ORDEN_LIMP'] == orden].iloc[0]
            ws_rac.cell(row=r_row, column=1, value=p['ORDEN_LIMP'])
            ws_rac.cell(row=r_row, column=2, value=p['NOMBRE_COMPLETO'])
            ws_rac.cell(row=r_row, column=3, value=p['GRADO'])
            ws_rac.cell(row=r_row, column=4, value=p['CE'])
            ws_rac.cell(row=r_row, column=5, value=p['DNI'])
            ws_rac.cell(row=r_row, column=6, value=p['AULA'])
            r_row += 1
            
        buf_rac = io.BytesIO()
        wb_rac.save(buf_rac)
        buf_rac.seek(0)
        
        st.download_button(
            label="📥 DESCARGAR PARTE DE RACIONAMIENTO (EXCEL)",
            data=buf_rac,
            file_name=f"RACIONAMIENTO_{datetime.now().strftime('%d%m%Y_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

# --- TAB: PLAN DE LLAMADA ---
with tab_plan:
    st.subheader("Plan de Llamada Interno de la Subunidad")
    st.info("Directorio de contactos de emergencia de la Guardia.")

# --- TAB: RESUMEN (PROCESAMIENTO Y GENERACIÓN DE EXCEL GENERAL) ---
with tab_res:
    st.subheader("📊 Generación del Parte Diario General")

    # 1.PROCESAMIENTO VIVO DE ASISTENCIA, NOVEDADES Y HORARIOS POR AULA
    totales_aulas = {}
    novedades_dict = {int(n["orden"]): n for n in st.session_state.novedades_lista}

    # Inicializar contadores por aula basados en la configuración viva
    for aula_id, conf in st.session_state.config_aulas.items():
        totales_aulas[aula_id] = {
            "total": 0,
            "presentes": 0,
            "ausentes": 0,
            "horario": conf.get("horario", "0600 a 0620 hs"),
            "tipo_ingreso": conf.get("tipo_ingreso", "Normal")
        }

    # Procesar cada aspirante cruzando su asistencia real e independiente
    for _, row in df.iterrows():
        ord_val = int(row["ORDEN_LIMP"])
        aula_val = str(row["AULA"]).strip().upper()
        
        if aula_val in totales_aulas:
            totales_aulas[aula_val]["total"] += 1
            # El presentismo manual manda sobre el reporte
            estado_asis = st.session_state.estado_asistencia.get(ord_val, "PRESENTE")
            
            if estado_asis == "PRESENTE":
                totales_aulas[aula_val]["presentes"] += 1
            else:
                totales_aulas[aula_val]["ausentes"] += 1

    # 2. AGRUPACIÓN DINÁMICA DE INGRESOS DIFERENCIALES CON CONTEO NUMÉRICO
    grupos_horarios = {}
    for aula_id, datos in totales_aulas.items():
        if datos["presentes"] > 0:  # Solo sumamos personal que efectivamente ingresa (Presentes)
            llave_horario = datos["horario"]
            if llave_horario not in grupos_horarios:
                grupos_horarios[llave_horario] = {"aulas": [], "personal_count": 0}
            grupos_horarios[llave_horario]["aulas"].append(aula_id)
            grupos_horarios[llave_horario]["personal_count"] += datos["presentes"]

    # 3. CONSTRUCCIÓN DEL DOCUMENTO EXCEL (OpenPyXL)
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from io import BytesIO

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PARTE"
    ws.views.sheetView[0].showGridLines = True

    # Encabezado Principal
    ws.merge_cells("A1:K1")
    ws["A1"] = f"PARTE DIARIO DEL ESCUADRÓN H - {FECHA_STR}"
    ws["A1"].font = Font(name="Calibri", size=14, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    # Cuadro de Fuerza Superior
    headers_fuerza = ["TOTAL", "PRESENTES", "AUSENTES", "NOVEDADES"]
    for col_idx, text in enumerate(headers_fuerza, start=2):
        cell = ws.cell(row=3, column=col_idx, value=text)
        cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="1E4620", end_color="1E4620", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    total_general = sum(d["total"] for d in totales_aulas.values())
    total_presentes = sum(d["presentes"] for d in totales_aulas.values())
    total_ausentes = sum(d["ausentes"] for d in totales_aulas.values())
    total_novedades = len(st.session_state.novedades_lista)

    valores_fuerza = [total_general, total_presentes, total_ausentes, total_novedades]
    for col_idx, val in enumerate(valores_fuerza, start=2):
        cell = ws.cell(row=4, column=col_idx, value=val)
        cell.font = Font(name="Calibri", size=11, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[3].height = 18
    ws.row_dimensions[4].height = 20

    # Cabecera de la Tabla de Personal
    headers_tabla = ["ORDEN", "GRADO", "APELLIDO Y NOMBRE", "DNI", "CE", "NOVEDAD", "PRESENTE / AUSENTE", "DESDE", "HASTA", "AULA"]
    for col_idx, text in enumerate(headers_tabla, start=1):
        cell = ws.cell(row=8, column=col_idx, value=text)
        cell.font = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="1E4620", end_color="1E4620", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[8].height = 22

    # Volcado de filas (Recorre los 139 aspirantes manteniendo independencia de estados)
    linea_actual = 9
    border_fino = Border(
        left=Side(style='thin', color='DDDDDD'), right=Side(style='thin', color='DDDDDD'),
        top=Side(style='thin', color='DDDDDD'), bottom=Side(style='thin', color='DDDDDD')
    )

    for _, row in df.iterrows():
        ord_int = int(row["ORDEN_LIMP"])
        
        # Consultar estados independientes correspondientes al número entero
        asis_real = st.session_state.estado_asistencia.get(ord_int, "PRESENTE")
        nov_viva = novedades_dict.get(ord_int, None)
        
        nov_texto = nov_viva["estado"].replace("<span>", "").replace("</span>", "") if nov_viva else ""
        desde_texto = nov_viva["fecha_ini"] if nov_viva else ""
        hasta_texto = nov_viva["fecha_fin"] if nov_viva else ""

        ws.cell(row=linea_actual, column=1, value=ord_int).alignment = Alignment(horizontal="center")
        ws.cell(row=linea_actual, column=2, value=row["GRADO"]).alignment = Alignment(horizontal="center")
        ws.cell(row=linea_actual, column=3, value=row["NOMBRE_COMPLETO"]).alignment = Alignment(horizontal="left")
        ws.cell(row=linea_actual, column=4, value=int(row["DNI"])).alignment = Alignment(horizontal="center")
        ws.cell(row=linea_actual, column=5, value=row["CE"]).alignment = Alignment(horizontal="center")
        ws.cell(row=linea_actual, column=6, value=nov_texto).alignment = Alignment(horizontal="center")
        
        # Celda de Presentismo Independiente con formato condicional soft
        cell_asis = ws.cell(row=linea_actual, column=7, value=asis_real)
        cell_asis.alignment = Alignment(horizontal="center")
        if asis_real == "AUSENTE":
            cell_asis.fill = PatternFill(start_color="FDE8E8", end_color="FDE8E8", fill_type="solid")
            cell_asis.font = Font(name="Calibri", color="9B1C1C", bold=True)
        else:
            cell_asis.fill = PatternFill(start_color="E1F5FE", end_color="E1F5FE", fill_type="solid")
            cell_asis.font = Font(name="Calibri", color="0288D1")

        ws.cell(row=linea_actual, column=8, value=desde_texto).alignment = Alignment(horizontal="center")
        ws.cell(row=linea_actual, column=9, value=hasta_texto).alignment = Alignment(horizontal="center")
        ws.cell(row=linea_actual, column=10, value=row["AULA"]).alignment = Alignment(horizontal="center")

        for c in range(1, 11):
            ws.cell(row=linea_actual, column=c).border = border_fino
            if c != 7:
                ws.cell(row=linea_actual, column=c).font = Font(name="Calibri", size=10)
        
        ws.row_dimensions[linea_actual].height = 19
        linea_actual += 1

    # 4. SECCIÓN OBSERVACIONES GENERALES E INGRESOS DIFERENCIALES VIVOS
    linea_actual += 2
    obs_label = ws.cell(row=linea_actual, column=1, value="OBSERVACIONES GENERALES E INGRESOS")
    obs_label.font = Font(name="Calibri", size=11, bold=True)
    linea_actual += 1

    # Renderizar ingresos basados puramente en la configuración de la pestaña Configuración
    for horario, datos in grupos_horarios.items():
        aulas_str = ", ".join(sorted(datos["aulas"]))
        count_pers = datos["personal_count"]
        
        # Formato dinámico: Ingreso horario diferencial [HORARIO]: Aulas [X] (Total: [N] Aspirantes Presentes).
        texto_ingreso = f"• Ingreso horario diferencial {horario}: Aulas {aulas_str} (Total: {count_pers} Aspirantes Presentes)."
        
        cell_ingreso = ws.cell(row=linea_actual, column=1, value=texto_ingreso)
        cell_ingreso.font = Font(name="Calibri", size=10, italic=True)
        ws.merge_cells(start_row=linea_actual, start_column=1, end_row=linea_actual, end_column=10)
        linea_actual += 1

    # Ajuste automático del ancho de las columnas para evitar cortes de texto
    for col in ws.columns:
        max_len = 0
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        for cell in col:
            if cell.row > 1 and cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_len + 3, 11)
    ws.column_dimensions["C"].width = 32  # Ancho fijo para nombres

    # 5. DESCARGA DEL ARCHIVO GENERADO
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.download_button(
        label="📥 Descargar Parte Diario General (Excel)",
        data=output,
        file_name=f"PARTE_DIARIO_ESCUADRÓN_H_{FECHA_STR}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    # ==============================================================================
    # 5. GENERACIÓN PROFESIONAL EXCEL MILITAR DEL PARTE DIARIO (AL RAS DEL MARGEN)
    # ==============================================================================
    st.divider()
    
    total_efectivos = len(df)
    total_novedades = len(novedades_sistema)
    presentes = total_efectivos - total_novedades

    # Creamos el libro oficial
    buffer_general = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PARTE"
    ws.views.sheetView[0].showGridLines = True

    # Estilos Gendarmería
    font_titulo = Font(name="Calibri", size=14, bold=True)
    font_subseccion = Font(name="Calibri", size=11, bold=True)
    font_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_datos = Font(name="Calibri", size=10)
    fill_verde_gendarme = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
    thin_border = Border(left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'),
                         top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF'))

    # Encabezado Cuadro Fuerza
    ws.merge_cells('A1:J1')
    ws['A1'] = f"PARTE DIARIO DEL ESCUADRÓN H - {datetime.now().strftime('%d%b%y').upper()}"
    ws['A1'].font = font_titulo
    ws['A1'].alignment = Alignment(horizontal="center", vertical="center")

    ws['B3'] = "TOTAL"
    ws['C3'] = "PRESENTES"
    ws['D3'] = "AUSENTES"
    ws['E3'] = "NOVEDADES"
    for col in ['B3', 'C3', 'D3', 'E3']:
        ws[col].font = font_header
        ws[col].fill = fill_verde_gendarme
        ws[col].alignment = Alignment(horizontal="center")

    ws['B4'] = total_efectivos
    ws['C4'] = presentes
    ws['D4'] = total_novedades
    ws['E4'] = total_novedades
    for col in ['B4', 'C4', 'D4', 'E4']:
        ws[col].font = Font(name="Calibri", size=11, bold=True)
        ws[col].border = thin_border
        ws[col].alignment = Alignment(horizontal="center")

    # Tabla Cuerpo de Novedades
    ws.cell(row=7, column=1, value="NOVEDADES DEL PERSONAL").font = font_subseccion
    headers_n = ["ORDEN", "GRADO", "APELLIDO Y NOMBRE", "DNI", "CE", "NOVEDAD", "PRESENTE / AUSENTE", "DESDE", "HASTA", "AULA"]
    for idx, h in enumerate(headers_n, 1):
        cell = ws.cell(row=8, column=idx, value=h)
        cell.font = font_header
        cell.fill = fill_verde_gendarme
        cell.alignment = Alignment(horizontal="center", vertical="center")

    row_idx = 9
    for nov in novedades_sistema:
        match = df[df['ORDEN_LIMP'] == nov.get("orden")]
        grado_c = match.iloc[0].get("GRADO", "") if not match.empty else ""
        dni_c = match.iloc[0].get("DNI", "") if not match.empty else ""
        ce_c = match.iloc[0].get("CE", "") if not match.empty else ""
        aula_c = match.iloc[0].get("AULA", "") if not match.empty else ""

        ws.cell(row=row_idx, column=1, value=nov.get("orden")).alignment = Alignment(horizontal="center")
        ws.cell(row=row_idx, column=2, value=grado_c).alignment = Alignment(horizontal="center")
        ws.cell(row=row_idx, column=3, value=nov.get("nombre"))
        ws.cell(row=row_idx, column=4, value=dni_c).alignment = Alignment(horizontal="center")
        ws.cell(row=row_idx, column=5, value=ce_c).alignment = Alignment(horizontal="center")
        ws.cell(row=row_idx, column=6, value=nov.get("estado", "S/D"))
        ws.cell(row=row_idx, column=7, value="AUSENTE").alignment = Alignment(horizontal="center")
        ws.cell(row=row_idx, column=8, value=nov.get("fecha_ini", "")).alignment = Alignment(horizontal="center")
        ws.cell(row=row_idx, column=9, value=nov.get("fecha_fin", "")).alignment = Alignment(horizontal="center")
        ws.cell(row=row_idx, column=10, value=aula_c).alignment = Alignment(horizontal="center")
        
        for col_b in range(1, 11):
            ws.cell(row=row_idx, column=col_b).font = font_datos
            ws.cell(row=row_idx, column=col_b).border = thin_border
        row_idx += 1

    row_idx += 2
    ws.cell(row=row_idx, column=1, value="OBSERVACIONES GENERALES E INGRESOS").font = font_subseccion
    row_idx += 1
    
    items_defecto = [
        "• Ingreso horario diferencial 0600 a 0620 hs: Aulas 18 TM, 24 TM y 23 TT.",
        "• Ingreso horario diferencial 0810 hs: Aulas 23 TM y 28 TM.",
        "• Ingreso horario diferencial 0900 a 0915 hs: Aula 26 TM.",
        "• Descanso servicio de armas nocturno: 32 ASP III Año.",
        "• Ingreso diferencial AOP: Aula 7 TT y Aula 8 TT."
    ]
    for item in items_defecto:
        ws.cell(row=row_idx, column=1, value=item).font = Font(italic=True, size=10)
        row_idx += 1

    anchos = {"A": 10, "B": 12, "C": 35, "D": 15, "E": 12, "F": 22, "G": 22, "H": 14, "I": 14, "J": 12}
    for l, w in anchos.items():
        ws.column_dimensions[l].width = w

    wb.save(buffer_general)
    buffer_general.seek(0)

    # 📥 BOTÓN PRINCIPAL GENERAL VISIBLE
    st.download_button(
        label="📊 DESCARGAR PARTE DIARIO GENERAL (EXCEL)",
        data=buffer_general,
        file_name=f"PARTE_DIARIO_ESCUADRON_H_{datetime.now().strftime('%d%m%Y_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )
