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
tab_nov, tab_res, tab_part = st.tabs(["📋 Novedades", "📊 Resumen", "🔍 Novedades Particulares"])

# --- TAB: NOVEDADES PARTICULARES (CAMPÁÑAS MASIVAS) ---
with tab_part:
    import sqlite3
    st.subheader("📋 Gestión de Novedades Particulares y Campañas")

    # Funciones de persistencia interna
    def db_ejecutar(query, params=()):
        conn = sqlite3.connect("parte_diario.db")
        cursor = conn.cursor()
        cursor.execute(query, params)
        conn.commit()
        conn.close()

    def db_consultar(query, params=()):
        conn = sqlite3.connect("parte_diario.db")
        cursor = conn.cursor()
        cursor.execute(query, params)
        res = cursor.fetchall()
        conn.close()
        return res

    # Selectores para controlar el flujo de la pantalla
    if "creando_campaña" not in st.session_state:
        st.session_state.creando_campaña = False
    if "ver_campaña_id" not in st.session_state:
        st.session_state.ver_campaña_id = None

    # --- ACCIÓN: BOTÓN NUEVA CAMPÁÑA / NOVEDAD GENERAL ---
    if not st.session_state.creando_campaña and st.session_state.ver_campaña_id is None:
        if st.button("➕ Crear Nueva Novedad Particular / Campaña", type="primary", use_container_width=True):
            st.session_state.creando_campaña = True
            st.rerun()

    # --- VISTA 1: FORMULARIO DE CREACIÓN EN LOTE ---
    if st.session_state.creando_campaña:
        st.write("### Nueva Novedad por Grupo")
        nuevo_titulo = st.text_input("Título de la Novedad Particular:", placeholder="Ej: Presentación de documentación Obra Social").upper()
        nueva_desc = st.text_area("Detalles / Instrucciones:", placeholder="Ej: Deben entregar la cartilla firmada antes del viernes.").upper()
        
        st.write("**Seleccionar integrantes para asignar en lote:**")
        # Multiselect dinámico usando el universo de tu dataframe 'df'
        lista_aspirantes = df.sort_values(by="NOMBRE_COMPLETO")["NOMBRE_COMPLETO"].tolist()
        seleccionados = st.multiselect("Buscar y seleccionar los integrantes afectados:", lista_aspirantes)

        c_btn1, c_btn2 = st.columns(2)
        with c_btn1:
            if st.button("💾 Grabar Novedad y Asignar Personal", type="primary", use_container_width=True):
                if nuevo_titulo and seleccionados:
                    # 1. Insertar cabecera de campaña
                    conn = sqlite3.connect("parte_diario.db")
                    cursor = conn.cursor()
                    cursor.execute("INSERT INTO campañas_novedades (titulo, descripcion, estado_general) VALUES (?, ?, ?)", 
                                   (nuevo_titulo, nueva_desc, "PENDIENTE"))
                    campaña_id = cursor.lastrowid
                    
                    # 2. Insertar lote de aspirantes vinculados
                    for nombre_sel in seleccionados:
                        # Extraemos los datos del dataframe de alumnos
                        row_asp = df[df["NOMBRE_COMPLETO"] == nombre_sel].iloc[0]
                        cursor.execute("""
                            INSERT INTO novedades_particulares (campaña_id, orden, nombre, aula, estado_individual) 
                            VALUES (?, ?, ?, ?, ?)
                        """, (campaña_id, int(row_asp["ORDEN_LIMP"]), nombre_sel, str(row_asp["AULA"]), "PENDIENTE"))
                    
                    conn.commit()
                    conn.close()
                    st.success("Novedad masiva registrada con éxito.")
                    st.session_state.creando_campaña = False
                    st.rerun()
                else:
                    st.error("Por favor, completa el título y selecciona al menos un integrante.")
        with c_btn2:
            if st.button("🚫 Cancelar", use_container_width=True):
                st.session_state.creando_campaña = False
                st.rerun()

    # --- VISTA 2: EDICIÓN Y GESTIÓN INTERNA DE UNA CAMPAÑA SELECCIONADA ---
    elif st.session_state.ver_campaña_id is not None:
        camp_id = st.session_state.ver_campaña_id
        info_camp = db_consultar("SELECT titulo, descripcion, estado_general FROM campañas_novedades WHERE id=?", (camp_id,))
        
        if info_camp:
            tit, desc, est_gen = info_camp[0]
            st.markdown(f"## ⚙️ Editando: {tit}")
            
            # Campos de edición de cabecera directos
            edit_tit = st.text_input("Editar Título General:", value=tit).upper()
            edit_desc = st.text_area("Editar Detalles / Notas:", value=desc).upper()
            edit_est_gen = st.selectbox("Estado General del Grupo:", ["PENDIENTE", "EN PROCESO", "FINALIZADO"], index=["PENDIENTE", "EN PROCESO", "FINALIZADO"].index(est_gen))
            
            if st.button("💾 Actualizar Datos Generales", use_container_width=True):
                db_ejecutar("UPDATE campañas_novedades SET titulo=?, descripcion=?, estado_general=? WHERE id=?", (edit_tit, edit_desc, edit_est_gen, camp_id))
                st.toast("Datos generales actualizados.")
                st.rerun()

            st.write("---")
            st.write("### Integrantes Asignados y Estados Individuales")
            
            # Listar integrantes asignados a esta campaña
            integrantes = db_consultar("SELECT id, orden, nombre, aula, estado_individual FROM novedades_particulares WHERE campaña_id=?", (camp_id,))
            
            for item_id, item_ord, item_nom, item_aula, item_est_ind in integrantes:
                col_i_nom, col_i_est, col_i_acc = st.columns([4, 2, 2])
                with col_i_nom:
                    st.markdown(f"**{item_nom}** (Aula: {item_aula})")
                with col_i_est:
                    # Cambiar estado individual de este integrante en caliente
                    opciones_ind = ["PENDIENTE", "PRESENTÓ", "COMPLETO", "CON OBSERVACIÓN"]
                    idx_ind = opciones_ind.index(item_est_ind) if item_est_ind in opciones_ind else 0
                    nuevo_est_ind = st.selectbox(f"Estado ({item_ord}):", opciones_ind, index=idx_ind, key=f"est_ind_{item_id}", label_visibility="collapsed")
                    if nuevo_est_ind != item_est_ind:
                        db_ejecutar("UPDATE novedades_particulares SET estado_individual=? WHERE id=?", (nuevo_est_ind, item_id))
                        st.toast(f"Estado de {item_nom} actualizado.")
                with col_i_acc:
                    if st.button("❌ Quitar", key=f"del_ind_{item_id}", use_container_width=True):
                        db_ejecutar("DELETE FROM novedades_particulares WHERE id=?", (item_id,))
                        st.rerun()

            st.write("---")
            # Agregar un integrante más que se hayan olvidado al armar el grupo
            lista_completa = df.sort_values(by="NOMBRE_COMPLETO")["NOMBRE_COMPLETO"].tolist()
            nuevo_miembro = st.selectbox("➕ Agregar rezagado a esta novedad:", ["-- SELECCIONAR --"] + lista_completa)
            if nuevo_miembro != "-- SELECCIONAR --":
                row_rez = df[df["NOMBRE_COMPLETO"] == nuevo_miembro].iloc[0]
                # Validar que no esté ya metido
                existe = db_consultar("SELECT id FROM novedades_particulares WHERE campaña_id=? AND orden=?", (camp_id, int(row_rez["ORDEN_LIMP"])))
                if not existe:
                    db_ejecutar("INSERT INTO novedades_particulares (campaña_id, orden, nombre, aula, estado_individual) VALUES (?, ?, ?, ?, ?)",
                                (camp_id, int(row_rez["ORDEN_LIMP"]), nuevo_miembro, str(row_rez["AULA"]), "PENDIENTE"))
                    st.rerun()
                else:
                    st.warning("El integrante ya se encuentra en esta lista.")

            if st.button("⬅️ Volver al Listado General", use_container_width=True):
                st.session_state.ver_campaña_id = None
                st.rerun()

    # --- VISTA 3: PANTALLA PRINCIPAL CON EL DASHBOARD DE NOVEDADES PARTICULARES ---
    else:
        campañas = db_consultar("SELECT id, titulo, descripcion, estado_general FROM campañas_novedades")
        
        if Hospital_campañas := campañas:
            for c_id, c_tit, c_desc, c_est in Hospital_campañas:
                # Contar cuántos efectivos integran la novedad y cuántos ya resolvieron
                totales = db_consultar("SELECT COUNT(*), SUM(CASE WHEN estado_individual IN ('PRESENTÓ', 'COMPLETO') THEN 1 ELSE 0 END) FROM novedades_particulares WHERE campaña_id=?", (c_id,))
                total_integrantes = totales[0][0] if totales else 0
                resueltos = totales[0][1] if totales and totales[0][1] is not None else 0

                # Tarjeta de visualización prolija por cada campaña
                with st.container():
                    c_col1, c_col2, c_col3 = st.columns([5, 2, 2])
                    with c_col1:
                        st.markdown(f"### 📂 {c_tit}")
                        st.markdown(f"*{c_desc}*")
                        st.caption(f"📊 Progreso de Efectivos: **{resueltos}/{total_integrantes} procesados**")
                    with c_col2:
                        st.info(f"Estado: {c_est}")
                    with c_col3:
                        c_sub1, c_sub2 = st.columns(2)
                        with c_sub1:
                            if st.button("👁️/✏️", key=f"ver_camp_{c_id}", help="Ver detalles y editar integrantes"):
                                st.session_state.ver_campaña_id = c_id
                                st.rerun()
                        with c_sub2:
                            if st.button("🗑️", key=f"del_camp_{c_id}", help="Eliminar campaña completa"):
                                db_ejecutar("DELETE FROM campañas_novedades WHERE id=?", (c_id,))
                                db_ejecutar("DELETE FROM novedades_particulares WHERE campaña_id=?", (c_id,))
                                st.rerun()
                st.markdown("---")
        else:
            st.info("No hay campañas de novedades particulares creadas hasta el momento.")
