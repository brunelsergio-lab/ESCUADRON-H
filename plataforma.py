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
        orden = data["orden"]
        nombre_asp = data["nombre"]
        estado_act = st.session_state.estado_asistencia.get(orden, "PRESENTE")

        col_btn_p, col_btn_a = st.columns(2)
        with col_btn_p:
            if st.button("✅ MARCAR PRESENTE", type="primary", use_container_width=True, key="btn_pres_edit"):
                actualizar_asistencia(FECHA_STR, orden, "PRESENTE")
                st.session_state.estado_asistencia[orden] = "PRESENTE"
                import sqlite3
                conn = sqlite3.connect("parte_diario.db")
                conn.execute("DELETE FROM novedades WHERE orden=?", (int(orden),))
                conn.commit()
                conn.close()
                st.session_state.novedades_lista = obtener_novedades()
                st.rerun()
        with col_btn_a:
            if st.button("❌ MARCAR AUSENTE", type="secondary", use_container_width=True, key="btn_aus_edit"):
                actualizar_asistencia(FECHA_STR, orden, "AUSENTE")
                st.session_state.estado_asistencia[orden] = "AUSENTE"
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
            orden = data["ORDEN_LIMP"]
            nombre_asp = data.get('NOMBRE_COMPLETO', data.get('nombre', 'Aspirante'))
            estado_act = st.session_state.estado_asistencia.get(orden, "PRESENTE")

            col_btn_p, col_btn_a, col_btn_c = st.columns([2, 2, 1])
            with col_btn_p:
                if st.button("✅ PRESENTE", type="primary", use_container_width=True, key="btn_pres_reg"):
                    actualizar_asistencia(FECHA_STR, orden, "PRESENTE")
                    st.session_state.estado_asistencia[orden] = "PRESENTE"
                    import sqlite3
                    conn = sqlite3.connect("parte_diario.db")
                    conn.execute("DELETE FROM novedades WHERE orden=?", (int(orden),))
                    conn.commit()
                    conn.close()
                    st.session_state.novedades_lista = obtener_novedades()
                    st.rerun()
            with col_btn_a:
                if st.button("❌ AUSENTE", type="secondary", use_container_width=True, key="btn_aus_reg"):
                    actualizar_asistencia(FECHA_STR, orden, "AUSENTE")
                    st.session_state.estado_asistencia[orden] = "AUSENTE"
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
                # CORRECCIÓN AQUÍ: Limpiamos las etiquetas html eliminando las strings '<span>' y '</span>' de la visualización
                est_limpio = str(nov['estado']).replace("<span>", "").replace("</span>", "")
                st.markdown(f"**{nov['nombre']}** | **[{est_limpio}]** | DNI: {nov['dni']} | Aula: {nov['aula']}")
                st.caption(f"📅 {nov['fecha_ini']} → {nov['fecha_fin']} | {nov['detalle']}")
            with col_edit:
                if st.button("✏️", key=f"edit_{idx}"):
                    st.session_state.editando_idx = idx
                    st.rerun()
            with col_borrar:
                if st.button("🗑️", key=f"del_{idx}"):
                    eliminar_novedad(nov['id'])
                    st.session_state.estado_asistencia[nov['orden']] = "PRESENTE"
                    actualizar_asistencia(FECHA_STR, nov['orden'], "PRESENTE")
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

# --- TAB: RESUMEN (AQUÍ CORREGIMOS EL PARTE GENERAL) ---
with tab_res:
    st.subheader("Resumen General y Novedades")

    # Reordenar y sincronizar orden general numérico del Excel original
    df = df.reset_index(drop=True)
    df = df.sort_values("ORDEN_LIMP").reset_index(drop=True)
    df["ORDEN_GENERAL"] = range(1, len(df) + 1)

    # Resumen estructurado por aula en la UI
    data_aulas = []
    for aula in AULAS_UNICAS:
        cfg = st.session_state.estado_aulas[aula]
        alumnos_aula = df[df['AULA'] == aula]
        total_aula = len(alumnos_aula)
        ausentes_aula = len({r['ORDEN_LIMP'] for _, r in alumnos_aula.iterrows() if st.session_state.estado_asistencia.get(r['ORDEN_LIMP']) == "AUSENTE"})
        presentes_aula = total_aula - ausentes_aula
        almuerzan_aula = sum(1 for _, r in alumnos_aula.iterrows() if r['ORDEN_LIMP'] in st.session_state.lista_almuerzo)

        data_aulas.append({
            "Aula": aula, "Total": total_aula, "Presentes": presentes_aula, "Ausentes": ausentes_aula,
            "Almuerzan": almuerzan_aula, "Estado": cfg["estado_m"]
        })

    st.dataframe(pd.DataFrame(data_aulas), use_container_width=True, hide_index=True)
    st.divider()

    # Tabla en pantalla de personal con Novedades activas
    st.markdown("### 📋 Personal en Situación de Ausencia / Licencia")
    data_ausentes = []
    novedades_sistema = st.session_state.get("novedades_lista", [])

    for nov in novedades_sistema:
        orden_limp = nov.get("orden")
        if orden_limp is not None:
            match = df[df['ORDEN_LIMP'] == orden_limp]
            if not match.empty:
                alumno = match.iloc[0]
                data_ausentes.append({
                    "Orden": int(alumno.get("ORDEN_GENERAL", 0)),
                    "Nombre": alumno.get("NOMBRE_COMPLETO", "S/N"),
                    "Motivo": nov.get("estado", "S/D"),
                    "Desde": nov.get("fecha_ini", ""),
                    "Hasta": nov.get("fecha_fin", "")
                })

    df_ausentes = pd.DataFrame(data_ausentes)
    if not df_ausentes.empty:
        st.dataframe(df_ausentes.sort_values("Orden"), use_container_width=True, hide_index=True)
    else:
        st.success("Sin personal ausente en los registros.")

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
