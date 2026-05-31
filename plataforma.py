# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font
from datetime import datetime
import os

# 🔹 IMPORTS CORREGIDOS (NO OMITIR NINGUNA FUNCIÓN)
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
# ... imports arriba ...

st.set_page_config(page_title="Gestión de Parte Diario - Escuadrón H", layout="wide")
# ↑↑↑ BUSCÁ ESTA LÍNEA ↑↑↑

# ==============================================================================
# 🎖️ PEGÁ EL BLOQUE INSTITUCIONAL ACÁ (justo después de set_page_config)
# ==============================================================================
st.markdown("""
<style>
/* ... todo el CSS ... */
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="header-box">
<!-- ... encabezado HTML ... -->
</div>
""", unsafe_allow_html=True)

st.components.v1.html("""
<script>
// ... reloj JS ...
</script>
""", height=0)
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

# ==============================================================================
# 2. INICIALIZACIÓN & DB
# ==============================================================================

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

# Novedades
if 'novedades_lista' not in st.session_state:
    st.session_state.novedades_lista = obtener_novedades()

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
    db_hor = obtener_horarios()
    st.session_state.horarios_config = {}
    for aula in AULAS_UNICAS:
        st.session_state.horarios_config[aula] = db_hor.get(aula, {
            "ent_m": "06:00", "sal_m": "12:00", "ent_t": "13:00", "sal_t": "19:00"
        })

# Asistencia diaria
if 'estado_asistencia' not in st.session_state:
    st.session_state.estado_asistencia = obtener_asistencia(FECHA_STR)

# Variables de control UI (¡ESTAS SON LAS QUE FALTABAN!)
if 'editando_idx' not in st.session_state:
    st.session_state.editando_idx = None

if 'sel_nov' not in st.session_state:
    st.session_state.sel_nov = None

# ==============================================================================
# 3. MÉTRICAS EN TIEMPO REAL (CON SINCRONIZACIÓN AUTOMÁTICA)
# ==============================================================================

# # 🔹 1. RECARGAR DATOS DESDE DB (Prioridad a session_state)
st.session_state.novedades_lista = obtener_novedades()
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
# 🔹 BOTÓN OCULTO PARA FORZAR SINCRONIZACIÓN (Por si las dudas)
if st.button("🔄 Sincronizar Datos", key="sync_btn", help="Fuerza la recarga desde base de datos"):
    st.session_state.novedades_lista = obtener_novedades()
    st.session_state.estado_asistencia = obtener_asistencia(FECHA_STR)
    st.session_state.lista_almuerzo = obtener_almuerzo(FECHA_STR)
    st.success("✅ Datos sincronizados correctamente")
    st.rerun()
# 🔹 BOTÓN DE EMERGENCIA: RESETEAR ASISTENCIA
if st.button("🚨 RESETEAR ASISTENCIA DEL DÍA", key="reset_asistencia", help="Pone a TODOS en PRESENTE"):
    import sqlite3
    conn = sqlite3.connect("parte_diario.db")
    conn.execute(f"DELETE FROM asistencia_diaria WHERE fecha='{FECHA_STR}'")
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
    "️ Configuración", "📝 Novedades", "🔍 Seguimiento", "️ Almuerzo", "📞 Plan de Llamada", "📊 Resumen"
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
        st.markdown(f"### Ajustar Asistencia para: **{data['nombre']}**")
        
        orden = data["orden"]
        nombre_asp = data["nombre"]
        estado_act = st.session_state.estado_asistencia.get(orden, "PRESENTE")

        col_btn_p, col_btn_a = st.columns(2)
        
        with col_btn_p:
            if st.button("✅ MARCAR PRESENTE", type="primary", use_container_width=True, key="btn_pres_edit", disabled=(estado_act=="PRESENTE")):
                actualizar_asistencia(FECHA_STR, orden, "PRESENTE")
                st.session_state.estado_asistencia[orden] = "PRESENTE"
                
                # Borra la novedad para que impacte en el numérico superior
                import sqlite3
                conn = sqlite3.connect("parte_diario.db")
                conn.execute("DELETE FROM novedades WHERE orden=?", (int(orden),))
                conn.commit()
                conn.close()
                
                st.session_state.novedades_lista = obtener_novedades()
                st.toast(f"✅ {nombre_asp} → PRESENTE y novedad quitada")
                st.rerun()
                
        with col_btn_a:
            if st.button("❌ MARCAR AUSENTE", type="secondary", use_container_width=True, key="btn_aus_edit", disabled=(estado_act=="AUSENTE")):
                actualizar_asistencia(FECHA_STR, orden, "AUSENTE")
                st.session_state.estado_asistencia[orden] = "AUSENTE"
                st.toast(f"❌ {nombre_asp} → AUSENTE")
                st.rerun()
        st.divider()

    else:
        # Lógica de búsqueda y selección (Modo Registro)
        search = st.text_input("🔍 Buscar aspirante:", placeholder="Nombre, DNI o CE", key="search_nov")
        if search.strip():
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
                            st.session_state.sel_nov = r.to_dict()
                            st.rerun()

        if st.session_state.sel_nov:
            data = st.session_state.sel_nov
            orden = data["ORDEN_LIMP"]
            nombre_asp = data.get('NOMBRE_COMPLETO', data.get('nombre', 'Aspirante'))
            estado_act = st.session_state.estado_asistencia.get(orden, "PRESENTE")

            st.divider()
            st.markdown(f"### 🎯 Asistencia Manual: **{nombre_asp}** | Estado actual: `{estado_act}`")

            col_btn_p, col_btn_a, col_btn_c = st.columns([2, 2, 1])

            with col_btn_p:
                if st.button("✅ PRESENTE", type="primary", use_container_width=True, key="btn_pres_reg", disabled=(estado_act=="PRESENTE")):
                    actualizar_asistencia(FECHA_STR, orden, "PRESENTE")
                    st.session_state.estado_asistencia[orden] = "PRESENTE"
                    
                    # Borra la novedad de la base de datos para limpiar el ausente
                    import sqlite3
                    conn = sqlite3.connect("parte_diario.db")
                    conn.execute("DELETE FROM novedades WHERE orden=?", (int(orden),))
                    conn.commit()
                    conn.close()
                    
                    st.session_state.novedades_lista = obtener_novedades()
                    st.toast(f"✅ {nombre_asp} → PRESENTE y novedad quitada")
                    st.rerun()

            with col_btn_a:
                if st.button("❌ AUSENTE", type="secondary", use_container_width=True, key="btn_aus_reg", disabled=(estado_act=="AUSENTE")):
                    actualizar_asistencia(FECHA_STR, orden, "AUSENTE")
                    st.session_state.estado_asistencia[orden] = "AUSENTE"
                    st.toast(f"❌ {nombre_asp} → AUSENTE")
                    st.rerun()

            with col_btn_c:
                if st.button("🔄", use_container_width=True, help="Cambiar selección", key="btn_clear_sel"):
                    st.session_state.sel_nov = None
                    st.rerun()
            st.divider()

    # 🔹 FORMULARIO DE NOVEDAD (Visible si hay data, tanto en edición como en registro)
    if data is not None:
        st.markdown("### ⚙️ Detalles de Novedad")
        c1, c2 = st.columns(2)
        with c1:
            opts = [
    "ART",
    "DAF",
    "LES",
    "SSD",
    "COMISIÓN",
    "AUTORIZADO",
    "ENTRANTE GUARDIA DIURNA",
    "ENTRANTE GUARDIA NOCTURNA",
    "DESCANSO DE GUARDIA"
]
 
            current_estado = data.get('estado', "ART")
            idx_opts = opts.index(current_estado) if current_estado in opts else 0
            est = st.selectbox("Situación:", opts, index=idx_opts, key="sel_estado")
        with c2:
            det = st.text_input("Detalle:", value=data.get('detalle', ''), key="txt_detalle")

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
                    actualizar_novedad(nov_id, {"estado": est, "detalle": det.upper(), "fecha_ini": fi, "fecha_fin": ff})
                    st.session_state.novedades_lista = obtener_novedades()
                    st.session_state.editando_idx = None
                    st.success("✅ Novedad y asistencia actualizadas")
                    st.rerun()
            else:
                if st.button("💾 Grabar Novedad", use_container_width=True, key="btn_save_new"):
                    nombre_asp = data.get('NOMBRE_COMPLETO', data.get('nombre', 'Aspirante'))
                    agregar_novedad({
                        "orden": int(data["ORDEN_LIMP"]), "grado": data["GRADO"],
                        "nombre": nombre_asp, "dni": data["DNI"], "ce": data["CE"],
                        "aula": data["AULA"], "estado": est, "detalle": det.upper(),
                        "fecha_ini": fi, "fecha_fin": ff
                    })
                    st.session_state.novedades_lista = obtener_novedades()
                    st.session_state.sel_nov = None
                    st.success(f"✅ Novedad grabada para {nombre_asp}")
                    st.rerun()
        with b2:
            if st.button("🚫 Cancelar", use_container_width=True, key="btn_cancel"):
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
                st.caption(f"📅 {nov['fecha_ini']} → {nov['fecha_fin']} | 📝 {nov['detalle']}")
            with col_edit:
                if st.button("✏️", key=f"edit_{idx}", use_container_width=True, help="Editar"):
                    st.session_state.editando_idx = idx
                    st.rerun()
            with col_borrar:
                if st.button("🗑️", key=f"del_{idx}", use_container_width=True, help="Eliminar"):
                    # 1. Eliminar de la base de datos
                    eliminar_novedad(nov['id'])
                    
                    # 2. Sincronizar asistencia a PRESENTE
                    st.session_state.estado_asistencia[nov['orden']] = "PRESENTE"
                    actualizar_asistencia(FECHA_STR, nov['orden'], "PRESENTE")
                    
                    # 3. Actualizar la lista y refrescar
                    st.session_state.novedades_lista = obtener_novedades()
                    st.toast("Novedad eliminada y asistencia actualizada")
                    st.rerun()
            st.markdown("<hr style='margin: 5px 0px; border-color: #333;'>", unsafe_allow_html=True)
        
        if st.button("🗑️ Vaciar Todas las Novedades", type="secondary", key="btn_clear_all_nov"):
            vaciar_novedades()
            for n in st.session_state.novedades_lista:
                st.session_state.estado_asistencia[n['orden']] = "PRESENTE"
                actualizar_asistencia(FECHA_STR, n['orden'], "PRESENTE")
            st.session_state.novedades_lista = []
            st.toast("Todo limpio: Novedades y asistencia reiniciadas")
            st.rerun()

# --- TAB: SEGUIMIENTO ---
with tab_seg:
    st.subheader("Control de Ingreso/Egreso - Seguimiento Diario")
    
    # Selector de turno
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
            st.markdown(f"**{aula}**")
            st.caption(f"Total: {total} | Ausentes: {ausentes}")
            
        with c2:
            is_inside = cfg[estado_key] == 'EN INSTITUTO'
            color = "green" if is_inside else "red"
            icon = "" if is_inside else "🔴"
            st.markdown(f"<div style='color:{color}; font-weight:bold; font-size:1.1em;'>{icon} {cfg[estado_key]}</div>", unsafe_allow_html=True)
            if cfg[salida_key]: 
                st.caption(f"🕒 Salida: {cfg[salida_key]}")
                
        with c3:
            if is_inside:
                ubicacion_actual = cfg.get(ubic_key, 'EN AULA')
                loc_cols = st.columns(4)
                locations = [("🏫 AULA", "EN AULA"), ("🏃 URF", "URF"), ("🏋️ ED. FÍSICA", "EDUCACIÓN FÍSICA"), ("🏛️ ACTIV.", "EN INSTITUTO")]
                for i, (label, value) in enumerate(locations):
                    with loc_cols[i]:
                        is_active = (ubicacion_actual == value)
                        btn_type = "primary" if is_active else "secondary"
                        if st.button(label, key=f"loc_{prefijo}_{aula}_{i}", type=btn_type, use_container_width=True):
                            st.session_state.estado_aulas[aula][ubic_key] = value
                            guardar_estado_aula(FECHA_STR, aula, cfg['estado_m'], cfg['estado_t'], cfg.get('salida_m'), cfg.get('salida_t'), cfg.get('ubicacion_m', 'EN AULA'), cfg.get('ubicacion_t', 'EN AULA'))
                            st.rerun()
            else:
                st.markdown(" _Ubicación no disponible_")
                st.caption("Marcar 'EN INSTITUTO' para activar")
                
        with c4:
            if is_inside:
                if st.button("🚪 Retirar", key=f"out_{prefijo}_{aula}", use_container_width=True):
                    st.session_state.estado_aulas[aula][estado_key] = 'FUERA'
                    st.session_state.estado_aulas[aula][salida_key] = datetime.now().strftime("%H:%M")
                    guardar_estado_aula(FECHA_STR, aula, cfg['estado_m'], cfg['estado_t'], cfg.get('salida_m'), cfg.get('salida_t'), cfg.get('ubicacion_m', 'EN AULA'), cfg.get('ubicacion_t', 'EN AULA'))
                    st.rerun()
            else:
                if st.button("🔙 Reingresar", key=f"in_{prefijo}_{aula}", use_container_width=True):
                    st.session_state.estado_aulas[aula][estado_key] = 'EN INSTITUTO'
                    st.session_state.estado_aulas[aula][salida_key] = None
                    guardar_estado_aula(FECHA_STR, aula, cfg['estado_m'], cfg['estado_t'], cfg.get('salida_m'), cfg.get('salida_t'), cfg.get('ubicacion_m', 'EN AULA'), cfg.get('ubicacion_t', 'EN AULA'))
                    st.rerun()

# --- TAB: ALMUERZO ---
with tab_alm:
    st.subheader("Control de Personal que Almuerza")
    search = st.text_input("🔍 Buscar aspirante:", placeholder="Nombre, DNI o CE", key="search_alm")
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
                    st.toast(f"✅ {row['NOMBRE_COMPLETO']} removido")
                    st.rerun()
            st.markdown("<hr style='margin: 3px 0; border-color: #444;'>", unsafe_allow_html=True)
        if st.button("🗑️ Vaciar lista completa", type="secondary", key="clear_all_lunch"):
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
        
        headers = ["ORDEN", "NOMBRE COMPLETO", "GRADO", "CE", "DNI", "AULA"]
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=4, column=col, value=h)
            cell.font = Font(bold=True, color="FFFFFF", size=9)
            cell.fill = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
            cell.alignment = Alignment(horizontal="center")
            cell.border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
            
        row = 5
        for orden in sorted(st.session_state.lista_almuerzo):
            p = df[df['ORDEN_LIMP'] == orden].iloc[0]
            ws.cell(row=row, column=1, value=p['ORDEN_LIMP'])
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
        wb.save(output)
        st.success(f"✅ Parte generado: **{output}**")

# ==================== PLAN DE LLAMADA (CONTACTOS) ====================
# --- TAB: PLAN DE LLAMADA ---
with tab_plan:
    st.subheader("📞 Plan de Llamada - Base de Contactos")
    st.info("Registra domicilios y contactos de emergencia para cada personal del escuadrón.")
    
    search = st.text_input("🔍 Buscar personal:", placeholder="Nombre, DNI, CE o Orden", key="search_plan")
    
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
            
            headers = ["ORDEN", "NOMBRE", "AULA", "DOMICILIO", "TEL. PERSONAL", "TEL. EMERGENCIA", "CONTACTO EMERG.", "OBSERV."]
            for col, h in enumerate(headers, 1):
                cell = ws.cell(row=3, column=col, value=h)
                cell.font = Font(bold=True, color="FFFFFF", size=9)
                cell.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                cell.alignment = Alignment(horizontal="center")
                cell.border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
            
            row = 4
            for cont in sorted(todos, key=lambda x: x['orden']):
                pers = df[df['ORDEN_LIMP'] == cont['orden']]
                nombre = pers.iloc[0]['NOMBRE_COMPLETO'] if not pers.empty else ""
                aula = pers.iloc[0]['AULA'] if not pers.empty else ""
                
                ws.cell(row=row, column=1, value=cont['orden'])
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
            wb.save(output)
            st.success(f"✅ Plan de llamada exportado: **{output}**")
    else:
        st.warning("⚠️ Aún no hay contactos registrados. Usa el buscador para cargar datos.")
  

# --- TAB: RESUMEN ---
with tab_res:
    st.subheader("Resumen General y Novedades")

    # Recarga de datos
    st.session_state.novedades_lista = obtener_novedades()
    st.session_state.estado_asistencia = obtener_asistencia(FECHA_STR)
    st.session_state.lista_almuerzo = obtener_almuerzo(FECHA_STR)

    # ==========================================================
    # AUSENTES (SOLO MANUALES)
    # ==========================================================

    ausentes_manuales_resumen = {
        orden
        for orden, estado in st.session_state.estado_asistencia.items()
        if estado == "AUSENTE"
    }

    presentes_manuales_resumen = {
        orden
        for orden, estado in st.session_state.estado_asistencia.items()
        if estado == "PRESENTE"
    }

    total_ausentes_resumen = (
        ausentes_manuales_resumen
        - presentes_manuales_resumen
    )

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
                if st.session_state.estado_asistencia.get(
                    row['ORDEN_LIMP']
                ) == "AUSENTE"
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

# Generar orden correlativo 1..N según posición en el Excel
df = df.reset_index(drop=True)
df = df.sort_values("ORDEN_LIMP").reset_index(drop=True)
df["ORDEN_GENERAL"] = range(1, len(df) + 1)

lista_nov = st.session_state.get("novedades_lista", [])

for nov in lista_nov:
    if not isinstance(nov, dict):
        continue
        
    orden_limp = nov.get("orden")
    if orden_limp is None:
        continue

    # --- NUEVA LÍNEA: Extraemos el estado/motivo desde la novedad ---
    # Cambia "estado" o "motivo" por la clave exacta que uses en tu diccionario 'nov'
    estado = nov.get("estado") # o nov.get("motivo") u otra clave que guardes ahí

    alumno_df = df[df['ORDEN_LIMP'] == orden_limp]

    if not alumno_df.empty:
        alumno = alumno_df.iloc[0]
        data_ausentes.append({
            "Orden": int(alumno.get("ORDEN_GENERAL", 0)),
            "Nombre": alumno.get("NOMBRE_COMPLETO", "S/N"),
            "Motivo": estado if estado else "S/D",  # ¡Ahora sí existirá la variable!
            "Desde": alumno.get("FECHA_INI", "S/D"),
            "Hasta": alumno.get("FECHA_FIN", "S/D"),
        })
# ==============================================================================
# PROCESAMIENTO DE AUSENTES (Limpio y Seguro)
# ==============================================================================
data_ausentes = []  # ⚠️ IMPORTANTE: Reiniciar lista para no acumular datos viejos

novedades = st.session_state.get("novedades_lista", [])

for nov in novedades:
    # Extraemos datos directamente del diccionario 'nov'
    orden_limp = nov.get("orden")
    motivo_nov = nov.get("estado", "S/D")
    fecha_ini = nov.get("fecha_ini", "")
    fecha_fin = nov.get("fecha_fin", "")
    
    if orden_limp is not None:
        # Buscamos al alumno en el DataFrame principal usando su orden
        match = df[df['ORDEN_LIMP'] == orden_limp]
        
        if not match.empty:
            alumno = match.iloc[0]
            data_ausentes.append({
                "Orden": int(alumno.get("ORDEN_GENERAL", 0)),
                "Nombre": alumno.get("NOMBRE_COMPLETO", "S/N"),
                "Motivo": motivo_nov,
                "Desde": fecha_ini,
                "Hasta": fecha_fin
            })

# Creamos el DataFrame final con los ausentes encontrados
df_ausentes = pd.DataFrame(data_ausentes)
if not df_ausentes.empty:
    st.dataframe(
        df_ausentes.sort_values("Orden"),
        use_container_width=True,
        hide_index=True
    )
else:
    st.success("Sin personal ausente.")
# ==============================================================================
# 5. BOTÓN Y LÓGICA DE EXPORTACIÓN (REPORTE EJECUTIVO)
# ==============================================================================
if st.button("📥 PREPARAR REPORTE EJECUTIVO (EXCEL)", type="primary", use_container_width=True):
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from datetime import datetime
    from io import BytesIO
    
    # Crear buffer en memoria (no se guarda en el servidor)
    buffer = BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PARTE DIARIO EJECUTIVO"

    # 🏛️ ENCABEZADO
    ws.merge_cells('A1:F1')
    ws['A1'] = "PARTE DIARIO - ESCUADRÓN H"
    ws['A1'].font = Font(bold=True, size=18, color="FFFFFF")
    ws['A1'].fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    ws['A1'].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells('A2:F2')
    fecha_rep = getattr(st.session_state, "fecha_reporte", datetime.now())
    ws['A2'] = f"Fecha: {fecha_rep.strftime('%d/%m/%Y')} | Generado: {datetime.now().strftime('%H:%M')}"
    ws['A2'].font = Font(italic=True, size=11, color="555555")
    ws['A2'].alignment = Alignment(horizontal="center")
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 20

    # 📋 ENCABEZADOS DE COLUMNA (Fila 4)
    headers = ["ORDEN", "NOMBRE COMPLETO", "MOTIVO", "DESDE", "HASTA", "OBSERVACIONES"]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="556B2F", end_color="556B2F", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=Side(style="thin"), right=Side(style="thin"), 
                             top=Side(style="thin"), bottom=Side(style="thin"))

    # 📊 CARGAR DATOS DE df_ausentes (empezando fila 5)
    if not df_ausentes.empty:
        for idx, row in df_ausentes.sort_values("Orden").iterrows():
            row_num = 5 + idx
            ws.cell(row=row_num, column=1, value=row["Orden"])
            ws.cell(row=row_num, column=2, value=row["Nombre"])
            ws.cell(row=row_num, column=3, value=row["Motivo"])
            ws.cell(row=row_num, column=4, value=row["Desde"])
            ws.cell(row=row_num, column=5, value=row["Hasta"])
            ws.cell(row=row_num, column=6, value="")  # Observaciones vacías
            
            # Aplicar bordes y alineación a todas las celdas
            for col in range(1, 7):
                cell = ws.cell(row=row_num, column=col)
                cell.alignment = Alignment(horizontal="left", vertical="center")
                cell.border = Border(left=Side(style="thin"), right=Side(style="thin"), 
                                     top=Side(style="thin"), bottom=Side(style="thin"))
                
                # Ajustar ancho de columnas
                if col == 1:  # Orden
                    ws.column_dimensions['A'].width = 8
                elif col == 2:  # Nombre
                    ws.column_dimensions['B'].width = 35
                elif col == 3:  # Motivo
                    ws.column_dimensions['C'].width = 20
                elif col == 4:  # Desde
                    ws.column_dimensions['D'].width = 12
                elif col == 5:  # Hasta
                    ws.column_dimensions['E'].width = 12
                elif col == 6:  # Observaciones
                    ws.column_dimensions['F'].width = 25

    # 💾 GUARDAR EN BUFFER (no en disco)
    wb.save(buffer)
    buffer.seek(0)
    
    # 📥 BOTÓN DE DESCARGA (aparece después de preparar)
    fecha_archivo = datetime.now().strftime("%Y%m%d_%H%M")
    st.success("✅ Reporte generado exitosamente")
    
    st.download_button(
        label="⬇️ DESCARGAR EXCEL AHORA",
        data=buffer,
        file_name=f"Parte_Diario_Escuadron_H_{fecha_archivo}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )
# ==============================================================================
# 6. PARTE DE RACIONAMIENTO (FORMATO EXACTO + SOLO DESIGNADOS)
# ==============================================================================

# ⚠️ AJUSTE CRÍTICO: Poné acá el nombre de tu variable donde guardás los seleccionados.
# Ej: "personal_designado", "lista_racion", etc. Si no tenés filtro, usá 'df' directo.
df_fuente = st.session_state.get("df_raciones", df)

# Validar si hay datos
if df_fuente.empty:
    st.error("⛔ No hay datos cargados en la lista.")
else:
    from io import BytesIO
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from datetime import datetime

    buffer = BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "RACIONAMIENTO"

    # 🎨 ESTILOS PREDEFINIDOS
    font_title = Font(bold=True, size=14, color="1F4E78")
    font_header = Font(bold=True, color="FFFFFF")
    fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")

    # 🏛️ ENCABEZADO (FILA 1)
    ws.merge_cells('A1:F1')
    ws['A1'] = "PARTE DE RACIONAMIENTO - ESCUADRÓN H"
    ws['A1'].font = Font(bold=True, size=16, color="1F4E78")
    ws['A1'].alignment = align_center

    # 📅 FECHA Y TOTAL (FILA 2)
    ws['A2'] = f"Fecha: {datetime.now().strftime('%d/%m/%Y')}"
    ws['A2'].font = Font(bold=True, size=11)
    
    # Contar cuántos registros vamos a exportar
    total_registros = len(df_fuente)
    ws['B2'] = f"Total: {total_registros}"
    ws['B2'].font = Font(bold=True, size=11)
    ws['B2'].alignment = align_left # Alineado a la izquierda de su celda (B2)

    # 📋 ENCABEZADOS DE TABLA (FILA 4)
    headers = ["ORDEN", "NOMBRE COMPLETO", "GRADO", "CE", "DNI", "AULA"]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col, value=header)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border

    # 📊 CARGAR DATOS (SOLO LOS DESIGNADOS/FILTRADOS)
    for idx, row in df_fuente.iterrows():
        row_num = 5 + idx
        # ⚠️ Ajustá los nombres de las columnas si en tu app se llaman distinto
        ws.cell(row=row_num, column=1, value=row.get("ORDEN", row.get("orden", "")))
        ws.cell(row=row_num, column=2, value=row.get("NOMBRE_COMPLETO", row.get("nombre", "")))
        ws.cell(row=row_num, column=3, value=row.get("GRADO", row.get("grado", "")))
        ws.cell(row=row_num, column=4, value=row.get("CE", row.get("ce", "")))
        ws.cell(row=row_num, column=5, value=row.get("DNI", row.get("dni", "")))
        ws.cell(row=row_num, column=6, value=row.get("AULA", row.get("aula", "")))

        # Aplicar bordes a toda la fila
        for c in range(1, 7):
            ws.cell(row=row_num, column=c).border = thin_border

    # 📏 AJUSTE DE ANCHOS DE COLUMNA
    ws.column_dimensions['A'].width = 10  # Orden
    ws.column_dimensions['B'].width = 35  # Nombre
    ws.column_dimensions['C'].width = 18  # Grado
    ws.column_dimensions['D'].width = 10  # CE
    ws.column_dimensions['E'].width = 12  # DNI
    ws.column_dimensions['F'].width = 10  # Aula

    # 💾 DESCARGA DIRECTA
    wb.save(buffer)
    buffer.seek(0)
    fecha_archivo = datetime.now().strftime("%Y%m%d_%H%M")

    st.success(f"✅ Reporte listo. Se exportarán {total_registros} registros.")
    st.download_button(
        label="📥 DESCARGAR PARTE DE RACIONAMIENTO",
        data=buffer,
        file_name=f"Racionamiento_Escuadron_H_{fecha_archivo}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )

# ==============================================================================
    # 🔹 SECCIÓN 3: PERSONAL QUE ALMUERZA
    ws['A18'] = "🍽️ PERSONAL QUE ALMUERZA"
    ws['A18'].font = Font(bold=True, size=13, color="1F4E78")

    alm_headers = ["ORDEN", "NOMBRE COMPLETO", "AULA"]
    for col, h in enumerate(alm_headers, 1):
        cell = ws.cell(row=19, column=col, value=h)
        cell.font = Font(bold=True, color="FFFFFF", size=10)
        cell.fill = PatternFill(start_color="E65100", end_color="E65100", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")
        cell.border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

    if st.session_state.lista_almuerzo:
        for i, orden in enumerate(sorted(st.session_state.lista_almuerzo)):
            row = 20 + i
            asp = df[df['ORDEN_LIMP'] == orden]
            if not asp.empty:
                ws.cell(row=row, column=1, value=orden)
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
    output = f"PARTE_EJECUTIVO_{datetime.now().strftime('%d%m%Y_%H%M')}.xlsx"
    wb.save(output)
    st.success(f"✅ Reporte ejecutivo generado: **{output}**")
