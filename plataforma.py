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
    obtener_todos_contactos, guardar_contacto,
    guardar_horarios_dia, obtener_horarios_dia, obtener_todos_horarios_dia
)

st.set_page_config(page_title="Gestión de Parte Diario - Escuadrón H", layout="wide")

# ==============================================================================
# 🎖️ ESTILOS E INTERFAZ INSTITUCIONAL
# ==============================================================================
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

st.markdown('<div class="header-box"></div>', unsafe_allow_html=True)

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

if 'db_iniciada' not in st.session_state:
    init_db()
    st.session_state.db_iniciada = True

df = cargar_personal()
if df.empty:
    st.stop()

TOTAL_ESCUADRON = 139
AULAS_UNICAS = sorted(df['AULA'].unique())

if 'fecha_reporte' not in st.session_state:
    st.session_state.fecha_reporte = datetime.now().date()
FECHA_STR = st.session_state.fecha_reporte.isoformat()

# ==============================================================================
# 2. INICIALIZACIÓN DE ESTADO
# ==============================================================================
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

if 'editando_idx' not in st.session_state:
    st.session_state.editando_idx = None
if 'sel_nov' not in st.session_state:
    st.session_state.sel_nov = None

# ==============================================================================
# 3. CÁLCULO DE MÉTRICAS - LÓGICA MAESTRA CORREGIDA (V2)
# ==============================================================================
st.session_state.novedades_lista = obtener_novedades()
st.session_state.lista_almuerzo = obtener_almuerzo(FECHA_STR)

DIAS_SEMANA_ES = ["lunes", "martes", "miercoles", "jueves", "viernes"]
dia_actual_idx = st.session_state.fecha_reporte.weekday()
dia_actual = DIAS_SEMANA_ES[dia_actual_idx] if dia_actual_idx < 5 else "lunes"

horarios_hoy = obtener_todos_horarios_dia(dia_actual)
asistencia_manual = obtener_asistencia(FECHA_STR)

def es_novedad_activa(nov, fecha_rep):
    try:
        from datetime import datetime as dt
        fi = dt.strptime(nov['fecha_ini'], '%d%b%y').date()
        ff = dt(2099, 12, 31).date() if nov['fecha_fin'] == "N/O" else dt.strptime(nov['fecha_fin'], '%d%b%y').date()
        return fi <= fecha_rep <= ff
    except:
        return True

novedades_dict = {int(n['orden']): n for n in st.session_state.novedades_lista}

# Inicialización de TODOS los contadores
presentes_escuadron_count = 0
presentes_instituto_count = 0
ausentes_count = 0
total_entrantes_gd = 0
total_entrantes_gn = 0
total_comision = 0

for _, row in df.iterrows():
    orden = int(row['ORDEN_LIMP'])
    estado_manual = asistencia_manual.get(orden, None)
    
    # 1. PRIORIDAD ABSOLUTA: Lo que marcaste manualmente en los botones
    if estado_manual == "PRESENTE_ESCUADRON":
        presentes_escuadron_count += 1
    elif estado_manual == "PRESENTE_INSTITUTO":
        presentes_instituto_count += 1
        # Refinamos la métrica específica según su novedad
        if orden in novedades_dict and es_novedad_activa(novedades_dict[orden], st.session_state.fecha_reporte):
            est_nov = novedades_dict[orden]['estado']
            if est_nov == 'ENTRANTE GUARDIA NOCTURNA':
                total_entrantes_gn += 1
            elif est_nov in ['COMISIÓN', 'AUTORIZADO']:
                total_comision += 1
            else:
                total_entrantes_gd += 1
        else:
            total_entrantes_gd += 1 # Por defecto si no hay novedad específica
            
    elif estado_manual == "AUSENTE":
        ausentes_count += 1
    else:
        # 2. SECUNDARIO: Si no hay marca manual, evaluamos la novedad activa
        if orden in novedades_dict:
            nov = novedades_dict[orden]
            if es_novedad_activa(nov, st.session_state.fecha_reporte):
                estado_nov = nov['estado']
                if estado_nov in ['ART', 'DAF', 'LES', 'SSD']:
                    ausentes_count += 1
                elif estado_nov == 'ENTRANTE GUARDIA DIURNA':
                    presentes_instituto_count += 1
                    total_entrantes_gd += 1
                elif estado_nov == 'ENTRANTE GUARDIA NOCTURNA':
                    presentes_instituto_count += 1 
                    total_entrantes_gn += 1
                elif estado_nov in ['COMISIÓN', 'AUTORIZADO']:
                    presentes_escuadron_count += 1 
                    total_comision += 1
                else:
                    presentes_escuadron_count += 1
            else:
                presentes_escuadron_count += 1
        else:
            # Sin novedad y sin manual -> PRESENTE EN ESCUADRÓN por defecto
            presentes_escuadron_count += 1

# Totales finales para la interfaz
presentes_en_instituto = presentes_escuadron_count + presentes_instituto_count
presentes_en_escuadron = presentes_escuadron_count
total_ausentes = ausentes_count # Ahora es un número entero, no un conjunto

# Cálculo de ubicación física (solo para los que están activos en el escuadrón)
en_instituto = 0
fuera_por_aula = 0
for _, row in df.iterrows():
    orden = row['ORDEN_LIMP']
    aula = row['AULA']
    
    # No contar movimiento si está ausente o marcado como solo instituto
    estado_asist = asistencia_manual.get(orden, None)
    if estado_asist in ["AUSENTE", "PRESENTE_INSTITUTO"]:
        continue
        
    if orden in novedades_dict and es_novedad_activa(novedades_dict[orden], st.session_state.fecha_reporte):
        if novedades_dict[orden]['estado'] in ['ENTRANTE GUARDIA DIURNA', 'ENTRANTE GUARDIA NOCTURNA']:
            continue 

    if st.session_state.estado_aulas.get(aula, {}).get('estado_m', 'EN INSTITUTO') == 'EN INSTITUTO':
        en_instituto += 1
    else:
        fuera_por_aula += 1

total_fuera = fuera_por_aula + total_ausentes

# Guardar en session state para usar en otras pestañas
st.session_state.dia_actual = dia_actual
st.session_state.horarios_hoy = horarios_hoy
st.session_state.total_ausentes = total_ausentes
st.session_state.presentes_en_escuadron = presentes_en_escuadron
st.session_state.presentes_en_instituto = presentes_en_instituto
st.session_state.total_entrantes_gd = total_entrantes_gd
st.session_state.total_entrantes_gn = total_entrantes_gn
st.session_state.total_comision = total_comision

# ==============================================================================
# 4. INTERFAZ: BARRA SUPERIOR Y MÉTRICAS VISUALES (ÚNICA)
# ==============================================================================
st.markdown('<div class="sticky-bar">', unsafe_allow_html=True)
st.markdown('<div class="header-title">ESCUADRÓN H "CABO MARCELO GODOY"</div>', unsafe_allow_html=True)

c1, c2, c3, c4, c5, c6, c7, c8 = st.columns(8)
with c1: st.metric("TOTAL", TOTAL_ESCUADRON)
with c2: 
    st.metric("EN INSTITUTO", st.session_state.presentes_en_instituto, 
              delta=f"-{st.session_state.total_ausentes} aus." if st.session_state.total_ausentes > 0 else None)
with c3: 
    st.metric("EN ESCUADRÓN", st.session_state.presentes_en_escuadron, 
              help="Presentes en instituto menos entrantes guardia diurna")
with c4: 
    st.metric("AUSENTES", st.session_state.total_ausentes, 
              delta_color="inverse" if st.session_state.total_ausentes > 0 else "normal")
with c5: 
    st.metric("GUARDIA D.", st.session_state.total_entrantes_gd, 
              help="En instituto, no en escuadrón")
with c6: 
    st.metric("GUARDIA N.", st.session_state.total_entrantes_gn)
with c7: 
    st.metric("COMISIÓN", st.session_state.total_comision)
with c8: 
    st.metric("DISPONIBLES", st.session_state.presentes_en_escuadron - fuera_por_aula)
st.markdown('</div>', unsafe_allow_html=True)

st.divider()

col_sync1, col_sync2 = st.columns(2)
with col_sync1:
    if st.button("🔄 Sincronizar Datos", key="sync_btn_unico_v3", use_container_width=True):
        st.session_state.novedades_lista = obtener_novedades()
        st.session_state.lista_almuerzo = obtener_almuerzo(FECHA_STR)
        st.session_state.estado_aulas = obtener_estado_aulas(FECHA_STR)
        st.success("✅ Datos sincronizados correctamente")
        st.rerun()

with col_sync2:
    if st.button("🚨 RESETEAR ASISTENCIA DEL DÍA", key="reset_asistencia_unico_v3", use_container_width=True):
        import sqlite3
        conn = sqlite3.connect("parte_diario.db")
        conn.execute(f"DELETE FROM asistencia_diaria WHERE fecha='{FECHA_STR}'")
        conn.commit()
        conn.close()
        st.success("✅ Asistencia manual reiniciada. Se recalculará en base a las novedades.")
        st.rerun()

st.divider()

col_sync1, col_sync2 = st.columns(2)
with col_sync1:
    if st.button("🔄 Sincronizar Datos", key="sync_btn_unico_v2", use_container_width=True):
        st.session_state.novedades_lista = obtener_novedades()
        st.session_state.lista_almuerzo = obtener_almuerzo(FECHA_STR)
        st.session_state.estado_aulas = obtener_estado_aulas(FECHA_STR)
        st.success("✅ Datos sincronizados correctamente")
        st.rerun()

with col_sync2:
    if st.button("🚨 RESETEAR ASISTENCIA DEL DÍA", key="reset_asistencia_unico_v2", use_container_width=True):
        import sqlite3
        conn = sqlite3.connect("parte_diario.db")
        conn.execute(f"DELETE FROM asistencia_diaria WHERE fecha='{FECHA_STR}'")
        conn.commit()
        conn.close()
        st.success("✅ Asistencia manual reiniciada. Se recalculará en base a las novedades.")
        st.rerun()
# ==============================================================================
# 4. INTERFAZ: BARRA SUPERIOR Y MÉTRICAS
# ==============================================================================
st.markdown('<div class="sticky-bar">', unsafe_allow_html=True)
st.markdown('<div class="header-title">ESCUADRÓN H "CABO MARCELO GODOY"</div>', unsafe_allow_html=True)

c1, c2, c3, c4, c5, c6, c7, c8 = st.columns(8)
with c1: st.metric("TOTAL", TOTAL_ESCUADRON)
with c2: 
    # CORREGIDO: total_ausentes ya es un número, no usamos len()
    st.metric("EN INSTITUTO", presentes_en_instituto, delta=f"-{total_ausentes} aus." if total_ausentes > 0 else None)
with c3: 
    st.metric("EN ESCUADRÓN", presentes_en_escuadron, help="Presentes en instituto menos entrantes guardia diurna")
with c4: 
    st.metric("AUSENTES", total_ausentes, delta_color="inverse" if total_ausentes > 0 else "normal")
with c5: 
    st.metric("GUARDIA D.", total_entrantes_gd, help="En instituto, no en escuadrón")
with c6: 
    st.metric("GUARDIA N.", total_entrantes_gn)
with c7: 
    st.metric("COMISIÓN", total_comision)
with c8: 
    st.metric("DISPONIBLES", presentes_en_escuadron - fuera_por_aula)
st.markdown('</div>', unsafe_allow_html=True)

st.divider()

col_sync1, col_sync2 = st.columns(2)
with col_sync1:
    if st.button("🔄 Sincronizar Datos", key="sync_btn", use_container_width=True):
        st.session_state.novedades_lista = obtener_novedades()
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
        st.success("✅ Asistencia manual reiniciada.")
        st.rerun()

# ==============================================================================
# 5. PESTAÑAS
# ==============================================================================
tab_config, tab_nov, tab_seg, tab_alm, tab_plan, tab_res = st.tabs([
    "⚙️ Configuración", "📝 Novedades", "🔍 Seguimiento", "🍴 Almuerzo", "📞 Plan de Llamada", "📊 Resumen"
])

# --- TAB: CONFIGURACIÓN ---
with tab_config:
    st.subheader("⚙️ Configuración de Horarios Semanales")
    st.session_state.fecha_reporte = st.date_input("📅 Fecha del Reporte", st.session_state.fecha_reporte)
    
    dias_semana_es = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]
    dias_semana_db = ["lunes", "martes", "miercoles", "jueves", "viernes"]
    dia_idx = st.session_state.fecha_reporte.weekday()
    
    if dia_idx < 5:
        st.info(f"📌 Día activo: **{dias_semana_es[dia_idx]}**. Los horarios de este día se usarán en el parte.")
    else:
        st.warning("⚠️ Fin de semana detectado. Se tomará Lunes como referencia.")
        dia_idx = 0
    
    st.divider()
    
    if 'horarios_semanales' not in st.session_state:
        st.session_state.horarios_semanales = {}
        for aula in AULAS_UNICAS:
            st.session_state.horarios_semanales[aula] = {}
            for d_db in dias_semana_db:
                st.session_state.horarios_semanales[aula][d_db] = obtener_horarios_dia(aula, d_db)
    
    for aula in AULAS_UNICAS:
        with st.expander(f"**🏫 Aula {aula}**", expanded=False):
            tab_dias = st.tabs(dias_semana_es)
            for i, (d_es, d_db) in enumerate(zip(dias_semana_es, dias_semana_db)):
                with tab_dias[i]:
                    cfg = st.session_state.horarios_semanales[aula][d_db]
                    c1, c2, c3, c4, c5 = st.columns([2, 2, 2, 2, 3])
                    cfg['ent_m'] = c1.text_input("Entrada M.", cfg['ent_m'], key=f"em_{aula}_{d_db}")
                    cfg['sal_m'] = c2.text_input("Salida M.", cfg['sal_m'], key=f"sm_{aula}_{d_db}")
                    cfg['ent_t'] = c3.text_input("Entrada T.", cfg['ent_t'], key=f"et_{aula}_{d_db}")
                    cfg['sal_t'] = c4.text_input("Salida T.", cfg['sal_t'], key=f"st_{aula}_{d_db}")
                    cfg['tipo_ingreso'] = c5.selectbox("Tipo Ingreso", ["Normal", "Diferencial"], index=0 if cfg.get('tipo_ingreso', 'Normal') == 'Normal' else 1, key=f"ti_{aula}_{d_db}")
    
    if st.button("💾 Guardar Horarios de Toda la Semana", type="primary"):
        for aula in AULAS_UNICAS:
            for d_db in dias_semana_db:
                guardar_horarios_dia(aula, d_db, st.session_state.horarios_semanales[aula][d_db])
        st.success("✅ Horarios semanales guardados correctamente")
        st.rerun()

# --- TAB: NOVEDADES ---
with tab_nov:
    edit_idx = st.session_state.editando_idx
    es_edicion = edit_idx is not None

    if es_edicion and not (0 <= edit_idx < len(st.session_state.novedades_lista)):
        st.session_state.editando_idx = None
        st.rerun()

    st.subheader("✏️ Editando Novedad" if es_edicion else "➕ Registrar Novedad")
    st.info("⚠️ Lo que marques acá será la REFERENCIA para todos los cálculos del parte (métricas, seguimiento y excel).")
    
    data = None
    if es_edicion:
        nov = st.session_state.novedades_lista[edit_idx]
        st.info(f"Editando a: **{nov['nombre']}**")
        data = nov
        orden = int(data["orden"])
        nombre_asp = data["nombre"]
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
            orden = int(data["ORDEN_LIMP"])
            nombre_asp = data.get('NOMBRE_COMPLETO', data.get('nombre', 'Aspirante'))

    # =========================================================================
    # SECCIÓN DE PRESENTISMO (3 ESTADOS)
    # =========================================================================
    if data is not None:
        st.markdown("### 🎯 1. Definir Estado de Presentismo")
        st.caption("Seleccione dónde se encuentra físicamente el aspirante. Esto afecta directamente las métricas del parte.")
        
        col_btn_pe, col_btn_pi, col_btn_a = st.columns(3)
        with col_btn_pe:
            if st.button("🟢 PRESENTE EN ESCUADRÓN", type="primary", use_container_width=True, key="btn_pe_reg" if not es_edicion else "btn_pe_edit"):
                actualizar_asistencia(FECHA_STR, orden, "PRESENTE_ESCUADRON")
                st.success(f"{nombre_asp} marcado como PRESENTE EN ESCUADRÓN")
                st.rerun()
        with col_btn_pi:
            if st.button("🏛️ PRESENTE EN INSTITUTO", type="secondary", use_container_width=True, key="btn_pi_reg" if not es_edicion else "btn_pi_edit"):
                actualizar_asistencia(FECHA_STR, orden, "PRESENTE_INSTITUTO")
                st.info(f"{nombre_asp} marcado como PRESENTE EN INSTITUTO (No suma al escuadrón)")
                st.rerun()
        with col_btn_a:
            if st.button("🔴 AUSENTE", type="secondary", use_container_width=True, key="btn_aus_reg" if not es_edicion else "btn_aus_edit"):
                actualizar_asistencia(FECHA_STR, orden, "AUSENTE")
                st.error(f"{nombre_asp} marcado como AUSENTE")
                st.rerun()
        
        st.divider()

        # =========================================================================
        # SECCIÓN DE DETALLES DE LA NOVEDAD
        # =========================================================================
        st.markdown("### ⚙️ 2. Detalles de la Novedad")
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
                if st.button("💾 Guardar Cambios de Novedad", type="primary", use_container_width=True, key="btn_save_edit"):
                    actualizar_novedad(data['id'], {"estado": est, "detalle": det.upper(), "fecha_ini": fi, "fecha_fin": ff})
                    st.session_state.novedades_lista = obtener_novedades()
                    st.session_state.editando_idx = None
                    st.success("✅ Novedad actualizada")
                    st.rerun()
            else:
                if st.button("💾 Grabar Novedad", use_container_width=True, key="btn_save_new"):
                    agregar_novedad({
                        "orden": int(data["ORDEN_LIMP"]), "grado": data["GRADO"], "nombre": nombre_asp,
                        "dni": data["DNI"], "ce": data["CE"], "aula": data["AULA"], "estado": est,
                        "detalle": det.upper(), "fecha_ini": fi, "fecha_fin": ff
                    })
                    # Auto-marcar presentismo según la novedad si el usuario no lo hizo manualmente antes
                    asistencia_actual = obtener_asistencia(FECHA_STR).get(orden)
                    if asistencia_actual is None:
                        if est == 'ENTRANTE GUARDIA DIURNA':
                            actualizar_asistencia(FECHA_STR, orden, "PRESENTE_INSTITUTO")
                        else:
                            actualizar_asistencia(FECHA_STR, orden, "AUSENTE")
                    
                    st.session_state.novedades_lista = obtener_novedades()
                    st.session_state.sel_nov = None
                    st.success("✅ Novedad registrada y presentismo actualizado")
                    st.rerun()
        with b2:
            if st.button("🚫 Cancelar / Limpiar", use_container_width=True, key="btn_cancel"):
                st.session_state.editando_idx = None
                st.session_state.sel_nov = None
                st.rerun()

    # =========================================================================
    # TABLA DE NOVEDADES ORDENADA
    # =========================================================================
    st.divider()
    if st.session_state.novedades_lista:
        st.subheader("📋 Novedades Registradas")
        st.info("💡 Tabla ordenada por: **Estado ➔ Aula ➔ Número deORDER**")
        
        novedades_ordenadas = sorted(
            st.session_state.novedades_lista, 
            key=lambda x: (
                x.get('estado', 'ZZZ'),
                x.get('aula', 'ZZZ'),
                int(x.get('orden', 999))
            )
        )
        
        datos_tabla = []
        for idx, nov in enumerate(novedades_ordenadas):
            orden = int(nov.get('orden'))
            asis_actual = obtener_asistencia(FECHA_STR).get(orden, "SIN_MARCAR")
            
            if asis_actual == "PRESENTE_ESCUADRON":
                estado_asis = "🟢 EN ESCUADRÓN"
            elif asis_actual == "PRESENTE_INSTITUTO":
                estado_asis = "🏛️ EN INSTITUTO"
            elif asis_actual == "AUSENTE":
                estado_asis = "🔴 AUSENTE"
            else:
                # Fallback si no hay marca manual, inferimos por la novedad
                estado_asis = "🏛️ EN INSTITUTO" if nov.get('estado') == 'ENTRANTE GUARDIA DIURNA' else "🔴 AUSENTE"
            
            datos_tabla.append({
                "N°": idx + 1,
                "Orden": orden,
                "Aula": nov.get('aula', ''),
                "Grado": nov.get('grado', ''),
                "Apellido y Nombre": nov.get('nombre', ''),
                "DNI": nov.get('dni', ''),
                "Novedad": str(nov.get('estado', '')).replace("<span>", "").replace("</span>", ""),
                "Detalle": nov.get('detalle', ''),
                "Desde": nov.get('fecha_ini', ''),
                "Hasta": nov.get('fecha_fin', ''),
                "Presentismo": estado_asis
            })
        
        df_novedades = pd.DataFrame(datos_tabla)
        
        st.dataframe(
            df_novedades,
            use_container_width=True,
            hide_index=True,
            column_config={
                "N°": st.column_config.NumberColumn("N°", width="small"),
                "Orden": st.column_config.NumberColumn("Orden", width="small"),
                "Aula": st.column_config.TextColumn("Aula", width="small"),
                "Grado": st.column_config.TextColumn("Grado", width="small"),
                "Apellido y Nombre": st.column_config.TextColumn("Apellido y Nombre", width="medium"),
                "DNI": st.column_config.TextColumn("DNI", width="small"),
                "Novedad": st.column_config.TextColumn("Novedad", width="small"),
                "Detalle": st.column_config.TextColumn("Detalle", width="medium"),
                "Desde": st.column_config.TextColumn("Desde", width="small"),
                "Hasta": st.column_config.TextColumn("Hasta", width="small"),
                "Presentismo": st.column_config.TextColumn("Presentismo", width="small")
            }
        )
        
        st.divider()
        st.markdown("#### 🔧 Acciones sobre Novedades")
        col_sel1, col_sel2 = st.columns([3, 1])
        with col_sel1:
            opciones_novedades = [
                f"#{i+1} | Orden {row['Orden']} | {row['Apellido y Nombre']} | {row['Novedad']} | {row['Aula']}"
                for i, row in df_novedades.iterrows()
            ]
            seleccion = st.selectbox(
                "Seleccionar novedad para editar/eliminar:",
                opciones_novedades,
                key="sel_nov_accion"
            )
        
        if seleccion:
            idx_seleccionado = opciones_novedades.index(seleccion)
            nov_seleccionada = novedades_ordenadas[idx_seleccionado]
            
            col_btn_edit, col_btn_del = st.columns(2)
            with col_btn_edit:
                if st.button("✏️ Editar Novedad Seleccionada", type="primary", use_container_width=True):
                    # Encontrar el índice real en la lista original no ordenada para editar correctamente
                    idx_original = next((i for i, n in enumerate(st.session_state.novedades_lista) if n['id'] == nov_seleccionada['id']), None)
                    st.session_state.editando_idx = idx_original
                    st.rerun()
            
            with col_btn_del:
                if st.button("🗑️ Eliminar Novedad Seleccionada", type="secondary", use_container_width=True):
                    eliminar_novedad(nov_seleccionada['id'])
                    
                    # Limpieza de seguridad: borrar también el registro de asistencia manual para evitar "fantasmas"
                    import sqlite3
                    conn = sqlite3.connect("parte_diario.db")
                    conn.execute(f"DELETE FROM asistencia_diaria WHERE fecha='{FECHA_STR}' AND orden={nov_seleccionada['orden']}")
                    conn.commit()
                    conn.close()
                    
                    st.session_state.novedades_lista = obtener_novedades()
                    st.success(f"✅ Novedad de **{nov_seleccionada['nombre']}** eliminada")
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
        ausentes = sum(1 for n in st.session_state.novedades_lista if n['aula'] == aula and n['estado'] in ['ART', 'DAF', 'LES', 'SSD'])
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
    st.info("Directorio de contactos de emergencia de la Guardia. (Funcionalidad a desarrollar)")

# --- TAB: RESUMEN Y EXCEL ---
with tab_res:
    st.subheader("📊 Reporte General del Parte Diario")
    
    col_a, col_b, col_c, col_d, col_e = st.columns(5)
    with col_a: st.metric("TOTAL", TOTAL_ESCUADRON)
    with col_b: st.metric("EN INSTITUTO", st.session_state.presentes_en_instituto)
    with col_c: st.metric("EN ESCUADRÓN", st.session_state.presentes_en_escuadron)
    with col_d: st.metric("AUSENTES", st.session_state.total_ausentes)
    with col_e: st.metric("GUARDIA D.", st.session_state.total_entrantes_gd)
    
    st.divider()
    
    # =========================================================================
    # GENERACIÓN DEL EXCEL (CON TODAS LAS VARIABLES DE ESTILO DEFINIDAS)
    # =========================================================================
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PARTE DIARIO"
    ws.views.sheetView[0].showGridLines = False
    
    # 1. DEFINICIÓN DE ESTILOS (Para evitar NameError)
    font_titulo = Font(name="Calibri", size=14, bold=True)
    font_subtitulo = Font(name="Calibri", size=11, bold=True)
    font_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_datos = Font(name="Calibri", size=10)
    font_observacion = Font(name="Calibri", size=10, italic=True)
    
    fill_verde = PatternFill(start_color="1E4620", end_color="1E4620", fill_type="solid")
    fill_rojo_claro = PatternFill(start_color="FDE8E8", end_color="FDE8E8", fill_type="solid")
    fill_azul_claro = PatternFill(start_color="E1F5FE", end_color="E1F5FE", fill_type="solid")
    
    thin_border = Border(
        left=Side(style='thin', color='BFBFBF'), 
        right=Side(style='thin', color='BFBFBF'),
        top=Side(style='thin', color='BFBFBF'), 
        bottom=Side(style='thin', color='BFBFBF')
    )
    
    # 2. ENCABEZADO
    ws.merge_cells('A1:J1')
    ws['A1'] = f"PARTE DIARIO DEL ESCUADRÓN H - {st.session_state.fecha_reporte.strftime('%d%b%y').upper()}"
    ws['A1'].font = font_titulo
    ws['A1'].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30
    
    ws['A2'] = f"Día: {st.session_state.dia_actual.capitalize()}"
    ws['A2'].font = Font(name="Calibri", size=10, italic=True)
    
    # 3. CUADRO DE FUERZA
    ws['B4'] = "TOTAL"
    ws['C4'] = "EN INSTITUTO"
    ws['D4'] = "EN ESCUADRÓN"
    ws['E4'] = "AUSENTES"
    ws['F4'] = "GUARDIA D."
    ws['G4'] = "GUARDIA N."
    ws['H4'] = "COMISIÓN"
    
    for col in ['B4', 'C4', 'D4', 'E4', 'F4', 'G4', 'H4']:
        ws[col].font = font_header
        ws[col].fill = fill_verde
        ws[col].alignment = Alignment(horizontal="center", vertical="center")
    
    ws['B5'] = TOTAL_ESCUADRON
    ws['C5'] = st.session_state.presentes_en_instituto
    ws['D5'] = st.session_state.presentes_en_escuadron
    ws['E5'] = st.session_state.total_ausentes
    ws['F5'] = st.session_state.total_entrantes_gd
    ws['G5'] = st.session_state.total_entrantes_gn
    ws['H5'] = st.session_state.total_comision
    
    for col in ['B5', 'C5', 'D5', 'E5', 'F5', 'G5', 'H5']:
        ws[col].font = Font(name="Calibri", size=12, bold=True)
        ws[col].border = thin_border
        ws[col].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[4].height = 22
    ws.row_dimensions[5].height = 25
    
    # 4. TABLA DE NOVEDADES
    ws['A8'] = "NOVEDADES DEL PERSONAL (AUSENTES JUSTIFICADOS)"
    ws['A8'].font = font_subtitulo
    
    headers_nov = ["ORDEN", "GRADO", "APELLIDO Y NOMBRE", "DNI", "CE", "NOVEDAD", "DETALLE", "DESDE", "HASTA", "AULA"]
    for idx, h in enumerate(headers_nov, 1):
        cell = ws.cell(row=9, column=idx, value=h)
        cell.font = font_header
        cell.fill = fill_verde
        cell.alignment = Alignment(horizontal="center", vertical="center")
    
    row_idx = 10
    novedades_ordenadas = sorted(st.session_state.novedades_lista, key=lambda x: (x.get('aula', ''), int(x.get('orden', 0))))
    
    for nov in novedades_ordenadas:
        orden = int(nov.get("orden"))
        match = df[df['ORDEN_LIMP'] == orden]
        if match.empty: 
            continue
        
        estado_real = obtener_asistencia(FECHA_STR).get(orden, "AUSENTE")
        
        ws.cell(row=row_idx, column=1, value=orden).alignment = Alignment(horizontal="center")
        ws.cell(row=row_idx, column=2, value=match.iloc[0].get("GRADO", "")).alignment = Alignment(horizontal="center")
        ws.cell(row=row_idx, column=3, value=nov.get("nombre"))
        ws.cell(row=row_idx, column=4, value=match.iloc[0].get("DNI", "")).alignment = Alignment(horizontal="center")
        ws.cell(row=row_idx, column=5, value=match.iloc[0].get("CE", "")).alignment = Alignment(horizontal="center")
        ws.cell(row=row_idx, column=6, value=nov.get("estado", "S/D")).alignment = Alignment(horizontal="center")
        ws.cell(row=row_idx, column=7, value=nov.get("detalle", ""))
        ws.cell(row=row_idx, column=8, value=nov.get("fecha_ini", "")).alignment = Alignment(horizontal="center")
        ws.cell(row=row_idx, column=9, value=nov.get("fecha_fin", "")).alignment = Alignment(horizontal="center")
        ws.cell(row=row_idx, column=10, value=match.iloc[0].get("AULA", "")).alignment = Alignment(horizontal="center")
        
        fill_fila = fill_rojo_claro if estado_real == "AUSENTE" else fill_azul_claro
        for col_b in range(1, 11):
            ws.cell(row=row_idx, column=col_b).font = font_datos
            ws.cell(row=row_idx, column=col_b).border = thin_border
            ws.cell(row=row_idx, column=col_b).fill = fill_fila
        row_idx += 1
    
    # 5. HORARIOS DE INGRESO
    row_idx += 2
    ws.cell(row=row_idx, column=1, value="HORARIOS DE INGRESO - " + st.session_state.dia_actual.upper()).font = font_subtitulo
    row_idx += 1
    
    grupos_ingreso = {}
    for aula in AULAS_UNICAS:
        hor = st.session_state.horarios_hoy.get(aula, {"ent_m": "06:00", "tipo_ingreso": "Normal"})
        hora_ent = hor.get('ent_m', '06:00')
        tipo = hor.get('tipo_ingreso', 'Normal')
        llave = (hora_ent, tipo)
        
        if llave not in grupos_ingreso:
            grupos_ingreso[llave] = {"aulas": [], "presentes": 0}
        
        grupos_ingreso[llave]["aulas"].append(aula)
        
        alumnos_aula = df[df['AULA'] == aula]
        for _, r in alumnos_aula.iterrows():
            orden_r = int(r['ORDEN_LIMP'])
            # Contar solo si no está ausente y no es entrante de guardia diurna
            if obtener_asistencia(FECHA_STR).get(orden_r) != "AUSENTE" and nov.get('estado') != 'ENTRANTE GUARDIA DIURNA':
                 # Nota: la lógica de conteo de presentes por aula se simplifica aquí para el reporte
                 pass # El conteo preciso ya está en las métricas principales

    # Texto estático de horarios basado en la configuración del día
    for aula in AULAS_UNICAS:
        hor = st.session_state.horarios_hoy.get(aula, {"ent_m": "06:00", "tipo_ingreso": "Normal"})
        texto = f"• Aula {aula}: Ingreso {hor['ent_m']} hs ({hor['tipo_ingreso']})."
        ws.cell(row=row_idx, column=1, value=texto).font = font_observacion
        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=10)
        row_idx += 1
    
    # 6. OBSERVACIONES GENERALES
    row_idx += 1
    ws.cell(row=row_idx, column=1, value="OBSERVACIONES").font = font_subtitulo
    row_idx += 1
    
    if st.session_state.total_entrantes_gd > 0:
        ws.cell(row=row_idx, column=1, value=f"• Se encuentran {st.session_state.total_entrantes_gd} aspirante(s) como ENTRANTE GUARDIA DIURNA (presentes en instituto, fuera del escuadrón).").font = font_observacion
        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=10)
        row_idx += 1
    
    if st.session_state.total_entrantes_gn > 0:
        ws.cell(row=row_idx, column=1, value=f"• Se encuentran {st.session_state.total_entrantes_gn} aspirante(s) como ENTRANTE GUARDIA NOCTURNA.").font = font_observacion
        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=10)
        row_idx += 1
    
    if st.session_state.total_comision > 0:
        ws.cell(row=row_idx, column=1, value=f"• {st.session_state.total_comision} aspirante(s) en COMISIÓN o AUTORIZADO.").font = font_observacion
        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=10)
        row_idx += 1
    
    # 7. AJUSTE DE ANCHOS DE COLUMNA
    anchos = {"A": 8, "B": 10, "C": 35, "D": 12, "E": 10, "F": 22, "G": 25, "H": 12, "I": 12, "J": 10}
    for l, w in anchos.items():
        ws.column_dimensions[l].width = w
    
    # 8. GUARDAR Y BOTÓN DE DESCARGA
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    st.download_button(
        label="📥 DESCARGAR PARTE DIARIO (EXCEL)",
        data=output,
        file_name=f"PARTE_DIARIO_ESCUADRON_H_{st.session_state.fecha_reporte.strftime('%d%m%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )
