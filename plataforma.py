# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font
from datetime import datetime
import os
from io import BytesIO
import base64
import json
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

EXCEL_FONT_NAME = "Arial"
EXCEL_OLIVE_DARK = "4B5320"
EXCEL_OLIVE = "556B2F"
EXCEL_OLIVE_LIGHT = "E8EAD7"
EXCEL_OLIVE_ROW = "F4F6ED"
EXCEL_TEXT_DARK = "1F2937"
EXCEL_TEXT_MUTED = "555555"
EXCEL_BORDER = "A6A078"
EXCEL_WHITE = "FFFFFF"


def excel_font(**kwargs):
    kwargs.setdefault("name", EXCEL_FONT_NAME)
    return Font(**kwargs)

def descargar_archivo_auto(data, file_name, mime):
    href = f"data:{mime};base64,{base64.b64encode(data).decode('ascii')}"
    components.html(
        f"""
        <script>
        const a = document.createElement('a');
        a.href = {json.dumps(href)};
        a.download = {json.dumps(file_name)};
        document.body.appendChild(a);
        a.click();
        a.remove();
        </script>
        """,
        height=0,
    )

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
        grado = "ASP III ANO" if es_tercer_anio(nov.get('grado', '')) else "ASP I"
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
        grado = "ASP III ANO" if es_tercer_anio(nov.get('grado', '')) else "ASP I"
        detalle = f" {nov.get('detalle', '').strip()}" if nov.get('detalle') else ""
        lineas.append(f"{idx}. {grado} {nov['nombre']}{detalle}")
    return "\n".join(lineas)

def generar_minuta_informativa():
    fecha_minuta = st.session_state.fecha_reporte.strftime('%d%b%y').upper()

    # La minuta debe reflejar la base actual, no el estado que quedo en pantalla.
    novedades = obtener_novedades(FECHA_STR)
    estado_asistencia_actual = obtener_asistencia(FECHA_STR)
    st.session_state.novedades_lista = novedades
    st.session_state.estado_asistencia = estado_asistencia_actual

    ambito_minuta = {
        n['orden']: ambito_efectivo(n)
        for n in novedades
    }
    ausentes_novedad = {orden for orden, ambito in ambito_minuta.items() if ambito == "AUSENTE"}
    presentes_instituto_novedad = {orden for orden, ambito in ambito_minuta.items() if ambito in {"INSTITUTO", "ESCUADRON"}}
    presentes_escuadron_novedad = {orden for orden, ambito in ambito_minuta.items() if ambito == "ESCUADRON"}
    ausentes_manuales = {orden for orden, estado in estado_asistencia_actual.items() if estado == "AUSENTE"}
    presentes_instituto_manuales = {
        orden for orden, estado in estado_asistencia_actual.items()
        if estado in {"PRESENTE", "PRESENTE EN INSTITUTO", "PRESENTE EN ESCUADRÓN"}
    }
    presentes_escuadron_manuales = {
        orden for orden, estado in estado_asistencia_actual.items()
        if estado in {"PRESENTE", "PRESENTE EN ESCUADRÓN"}
    }

    total_ausentes_minuta = ausentes_novedad | (ausentes_manuales - presentes_instituto_manuales)

    df_presentes_primera_minuta = df[
        (~df['ORDEN_LIMP'].isin(total_ausentes_minuta)) &
        (
            df['ORDEN_LIMP'].isin(presentes_escuadron_novedad) |
            df['ORDEN_LIMP'].isin(presentes_escuadron_manuales) |
            ((~df['ORDEN_LIMP'].isin(ambito_minuta.keys())) &
             (~df['ORDEN_LIMP'].isin(presentes_instituto_manuales)) &
             (df['AULA'].map(lambda aula: st.session_state.estado_aulas.get(aula, {}).get('estado_m', 'EN INSTITUTO')) == 'EN INSTITUTO'))
        )
    ]

    df_tercero = df[df['GRADO'].map(es_tercer_anio)]
    df_aop = df[df['GRADO'].map(es_aop)]
    ausentes_tercero = set(df_tercero['ORDEN_LIMP']) & total_ausentes_minuta
    ausentes_aop = set(df_aop['ORDEN_LIMP']) & total_ausentes_minuta
    formados_tercero = len(df_presentes_primera_minuta[df_presentes_primera_minuta['GRADO'].map(es_tercer_anio)])
    formados_aop = len(df_presentes_primera_minuta[df_presentes_primera_minuta['GRADO'].map(es_aop)])
    disponibles_minuta = TOTAL_ESCUADRON - len(total_ausentes_minuta)
    primera_total_minuta = len(df_presentes_primera_minuta)

    lineas = [
        f'MINUTA INFORMATIVA DEL ESCUADRÓN H "CABO MARCELO GODOY" DEL DÍA {fecha_minuta}',
        "",
        f"FE: {TOTAL_ESCUADRON}",
        f"P: {disponibles_minuta}",
        f"A: {len(total_ausentes_minuta)}",
        f"FORMADOS A PRIMERA OBLIGACIÓN: {primera_total_minuta}",
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
import db_manager as _db_manager
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

actualizar_movimiento = getattr(_db_manager, "actualizar_movimiento", None)
eliminar_movimiento = getattr(_db_manager, "eliminar_movimiento", None)

if actualizar_movimiento is None:
    def actualizar_movimiento(id_mov, data):
        with _db_manager.get_db() as conn:
            _db_manager.run(conn, """UPDATE movimientos SET
                fecha_parte=?, modulo=?, accion=?, orden=?, nombre=?, aula=?, detalle=?
                WHERE id=?""",
                (
                    data.get('fecha_parte'), data.get('modulo'), data.get('accion'),
                    data.get('orden'), data.get('nombre'), data.get('aula'),
                    data.get('detalle', ''), id_mov
                ))
            conn.commit()

if eliminar_movimiento is None:
    def eliminar_movimiento(id_mov):
        with _db_manager.get_db() as conn:
            _db_manager.run(conn, "DELETE FROM movimientos WHERE id=?", (id_mov,))
            conn.commit()


if not hasattr(_db_manager, "contar_usuarios"):
    import hashlib as _hashlib
    import hmac as _hmac
    import os as _os
    from datetime import datetime as _dt_auth

    def _auth_hash_password(password, salt=None):
        if salt is None:
            salt = _os.urandom(16).hex()
        digest = _hashlib.pbkdf2_hmac("sha256", str(password).encode("utf-8"), bytes.fromhex(salt), 120000)
        return salt, digest.hex()

    def _auth_ensure_table():
        with _db_manager.get_db() as conn:
            _db_manager.run(conn, """CREATE TABLE IF NOT EXISTS usuarios (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                usuario TEXT NOT NULL UNIQUE,
                password_hash TEXT NOT NULL,
                salt TEXT NOT NULL,
                rol TEXT NOT NULL DEFAULT 'usuario',
                activo INTEGER NOT NULL DEFAULT 1,
                creado_en TEXT NOT NULL
            )""")
            conn.commit()

    def _contar_usuarios():
        _auth_ensure_table()
        with _db_manager.get_db() as conn:
            cur = _db_manager.run(conn, "SELECT COUNT(*) AS total FROM usuarios")
            row = cur.fetchone()
            return int(row["total"] if isinstance(row, dict) else row[0])

    def _crear_usuario(usuario, password, rol="usuario", activo=True):
        _auth_ensure_table()
        usuario = str(usuario).strip().lower()
        salt, password_hash = _auth_hash_password(password)
        with _db_manager.get_db() as conn:
            _db_manager.run(conn, """INSERT INTO usuarios (usuario, password_hash, salt, rol, activo, creado_en)
                VALUES (?, ?, ?, ?, ?, ?)""",
                (usuario, password_hash, salt, rol, 1 if activo else 0, _dt_auth.now().isoformat(timespec="seconds")))
            conn.commit()

    def _obtener_usuario(usuario):
        _auth_ensure_table()
        with _db_manager.get_db() as conn:
            cur = _db_manager.run(conn, "SELECT * FROM usuarios WHERE usuario=?", (str(usuario).strip().lower(),))
            row = cur.fetchone()
            return dict(row) if row else None

    def _autenticar_usuario(usuario, password):
        user = _obtener_usuario(usuario)
        if not user or int(user.get("activo", 0)) != 1:
            return None
        _, password_hash = _auth_hash_password(password, user["salt"])
        if not _hmac.compare_digest(password_hash, user["password_hash"]):
            return None
        return {"id": user["id"], "usuario": user["usuario"], "rol": user.get("rol", "usuario")}

    def _listar_usuarios():
        _auth_ensure_table()
        with _db_manager.get_db() as conn:
            cur = _db_manager.run(conn, "SELECT id, usuario, rol, activo, creado_en FROM usuarios ORDER BY usuario")
            return _db_manager.fetch_all(cur)

    def _actualizar_password_usuario(id_usuario, password):
        salt, password_hash = _auth_hash_password(password)
        with _db_manager.get_db() as conn:
            _db_manager.run(conn, "UPDATE usuarios SET password_hash=?, salt=? WHERE id=?", (password_hash, salt, id_usuario))
            conn.commit()

    def _actualizar_estado_usuario(id_usuario, activo):
        with _db_manager.get_db() as conn:
            _db_manager.run(conn, "UPDATE usuarios SET activo=? WHERE id=?", (1 if activo else 0, id_usuario))
            conn.commit()

    def _eliminar_usuario(id_usuario):
        with _db_manager.get_db() as conn:
            _db_manager.run(conn, "DELETE FROM usuarios WHERE id=?", (id_usuario,))
            conn.commit()

    _db_manager.contar_usuarios = _contar_usuarios
    _db_manager.crear_usuario = _crear_usuario
    _db_manager.obtener_usuario = _obtener_usuario
    _db_manager.autenticar_usuario = _autenticar_usuario
    _db_manager.listar_usuarios = _listar_usuarios
    _db_manager.actualizar_password_usuario = _actualizar_password_usuario
    _db_manager.actualizar_estado_usuario = _actualizar_estado_usuario
    _db_manager.eliminar_usuario = _eliminar_usuario



def usuario_actual():
    return st.session_state.get("usuario_actual")


def es_admin():
    user = usuario_actual() or {}
    return user.get("rol") == "admin"


def cerrar_sesion():
    for key in ("usuario_actual", "autenticado"):
        if key in st.session_state:
            del st.session_state[key]


def requerir_login():
    total_usuarios = _db_manager.contar_usuarios()

    if total_usuarios == 0:
        st.markdown("## Configuracion inicial de seguridad")
        st.info("Crea tu usuario administrador. Despues de esto, nadie podra ingresar sin usuario y contrasena.")
        with st.form("crear_admin_inicial"):
            usuario = st.text_input("Usuario administrador", value="admin")
            password = st.text_input("Contrasena", type="password")
            password2 = st.text_input("Repetir contrasena", type="password")
            crear = st.form_submit_button("Crear administrador", type="primary", use_container_width=True)
        if crear:
            if not usuario.strip() or not password:
                st.error("Completa usuario y contrasena.")
            elif password != password2:
                st.error("Las contrasenas no coinciden.")
            elif len(password) < 6:
                st.error("La contrasena debe tener al menos 6 caracteres.")
            else:
                _db_manager.crear_usuario(usuario, password, rol="admin", activo=True)
                st.success("Administrador creado. Inicia sesion para continuar.")
                st.rerun()
        st.stop()

    if not st.session_state.get("autenticado"):
        st.markdown("## Acceso restringido")
        st.caption('Escuadron H "Cabo Marcelo Godoy"')
        with st.form("login_form"):
            usuario = st.text_input("Usuario")
            password = st.text_input("Contrasena", type="password")
            entrar = st.form_submit_button("Ingresar", type="primary", use_container_width=True)
        if entrar:
            user = _db_manager.autenticar_usuario(usuario, password)
            if user:
                st.session_state.autenticado = True
                st.session_state.usuario_actual = user
                st.rerun()
            else:
                st.error("Usuario o contrasena incorrectos.")
        st.stop()

    with st.sidebar:
        user = usuario_actual() or {}
        st.caption(f"Usuario: {user.get('usuario', '-')}")
        st.caption(f"Rol: {user.get('rol', '-')}")
        if st.button("Cerrar sesion", use_container_width=True):
            cerrar_sesion()
            st.rerun()


def panel_admin_usuarios():
    st.subheader("Administracion de usuarios")
    st.caption("Crea usuarios para quienes puedan ingresar a la aplicacion.")

    with st.form("crear_usuario_form"):
        c1, c2, c3 = st.columns([2, 2, 1])
        with c1:
            nuevo_usuario = st.text_input("Nuevo usuario")
        with c2:
            nueva_password = st.text_input("Contrasena inicial", type="password")
        with c3:
            nuevo_rol = st.selectbox("Rol", ["usuario", "admin"])
        crear_usuario_btn = st.form_submit_button("Crear usuario", type="primary", use_container_width=True)
    if crear_usuario_btn:
        if not nuevo_usuario.strip() or not nueva_password:
            st.error("Completa usuario y contrasena.")
        elif len(nueva_password) < 6:
            st.error("La contrasena debe tener al menos 6 caracteres.")
        else:
            try:
                _db_manager.crear_usuario(nuevo_usuario, nueva_password, rol=nuevo_rol, activo=True)
                st.success("Usuario creado.")
                st.rerun()
            except Exception as e:
                st.error(f"No se pudo crear el usuario: {e}")

    usuarios = _db_manager.listar_usuarios()
    if not usuarios:
        st.info("No hay usuarios cargados.")
        return

    st.divider()
    for user in usuarios:
        with st.container(border=True):
            c_info, c_estado, c_pass, c_del = st.columns([3, 1.3, 2, 1])
            with c_info:
                estado = "Activo" if int(user.get("activo", 0)) == 1 else "Inactivo"
                st.markdown(f"**{user['usuario']}** | {user.get('rol', 'usuario')} | {estado}")
                st.caption(f"Creado: {user.get('creado_en', '-')}")
            with c_estado:
                activo = int(user.get("activo", 0)) == 1
                nuevo_estado = st.toggle("Activo", value=activo, key=f"usr_activo_{user['id']}")
                if nuevo_estado != activo:
                    _db_manager.actualizar_estado_usuario(user["id"], nuevo_estado)
                    st.rerun()
            with c_pass:
                nueva = st.text_input("Nueva contrasena", type="password", key=f"usr_pass_{user['id']}")
                if st.button("Cambiar", key=f"usr_cambiar_{user['id']}", use_container_width=True):
                    if len(nueva) < 6:
                        st.warning("Minimo 6 caracteres.")
                    else:
                        _db_manager.actualizar_password_usuario(user["id"], nueva)
                        st.success("Contrasena actualizada.")
            with c_del:
                if user.get("usuario") == (usuario_actual() or {}).get("usuario"):
                    st.caption("Tu usuario")
                elif st.button("Eliminar", key=f"usr_del_{user['id']}", use_container_width=True):
                    _db_manager.eliminar_usuario(user["id"])
                    st.rerun()

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

st.markdown("""
<style>
@media (max-width: 720px) {
    .main .block-container {
        padding-top: 0.6rem !important;
        padding-left: 0.55rem !important;
        padding-right: 0.55rem !important;
        padding-bottom: 2rem !important;
    }
    div[data-testid="stTabs"] > div[role="tablist"] {
        gap: 0.35rem !important;
        overflow-x: auto !important;
        flex-wrap: nowrap !important;
        padding-bottom: 0.35rem !important;
    }
    div[data-testid="stTabs"] button[role="tab"],
    div[data-testid="stTabs"] button[data-baseweb="tab"] {
        min-width: max-content !important;
        padding: 0.62rem 0.85rem !important;
        border-radius: 999px !important;
        border: 1px solid rgba(143, 166, 80, 0.35) !important;
        background: rgba(75, 83, 32, 0.22) !important;
        font-size: 0.9rem !important;
        font-weight: 800 !important;
    }
    div[data-testid="stTabs"] button[aria-selected="true"],
    div[data-testid="stTabs"] button[data-baseweb="tab"][aria-selected="true"] {
        background: #4B5320 !important;
        color: #FFFFFF !important;
        border-color: #8FA650 !important;
    }
    hr {
        margin: 0.85rem 0 !important;
    }
    .stButton {
        margin-top: 0.55rem !important;
        margin-bottom: 0.55rem !important;
    }
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

requerir_login()

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


def estado_asistencia_por_ambito(ambito):
    if ambito == "AUSENTE":
        return "AUSENTE"
    if ambito == "INSTITUTO":
        return "PRESENTE EN INSTITUTO"
    return "PRESENTE EN ESCUADR?N"

def parsear_detalle_movimiento(detalle):
    datos = {"Motivo": "", "Presencia": "", "Desde": "", "Hasta": "", "Observación": detalle or ""}
    partes = [p.strip() for p in str(detalle or "").split("|")]
    if not partes:
        return datos

    datos["Motivo"] = partes[0] if partes else ""
    if len(partes) >= 2:
        if " a " in partes[1]:
            datos["Desde"], datos["Hasta"] = [p.strip() for p in partes[1].split(" a ", 1)]
            if len(partes) >= 3:
                datos["Observación"] = partes[2]
        else:
            datos["Presencia"] = partes[1]
            if len(partes) >= 3 and " a " in partes[2]:
                datos["Desde"], datos["Hasta"] = [p.strip() for p in partes[2].split(" a ", 1)]
                if len(partes) >= 4:
                    datos["Observación"] = partes[3]
            elif len(partes) >= 3:
                datos["Observación"] = partes[2]
    return datos

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

total_ausentes = ausentes_fijos | (ausentes_manuales - presentes_instituto_manuales)

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

def mostrar_monitor_novedades():
    st.markdown(f"""
    <div class="nov-monitor">
        <div class="nov-monitor-head">
            <div class="nov-monitor-title">Monitor de novedades</div>
            <div class="nov-monitor-sub">{len(st.session_state.novedades_lista)} activa(s)</div>
        </div>
        <div class="nov-monitor-grid">{monitor_items_html}</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown(f"""
<style>
    .header-title {{
        display: none !important;
    }}
    .rrhh-panel {{
        background: linear-gradient(135deg, #2f3717 0%, #4B5320 48%, #111827 100%);
        border: 1px solid rgba(166, 160, 120, 0.45);
        border-radius: 8px;
        padding: 1rem;
        margin: 0.25rem 0 1rem 0;
        box-shadow: 0 10px 28px rgba(0, 0, 0, 0.28);
    }}
    .rrhh-head {{
        display: flex;
        justify-content: space-between;
        gap: 1rem;
        align-items: flex-start;
        margin-bottom: 0.85rem;
    }}
    .rrhh-brand {{
        display: flex;
        gap: 0.75rem;
        align-items: center;
    }}
    .rrhh-emblem {{
        width: 48px;
        height: 48px;
        border-radius: 8px;
        display: grid;
        place-items: center;
        background: rgba(232, 234, 215, 0.14);
        border: 1px solid rgba(232, 234, 215, 0.28);
        color: #F9FAFB;
        font-weight: 900;
        letter-spacing: 0.04em;
    }}
    .rrhh-eyebrow {{
        margin: 0 0 0.2rem 0;
        color: #D8DEC2;
        font-size: 0.72rem;
        letter-spacing: 0.12em;
        text-transform: uppercase;
        font-weight: 800;
    }}
    .rrhh-title {{
        margin: 0;
        color: #FFFFFF;
        font-size: 1.35rem;
        line-height: 1.15;
        font-weight: 900;
    }}
    .rrhh-title-short {{
        display: none;
    }}
    .rrhh-subtitle {{
        margin: 0.28rem 0 0 0;
        color: #E8EAD7;
        font-size: 0.9rem;
    }}
    .rrhh-status-box {{
        display: flex;
        flex-direction: column;
        gap: 0.35rem;
        align-items: flex-end;
    }}
    .rrhh-date, .rrhh-live {{
        border: 1px solid rgba(232, 234, 215, 0.34);
        border-radius: 999px;
        color: #F9FAFB;
        background: rgba(17, 24, 39, 0.42);
        padding: 0.34rem 0.7rem;
        font-size: 0.8rem;
        white-space: nowrap;
        font-weight: 700;
    }}
    .rrhh-live {{
        color: #E8EAD7;
    }}
    .rrhh-access {{
        display: grid;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        gap: 0.55rem;
        margin-bottom: 0.8rem;
    }}
    .rrhh-access-card {{
        border: 1px solid rgba(232, 234, 215, 0.22);
        background: rgba(17, 24, 39, 0.34);
        border-radius: 8px;
        padding: 0.65rem 0.75rem;
        min-height: 68px;
    }}
    .rrhh-access-card span {{
        display: block;
        color: #D8DEC2;
        font-size: 0.72rem;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        font-weight: 800;
    }}
    .rrhh-access-card strong {{
        display: block;
        color: #FFFFFF;
        margin-top: 0.25rem;
        font-size: 0.95rem;
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
        border: 1px solid rgba(232, 234, 215, 0.16);
        background: rgba(255, 255, 255, 0.045);
        border-radius: 8px;
        padding: 0.75rem;
        min-height: 96px;
    }}
    .rrhh-kpi span {{
        display: block;
        color: #D8DEC2;
        font-size: 0.72rem;
        text-transform: uppercase;
        letter-spacing: 0.06em;
        font-weight: 800;
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
        color: #CBD5E1;
        font-size: 0.78rem;
        margin-top: 0.3rem;
    }}
    .rrhh-kpi-ok {{ border-left: 3px solid #8FA650; }}
    .rrhh-kpi-warn {{ border-left: 3px solid #D8B55A; }}
    .rrhh-kpi-alert {{ border-left: 3px solid #C05646; }}
    @media (max-width: 720px) {{
        .rrhh-panel {{
            padding: 0.55rem 0.6rem;
            margin: 0 0 0.55rem 0;
            box-shadow: none;
            background: linear-gradient(135deg, #344018 0%, #111827 100%);
        }}
        .rrhh-head {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 0.5rem;
            margin-bottom: 0.45rem;
        }}
        .rrhh-brand {{ gap: 0.45rem; min-width: 0; }}
        .rrhh-emblem {{ width: 32px; height: 32px; border-radius: 7px; font-size: 0.78rem; flex: 0 0 auto; }}
        .rrhh-eyebrow {{ display: none; }}
        .rrhh-title {{ font-size: 1rem; line-height: 1.1; white-space: nowrap; overflow: visible; max-width: none; }}
        .rrhh-title-full {{ display: none; }}
        .rrhh-title-short {{ display: inline; }}
        .rrhh-subtitle {{ display: none; }}
        .rrhh-status-box {{ align-items: flex-end; margin-top: 0; gap: 0.25rem; }}
        .rrhh-live {{ display: none; }}
        .rrhh-date {{ padding: 0.24rem 0.48rem; font-size: 0.7rem; }}
        .rrhh-access {{ display: none; }}
        .rrhh-kpi-grid {{ grid-template-columns: repeat(4, minmax(68px, 1fr)); gap: 0.35rem; }}
        .rrhh-kpi {{ min-height: 58px; padding: 0.45rem; border-radius: 7px; }}
        .rrhh-kpi span {{ font-size: 0.58rem; letter-spacing: 0.03em; }}
        .rrhh-kpi strong {{ font-size: 1.08rem; margin-top: 0.15rem; }}
        .rrhh-kpi small {{ display: none; }}
        .nov-monitor {{ padding: 0.55rem; position: sticky; top: 0.2rem; z-index: 10; }}
        .nov-monitor-head {{ display: block; margin-bottom: 0.4rem; }}
        .nov-monitor-sub {{ display: block; margin-top: 0.15rem; }}
        .nov-monitor-grid {{ display: flex; overflow-x: auto; gap: 0.4rem; padding-bottom: 0.1rem; }}
        .nov-card {{ min-width: 210px; flex: 0 0 auto; min-height: 62px; padding: 0.52rem; }}
        .nov-card strong {{ font-size: 1.2rem; }}
    }}
</style>
<section class="rrhh-panel">
    <div class="rrhh-head">
        <div class="rrhh-brand">
            <div class="rrhh-emblem">EH</div>
            <div>
                <p class="rrhh-eyebrow">Sistema compartido de control de personal</p>
                <h1 class="rrhh-title"><span class="rrhh-title-full">Escuadron H "Cabo Marcelo Godoy"</span><span class="rrhh-title-short">Escuadron H</span></h1>
                <p class="rrhh-subtitle">Parte diario, novedades, presentismo, ubicacion y reportes en tiempo real.</p>
            </div>
        </div>
        <div class="rrhh-status-box">
            <div class="rrhh-live">Sistema operativo</div>
            <div class="rrhh-date">Parte: {st.session_state.fecha_reporte.strftime('%d/%m/%Y')}</div>
        </div>
    </div>
    <div class="rrhh-access">
        <div class="rrhh-access-card"><span>Trabajo multiusuario</span><strong>Datos centralizados para guardia y control</strong></div>
        <div class="rrhh-access-card"><span>Seguimiento activo</span><strong>Novedades vigentes hasta su vencimiento</strong></div>
        <div class="rrhh-access-card"><span>Reportes</span><strong>Excel institucional con formato uniforme</strong></div>
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
tab_labels = ["Dia y horarios", "Novedades", "Ubicacion", "Racionamiento", "Legajos y contactos", "Reportes"]
if es_admin():
    tab_labels.append("Usuarios")
tabs_creadas = st.tabs(tab_labels)
tab_config, tab_nov, tab_seg, tab_alm, tab_plan, tab_res = tabs_creadas[:6]
tab_usuarios = tabs_creadas[6] if es_admin() else None

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

if tab_usuarios is not None:
    with tab_usuarios:
        panel_admin_usuarios()

# --- TAB: NOVEDADES ---
with tab_nov:
    st.divider()
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
                    actualizar_asistencia(FECHA_STR, data.get("orden"), estado_asistencia_por_ambito(ambito))
                    st.session_state.estado_asistencia[data.get("orden")] = estado_asistencia_por_ambito(ambito)
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
                    actualizar_asistencia(FECHA_STR, int(data["ORDEN_LIMP"]), estado_asistencia_por_ambito(ambito))
                    st.session_state.estado_asistencia[int(data["ORDEN_LIMP"])] = estado_asistencia_por_ambito(ambito)
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

    st.divider()
    st.subheader("Novedades activas")
    if st.session_state.novedades_lista:
        st.caption("Panel editable para modificar motivo, presencia real, fechas o cerrar una novedad antes de su vencimiento.")
        for idx, nov in enumerate(st.session_state.novedades_lista):
            ambito_lbl = AMBITOS_NOVEDAD.get(ambito_efectivo(nov), "Sin definir")
            with st.container(border=True):
                c_info, c_edit, c_del = st.columns([6, 1, 1])
                with c_info:
                    st.markdown(f"**{nov['nombre']}** | **{nov['estado']}**")
                    st.caption(f"{nov['fecha_ini']} a {nov['fecha_fin']} | {ambito_lbl} | Aula {nov.get('aula', '-')} | DNI {nov.get('dni', '-')} | CE {nov.get('ce', '-')}")
                    if nov.get('detalle'):
                        st.caption(f"Detalle: {nov['detalle']}")
                with c_edit:
                    if st.button("Editar", key=f"edit_nov_activa_{idx}", use_container_width=True):
                        limpiar_form_novedad()
                        st.session_state.editando_idx = idx
                        st.session_state.sel_nov = None
                        st.rerun()
                with c_del:
                    if st.button("Quitar", key=f"del_nov_activa_{idx}", use_container_width=True):
                        log_movimiento("Novedades", "ELIMINAR NOVEDAD", nov.get("orden"), nov.get("nombre"), nov.get("aula"), f"{nov.get('estado')} | {nov.get('fecha_ini')} a {nov.get('fecha_fin')} | {nov.get('detalle')}")
                        eliminar_novedad(nov['id'])
                        st.session_state.estado_asistencia[nov['orden']] = "PRESENTE"
                        actualizar_asistencia(FECHA_STR, nov['orden'], "PRESENTE")
                        st.session_state.novedades_lista = obtener_novedades(FECHA_STR)
                        st.toast("Novedad eliminada y asistencia actualizada")
                        st.rerun()
    else:
        st.info("No hay novedades activas para la fecha seleccionada.")

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
    if st.session_state.lista_almuerzo:
        from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "RACIONAMIENTO"
        
        ws.merge_cells('A1:F1')
        ws['A1'] = "PARTE DE RACIONAMIENTO - ESCUADRÓN H"
        ws['A1'].font = excel_font(bold=True, size=15, color=EXCEL_WHITE)
        ws['A1'].fill = PatternFill(start_color=EXCEL_OLIVE_DARK, end_color=EXCEL_OLIVE_DARK, fill_type="solid")
        ws['A1'].alignment = Alignment(horizontal="center", vertical="center")
        ws['A2'] = f"Fecha: {st.session_state.fecha_reporte.strftime('%d/%m/%Y')} | Total: {len(st.session_state.lista_almuerzo)}"
        ws['A2'].font = excel_font(italic=True, size=10, color=EXCEL_TEXT_MUTED)
        ws.row_dimensions[1].height = 25
        
        headers = ["Nro", "NOMBRE COMPLETO", "GRADO", "CE", "DNI", "AULA"]
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=4, column=col, value=h)
            cell.font = excel_font(bold=True, color=EXCEL_WHITE, size=9)
            cell.fill = PatternFill(start_color=EXCEL_OLIVE, end_color=EXCEL_OLIVE, fill_type="solid")
            cell.alignment = Alignment(horizontal="center")
            cell.border = Border(left=Side(style="thin", color=EXCEL_BORDER), right=Side(style="thin", color=EXCEL_BORDER), top=Side(style="thin", color=EXCEL_BORDER), bottom=Side(style="thin", color=EXCEL_BORDER))
            
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
                ws.cell(row=row, column=c).border = Border(left=Side(style="thin", color=EXCEL_BORDER), right=Side(style="thin", color=EXCEL_BORDER), top=Side(style="thin", color=EXCEL_BORDER), bottom=Side(style="thin", color=EXCEL_BORDER))
                ws.cell(row=row, column=c).alignment = Alignment(horizontal="center" if c in [1,3,6] else "left")
            row += 1
            
        for col, w in zip("ABCDEF", [10, 35, 12, 12, 15, 12]): 
            ws.column_dimensions[col].width = w
        output = f"RACIONAMIENTO_{datetime.now().strftime('%d%m%Y_%H%M')}.xlsx"
        st.download_button(
            "GENERAR PARTE DE RACIONAMIENTO",
            data=excel_bytes(wb),
            file_name=output,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
            key="download_parte_racionamiento"
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
            ws['A1'].font = excel_font(bold=True, size=16, color=EXCEL_WHITE)
            ws['A1'].fill = PatternFill(start_color=EXCEL_OLIVE_DARK, end_color=EXCEL_OLIVE_DARK, fill_type="solid")
            ws['A1'].alignment = Alignment(horizontal="center")
            
            headers = ["Nro", "NOMBRE", "AULA", "DOMICILIO", "TEL. PERSONAL", "TEL. EMERGENCIA", "CONTACTO EMERG.", "OBSERV."]
            for col, h in enumerate(headers, 1):
                cell = ws.cell(row=3, column=col, value=h)
                cell.font = excel_font(bold=True, color=EXCEL_WHITE, size=9)
                cell.fill = PatternFill(start_color=EXCEL_OLIVE_DARK, end_color=EXCEL_OLIVE_DARK, fill_type="solid")
                cell.alignment = Alignment(horizontal="center")
                cell.border = Border(left=Side(style="thin", color=EXCEL_BORDER), right=Side(style="thin", color=EXCEL_BORDER), top=Side(style="thin", color=EXCEL_BORDER), bottom=Side(style="thin", color=EXCEL_BORDER))
            
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
                    ws.cell(row=row, column=c).border = Border(left=Side(style="thin", color=EXCEL_BORDER), right=Side(style="thin", color=EXCEL_BORDER), top=Side(style="thin", color=EXCEL_BORDER), bottom=Side(style="thin", color=EXCEL_BORDER))
                    ws.cell(row=row, column=c).alignment = Alignment(horizontal="center" if c in [1,3] else "left")
                row += 1
            
            for col, w in zip("ABCDEFGH", [8, 30, 10, 35, 15, 15, 25, 30]):
                ws.column_dimensions[col].width = w
            
            output = f"PLAN_LLAMADA_{datetime.now().strftime('%d%m%Y')}.xlsx"
            descargar_archivo_auto(excel_bytes(wb), output, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.success(f"✅ Plan de llamada descargado: **{output}**")
    else:
        st.warning("⚠️ Aún no hay contactos registrados. Usa el buscador para cargar datos.")
  

# --- TAB: RESUMEN ---
with tab_res:
    st.subheader("Resumen General y Novedades")
    mostrar_monitor_novedades()
    st.divider()

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
    total_ausentes_resumen = ausentes_resumen | (ausentes_manuales_resumen - presentes_manuales_resumen)

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
    st.text_area("Minuta generada automáticamente desde Novedades", value=minuta_texto, height=520)
    nombre_minuta = f"MINUTA_ESCUADRON_H_{st.session_state.fecha_reporte.strftime('%d%m%Y')}.txt"
    minuta_bytes = minuta_texto.encode("utf-8-sig")
    st.download_button(
        "Descargar minuta (.txt)",
        data=minuta_bytes,
        file_name=nombre_minuta,
        mime="text/plain; charset=utf-8",
        key=f"descargar_minuta_{st.session_state.fecha_reporte.strftime('%Y%m%d')}_{len(minuta_bytes)}",
        on_click="ignore",
        use_container_width=True
    )


    st.divider()
    st.subheader("Historial de movimientos")
    col_fi, col_ff = st.columns(2)
    with col_fi:
        fecha_desde_hist = st.date_input("Desde", st.session_state.fecha_reporte, key="fecha_historial_desde")
    with col_ff:
        fecha_hasta_hist = st.date_input("Hasta", st.session_state.fecha_reporte, key="fecha_historial_hasta")

    movimientos = obtener_movimientos()
    if movimientos:
        df_mov = pd.DataFrame(movimientos)
        df_mov = df_mov.rename(columns={
            "fecha_hora": "Fecha/hora",
            "fecha_parte": "Fecha parte",
            "modulo": "Módulo",
            "accion": "Acción",
            "orden": "Orden interno",
            "nombre": "Nombre completo",
            "aula": "Aula",
            "detalle": "Detalle",
        })
        detalle_cols = df_mov["Detalle"].apply(parsear_detalle_movimiento).apply(pd.Series)
        df_mov = pd.concat([df_mov, detalle_cols], axis=1)
        nombres_separados = df_mov["Nombre completo"].fillna("").astype(str).str.strip().str.split(n=1, expand=True)
        df_mov["Apellido"] = nombres_separados[0].fillna("").astype(str).str.strip().str.rstrip(",")
        df_mov["Nombre"] = nombres_separados[1].fillna("") if nombres_separados.shape[1] > 1 else ""
        df_mov["Nombre"] = df_mov["Nombre"].astype(str).str.strip().str.lstrip(",").str.strip()
        datos_personal = df[["ORDEN_LIMP", "DNI", "CE"]].copy()
        df_mov["Orden personal"] = pd.to_numeric(df_mov["Orden interno"], errors="coerce")
        df_mov = df_mov.merge(datos_personal, left_on="Orden personal", right_on="ORDEN_LIMP", how="left")
        df_mov["DNI"] = df_mov["DNI"].fillna("").astype(str).str.replace(".0", "", regex=False)
        df_mov["CE"] = df_mov["CE"].fillna("").astype(str).str.replace(".0", "", regex=False)
        df_mov["Fecha parte dt"] = pd.to_datetime(df_mov["Fecha parte"], errors="coerce").dt.date
        df_mov = df_mov[
            (df_mov["Fecha parte dt"] >= fecha_desde_hist) &
            (df_mov["Fecha parte dt"] <= fecha_hasta_hist)
        ].copy()
        df_mov["Clave aspirante"] = df_mov["Orden interno"].fillna(df_mov["Nombre completo"]).astype(str)
        df_mov = df_mov.sort_values("Fecha/hora", ascending=False).copy()

        if not df_mov.empty:
            fc1, fc2, fc3 = st.columns(3)
            with fc1:
                motivos_sel = st.multiselect("Motivo", sorted([m for m in df_mov["Motivo"].dropna().unique() if m]), placeholder="Filtrar")
            with fc2:
                presencias_sel = st.multiselect("Presencia", sorted([p for p in df_mov["Presencia"].dropna().unique() if p]), placeholder="Filtrar")
            with fc3:
                aulas_sel = st.multiselect("Aula", sorted([a for a in df_mov["Aula"].dropna().unique() if a]), placeholder="Filtrar")

            if motivos_sel:
                df_mov = df_mov[df_mov["Motivo"].isin(motivos_sel)]
            if presencias_sel:
                df_mov = df_mov[df_mov["Presencia"].isin(presencias_sel)]
            if aulas_sel:
                df_mov = df_mov[df_mov["Aula"].isin(aulas_sel)]

            if not df_mov.empty:
                opciones_asp = df_mov[["Clave aspirante", "Nombre completo", "Aula"]].drop_duplicates("Clave aspirante")
                opciones_asp = opciones_asp.sort_values("Nombre completo")
                etiquetas_asp = {
                    f"{row['Nombre completo']} | Aula: {row['Aula']}": row["Clave aspirante"]
                    for _, row in opciones_asp.iterrows()
                }
                aspirante_sel = st.selectbox(
                    "Buscar aspirante",
                    list(etiquetas_asp.keys()),
                    index=None,
                    placeholder="Escribi para buscar o deja vacio para ver todo",
                    key="historial_aspirante_sel",
                    help="Si no seleccionas un aspirante, se muestra todo el listado del rango de fechas"
                )
                if aspirante_sel:
                    df_mov = df_mov[df_mov["Clave aspirante"] == etiquetas_asp[aspirante_sel]]

        columnas = ["Fecha/hora", "Apellido", "Nombre", "DNI", "CE", "Aula", "Motivo", "Presencia", "Desde", "Hasta", "Observación"]

        if not df_mov.empty:
            st.caption(f"{len(df_mov)} movimiento(s) encontrados")
            st.dataframe(df_mov[columnas], use_container_width=True, hide_index=True)

            edit_mov_id = st.session_state.get("editando_movimiento_id")
            if edit_mov_id:
                mov_edit = df_mov[df_mov["id"] == edit_mov_id]
                if not mov_edit.empty:
                    mov = mov_edit.iloc[0]
                    with st.form("form_editar_movimiento"):
                        st.markdown(f"### Editar movimiento de {mov.get('Nombre completo', '-')}")
                        e1, e2, e3 = st.columns(3)
                        with e1:
                            mov_motivo = st.text_input("Motivo", value=str(mov.get("Motivo", "")), key="mov_edit_motivo")
                        with e2:
                            opciones_presencia = ["Ausente", "Presente en instituto", "Presente en escuadrón"]
                            presencia_actual = str(mov.get("Presencia", "Ausente"))
                            indice_presencia = opciones_presencia.index(presencia_actual) if presencia_actual in opciones_presencia else 0
                            mov_presencia = st.selectbox("Presencia", opciones_presencia, index=indice_presencia, key="mov_edit_presencia")
                        with e3:
                            mov_aula = st.text_input("Aula", value=str(mov.get("Aula", "")), key="mov_edit_aula")
                        e4, e5 = st.columns(2)
                        with e4:
                            mov_desde = st.text_input("Desde", value=str(mov.get("Desde", "")), key="mov_edit_desde")
                        with e5:
                            mov_hasta = st.text_input("Hasta", value=str(mov.get("Hasta", "")), key="mov_edit_hasta")
                        mov_obs = st.text_input("Observación", value=str(mov.get("Observación", "")), key="mov_edit_obs")
                        guardar_mov, cancelar_mov = st.columns([3, 1])
                        with guardar_mov:
                            guardar = st.form_submit_button("Guardar movimiento", type="primary", use_container_width=True)
                        with cancelar_mov:
                            cancelar = st.form_submit_button("Cancelar", use_container_width=True)
                        if guardar:
                            detalle_nuevo = f"{mov_motivo.upper()} | {mov_presencia} | {mov_desde.upper()} a {mov_hasta.upper()} | {mov_obs.upper()}"
                            orden_valor = mov.get("Orden interno")
                            try:
                                orden_valor = int(orden_valor) if pd.notna(orden_valor) and str(orden_valor).strip() else None
                            except Exception:
                                orden_valor = None
                            actualizar_movimiento(int(mov["id"]), {
                                "fecha_parte": mov.get("Fecha parte"),
                                "modulo": mov.get("Módulo", "Historial"),
                                "accion": "EDITAR MOVIMIENTO",
                                "orden": orden_valor,
                                "nombre": mov.get("Nombre completo"),
                                "aula": mov_aula.upper(),
                                "detalle": detalle_nuevo,
                            })
                            st.session_state.editando_movimiento_id = None
                            st.success("Movimiento actualizado")
                            st.rerun()
                        if cancelar:
                            st.session_state.editando_movimiento_id = None
                            st.rerun()

            with st.expander("Administrar movimientos filtrados", expanded=False):
                st.caption("Desde acá podés editar o borrar registros de prueba del historial.")
                for _, mov in df_mov.iterrows():
                    mov_id = int(mov["id"])
                    c_info, c_edit, c_del = st.columns([6, 1, 1])
                    with c_info:
                        st.markdown(f"**{mov.get('Apellido', '')} {mov.get('Nombre', '')}** | {mov.get('Motivo', '-')} | {mov.get('Presencia', '-')}")
                        st.caption(f"{mov.get('Fecha/hora', '-')} | Aula {mov.get('Aula', '-')} | {mov.get('Desde', '-')} a {mov.get('Hasta', '-')} | {mov.get('Observación', '')}")
                    with c_edit:
                        if st.button("Editar", key=f"edit_mov_{mov_id}", use_container_width=True):
                            st.session_state.editando_movimiento_id = mov_id
                            st.rerun()
                    with c_del:
                        if st.button("Eliminar", key=f"del_mov_{mov_id}", use_container_width=True):
                            eliminar_movimiento(mov_id)
                            if st.session_state.get("editando_movimiento_id") == mov_id:
                                st.session_state.editando_movimiento_id = None
                            st.toast("Movimiento eliminado")
                            st.rerun()
        else:
            st.info("No hay movimientos para los filtros seleccionados.")

        historial_df = df_mov[columnas] if not df_mov.empty else pd.DataFrame(columns=columnas)
        wb_hist = openpyxl.Workbook()
        ws_hist = wb_hist.active
        ws_hist.title = "HISTORIAL"
        from openpyxl.styles import Alignment, PatternFill, Border, Side

        titulo_fill = PatternFill(start_color=EXCEL_OLIVE_DARK, end_color=EXCEL_OLIVE_DARK, fill_type="solid")
        subtitulo_fill = PatternFill(start_color=EXCEL_OLIVE_LIGHT, end_color=EXCEL_OLIVE_LIGHT, fill_type="solid")
        header_fill = PatternFill(start_color=EXCEL_OLIVE, end_color=EXCEL_OLIVE, fill_type="solid")
        thin = Side(style="thin", color=EXCEL_BORDER)
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        ws_hist.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(columnas))
        ws_hist.cell(1, 1, "ESCUADRON H \"Cabo Marcelo Godoy\"")
        ws_hist.cell(1, 1).font = excel_font(bold=True, size=16, color=EXCEL_WHITE)
        ws_hist.cell(1, 1).alignment = Alignment(horizontal="center", vertical="center")
        ws_hist.cell(1, 1).fill = titulo_fill
        ws_hist.row_dimensions[1].height = 28

        ws_hist.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(columnas))
        ws_hist.cell(2, 1, "Historial De Movimientos de Aspirantes")
        ws_hist.cell(2, 1).font = excel_font(bold=True, size=13, color=EXCEL_OLIVE_DARK)
        ws_hist.cell(2, 1).alignment = Alignment(horizontal="center", vertical="center")
        ws_hist.cell(2, 1).fill = subtitulo_fill
        ws_hist.row_dimensions[2].height = 24

        ws_hist.merge_cells(start_row=3, start_column=1, end_row=3, end_column=len(columnas))
        ws_hist.cell(3, 1, f"Periodo: {fecha_desde_hist.strftime('%d/%m/%Y')} al {fecha_hasta_hist.strftime('%d/%m/%Y')} | Registros: {len(historial_df)}")
        ws_hist.cell(3, 1).font = excel_font(italic=True, size=10, color=EXCEL_TEXT_MUTED)
        ws_hist.cell(3, 1).alignment = Alignment(horizontal="center", vertical="center")

        header_row = 5
        for col_idx, col_name in enumerate(columnas, 1):
            cell = ws_hist.cell(header_row, col_idx, col_name)
            cell.font = excel_font(bold=True, color=EXCEL_WHITE)
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border

        for row_idx, (_, row) in enumerate(historial_df.iterrows(), header_row + 1):
            for col_idx, col_name in enumerate(columnas, 1):
                cell = ws_hist.cell(row_idx, col_idx, row.get(col_name, ""))
                cell.alignment = Alignment(vertical="center", wrap_text=True)
                cell.border = border
                if row_idx % 2 == 0:
                    cell.fill = PatternFill(start_color=EXCEL_OLIVE_ROW, end_color=EXCEL_OLIVE_ROW, fill_type="solid")

        widths = {
            "Fecha/hora": 20, "Apellido": 18, "Nombre": 28, "DNI": 14, "CE": 12,
            "Aula": 12, "Motivo": 18, "Presencia": 22, "Desde": 14, "Hasta": 14, "Observación": 36
        }
        for col_idx, col_name in enumerate(columnas, 1):
            ws_hist.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = widths.get(col_name, 16)
        ws_hist.freeze_panes = "A6"
        ws_hist.auto_filter.ref = f"A{header_row}:{openpyxl.utils.get_column_letter(len(columnas))}{max(header_row, header_row + len(historial_df))}"

        st.download_button(
            "descargar historial de movimiento",
            data=excel_bytes(wb_hist),
            file_name=f"HISTORIAL_MOVIMIENTOS_{fecha_desde_hist.strftime('%d%m%Y')}_{fecha_hasta_hist.strftime('%d%m%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    else:
        st.info("No hay movimientos registrados.")

with tab_res:
    if True:
        from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "PARTE DIARIO"
    
        thin = Side(style="thin", color=EXCEL_BORDER)
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        header_fill = PatternFill(start_color=EXCEL_OLIVE_DARK, end_color=EXCEL_OLIVE_DARK, fill_type="solid")
        sub_fill = PatternFill(start_color=EXCEL_OLIVE_LIGHT, end_color=EXCEL_OLIVE_LIGHT, fill_type="solid")
        fecha_titulo = st.session_state.fecha_reporte.strftime('%d%b%y').upper()
        dia_reporte = DIAS_SEMANA[st.session_state.fecha_reporte.weekday()]
    
        ws.merge_cells('A1:J1')
        ws['A1'] = f"PARTE DIARIO DEL ESCUADRÓN H - {fecha_titulo}"
        ws['A1'].font = excel_font(bold=True, size=16, color=EXCEL_WHITE)
        ws['A1'].fill = header_fill
        ws['A1'].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 26
    
        ws.merge_cells('A2:J2')
        ws['A2'] = f"Día: {dia_reporte} | Primera obligación: 06:00 hs | Generado: {datetime.now().strftime('%H:%M')}"
        ws['A2'].font = excel_font(italic=True, size=11, color=EXCEL_TEXT_MUTED)
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
            cell.font = excel_font(bold=True, size=9, color=EXCEL_TEXT_DARK)
            cell.fill = sub_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border
            val = ws.cell(row=5, column=col, value=metric_values[col - 1])
            val.font = excel_font(bold=True, size=12)
            val.alignment = Alignment(horizontal="center")
            val.border = border
    
        ws.merge_cells('A8:J8')
        ws['A8'] = "NOVEDADES DEL PERSONAL (AUSENTES JUSTIFICADOS)"
        ws['A8'].font = excel_font(bold=True, size=12, color=EXCEL_OLIVE_DARK)
    
        nov_headers = ["Nro", "GRADO", "APELLIDO Y NOMBRE", "DNI", "CE", "NOVEDAD", "DETALLE", "DESDE", "HASTA", "AULA"]
        for col, h in enumerate(nov_headers, 1):
            cell = ws.cell(row=9, column=col, value=h)
            cell.font = excel_font(bold=True, color=EXCEL_WHITE, size=9)
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
            ws.cell(row=current_row, column=1, value="Sin novedades registradas en la guardia").font = excel_font(italic=True, color=EXCEL_TEXT_MUTED)
            ws.cell(row=current_row, column=1).alignment = Alignment(horizontal="center")
            current_row += 1
    
        current_row += 2
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=10)
        ws.cell(row=current_row, column=1, value=f"HORARIOS DE INGRESO - {dia_reporte.upper()}").font = excel_font(bold=True, size=12, color=EXCEL_OLIVE_DARK)
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
        ws.cell(row=current_row, column=1, value="OBSERVACIONES").font = excel_font(bold=True, size=12, color=EXCEL_OLIVE_DARK)
    
        for col, width in enumerate([10, 14, 34, 12, 10, 18, 24, 12, 12, 12], 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = width
    
        output = f"PARTE_DIARIO_ESCUADRON_H_{datetime.now().strftime('%d%m%Y_%H%M')}.xlsx"
        st.download_button(
            "GENERAR PARTE DIARIO (EXCEL)",
            data=excel_bytes(wb),
            file_name=output,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
            key="download_parte_diario_formal"
        )
    
    # ==============================================================================
    # 5. EXPORTAR EXCEL (PARTE DIARIO DETALLADO)
    # ==============================================================================
    if True:
        from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "PARTE DIARIO DETALLADO"

        thin = Side(style="thin", color=EXCEL_BORDER)
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        header_fill = PatternFill(start_color=EXCEL_OLIVE_DARK, end_color=EXCEL_OLIVE_DARK, fill_type="solid")
        sub_fill = PatternFill(start_color=EXCEL_OLIVE_LIGHT, end_color=EXCEL_OLIVE_LIGHT, fill_type="solid")
        fecha_titulo = st.session_state.fecha_reporte.strftime('%d%b%y').upper()
        dia_reporte = DIAS_SEMANA[st.session_state.fecha_reporte.weekday()]

        ws.merge_cells('A1:J1')
        ws['A1'] = f"PARTE DIARIO DETALLADO DEL ESCUADRÓN H - {fecha_titulo}"
        ws['A1'].font = excel_font(bold=True, size=16, color=EXCEL_WHITE)
        ws['A1'].fill = header_fill
        ws['A1'].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 26

        ws.merge_cells('A2:J2')
        ws['A2'] = f"Día: {dia_reporte} | Primera obligación: 06:00 hs | Generado: {datetime.now().strftime('%H:%M')}"
        ws['A2'].font = excel_font(italic=True, size=11, color=EXCEL_TEXT_MUTED)
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
            cell.font = excel_font(bold=True, size=9, color=EXCEL_TEXT_DARK)
            cell.fill = sub_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border
            val = ws.cell(row=5, column=col, value=metric_values[col - 1])
            val.font = excel_font(bold=True, size=12)
            val.alignment = Alignment(horizontal="center")
            val.border = border

        ws.merge_cells('A8:J8')
        ws['A8'] = "NOVEDADES DEL PERSONAL"
        ws['A8'].font = excel_font(bold=True, size=12, color=EXCEL_OLIVE_DARK)

        nov_headers = ["Nro", "GRADO", "APELLIDO Y NOMBRE", "DNI", "CE", "NOVEDAD", "PRESENCIA", "DETALLE", "DESDE/HASTA", "AULA"]
        for col, h in enumerate(nov_headers, 1):
            cell = ws.cell(row=9, column=col, value=h)
            cell.font = excel_font(bold=True, color=EXCEL_WHITE, size=9)
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")
            cell.border = border

        current_row = 10
        if st.session_state.novedades_lista:
            for i, nov in enumerate(st.session_state.novedades_lista, 1):
                row = current_row + i - 1
                values = [
                    i, nov.get('grado', ''), nov['nombre'], nov.get('dni', ''), nov.get('ce', ''),
                    nov['estado'], AMBITOS_NOVEDAD.get(ambito_efectivo(nov), ""), nov['detalle'],
                    f"{nov['fecha_ini']} a {nov['fecha_fin']}", nov.get('aula', '-')
                ]
                for col, value in enumerate(values, 1):
                    cell = ws.cell(row=row, column=col, value=value)
                    cell.border = border
                    cell.alignment = Alignment(horizontal="center" if col in [1, 2, 4, 5, 6, 7, 9, 10] else "left")
            current_row += len(st.session_state.novedades_lista)
        else:
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=10)
            ws.cell(row=current_row, column=1, value="Sin novedades registradas en la guardia").font = excel_font(italic=True, color=EXCEL_TEXT_MUTED)
            ws.cell(row=current_row, column=1).alignment = Alignment(horizontal="center")
            current_row += 1

        current_row += 2
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=10)
        ws.cell(row=current_row, column=1, value="PERSONAL QUE ALMUERZA").font = excel_font(bold=True, size=12, color=EXCEL_OLIVE_DARK)
        current_row += 1
        alm_headers = ["Nro", "APELLIDO Y NOMBRE", "AULA", "CE", "DNI"]
        for col, h in enumerate(alm_headers, 1):
            cell = ws.cell(row=current_row, column=col, value=h)
            cell.font = excel_font(bold=True, color=EXCEL_WHITE, size=9)
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")
            cell.border = border
        current_row += 1
        if st.session_state.lista_almuerzo:
            for i, orden in enumerate(sorted(st.session_state.lista_almuerzo), 1):
                asp = df[df['ORDEN_LIMP'] == orden]
                if not asp.empty:
                    p = asp.iloc[0]
                    values = [i, p['NOMBRE_COMPLETO'], p['AULA'], p['CE'], p['DNI']]
                    for col, value in enumerate(values, 1):
                        cell = ws.cell(row=current_row, column=col, value=value)
                        cell.border = border
                        cell.alignment = Alignment(horizontal="center" if col != 2 else "left")
                    current_row += 1
        else:
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=5)
            ws.cell(row=current_row, column=1, value="Sin personal cargado para almuerzo").font = excel_font(italic=True, color=EXCEL_TEXT_MUTED)
            current_row += 1

        current_row += 2
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=10)
        ws.cell(row=current_row, column=1, value=f"HORARIOS DE INGRESO - {dia_reporte.upper()}").font = excel_font(bold=True, size=12, color=EXCEL_OLIVE_DARK)
        current_row += 1
        hor_headers = ["AULA", "ENT. MAÑANA", "SAL. MAÑANA", "ENT. TARDE", "SAL. TARDE"]
        for col, h in enumerate(hor_headers, 1):
            cell = ws.cell(row=current_row, column=col, value=h)
            cell.font = excel_font(bold=True, color=EXCEL_WHITE, size=9)
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")
            cell.border = border
        current_row += 1
        for aula in AULAS_UNICAS:
            hor = st.session_state.horarios_config[aula]
            values = [aula, hor.get('ent_m'), hor.get('sal_m'), hor.get('ent_t'), hor.get('sal_t')]
            for col, value in enumerate(values, 1):
                cell = ws.cell(row=current_row, column=col, value=value)
                cell.border = border
                cell.alignment = Alignment(horizontal="center")
            current_row += 1

        current_row += 2
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=10)
        ws.cell(row=current_row, column=1, value="CONTROL DE AULAS").font = excel_font(bold=True, size=12, color=EXCEL_OLIVE_DARK)
        current_row += 1
        aula_headers = ["AULA", "TOTAL", "PRESENTES", "AUSENTES", "ALMUERZAN", "UBICACIÓN", "ESTADO M", "ESTADO T"]
        for col, h in enumerate(aula_headers, 1):
            cell = ws.cell(row=current_row, column=col, value=h)
            cell.font = excel_font(bold=True, color=EXCEL_WHITE, size=9)
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")
            cell.border = border
        current_row += 1
        for aula in AULAS_UNICAS:
            cfg = st.session_state.estado_aulas[aula]
            alumnos = df[df['AULA'] == aula]
            ausentes_aula = len({row['ORDEN_LIMP'] for _, row in alumnos.iterrows() if row['ORDEN_LIMP'] in total_ausentes})
            almuerzan = sum(1 for _, row in alumnos.iterrows() if row['ORDEN_LIMP'] in st.session_state.lista_almuerzo)
            values = [
                aula, len(alumnos), len(alumnos) - ausentes_aula, ausentes_aula, almuerzan,
                cfg.get("ubicacion_m", "-") if cfg["estado_m"] == "EN INSTITUTO" else "FUERA",
                cfg["estado_m"], cfg["estado_t"]
            ]
            for col, value in enumerate(values, 1):
                cell = ws.cell(row=current_row, column=col, value=value)
                cell.border = border
                cell.alignment = Alignment(horizontal="center")
            current_row += 1

        for col, width in enumerate([10, 16, 34, 12, 10, 18, 18, 24, 18, 12], 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = width

        output = f"PARTE_DIARIO_DETALLADO_{datetime.now().strftime('%d%m%Y_%H%M')}.xlsx"
        st.download_button(
            "GENERAR PARTE DIARIO DETALLADO (EXCEL)",
            data=excel_bytes(wb),
            file_name=output,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="secondary",
            use_container_width=True,
            key="download_parte_diario_detallado"
        )
