import os
import sqlite3
import hashlib
import hmac
from datetime import datetime

DB_PATH = "parte_diario.db"


def get_database_url():
    if os.getenv("DATABASE_URL"):
        return os.getenv("DATABASE_URL")

    try:
        import streamlit as st
        return st.secrets.get("DATABASE_URL")
    except Exception:
        return None


def is_postgres():
    url = get_database_url()
    return bool(url and url.startswith(("postgresql://", "postgres://")))


def get_db():
    if is_postgres():
        import psycopg2
        from psycopg2.extras import RealDictCursor

        return psycopg2.connect(get_database_url(), cursor_factory=RealDictCursor)

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def run(conn, sql, params=()):
    if is_postgres():
        cur = conn.cursor()
        cur.execute(sql.replace("?", "%s"), params)
        return cur
    return conn.execute(sql, params)


def fetch_all(cur):
    return [dict(r) for r in cur.fetchall()]


def parse_fecha(value):
    if not value or str(value).upper() == "N/O":
        return None

    text = str(value).strip().upper()
    for fmt in ("%Y-%m-%d", "%d%b%y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return None


def novedad_activa_en_fecha(novedad, fecha):
    fecha_obj = parse_fecha(fecha)
    ini = parse_fecha(novedad.get("fecha_ini"))
    fin = parse_fecha(novedad.get("fecha_fin"))

    if not fecha_obj or not ini:
        return True
    if not fin:
        return ini <= fecha_obj
    return ini <= fecha_obj <= fin


def init_db():
    with get_db() as conn:
        if is_postgres():
            _init_postgres(conn)
        else:
            _init_sqlite(conn)
        conn.commit()


def _init_sqlite(conn):
    c = conn.cursor()
    c.execute("""CREATE TABLE IF NOT EXISTS novedades (
        id INTEGER PRIMARY KEY AUTOINCREMENT, orden INTEGER NOT NULL, grado TEXT, nombre TEXT NOT NULL,
        dni TEXT, ce TEXT, aula TEXT, estado TEXT NOT NULL, detalle TEXT, fecha_ini TEXT NOT NULL, fecha_fin TEXT NOT NULL,
        ambito TEXT DEFAULT 'AUSENTE',
        UNIQUE(orden, fecha_ini))""")
    cols = [r[1] for r in c.execute("PRAGMA table_info(novedades)").fetchall()]
    if "ambito" not in cols:
        c.execute("ALTER TABLE novedades ADD COLUMN ambito TEXT DEFAULT 'AUSENTE'")

    c.execute("""CREATE TABLE IF NOT EXISTS estado_aulas (
        fecha TEXT NOT NULL, aula TEXT NOT NULL, estado_m TEXT DEFAULT 'EN INSTITUTO', estado_t TEXT DEFAULT 'EN INSTITUTO',
        salida_m TEXT, salida_t TEXT, ubicacion_m TEXT DEFAULT 'EN AULA', ubicacion_t TEXT DEFAULT 'EN AULA',
        PRIMARY KEY (fecha, aula))""")

    c.execute("""CREATE TABLE IF NOT EXISTS almuerzo (
        fecha TEXT NOT NULL, orden INTEGER NOT NULL, PRIMARY KEY (fecha, orden))""")

    c.execute("""CREATE TABLE IF NOT EXISTS almuerzo_historial (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fecha_hora TEXT NOT NULL,
        fecha TEXT NOT NULL,
        orden INTEGER NOT NULL,
        nombre TEXT,
        aula TEXT
    )""")
    cols_alm_hist = [r[1] for r in c.execute("PRAGMA table_info(almuerzo_historial)").fetchall()]
    migraciones_alm_hist = {
        "fecha_hora": "TEXT",
        "fecha": "TEXT",
        "orden": "INTEGER",
        "nombre": "TEXT",
        "aula": "TEXT",
    }
    for col, tipo in migraciones_alm_hist.items():
        if col not in cols_alm_hist:
            c.execute(f"ALTER TABLE almuerzo_historial ADD COLUMN {col} {tipo}")
    if "fecha" not in cols_alm_hist:
        c.execute("""UPDATE almuerzo_historial
            SET fecha = substr(fecha_hora, 1, 10)
            WHERE fecha IS NULL AND fecha_hora IS NOT NULL""")

    c.execute("""CREATE TABLE IF NOT EXISTS horarios_config (
        aula TEXT PRIMARY KEY, ent_m TEXT DEFAULT '06:00', sal_m TEXT DEFAULT '12:00',
        ent_t TEXT DEFAULT '13:00', sal_t TEXT DEFAULT '19:00')""")

    c.execute("""CREATE TABLE IF NOT EXISTS horarios_diarios (
        aula TEXT NOT NULL, dia TEXT NOT NULL, ent_m TEXT DEFAULT '06:00', sal_m TEXT DEFAULT '12:00',
        ent_t TEXT DEFAULT '13:00', sal_t TEXT DEFAULT '19:00',
        PRIMARY KEY (aula, dia))""")

    c.execute("""CREATE TABLE IF NOT EXISTS asistencia_diaria (
        fecha TEXT NOT NULL, orden INTEGER NOT NULL, estado TEXT DEFAULT 'PRESENTE',
        PRIMARY KEY (fecha, orden))""")

    c.execute("""CREATE TABLE IF NOT EXISTS plan_llamada (
        orden INTEGER PRIMARY KEY, domicilio TEXT, telefono_personal TEXT, telefono_emergencia TEXT,
        nombre_emergencia TEXT, parentesco_emergencia TEXT, observaciones TEXT)""")

    c.execute("""CREATE TABLE IF NOT EXISTS movimientos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fecha_hora TEXT NOT NULL,
        fecha_parte TEXT NOT NULL,
        modulo TEXT NOT NULL,
        accion TEXT NOT NULL,
        orden INTEGER,
        nombre TEXT,
        aula TEXT,
        detalle TEXT
    )""")

    c.execute("""CREATE TABLE IF NOT EXISTS usuarios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        usuario TEXT NOT NULL UNIQUE,
        password_hash TEXT NOT NULL,
        salt TEXT NOT NULL,
        rol TEXT NOT NULL DEFAULT 'usuario',
        activo INTEGER NOT NULL DEFAULT 1,
        creado_en TEXT NOT NULL
    )""")


def _init_postgres(conn):
    cur = conn.cursor()
    cur.execute("""CREATE TABLE IF NOT EXISTS novedades (
        id INTEGER GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
        orden INTEGER NOT NULL,
        grado TEXT,
        nombre TEXT NOT NULL,
        dni TEXT,
        ce TEXT,
        aula TEXT,
        estado TEXT NOT NULL,
        detalle TEXT,
        fecha_ini TEXT NOT NULL,
        fecha_fin TEXT NOT NULL,
        ambito TEXT DEFAULT 'AUSENTE',
        UNIQUE(orden, fecha_ini)
    )""")

    cur.execute("""CREATE TABLE IF NOT EXISTS estado_aulas (
        fecha TEXT NOT NULL,
        aula TEXT NOT NULL,
        estado_m TEXT DEFAULT 'EN INSTITUTO',
        estado_t TEXT DEFAULT 'EN INSTITUTO',
        salida_m TEXT,
        salida_t TEXT,
        ubicacion_m TEXT DEFAULT 'EN AULA',
        ubicacion_t TEXT DEFAULT 'EN AULA',
        PRIMARY KEY (fecha, aula)
    )""")

    cur.execute("""CREATE TABLE IF NOT EXISTS almuerzo (
        fecha TEXT NOT NULL,
        orden INTEGER NOT NULL,
        PRIMARY KEY (fecha, orden)
    )""")

    cur.execute("""CREATE TABLE IF NOT EXISTS almuerzo_historial (
        id INTEGER GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
        fecha_hora TEXT NOT NULL,
        fecha TEXT NOT NULL,
        orden INTEGER NOT NULL,
        nombre TEXT,
        aula TEXT
    )""")
    cur.execute("ALTER TABLE almuerzo_historial ADD COLUMN IF NOT EXISTS fecha_hora TEXT")
    cur.execute("ALTER TABLE almuerzo_historial ADD COLUMN IF NOT EXISTS fecha TEXT")
    cur.execute("ALTER TABLE almuerzo_historial ADD COLUMN IF NOT EXISTS orden INTEGER")
    cur.execute("ALTER TABLE almuerzo_historial ADD COLUMN IF NOT EXISTS nombre TEXT")
    cur.execute("ALTER TABLE almuerzo_historial ADD COLUMN IF NOT EXISTS aula TEXT")
    cur.execute("""UPDATE almuerzo_historial
        SET fecha = substr(fecha_hora, 1, 10)
        WHERE fecha IS NULL AND fecha_hora IS NOT NULL""")

    cur.execute("""CREATE TABLE IF NOT EXISTS horarios_config (
        aula TEXT PRIMARY KEY,
        ent_m TEXT DEFAULT '06:00',
        sal_m TEXT DEFAULT '12:00',
        ent_t TEXT DEFAULT '13:00',
        sal_t TEXT DEFAULT '19:00'
    )""")

    cur.execute("""CREATE TABLE IF NOT EXISTS horarios_diarios (
        aula TEXT NOT NULL,
        dia TEXT NOT NULL,
        ent_m TEXT DEFAULT '06:00',
        sal_m TEXT DEFAULT '12:00',
        ent_t TEXT DEFAULT '13:00',
        sal_t TEXT DEFAULT '19:00',
        PRIMARY KEY (aula, dia)
    )""")

    cur.execute("""CREATE TABLE IF NOT EXISTS asistencia_diaria (
        fecha TEXT NOT NULL,
        orden INTEGER NOT NULL,
        estado TEXT DEFAULT 'PRESENTE',
        PRIMARY KEY (fecha, orden)
    )""")

    cur.execute("""CREATE TABLE IF NOT EXISTS plan_llamada (
        orden INTEGER PRIMARY KEY,
        domicilio TEXT,
        telefono_personal TEXT,
        telefono_emergencia TEXT,
        nombre_emergencia TEXT,
        parentesco_emergencia TEXT,
        observaciones TEXT
    )""")

    cur.execute("""CREATE TABLE IF NOT EXISTS movimientos (
        id INTEGER GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
        fecha_hora TEXT NOT NULL,
        fecha_parte TEXT NOT NULL,
        modulo TEXT NOT NULL,
        accion TEXT NOT NULL,
        orden INTEGER,
        nombre TEXT,
        aula TEXT,
        detalle TEXT
    )""")

    cur.execute("""CREATE TABLE IF NOT EXISTS usuarios (
        id INTEGER GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
        usuario TEXT NOT NULL UNIQUE,
        password_hash TEXT NOT NULL,
        salt TEXT NOT NULL,
        rol TEXT NOT NULL DEFAULT 'usuario',
        activo INTEGER NOT NULL DEFAULT 1,
        creado_en TEXT NOT NULL
    )""")



# --- USUARIOS Y SEGURIDAD ---
def _hash_password(password, salt=None):
    if salt is None:
        salt = os.urandom(16).hex()
    digest = hashlib.pbkdf2_hmac("sha256", str(password).encode("utf-8"), bytes.fromhex(salt), 120000)
    return salt, digest.hex()


def contar_usuarios():
    with get_db() as conn:
        cur = run(conn, "SELECT COUNT(*) AS total FROM usuarios")
        row = cur.fetchone()
        return int(row["total"] if isinstance(row, dict) else row[0])


def crear_usuario(usuario, password, rol="usuario", activo=True):
    usuario = str(usuario).strip().lower()
    salt, password_hash = _hash_password(password)
    with get_db() as conn:
        run(conn, """INSERT INTO usuarios (usuario, password_hash, salt, rol, activo, creado_en)
            VALUES (?, ?, ?, ?, ?, ?)""",
            (usuario, password_hash, salt, rol, 1 if activo else 0, datetime.now().isoformat(timespec="seconds")))
        conn.commit()


def obtener_usuario(usuario):
    with get_db() as conn:
        cur = run(conn, "SELECT * FROM usuarios WHERE usuario=?", (str(usuario).strip().lower(),))
        row = cur.fetchone()
        return dict(row) if row else None


def autenticar_usuario(usuario, password):
    user = obtener_usuario(usuario)
    if not user or int(user.get("activo", 0)) != 1:
        return None
    _, password_hash = _hash_password(password, user["salt"])
    if not hmac.compare_digest(password_hash, user["password_hash"]):
        return None
    return {"id": user["id"], "usuario": user["usuario"], "rol": user.get("rol", "usuario")}


def listar_usuarios():
    with get_db() as conn:
        cur = run(conn, "SELECT id, usuario, rol, activo, creado_en FROM usuarios ORDER BY usuario")
        return fetch_all(cur)


def actualizar_password_usuario(id_usuario, password):
    salt, password_hash = _hash_password(password)
    with get_db() as conn:
        run(conn, "UPDATE usuarios SET password_hash=?, salt=? WHERE id=?", (password_hash, salt, id_usuario))
        conn.commit()


def actualizar_estado_usuario(id_usuario, activo):
    with get_db() as conn:
        run(conn, "UPDATE usuarios SET activo=? WHERE id=?", (1 if activo else 0, id_usuario))
        conn.commit()


def eliminar_usuario(id_usuario):
    with get_db() as conn:
        run(conn, "DELETE FROM usuarios WHERE id=?", (id_usuario,))
        conn.commit()

# --- NOVEDADES ---
def obtener_novedades(fecha=None):
    with get_db() as conn:
        cur = run(conn, "SELECT * FROM novedades ORDER BY orden")
        novedades = fetch_all(cur)
        if fecha:
            return [n for n in novedades if novedad_activa_en_fecha(n, fecha)]
        return novedades


def agregar_novedad(data):
    with get_db() as conn:
        if is_postgres():
            run(conn, """INSERT INTO novedades
                (orden, grado, nombre, dni, ce, aula, estado, detalle, fecha_ini, fecha_fin, ambito)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT (orden, fecha_ini) DO NOTHING""",
                (data['orden'], data['grado'], data['nombre'], data['dni'], data['ce'],
                 data['aula'], data['estado'], data['detalle'], data['fecha_ini'], data['fecha_fin'],
                 data.get('ambito', 'AUSENTE')))
        else:
            run(conn, """INSERT OR IGNORE INTO novedades
                (orden, grado, nombre, dni, ce, aula, estado, detalle, fecha_ini, fecha_fin, ambito)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (data['orden'], data['grado'], data['nombre'], data['dni'], data['ce'],
                 data['aula'], data['estado'], data['detalle'], data['fecha_ini'], data['fecha_fin'],
                 data.get('ambito', 'AUSENTE')))
        conn.commit()


def actualizar_novedad(id_nov, data):
    with get_db() as conn:
        run(conn, """UPDATE novedades SET estado=?, detalle=?, fecha_ini=?, fecha_fin=?, ambito=? WHERE id=?""",
            (data['estado'], data['detalle'], data['fecha_ini'], data['fecha_fin'], data.get('ambito', 'AUSENTE'), id_nov))
        conn.commit()


def eliminar_novedad(id_nov):
    with get_db() as conn:
        run(conn, "DELETE FROM novedades WHERE id=?", (id_nov,))
        conn.commit()


def vaciar_novedades():
    with get_db() as conn:
        run(conn, "DELETE FROM novedades")
        conn.commit()


# --- HISTORIAL DE MOVIMIENTOS ---
def registrar_movimiento(fecha_parte, modulo, accion, orden=None, nombre=None, aula=None, detalle=""):
    with get_db() as conn:
        run(conn, """INSERT INTO movimientos
            (fecha_hora, fecha_parte, modulo, accion, orden, nombre, aula, detalle)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (datetime.now().isoformat(timespec="seconds"), fecha_parte, modulo, accion, orden, nombre, aula, detalle))
        conn.commit()


def obtener_movimientos(fecha_parte=None):
    with get_db() as conn:
        if fecha_parte:
            cur = run(conn, """SELECT * FROM movimientos
                WHERE fecha_parte=? ORDER BY fecha_hora DESC, id DESC""", (fecha_parte,))
        else:
            cur = run(conn, "SELECT * FROM movimientos ORDER BY fecha_hora DESC, id DESC")
        return fetch_all(cur)


def actualizar_movimiento(id_mov, data):
    with get_db() as conn:
        run(conn, """UPDATE movimientos SET
            fecha_parte=?, modulo=?, accion=?, orden=?, nombre=?, aula=?, detalle=?
            WHERE id=?""",
            (
                data.get('fecha_parte'), data.get('modulo'), data.get('accion'),
                data.get('orden'), data.get('nombre'), data.get('aula'),
                data.get('detalle', ''), id_mov
            ))
        conn.commit()


def eliminar_movimiento(id_mov):
    with get_db() as conn:
        run(conn, "DELETE FROM movimientos WHERE id=?", (id_mov,))
        conn.commit()


# --- ESTADO AULAS ---
def obtener_estado_aulas(fecha):
    with get_db() as conn:
        cur = run(conn, "SELECT * FROM estado_aulas WHERE fecha=?", (fecha,))
        return {r['aula']: dict(r) for r in cur.fetchall()}


def guardar_estado_aula(fecha, aula, estado_m, estado_t, salida_m, salida_t, ubicacion_m='EN AULA', ubicacion_t='EN AULA'):
    with get_db() as conn:
        if is_postgres():
            run(conn, """INSERT INTO estado_aulas
                (fecha, aula, estado_m, estado_t, salida_m, salida_t, ubicacion_m, ubicacion_t)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT (fecha, aula) DO UPDATE SET
                    estado_m=EXCLUDED.estado_m, estado_t=EXCLUDED.estado_t,
                    salida_m=EXCLUDED.salida_m, salida_t=EXCLUDED.salida_t,
                    ubicacion_m=EXCLUDED.ubicacion_m, ubicacion_t=EXCLUDED.ubicacion_t""",
                (fecha, aula, estado_m, estado_t, salida_m, salida_t, ubicacion_m, ubicacion_t))
        else:
            run(conn, """INSERT OR REPLACE INTO estado_aulas
                (fecha, aula, estado_m, estado_t, salida_m, salida_t, ubicacion_m, ubicacion_t)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                (fecha, aula, estado_m, estado_t, salida_m, salida_t, ubicacion_m, ubicacion_t))
        conn.commit()


# --- ALMUERZO ---
def obtener_almuerzo(fecha):
    with get_db() as conn:
        cur = run(conn, "SELECT orden FROM almuerzo WHERE fecha=?", (fecha,))
        return {r['orden'] for r in cur.fetchall()}


def agregar_almuerzo(fecha, orden):
    with get_db() as conn:
        if is_postgres():
            run(conn, "INSERT INTO almuerzo (fecha, orden) VALUES (?, ?) ON CONFLICT (fecha, orden) DO NOTHING", (fecha, orden))
        else:
            run(conn, "INSERT OR IGNORE INTO almuerzo (fecha, orden) VALUES (?, ?)", (fecha, orden))
        conn.commit()


def quitar_almuerzo(fecha, orden):
    with get_db() as conn:
        run(conn, "DELETE FROM almuerzo WHERE fecha=? AND orden=?", (fecha, orden))
        conn.commit()


def registrar_almuerzo_historial(fecha, orden, nombre=None, aula=None):
    with get_db() as conn:
        run(conn, """INSERT INTO almuerzo_historial
            (fecha_hora, fecha, orden, nombre, aula)
            VALUES (?, ?, ?, ?, ?)""",
            (datetime.now().isoformat(timespec="seconds"), fecha, orden, nombre, aula))
        conn.commit()


def obtener_historial_almuerzo(fecha_desde=None, fecha_hasta=None):
    with get_db() as conn:
        if fecha_desde and fecha_hasta:
            cur = run(conn, """SELECT * FROM almuerzo_historial
                WHERE fecha BETWEEN ? AND ?
                ORDER BY fecha DESC, fecha_hora DESC, id DESC""", (fecha_desde, fecha_hasta))
        else:
            cur = run(conn, "SELECT * FROM almuerzo_historial ORDER BY fecha DESC, fecha_hora DESC, id DESC")
        return fetch_all(cur)


# --- HORARIOS ---
def obtener_horarios():
    with get_db() as conn:
        cur = run(conn, "SELECT * FROM horarios_config")
        return {r['aula']: dict(r) for r in cur.fetchall()}


def obtener_horarios_dia(dia):
    with get_db() as conn:
        cur = run(conn, "SELECT * FROM horarios_diarios WHERE dia=?", (dia,))
        return {r['aula']: dict(r) for r in cur.fetchall()}


def guardar_horarios(aula, data):
    with get_db() as conn:
        if is_postgres():
            run(conn, """INSERT INTO horarios_config
                (aula, ent_m, sal_m, ent_t, sal_t) VALUES (?, ?, ?, ?, ?)
                ON CONFLICT (aula) DO UPDATE SET
                    ent_m=EXCLUDED.ent_m, sal_m=EXCLUDED.sal_m,
                    ent_t=EXCLUDED.ent_t, sal_t=EXCLUDED.sal_t""",
                (aula, data['ent_m'], data['sal_m'], data['ent_t'], data['sal_t']))
        else:
            run(conn, """INSERT OR REPLACE INTO horarios_config
                (aula, ent_m, sal_m, ent_t, sal_t) VALUES (?, ?, ?, ?, ?)""",
                (aula, data['ent_m'], data['sal_m'], data['ent_t'], data['sal_t']))
        conn.commit()


def guardar_horarios_dia(aula, dia, data):
    with get_db() as conn:
        if is_postgres():
            run(conn, """INSERT INTO horarios_diarios
                (aula, dia, ent_m, sal_m, ent_t, sal_t) VALUES (?, ?, ?, ?, ?, ?)
                ON CONFLICT (aula, dia) DO UPDATE SET
                    ent_m=EXCLUDED.ent_m, sal_m=EXCLUDED.sal_m,
                    ent_t=EXCLUDED.ent_t, sal_t=EXCLUDED.sal_t""",
                (aula, dia, data['ent_m'], data['sal_m'], data['ent_t'], data['sal_t']))
        else:
            run(conn, """INSERT OR REPLACE INTO horarios_diarios
                (aula, dia, ent_m, sal_m, ent_t, sal_t) VALUES (?, ?, ?, ?, ?, ?)""",
                (aula, dia, data['ent_m'], data['sal_m'], data['ent_t'], data['sal_t']))
        conn.commit()


# --- ASISTENCIA ---
def obtener_asistencia(fecha):
    with get_db() as conn:
        cur = run(conn, "SELECT * FROM asistencia_diaria WHERE fecha=?", (fecha,))
        return {r['orden']: r['estado'] for r in cur.fetchall()}


def actualizar_asistencia(fecha, orden, estado):
    with get_db() as conn:
        if is_postgres():
            run(conn, """INSERT INTO asistencia_diaria (fecha, orden, estado) VALUES (?, ?, ?)
                ON CONFLICT (fecha, orden) DO UPDATE SET estado=EXCLUDED.estado""",
                (fecha, orden, estado))
        else:
            run(conn, """INSERT OR REPLACE INTO asistencia_diaria (fecha, orden, estado) VALUES (?, ?, ?)""",
                (fecha, orden, estado))
        conn.commit()


# --- PLAN DE LLAMADA ---
def obtener_contacto(orden):
    with get_db() as conn:
        cur = run(conn, "SELECT * FROM plan_llamada WHERE orden=?", (orden,))
        row = cur.fetchone()
        return dict(row) if row else None


def guardar_contacto(data):
    with get_db() as conn:
        params = (
            data['orden'], data.get('domicilio', ''), data.get('telefono_personal', ''),
            data.get('telefono_emergencia', ''), data.get('nombre_emergencia', ''),
            data.get('parentesco_emergencia', ''), data.get('observaciones', '')
        )
        if is_postgres():
            run(conn, """INSERT INTO plan_llamada
                (orden, domicilio, telefono_personal, telefono_emergencia,
                 nombre_emergencia, parentesco_emergencia, observaciones)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT (orden) DO UPDATE SET
                    domicilio=EXCLUDED.domicilio,
                    telefono_personal=EXCLUDED.telefono_personal,
                    telefono_emergencia=EXCLUDED.telefono_emergencia,
                    nombre_emergencia=EXCLUDED.nombre_emergencia,
                    parentesco_emergencia=EXCLUDED.parentesco_emergencia,
                    observaciones=EXCLUDED.observaciones""", params)
        else:
            run(conn, """INSERT OR REPLACE INTO plan_llamada
                (orden, domicilio, telefono_personal, telefono_emergencia,
                 nombre_emergencia, parentesco_emergencia, observaciones)
                VALUES (?, ?, ?, ?, ?, ?, ?)""", params)
        conn.commit()


def obtener_todos_contactos():
    with get_db() as conn:
        cur = run(conn, "SELECT * FROM plan_llamada ORDER BY orden")
        return fetch_all(cur)
