import sqlite3
from datetime import datetime

DB_PATH = "parte_diario.db"

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with get_db() as conn:
        c = conn.cursor()
        c.execute("""CREATE TABLE IF NOT EXISTS novedades (
            id INTEGER PRIMARY KEY AUTOINCREMENT, orden INTEGER NOT NULL, grado TEXT, nombre TEXT NOT NULL,
            dni TEXT, ce TEXT, aula TEXT, estado TEXT NOT NULL, detalle TEXT, fecha_ini TEXT NOT NULL, fecha_fin TEXT NOT NULL,
            UNIQUE(orden, fecha_ini))""")
            
        c.execute("""CREATE TABLE IF NOT EXISTS estado_aulas (
            fecha TEXT NOT NULL, aula TEXT NOT NULL, estado_m TEXT DEFAULT 'EN INSTITUTO', estado_t TEXT DEFAULT 'EN INSTITUTO',
            salida_m TEXT, salida_t TEXT, ubicacion_m TEXT DEFAULT 'EN AULA', ubicacion_t TEXT DEFAULT 'EN AULA',
            PRIMARY KEY (fecha, aula))""")
            
        c.execute("""CREATE TABLE IF NOT EXISTS almuerzo (
            fecha TEXT NOT NULL, orden INTEGER NOT NULL, PRIMARY KEY (fecha, orden))""")
            
        c.execute("""CREATE TABLE IF NOT EXISTS horarios_config (
            aula TEXT PRIMARY KEY, ent_m TEXT DEFAULT '06:00', sal_m TEXT DEFAULT '12:00',
            ent_t TEXT DEFAULT '13:00', sal_t DEFAULT '19:00')""")
            
        c.execute("""CREATE TABLE IF NOT EXISTS asistencia_diaria (
            fecha TEXT NOT NULL, orden INTEGER NOT NULL, estado TEXT DEFAULT 'PRESENTE',
            PRIMARY KEY (fecha, orden))""")

        c.execute("""CREATE TABLE IF NOT EXISTS plan_llamada (
            orden INTEGER PRIMARY KEY, domicilio TEXT, telefono_personal TEXT, telefono_emergencia TEXT,
            nombre_emergencia TEXT, parentesco_emergencia TEXT, observaciones TEXT)""")
            
        conn.commit()
        init_horarios_dia()
# --- NOVEDADES ---
def obtener_novedades():
    with get_db() as conn:
        cur = conn.execute("SELECT * FROM novedades ORDER BY orden")
        return [dict(r) for r in cur.fetchall()]

def agregar_novedad(data):
    with get_db() as conn:
        conn.execute("""INSERT OR IGNORE INTO novedades 
            (orden, grado, nombre, dni, ce, aula, estado, detalle, fecha_ini, fecha_fin)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (data['orden'], data['grado'], data['nombre'], data['dni'], data['ce'], 
             data['aula'], data['estado'], data['detalle'], data['fecha_ini'], data['fecha_fin']))
        conn.commit()

def actualizar_novedad(id_nov, data):
    with get_db() as conn:
        conn.execute("""UPDATE novedades SET estado=?, detalle=?, fecha_ini=?, fecha_fin=? WHERE id=?""",
            (data['estado'], data['detalle'], data['fecha_ini'], data['fecha_fin'], id_nov))
        conn.commit()

def eliminar_novedad(id_nov):
    with get_db() as conn:
        conn.execute("DELETE FROM novedades WHERE id=?", (id_nov,))
        conn.commit()

def vaciar_novedades():
    with get_db() as conn:
        conn.execute("DELETE FROM novedades")
        conn.commit()

# --- ESTADO AULAS ---
def obtener_estado_aulas(fecha):
    with get_db() as conn:
        cur = conn.execute("SELECT * FROM estado_aulas WHERE fecha=?", (fecha,))
        return {r['aula']: dict(r) for r in cur.fetchall()}

def guardar_estado_aula(fecha, aula, estado_m, estado_t, salida_m, salida_t, ubicacion_m='EN AULA', ubicacion_t='EN AULA'):
    with get_db() as conn:
        conn.execute("""INSERT OR REPLACE INTO estado_aulas 
            (fecha, aula, estado_m, estado_t, salida_m, salida_t, ubicacion_m, ubicacion_t)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (fecha, aula, estado_m, estado_t, salida_m, salida_t, ubicacion_m, ubicacion_t))
        conn.commit()

# --- ALMUERZO ---
def obtener_almuerzo(fecha):
    with get_db() as conn:
        cur = conn.execute("SELECT orden FROM almuerzo WHERE fecha=?", (fecha,))
        return {r['orden'] for r in cur.fetchall()}

def agregar_almuerzo(fecha, orden):
    with get_db() as conn:
        conn.execute("INSERT OR IGNORE INTO almuerzo (fecha, orden) VALUES (?, ?)", (fecha, orden))
        conn.commit()

def quitar_almuerzo(fecha, orden):
    with get_db() as conn:
        conn.execute("DELETE FROM almuerzo WHERE fecha=? AND orden=?", (fecha, orden))
        conn.commit()

# --- HORARIOS ---
def obtener_horarios():
    with get_db() as conn:
        cur = conn.execute("SELECT * FROM horarios_config")
        return {r['aula']: dict(r) for r in cur.fetchall()}

def guardar_horarios(aula, data):
    with get_db() as conn:
        conn.execute("""INSERT OR REPLACE INTO horarios_config 
            (aula, ent_m, sal_m, ent_t, sal_t) VALUES (?, ?, ?, ?, ?)""",
            (aula, data['ent_m'], data['sal_m'], data['ent_t'], data['sal_t']))
        conn.commit()

# --- ASISTENCIA ---
def obtener_asistencia(fecha):
    with get_db() as conn:
        cur = conn.execute("SELECT * FROM asistencia_diaria WHERE fecha=?", (fecha,))
        return {r['orden']: r['estado'] for r in cur.fetchall()}

def actualizar_asistencia(fecha, orden, estado):
    with get_db() as conn:
        conn.execute("""INSERT OR REPLACE INTO asistencia_diaria (fecha, orden, estado) VALUES (?, ?, ?)""",
                     (fecha, orden, estado))
        conn.commit()

# --- PLAN DE LLAMADA ---
def obtener_contacto(orden):
    with get_db() as conn:
        cur = conn.execute("SELECT * FROM plan_llamada WHERE orden=?", (orden,))
        row = cur.fetchone()
        return dict(row) if row else None

def guardar_contacto(data):
    with get_db() as conn:
        conn.execute("""INSERT OR REPLACE INTO plan_llamada 
            (orden, domicilio, telefono_personal, telefono_emergencia, 
             nombre_emergencia, parentesco_emergencia, observaciones)
            VALUES (?, ?, ?, ?, ?, ?, ?)""",
            (data['orden'], data.get('domicilio',''), data.get('telefono_personal',''),
             data.get('telefono_emergencia',''), data.get('nombre_emergencia',''),
             data.get('parentesco_emergencia',''), data.get('observaciones','')))
        conn.commit()

def obtener_todos_contactos():
    with get_db() as conn:
        cur = conn.execute("SELECT * FROM plan_llamada ORDER BY orden")
        return [dict(r) for r in cur.fetchall()]
        # ==============================================================================
# 🆕 FUNCIONES PARA HORARIOS POR DÍA DE LA SEMANA
# ==============================================================================

def init_horarios_dia():
    """Crea la tabla de horarios semanales si no existe"""
    conn = sqlite3.connect("parte_diario.db")
    conn.execute("""
        CREATE TABLE IF NOT EXISTS horarios_dia (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            aula TEXT NOT NULL,
            dia TEXT NOT NULL,
            ent_m TEXT DEFAULT '06:00',
            sal_m TEXT DEFAULT '12:00',
            ent_t TEXT DEFAULT '13:00',
            sal_t TEXT DEFAULT '19:00',
            tipo_ingreso TEXT DEFAULT 'Normal',
            UNIQUE(aula, dia)
        )
    """)
    conn.commit()
    conn.close()

def guardar_horarios_dia(aula, dia, horarios):
    """
    Guarda los horarios de un aula para un día específico.
    dia: 'lunes', 'martes', 'miercoles', 'jueves', 'viernes'
    horarios: dict con ent_m, sal_m, ent_t, sal_t, tipo_ingreso
    """
    conn = sqlite3.connect("parte_diario.db")
    conn.execute("""
        INSERT OR REPLACE INTO horarios_dia 
        (aula, dia, ent_m, sal_m, ent_t, sal_t, tipo_ingreso)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (
        aula, dia,
        horarios.get('ent_m', '06:00'),
        horarios.get('sal_m', '12:00'),
        horarios.get('ent_t', '13:00'),
        horarios.get('sal_t', '19:00'),
        horarios.get('tipo_ingreso', 'Normal')
    ))
    conn.commit()
    conn.close()

def obtener_horarios_dia(aula, dia):
    """Obtiene los horarios de un aula para un día específico"""
    conn = sqlite3.connect("parte_diario.db")
    cursor = conn.execute("""
        SELECT ent_m, sal_m, ent_t, sal_t, tipo_ingreso 
        FROM horarios_dia WHERE aula=? AND dia=?
    """, (aula, dia))
    row = cursor.fetchone()
    conn.close()
    if row:
        return {
            "ent_m": row[0], "sal_m": row[1], 
            "ent_t": row[2], "sal_t": row[3],
            "tipo_ingreso": row[4]
        }
    return {"ent_m": "06:00", "sal_m": "12:00", "ent_t": "13:00", "sal_t": "19:00", "tipo_ingreso": "Normal"}

def obtener_todos_horarios_dia(dia):
    """Obtiene los horarios de TODAS las aulas para un día específico"""
    conn = sqlite3.connect("parte_diario.db")
    cursor = conn.execute("""
        SELECT aula, ent_m, sal_m, ent_t, sal_t, tipo_ingreso 
        FROM horarios_dia WHERE dia=?
    """, (dia,))
    rows = cursor.fetchall()
    conn.close()
    resultado = {}
    for row in rows:
        resultado[row[0]] = {
            "ent_m": row[1], "sal_m": row[2],
            "ent_t": row[3], "sal_t": row[4],
            "tipo_ingreso": row[5]
        }
    return resultado
