"""Migra datos desde parte_diario.db local hacia PostgreSQL.

Antes de usar:
1) Configura la variable de entorno DATABASE_URL con tu URL real.
2) Ejecuta desde la carpeta del proyecto:

   python herramientas/migrar_sqlite_a_postgres.py

El script crea las tablas usando db_manager.init_db() y copia registros sin duplicar claves primarias.
"""
from pathlib import Path
import os
import sqlite3
import sys

BASE = Path(__file__).resolve().parents[1]
DB = BASE / "parte_diario.db"
sys.path.insert(0, str(BASE))

DATABASE_URL = os.getenv("DATABASE_URL")
if not DATABASE_URL:
    print("Falta DATABASE_URL. Configura la variable de entorno con tu base PostgreSQL.")
    raise SystemExit(1)

if not DB.exists():
    print(f"No existe la base SQLite local: {DB}")
    raise SystemExit(1)

import db_manager as db  # noqa: E402

if not db.is_postgres():
    print("db_manager no detectó PostgreSQL. Revisa que DATABASE_URL empiece con postgresql:// o postgres://")
    raise SystemExit(1)

db.init_db()

TABLAS = {
    "novedades": ["id", "orden", "grado", "nombre", "dni", "ce", "aula", "estado", "detalle", "fecha_ini", "fecha_fin", "ambito"],
    "estado_aulas": ["fecha", "aula", "estado_m", "estado_t", "salida_m", "salida_t", "ubicacion_m", "ubicacion_t"],
    "almuerzo": ["fecha", "orden"],
    "almuerzo_historial": ["id", "fecha_hora", "fecha", "orden", "nombre", "aula"],
    "horarios_config": ["aula", "ent_m", "sal_m", "ent_t", "sal_t"],
    "horarios_diarios": ["aula", "dia", "ent_m", "sal_m", "ent_t", "sal_t"],
    "asistencia_diaria": ["fecha", "orden", "estado"],
    "plan_llamada": ["orden", "domicilio", "telefono_personal", "telefono_emergencia", "nombre_emergencia", "parentesco_emergencia", "observaciones"],
    "movimientos": ["id", "fecha_hora", "fecha_parte", "modulo", "accion", "orden", "nombre", "aula", "detalle"],
    "usuarios": ["id", "usuario", "password_hash", "salt", "rol", "activo", "creado_en"],
    "app_config": ["clave", "valor"],
}

sqlite_conn = sqlite3.connect(DB)
sqlite_conn.row_factory = sqlite3.Row

copiados_total = 0
with db.get_db() as pg_conn:
    for tabla, columnas in TABLAS.items():
        existe = sqlite_conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name=?", (tabla,)
        ).fetchone()
        if not existe:
            print(f"{tabla}: no existe en SQLite, se omite")
            continue

        reales = {r[1] for r in sqlite_conn.execute(f'PRAGMA table_info("{tabla}")').fetchall()}
        cols = [c for c in columnas if c in reales]
        if not cols:
            print(f"{tabla}: sin columnas compatibles")
            continue

        rows = sqlite_conn.execute(f'SELECT {", ".join(cols)} FROM "{tabla}"').fetchall()
        if not rows:
            print(f"{tabla}: sin datos")
            continue

        placeholders = ", ".join(["?"] * len(cols))
        col_sql = ", ".join(cols)
        conflict = "ON CONFLICT DO NOTHING"
        sql = f'INSERT INTO {tabla} ({col_sql}) VALUES ({placeholders}) {conflict}'

        copiados = 0
        for row in rows:
            db.run(pg_conn, sql, tuple(row[c] for c in cols))
            copiados += 1
        pg_conn.commit()
        copiados_total += copiados
        print(f"{tabla}: {copiados} registro(s) procesado(s)")

sqlite_conn.close()
print(f"\nMigración terminada. Registros procesados: {copiados_total}")
