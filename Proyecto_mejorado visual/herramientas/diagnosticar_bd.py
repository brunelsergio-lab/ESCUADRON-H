"""Diagnóstico rápido de la base SQLite local.

Uso:
    python herramientas/diagnosticar_bd.py
"""
from pathlib import Path
import sqlite3

BASE = Path(__file__).resolve().parents[1]
DB = BASE / "parte_diario.db"

if not DB.exists():
    print(f"No existe la base local: {DB}")
    raise SystemExit(1)

conn = sqlite3.connect(DB)
print(f"Base local: {DB}")
print("\nTablas encontradas:")
for (tabla,) in conn.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name"):
    try:
        total = conn.execute(f'SELECT COUNT(*) FROM "{tabla}"').fetchone()[0]
    except Exception as exc:
        total = f"error: {exc}"
    print(f"- {tabla}: {total}")

print("\nDatos agrupados por fecha:")
for tabla, campo in [
    ("asistencia_diaria", "fecha"),
    ("almuerzo", "fecha"),
    ("estado_aulas", "fecha"),
    ("almuerzo_historial", "fecha"),
    ("movimientos", "fecha_parte"),
]:
    print(f"\n{tabla}:")
    try:
        rows = list(conn.execute(f'SELECT "{campo}", COUNT(*) FROM "{tabla}" GROUP BY "{campo}" ORDER BY "{campo}"'))
        if not rows:
            print("  sin datos")
        for fecha, total in rows:
            print(f"  {fecha}: {total}")
    except Exception as exc:
        print(f"  no disponible: {exc}")

print("\nUsuarios:")
try:
    rows = list(conn.execute("SELECT usuario, rol, activo, creado_en FROM usuarios ORDER BY usuario"))
    if not rows:
        print("  sin usuarios")
    for usuario, rol, activo, creado_en in rows:
        estado = "activo" if int(activo) == 1 else "inactivo"
        print(f"  {usuario} | {rol} | {estado} | {creado_en}")
except Exception as exc:
    print(f"  no disponible: {exc}")

conn.close()
