"""Resetea o crea un administrador en la base SQLite local.

Uso desde la carpeta del proyecto:
    python herramientas/reset_admin_local.py miusuario miclave123

Opcional, para dejar solo ese usuario como admin:
    python herramientas/reset_admin_local.py miusuario miclave123 --solo-admin
"""
from pathlib import Path
import sys

BASE = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE))

if len(sys.argv) < 3:
    print("Uso: python herramientas/reset_admin_local.py USUARIO CONTRASEÑA [--solo-admin]")
    raise SystemExit(1)

usuario = sys.argv[1].strip()
password = sys.argv[2].strip()
solo_admin = "--solo-admin" in sys.argv

if len(password) < 6:
    print("La contraseña debe tener al menos 6 caracteres.")
    raise SystemExit(1)

import db_manager as db

db.init_db()
user = db.crear_o_actualizar_usuario_admin(usuario, password, solo_admin=solo_admin)
print(f"Administrador actualizado: {user['usuario']} | rol={user['rol']} | activo={user['activo']}")
if solo_admin:
    print("Los demás administradores fueron cambiados a rol usuario.")
