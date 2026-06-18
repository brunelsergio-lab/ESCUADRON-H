"""Lanzador local de la app Escuadrón H.

Uso:
    python EJECUTAR_APP_LOCAL.py
"""
import os
import subprocess
import sys
import time
import webbrowser
from pathlib import Path

BASE = Path(__file__).resolve().parent
APP = BASE / "plataforma.py"

if not APP.exists():
    print("No se encontro plataforma.py")
    raise SystemExit(1)

cmd = [
    sys.executable, "-m", "streamlit", "run", str(APP),
    "--server.address", "localhost",
    "--server.port", "8501",
    "--browser.gatherUsageStats", "false",
]

print("Iniciando Escuadron H en http://localhost:8501 ...")
proc = subprocess.Popen(cmd, cwd=str(BASE))
time.sleep(2)
webbrowser.open("http://localhost:8501")
try:
    proc.wait()
except KeyboardInterrupt:
    proc.terminate()
