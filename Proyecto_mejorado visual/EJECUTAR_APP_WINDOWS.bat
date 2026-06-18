@echo off
setlocal
cd /d "%~dp0"

echo ================================================
echo   ESCUADRON H - APP LOCAL DESCARGABLE
echo ================================================
echo.

where py >nul 2>nul
if %errorlevel%==0 (
    set PY_CMD=py -3
) else (
    where python >nul 2>nul
    if %errorlevel%==0 (
        set PY_CMD=python
    ) else (
        echo ERROR: No se encontro Python instalado.
        echo Instale Python 3.11 o superior desde https://www.python.org/downloads/
        pause
        exit /b 1
    )
)

if not exist ".venv\Scripts\activate.bat" (
    echo Creando entorno virtual...
    %PY_CMD% -m venv .venv
)

call ".venv\Scripts\activate.bat"

echo Instalando/actualizando dependencias...
python -m pip install --upgrade pip
pip install -r requirements.txt

if not exist ".streamlit\secrets.toml" (
    echo.
    echo ADVERTENCIA: No existe .streamlit\secrets.toml
    echo Copie .streamlit\secrets.toml.plantilla como .streamlit\secrets.toml
    echo y coloque su DATABASE_URL real antes de usar la app en produccion.
    echo.
)

echo Iniciando aplicacion...
echo Cuando abra el navegador, use la app normalmente.
echo Para cerrar, vuelva a esta ventana y presione CTRL + C.
echo.
streamlit run plataforma.py --server.address localhost --server.port 8501 --browser.gatherUsageStats false

pause
