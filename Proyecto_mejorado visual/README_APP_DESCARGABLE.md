# Escuadrón H - versión descargable local

Esta versión permite ejecutar la app en una PC Windows con doble clic usando:

```text
EJECUTAR_APP_WINDOWS.bat
```

La app se abrirá en el navegador en:

```text
http://localhost:8501
```

## Requisitos

1. Tener instalado Python 3.11 o superior.
2. Tener internet si se usará Supabase/PostgreSQL.
3. Tener configurado `.streamlit/secrets.toml` con `DATABASE_URL`.

## Configuración inicial

Dentro de la carpeta `.streamlit`, copie:

```text
secrets.toml.plantilla
```

con el nombre:

```text
secrets.toml
```

Luego edite `secrets.toml` y coloque su `DATABASE_URL` real de Supabase/PostgreSQL.

Ejemplo:

```toml
DATABASE_URL = "postgresql://USUARIO:PASSWORD@HOST:6543/postgres?sslmode=require"
OWNER_ADMIN_ENABLED = "false"
OWNER_ADMIN_USER = "admin"
OWNER_ADMIN_PASSWORD = "CAMBIAR_ESTA_CLAVE"
OWNER_ADMIN_SOLO = "true"
```

## Cómo ejecutar

Haga doble clic en:

```text
EJECUTAR_APP_WINDOWS.bat
```

La primera vez puede demorar porque crea el entorno virtual e instala dependencias. Las siguientes veces inicia más rápido.

## Sobre rendimiento

- La interfaz corre en la PC local.
- Los datos siguen guardándose en Supabase/PostgreSQL si `DATABASE_URL` está configurado.
- Si varias PCs usan la misma `DATABASE_URL`, todas comparten la misma base de datos.

## Seguridad

No distribuya la carpeta con un `secrets.toml` que contenga credenciales reales a personas no autorizadas. En una app local, quien tenga acceso a la carpeta puede leer ese archivo.

Para usuarios no confiables, es más seguro mantener la app publicada en Streamlit Cloud y no entregar credenciales de base de datos.
