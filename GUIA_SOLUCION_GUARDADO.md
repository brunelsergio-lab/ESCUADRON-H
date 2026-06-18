# Solución para que la app no quede vacía al día siguiente

## Qué estaba pasando

Se encontraron dos problemas principales:

1. **La app usa `parte_diario.db` como SQLite local.**
   - En una PC local puede servir.
   - En Streamlit Cloud, Codespaces, Render sin disco persistente u otro servidor temporal, ese archivo puede perderse al reiniciar, reconstruir o redeployar la app.
   - Por eso pueden desaparecer usuarios, administradores, novedades, racionamiento y asistencia.

2. **Varias pantallas guardaban datos con la fecha real del día (`FECHA_STR`) y no con la `Fecha del Reporte` (`FECHA_PARTE_STR`).**
   - Si ayer cargabas novedades para el día siguiente, parte de la información quedaba guardada con fecha de ayer.
   - Al abrir hoy, la app buscaba datos con la fecha de hoy y parecía vacía.

## Qué se corrigió en esta versión

- `Novedades`, `Asistencia`, `Racionamiento`, `Movimientos` y `Resumen` ahora trabajan con la **Fecha del Reporte**.
- El formulario de novedades ahora propone por defecto la **Fecha del Reporte**, no la fecha actual de la computadora/servidor.
- El botón **Reiniciar asistencia** ya no usa `sqlite3.connect("parte_diario.db")` directamente; ahora usa el administrador de base de datos, por lo que también funciona si usas PostgreSQL/Supabase.
- `db_manager.py` ahora usa una ruta absoluta para SQLite local, evitando que se cree otra base vacía si se ejecuta desde otra carpeta.
- La app muestra un aviso en el panel lateral si está usando SQLite local, porque en despliegues web no es almacenamiento permanente.

## Solución definitiva para producción

Para que los datos no se pierdan nunca, usa una base PostgreSQL externa, por ejemplo Supabase, Neon, Railway, Render PostgreSQL, etc.

Tu código ya está preparado para eso. Solo debes configurar una variable/secret llamada:

```toml
DATABASE_URL = "postgresql://USUARIO:CONTRASEÑA@HOST:PUERTO/BASE?sslmode=require"
```

En Streamlit Cloud se coloca en:

**App → Settings → Secrets**

Después reinicia la app. Si la conexión está bien, en la barra lateral debe aparecer:

```text
✅ Base de datos persistente conectada
```

Si aparece:

```text
⚠️ Base local SQLite...
```

significa que todavía estás usando `parte_diario.db` local y puedes volver a perder datos en un reinicio/redeploy.

## Cómo revisar si los datos de ayer quedaron guardados con otra fecha

En la carpeta del proyecto, puedes ejecutar:

```bash
python - <<'PY'
import sqlite3
conn = sqlite3.connect('parte_diario.db')
for tabla, campo in [
    ('asistencia_diaria', 'fecha'),
    ('almuerzo', 'fecha'),
    ('estado_aulas', 'fecha'),
    ('movimientos', 'fecha_parte'),
]:
    print('\n', tabla)
    try:
        for row in conn.execute(f"SELECT {campo}, COUNT(*) FROM {tabla} GROUP BY {campo} ORDER BY {campo}"):
            print(row)
    except Exception as e:
        print('No se pudo leer:', e)
print('\nUsuarios:')
try:
    for row in conn.execute("SELECT usuario, rol, activo, creado_en FROM usuarios"):
        print(row)
except Exception as e:
    print('No hay tabla usuarios o está vacía:', e)
conn.close()
PY
```

Si ves datos en una fecha equivocada, se pueden recuperar migrándolos a la fecha correcta antes de seguir trabajando.


## Recuperar o cambiar el administrador propietario

Si otra persona creó el primer administrador y necesitas que tu usuario sea el administrador, hay dos formas:

### Opción A: por Secrets en Streamlit/hosting

Agrega estos Secrets:

```toml
OWNER_ADMIN_ENABLED = "true"
OWNER_ADMIN_USER = "tu_usuario"
OWNER_ADMIN_PASSWORD = "tu_contraseña_segura"
OWNER_ADMIN_SOLO = "true"
```

- `OWNER_ADMIN_ENABLED=true` activa la recuperación.
- `OWNER_ADMIN_USER` será creado o actualizado como administrador.
- `OWNER_ADMIN_PASSWORD` será la nueva contraseña.
- `OWNER_ADMIN_SOLO=true` es opcional; cambia a los demás administradores al rol usuario.

Después de entrar correctamente, cambia `OWNER_ADMIN_ENABLED` a `false` o elimina esos Secrets para no seguir reseteando la contraseña en cada reinicio.

### Opción B: si estás trabajando localmente con SQLite

Ejecuta desde la carpeta del proyecto:

```bash
python herramientas/reset_admin_local.py tu_usuario tu_contraseña_segura --solo-admin
```

Luego inicia la app y entra con ese usuario.
