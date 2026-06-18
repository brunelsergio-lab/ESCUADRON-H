# Instalar Escuadrón H en el teléfono

Esta versión está preparada para verse mejor en celular y para agregarse a la pantalla principal como app web.

## Android / Chrome

1. Abrir la URL publicada de la app en Chrome.
2. Tocar los tres puntos del navegador.
3. Elegir **Agregar a pantalla principal** o **Instalar app**.
4. Confirmar el nombre **Escuadrón H**.
5. Abrirla desde el ícono del teléfono.

## iPhone / Safari

1. Abrir la URL publicada de la app en Safari.
2. Tocar el botón **Compartir**.
3. Elegir **Agregar a inicio**.
4. Confirmar el nombre **Escuadrón H**.
5. Abrirla desde el ícono del iPhone.

## Notas importantes

- La app sigue corriendo en Streamlit Cloud.
- Los datos siguen guardándose en Supabase/PostgreSQL.
- No hay que poner `DATABASE_URL` en el teléfono.
- Esta opción no expone credenciales de base de datos a los usuarios.
- En Streamlit, la instalación móvil puede comportarse como acceso web avanzado; el soporte PWA completo depende del navegador y de la plataforma.

## Mejora incluida

- Manifest básico para instalación móvil.
- Íconos 192x192 y 512x512.
- Estilos CSS responsive para pantallas chicas.
- Configuración de archivos estáticos de Streamlit.
