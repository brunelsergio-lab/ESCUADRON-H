# Cambios visuales aplicados

Esta versión mejora la interfaz de la plataforma Streamlit del Escuadrón H sin cambiar la lógica principal.

## Archivos agregados o modificados

- `plataforma.py`: se mejoró la carga de tema, login, sidebar, encabezado principal, pestañas y acciones rápidas.
- `assets/styles.css`: se centralizó el diseño visual para evitar CSS repetido dentro del archivo principal.
- `.streamlit/config.toml`: se agregó configuración de tema oscuro institucional.
- `live_search_component/index.html`: se ajustó el estilo del buscador para que combine con el tema.

## Mejoras visibles

- Login más profesional con logo.
- Sidebar con identidad visual.
- Dashboard superior con logo institucional y tarjetas KPI.
- Botones más modernos.
- Pestañas con iconos.
- Acciones de sincronización/reinicio ocultas en un panel de mantenimiento.
- Mejor diseño responsive para celular.

## Cómo ejecutar

```bash
pip install -r requirements.txt
streamlit run plataforma.py
```
