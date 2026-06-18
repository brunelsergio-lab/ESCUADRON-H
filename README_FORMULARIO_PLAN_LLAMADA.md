# Formulario público de Plan de Llamada / Licencia

Esta versión agrega un formulario público protegido por token para que cada aspirante complete sus datos de licencia y contacto sin entrar al panel administrativo.

## Importante

El formulario valida los datos ya cargados en `alumnos.csv` mediante DNI o CE. No modifica la base de personal existente: apellido y nombres, DNI, CE, aula y orden se muestran como datos validados y bloqueados. Solo se agregan o actualizan los campos nuevos del plan de licencia.

Si ya había datos guardados en `plan_llamada`, se conservan. Al guardar, solo se actualizan los campos enviados por el formulario.

## Activar en Streamlit Secrets

Agregar o mantener:

```toml
PLAN_FORM_ENABLED = "true"
PLAN_FORM_TOKEN = "licencia2026"
```

Puede cambiar `licencia2026` por cualquier clave/token que quiera.

## Enlace para compartir

Si la URL de la app es:

```text
https://TU-APP.streamlit.app/
```

El enlace del formulario será:

```text
https://TU-APP.streamlit.app/?form=plan_llamada&token=licencia2026
```

## Seguridad

- El formulario no muestra la app completa.
- Valida por DNI o CE contra `alumnos.csv`.
- Guarda directamente en la tabla `plan_llamada` de Supabase/PostgreSQL.
- No requiere usuario/contraseña de administrador.
- El token evita que el formulario quede abierto al público general.

## Campos nuevos para completar

Cada aspirante completa:

- Si viaja en transporte público: SI/NO.
- Si viaja en vehículo particular: SI/NO.
- Si viaja en vehículo particular: marca, modelo, dominio y titular del vehículo.
- Lugar de licencia.
- Dirección.
- Barrio.
- Calle.
- Número.
- Unidad más próxima de GN.
- Teléfono particular.
- Teléfono de emergencia para licencia.
- Observaciones opcionales.

## Exportación

En la pestaña **Legajos y contactos** el botón **EXPORTAR PLAN DE LLAMADA (EXCEL)** genera un Excel con todos los campos nuevos.

## Vista administrativa

En la pestaña **Legajos y contactos** queda un bloque con el estado del formulario y el final de URL que se debe compartir.
