# Loop AFIP / ARCA → App estudio (Cursor)

Objetivo: la web Streamlit (`estudiocontablemdp.streamlit.app`) solo **crea y muestra trabajos**. Un worker local (Cursor / PC RECEPCION) **ejecuta** AFIP, descarga/emite y guarda en rutas UNC. Nunca guardar claves en Excel ni en la web.

## Barra top-level

`Devengamiento · Conciliación · Préstamos · Herramientas · ARCA`

- **ARCA** es módulo top-level (no dentro de Herramientas).
- Monotributo salió de la barra.

## Principio

```
Usuario en ARCA (app)
  → Job (JSON) en cola
    → Worker Cursor en RECEPCION
      → Chrome AFIP (autofill / 2FA humano una vez)
        → Archivos en \\TANGOSRV\...
          → Job = done | needs_auth | error
```

## Auth (regla dura)

- Registry `jobs/cuit_registry.json`: `ready | needs_admin | unknown | failed`.
- Nunca claves en Excel ni Streamlit.
- CUIT nuevo → `needs_admin` / job `needs_auth` + handoff admin.
- Claves solo Chrome autofill en la PC del worker.

## Orden

1. Dry-run del loop ✅
2. UI ARCA (encolar + cola + registry) ✅
3. `bajar_comprobantes` Playwright
4. `bajar_veps`
5. `emitir_fcc`
6. Badge auth + handoff 2FA (registry) ✅ base
