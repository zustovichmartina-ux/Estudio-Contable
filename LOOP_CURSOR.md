# Loop AFIP / ARCA → App estudio (Cursor)

Objetivo: la web Streamlit (`estudiocontablemdp.streamlit.app`) solo **crea y muestra trabajos**. Un worker local (Cursor / PC RECEPCION) **ejecuta** AFIP, descarga/emite y guarda en rutas UNC. Nunca guardar claves en Excel ni en la web.

## Barra top-level

`Devengamiento · Conciliación · Préstamos · Herramientas · ARCA`

- **ARCA** es módulo top-level (no dentro de Herramientas).
- Monotributo salió de la barra.

## Principio

```
Usuario en ARCA (app en la nube)
  → HTTPS (túnel Cloudflare) + token
    → API en RECEPCION (`iniciar_afip_worker.bat --serve`)
      → Job JSON en jobs/pending
        → Si CUIT Listo: ejecuta solo
        → Si Pedir acceso: needs_auth + 2FA
          → Archivos en \\TANGOSRV\...

```

El worker **pollea solo** `jobs/pending` cada ~3 s. La web en Streamlit Cloud **no escribe en tu disco**: manda el job por el túnel.

Secrets Cloud: `AFIP_WORKER_URL` + `AFIP_WORKER_TOKEN` (ver `jobs/cloud_bridge.txt` al arrancar).

## Auth (regla dura)

- Registry `jobs/cuit_registry.json`: `ready | needs_admin | unknown | failed`.
- Nunca claves en Excel ni Streamlit.
- CUIT nuevo → `needs_admin` / job `needs_auth` + handoff admin.
- Claves solo Chrome autofill en la PC del worker.
- Arranque: `iniciar_afip_worker.bat` (dry-run hasta Playwright live).

## Orden

1. Dry-run del loop ✅
2. UI ARCA (encolar + cola + registry) ✅
3. Worker autónomo + API/túnel (`iniciar_afip_worker.bat`) ✅
4. `bajar_comprobantes` Playwright
5. `bajar_veps`
6. `emitir_fcc`
7. Badge auth + handoff 2FA (registry) ✅ base
