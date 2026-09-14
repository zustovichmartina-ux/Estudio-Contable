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

## Túnel Cloudflare estable (recomendado)

Sin config extra el worker abre un **túnel rápido** (`*.trycloudflare.com`): la URL cambia en cada restart y hay que actualizar Secrets. Para una URL fija, crear **una vez** un túnel con nombre:

1. Instalar: `winget install --id Cloudflare.cloudflared`
2. Login (elige el dominio de Cloudflare): `cloudflared tunnel login`
3. Crear: `cloudflared tunnel create afip-worker`  
   Copia el JSON de credenciales a `jobs/cloudflared/<UUID>.json` (gitignored). En Windows, si no lo encuentra, usá ruta absoluta en `credentials-file`.
4. Hostname público, una de estas:
   - DNS: `cloudflared tunnel route dns afip-worker afip-worker.TU-DOMINIO.com`
   - o en Cloudflare Zero Trust → Networks → Tunnels → Public hostname → origen `http://127.0.0.1:8765`
5. Copiar `jobs/cloudflared/config.yml.example` → `jobs/cloudflared/config.yml` y completar UUID + hostname.  
   Alternativa: `AFIP_CLOUDFLARED_CONFIG` (ruta al YAML) y/o `AFIP_TUNNEL_HOSTNAME` (URL pública).  
   Token de Zero Trust: `AFIP_CLOUDFLARED_TOKEN` + `AFIP_TUNNEL_HOSTNAME` (sin YAML).
6. Arrancar `iniciar_afip_worker.bat`. Pegar **una vez** URL + token de `jobs/cloud_bridge.txt` en Streamlit Secrets. Al reiniciar el worker la URL no cambia.

No subir `jobs/cloudflared/*.json` ni `config.yml` al repo. El token de la API (`AFIP_WORKER_TOKEN`) sigue siendo obligatorio.

## Auth (regla dura)

- Registry `jobs/cuit_registry.json`: `ready | needs_admin | unknown | failed`.
- Nunca claves en Excel ni Streamlit.
- CUIT nuevo → `needs_admin` / job `needs_auth` + handoff admin.
- Claves solo Chrome autofill en la PC del worker.
- Arranque: `iniciar_afip_worker.bat` (`--live`: Chrome/AFIP para `bajar_comprobantes`).

## Orden

1. Dry-run del loop ✅
2. UI ARCA (encolar + cola + registry) ✅
3. Worker autónomo + API/túnel (`iniciar_afip_worker.bat`) ✅
4. `bajar_comprobantes` Playwright ✅ (Comprobantes en Línea → PDFs)
5. `bajar_veps`
6. `emitir_fcc`
7. Badge auth + handoff 2FA (registry) ✅ base
