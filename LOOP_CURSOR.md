# Loop AFIP → App estudio (Cursor)

Objetivo: la web Streamlit (`estudiocontablemdp.streamlit.app`) solo **crea y muestra trabajos**. Un worker local (Cursor / PC RECEPCION) **ejecuta** AFIP, descarga/emite y guarda en rutas UNC. Nunca guardar claves en Excel ni en la web.

## Principio

```
Usuario en la app
  → Job (JSON) en cola
    → Worker Cursor en RECEPCION
      → Chrome AFIP (autofill / 2FA humano una vez)
        → Archivos en \\TANGOSRV\...
          → Job = done | needs_auth | error
```

## Modelo de Job

Ruta: `jobs/pending/{job_id}.json` → al terminar `jobs/done/` o `jobs/error/`.

Campos: `id`, `created_at`, `requested_by`, `cuit`, `razon_social`, `action`
(`emitir_fcc` | `bajar_veps` | `bajar_comprobantes`), `params`, `auth`, `status`, `result`.

## Auth (regla dura)

- Nunca claves en Excel ni en Streamlit secrets de la plantilla.
- Preferir Chrome autofill ya guardado en la PC del worker.
- Si sesión caída / CUIT nuevo / 2FA → `needs_auth` y handoff humano.

## Orden de construcción

1. Dry-run del loop (sin AFIP) ✅
2. UI de encolar ✅
3. `bajar_comprobantes` (Playwright local)
4. `bajar_veps`
5. `emitir_fcc`
6. Badge de auth + handoff 2FA

## Fuera de alcance v1

Sueldos, proyección Ganancias, Visa, claves, multi-usuario concurrente sobre el mismo CUIT.
