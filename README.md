# Estudio Contable

Aplicacion Streamlit y herramientas del estudio contable (extractos, conciliacion, IVA, sueldos, auditoria de prestamos, etc.).

## Requisitos

- Python 3.10+
- Dependencias: ver `requirements.txt`

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

## Contenido del repo

- Codigo fuente (`.py`), scripts de arranque, tests
- `requirements.txt`, `.streamlit/config.toml`
- Reglas de Cursor en `.cursor/rules/`
- Plantillas en `plantillas/` (si aplica)

**No** se suben datos de clientes, Excel/PDF de trabajo, bases `.db`, ni secrets (`.env`, `.streamlit/secrets.toml`).

## Streamlit Cloud (login + cifrado)

1. En Streamlit Cloud dejá la app **Public** (el link lo pueden abrir los compañeros).
2. La app igual exige **usuario + PIN** antes de usarla.
3. Pegá el bloque de Secrets (ver `.streamlit/secrets.toml.example`):
   - `DATA_ENCRYPTION_KEY` (Fernet) para cifrar planes/balances subidos
   - `[oficina_usuarios.*]` con PIN de cada persona
4. Tras cambiar Secrets, **Reboot** / redeploy y refrescá el navegador.
5. Un **admin** puede crear más usuarios desde Menú → Usuarios de la oficina.

Cifrado: protege archivos en disco del contenedor (`data/secure/<usuario>/`). No reemplaza el login ni HTTPS.

## Facturación ARCA (solapa ARCA → Facturación)

Emite comprobantes A/B/C, notas de crédito y débito por WSFEv1 (el paquete `arca/` y `emitir.py`). El certificado es del estudio; cada cliente delega Facturación Electrónica a la CUIT 23-42284343-4. El instructivo de delegaciones y puntos de venta está en `docs/PASOS_ARCA.md`.

### Secretos en Streamlit Cloud

En **Manage app → Settings → Secrets**, pegá el bloque de `.streamlit/secrets.toml.example`. Para facturar hacen falta:

```toml
[arca]
cert_homo = """
-----BEGIN CERTIFICATE-----
... PEM de homologación (WSASS) ...
-----END CERTIFICATE-----
"""
cert_prod = """
-----BEGIN CERTIFICATE-----
... PEM de producción ...
-----END CERTIFICATE-----
"""
key = """
-----BEGIN PRIVATE KEY-----
... clave privada del estudio, la misma para los dos ambientes ...
-----END PRIVATE KEY-----
"""
```

Incluí las líneas `BEGIN` / `END`. No subas esos PEM al repo (`.gitignore` ignora `*.pem`, `*.key` y `*.crt`).

Después de guardar los Secrets, **Reboot** de la app.

### Que no se pierdan los CAE

Los números emitidos, el ticket WSAA (vale unas 12 horas; no hay que pedir otro antes) y los datos del emisor para el PDF se guardan con la misma base que el resto del estudio: `database.obtener_conexion()`.

- Si en Secrets están `TURSO_DATABASE_URL` y `TURSO_AUTH_TOKEN`, eso es Turso (SQLite remoto, el plan gratuito alcanza). Sobrevive a un redeploy de Streamlit Cloud.
- Si no están, se usa `estudio_contable.db` en el disco del contenedor. En Cloud ese disco es efímero: sin Turso, un redeploy borra los CAE y el ticket. Conviene cargar esas dos claves (la app ya las usa para clientes y sueldos).

### Uso

1. Elegí Homologación o Producción. Producción pide la casilla y escribir `EMITIR`.
2. Cargá un comprobante a mano o subí el Excel (el modelo se descarga en la misma pantalla).
3. **Validar y armar borrador** muestra la tabla y el PDF sin enviar nada.
4. Recién **Confirmar emisión** pide el CAE. Después se descargan el PDF (zip si son varios) y el Excel con la hoja Resultado.

Filas marcadas `EJEMPLO` no salen a producción. Un comprobante ya aprobado con el mismo contenido no se reenvía. Si se corta la conexión, se consulta en ARCA antes de darlo por perdido. La condición de IVA del receptor es obligatoria.

Pruebas sin certificado ni red: `python -m pytest -q tests/test_offline.py tests/test_arca_servicio.py`

## Rutinas (solapa Rutinas)

La oficina pide a mano tareas que el asistente corre en la PC del estudio (ahí están los archivos y ARCA). La web **no** las ejecuta: apretar **Ejecutar** deja un pedido en la tabla `rutina_pedidos`. No hay un segundo paso de aprobación.

El catálogo (nombre y descripción) está en `rutinas.py`, en la lista `RUTINAS`. Cada pedido puede llevar un texto libre de parámetros (período, cliente) y el nombre de quien lo pide.

Códigos que viajan en el JSON (`rutina`):

- `seguimiento_balances_urgencia` — Seguimiento balances urgencia
- `control_fcc_portal_iva` — Control FCC Portal IVA
- `fcc_monotributistas` — FCC monotributistas
- `aviso_bazan_bajar_archivos` — Aviso Bazan bajar archivos
- `bazan_detalle_items` — Bazan Detalle Items

Estados: `PENDIENTE` → `EN_CURSO` → `OK` o `ERROR`. Un pedido `PENDIENTE` se puede cancelar. Si esa rutina ya tiene uno `PENDIENTE` o `EN_CURSO`, la web avisa y no crea otro.

### Asistente externo

Desde la raíz del repo, con las mismas variables que Streamlit Cloud (`TURSO_DATABASE_URL` y `TURSO_AUTH_TOKEN`). Sin esas variables lee el SQLite local y no ve la cola de la web.

```bash
python scripts/cola_rutinas.py listar --estado PENDIENTE
python scripts/cola_rutinas.py tomar 12
python scripts/cola_rutinas.py terminar 12 --estado OK --resultado "Listo" --archivos "C:\ruta\salida.xlsx"
```

`listar` imprime un array JSON. `tomar` pasa a `EN_CURSO` solo si seguía `PENDIENTE`. `terminar` cierra un `EN_CURSO` en `OK` o `ERROR` (`--resultado` es obligatorio; `--archivos` es opcional, rutas o links). Si la operación no corresponde, sale con código 1 y un JSON `{"ok": false, "error": "..."}`.

Pruebas de la cola, sin Turso: `python -m pytest -q tests/test_rutinas.py`
