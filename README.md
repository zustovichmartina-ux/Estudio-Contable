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
