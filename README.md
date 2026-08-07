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
