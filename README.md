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
