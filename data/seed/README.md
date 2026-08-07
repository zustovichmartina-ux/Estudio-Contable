# Seed Cloud — sociedades PJ

`sociedades_pj.json` siembra las Personas Jurídicas al iniciar la app cuando la BD SQLite está vacía (p. ej. Streamlit Cloud).

Los planes propios viven en `data/planes_cuentas/plan_{CUIT}.xlsx`.
Si una sociedad no tiene plan propio, se usa `plan_default.xlsx` hasta que se suba uno desde la UI.
