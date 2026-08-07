# Roadmap de refactor arquitectonico

Contexto: `app.py` (~12.600 lineas) y `procesador.py` (~14.000 lineas)
concentran casi todo el proyecto. Este documento registra el plan para
seguir dividiendolos por dominio, sin romper la app en producción.

## Hecho (esta sesion)

- Seguridad de login: PIN con hash + salt (PBKDF2-HMAC-SHA256, 600k
  iteraciones) en vez de SHA-256 sin salt. Migra en caliente los hashes
  legacy la primera vez que cada usuario hace login (nadie necesita
  resetear su PIN). Rate limiting: 5 intentos fallidos bloquean el usuario
  15 minutos.
- `requirements.txt` con versiones fijas (antes `>=`).
- Scripts sueltos de un solo uso (`_extraer_*`, `_cruce_*`, `diagnostico_*`,
  etc., ~30 archivos) movidos a `scripts/`. No se tocó nada que
  `app.py`/`procesador.py` importen.
- Extraido `auth_oficina.py`: toda la logica de usuarios de oficina
  (login, PIN, roles, alta/baja) que antes vivia en `database.py`.
  `database.py` quedo con wrappers finos que delegan a este modulo nuevo,
  para no tener que tocar los ~10 call sites de `app.py` en este primer
  paso.

## Como seguir (próximos pasos, en orden sugerido)

1. **app.py → auth_oficina directo**: reemplazar los `db.verificar_login_oficina`,
   `db.listar_usuarios_oficina`, etc. en `app.py` (funciones
   `_pantalla_login_oficina` y `_seccion_usuarios_oficina`, ~línea 12350+)
   por `import auth_oficina` y llamadas directas. Después, borrar los
   wrappers de compatibilidad en `database.py`.

2. **Extraer dominios de `procesador.py`** (14k líneas, importado por 23
   módulos). Candidatos, de más aislado a menos:
   - Parsers de bancos (`PERFILES_BANCO`, `BANK_REGISTRY`, funciones
     `_extraer_*` de PDFs/extractos) → `parsers_bancos.py`
   - Generación de Excel/Tango (`generar_excel_tango_nativo`,
     `generar_txt_tango`, etc.) → `exportacion_tango.py`
   - Lógica de balance/plan de cuentas → `balance.py`
   Cada extracción debe seguir el mismo patrón que `auth_oficina.py`:
   mover el código, dejar wrappers de compatibilidad si hace falta, correr
   `test_*.py` antes/después y comparar.

3. **Dividir `app.py` por pantalla/módulo** (IVA, IIBB, sueldos, bancos,
   préstamos, TISH) en archivos bajo algo como `paginas/`, usando
   `st.session_state` igual que ahora pero con cada pantalla en su propio
   archivo. Esto es lo más riesgoso porque hay mucho estado compartido
   (`_flush_estado_*`, claves de `session_state`) — conviene hacerlo de a
   una pantalla por vez y probar manualmente cada una en Streamlit Cloud
   o local antes de seguir con la próxima.

4. **Tests reales**: convertir los `test_*.py` (actualmente scripts que
   imprimen resultados) a `pytest` con `assert` y fixtures, y agregar un
   workflow de GitHub Actions que los corra en cada push. Esto es lo que
   más falta para poder refactorizar con confianza — sin esto, cada paso
   de división de `procesador.py`/`app.py` depende de revisión manual.

## Por qué no se hizo todo de una

`procesador.py` y `app.py` no tienen tests con asserts reales (son scripts
que imprimen resultados para inspección manual) y manejan datos contables
reales de clientes. Dividir 26.000 líneas de una sola vez, sin una red de
tests confiable, es el tipo de cambio que puede introducir un bug sutil en
el cálculo de un impuesto sin que nadie lo note hasta que ya salió mal.
Ir dominio por dominio, con `git diff` chico y pruebas después de cada
paso, es más lento pero mucho más seguro para una herramienta que ya está
en uso.
