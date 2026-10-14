Scripts de un solo uso / diagnostico (no importados por app.py ni procesador.py). Ejecutar desde la raiz del repo: python scripts/<nombre>.py.

`cola_rutinas.py` es la excepcion: lo usa el asistente externo para leer la cola de la solapa Rutinas (misma base que la web, Turso si hay TURSO_DATABASE_URL y TURSO_AUTH_TOKEN). Ver README.

`generar_instructivo_deducciones_ganancias.py` regenera `plantillas/Instructivo_Deducciones_Ganancias_1hoja.pdf` (checklist documentación, sin montos).
