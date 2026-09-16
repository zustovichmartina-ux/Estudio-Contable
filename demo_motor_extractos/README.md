# Demo motor extractos (Galicia)

Demo de ayer: buzón PDF con OCR, parseo por coordenada Y (importe + saldo) y grilla imputable.

- Generales (IVA, IIBB, Ley 25.413, comisiones, VEP, Visa): **cuenta fija**, no se tocan.
- Transferencias: se cruzan con **deudores / proveedores**.
- Asiento Tango y papeles de trabajo: ganchos; esperan plantillas vacías.

## Cómo abrirlo

Desde esta carpeta:

```
python server.py
```

Queda en http://127.0.0.1:8765/

- `index.html` — buzón OCR (Abril 2026 digital / Junio 2025 escaneado).
- `imputacion.html` — misma grilla del banco, con imputación.

Los PDF de muestra salen de la carpeta Galicia de Sunny Beach en el servidor. Si no está montada, subí un PDF en el buzón.
