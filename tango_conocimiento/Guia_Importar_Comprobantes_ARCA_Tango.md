# Importar comprobantes de ARCA a Tango (proceso)

Esto es un **proceso** del Liquidador de IVA, no una fórmula de sueldos.

Menú en Tango: **Liquidador de IVA → Comprobantes → Importación de comprobantes desde ARCA → Importación de comprobantes desde ARCA - Portal IVA**.

Modelos que usa el importador (no los cambies salvo cuentas contables): **PIVAC** (compras) y **PIVAV** (ventas).

## 1. Bajar el CSV desde ARCA (Portal IVA)

1. Entrá a **Portal IVA** con clave fiscal (ARCA).
2. **Nueva declaración jurada** → Ingresar.
3. Elegí el período → Continuar.
4. **Libro de IVA y declaración jurada de IVA**.
5. Elegí **Compras** o **Ventas**.
6. Botón **Importar** → **Importar comprobantes desde ARCA**.
7. Icono **CSV** → descargá el ZIP.
8. **Descomprimí** el ZIP. Tango no acepta el ZIP: hay que importar el **.CSV**.

No edites la estructura del CSV (tiene que quedar como lo bajó ARCA).

## 2. Importar el CSV en Tango

1. Abrí **Liquidador de IVA → Comprobantes → Importación de comprobantes desde ARCA**.
2. Origen: **Portal IVA – Libro IVA compras** o **Portal IVA – Libro IVA ventas**.
3. **Archivo externo**: examiná y elegí el `.csv`.
4. Dejá activa la conexión con el web service de ARCA si querés que complete clientes/proveedores.
5. Completá las **equivalencias** de tipos de comprobante ARCA ↔ tipos de Tango (la primera vez, o si aparece un tipo nuevo).
6. Si hay duplicados, elegí: solo nuevos, o nuevos + actualizar existentes.
7. Confirmá el resumen (cantidad de comprobantes, clientes/proveedores, tipos).
8. Al terminar, Tango arma un **Excel** con cada comprobante: si entró o por qué se rechazó.

## 3. Controlar lo importado

- **Registración de comprobantes**: buscá el número y, si hace falta, editá.
- Libro IVA Compras / Ventas.
- Live: Subdiario de Compras / Ventas.

El archivo de ARCA **no abre percepciones e impuestos**. Esas líneas se reimputan a mano en la registración.

## 4. Asientos (no se generan solos)

La importación **no** genera el asiento contable. Después corré **Liquidador de IVA → Procesos periódicos → Contabilización → Generación de asientos contables**, filtrando modelos **PIVAC** / **PIVAV**.

Antes, en esos modelos, tienen que estar las **cuentas contables** de cada fórmula.

## Primera vez (parametrización)

En **Parámetros de Liquidador de IVA**:

- Solapa **Parámetros Portal IVA** (compras y ventas): ocasionales, jurisdicción IIBB, actividad, fórmulas de percepciones.
- Solapa **RG 3711**: relaciones de fórmulas para Libro IVA Digital.

No toques las fórmulas de PIVAC/PIVAV; solo las cuentas.

## Si un comprobante no entra

Suele ser: falta equivalencia de tipo, falta cliente/proveedor, CUIT duplicado, falta fórmula para un importe del CSV, falta jurisdicción IIBB, o la fecha está fuera del rango de Procesos generales.
