# Reglas Tango del Estudio Contable

Fuente: motor de la app (procesador.py) + parametrización del estudio.
No incluye claves SQL ni TANGO.INI.

## Exportación de asientos desde la web

- Template: `Asientos contables VACIO PARA LLENAR E IMPORTAR.xlsx` (hojas Asientos, Renglones, Cotizaciones).
- Fecha en TXT: `AAAAMMDD`. En Excel nativo, fecha de celda.
- Moneda: `PES`. No dejar filas de cotización DOL/USD/U$S: Tango las rechaza.
- Leyenda de renglones: vacía. El concepto va en la cabecera del asiento.
- Debe = Haber. Si hay diferencia de centavos, el motor la balancea (tolerancia chica).

## Tipos de asiento al exportar

Códigos válidos del template (ejemplos): `SUELDOS`, `SUELDOSRES`, `VARIOS`, `VARIOSRES`, `IVA`, `ACT`, etc.

Mapeo del estudio al importar:

- IVA, IIBB, IIBB_ARBA, IIBB_CM03, TISH/TSH, CM → **VARIOS**
- BANCO (conciliación) → **CN**
- SUELDOS → **SUELDOS**

Motivo IVA→VARIOS: en muchas empresas el tipo `IVA` no está habilitado para Contabilidad y Tango rechaza la importación.
Motivo BANCO→CN: así lo define el instructivo de asientos / Conceptos Bancos (Clase Básico, tipo CN, moneda PES, fecha último día del mes).

Para determinación mensual *dentro* de Tango (no el Excel de la web) el estudio usa tipos propios: `DETIVA`, `DETIIBB`, `DETTISH` (ver guía de modelos).

## Cuentas

- No imputar cuentas **madre/rubro** (código prefijo de otras, o flag imputable = No).
- Código `99999` = sin asignar: bloquea la exportación.
- Si la cuenta **usa auxiliares** y el Excel no manda apropiaciones, Tango dice: «La cuenta no tiene asignado ningún tipo de auxiliar». Hay que asignar auxiliar al 100% en el plan, o elegir otra cuenta.

## Cuentas típicas del plan genérico del estudio

| Rol | Código |
|---|---|
| IVA crédito 10,5% | 11401 |
| IVA crédito 21% | 11402 |
| IVA percepciones | 11403 |
| IIBB percepciones | 11404 |
| IVA crédito 27% | 11409 |
| IVA retenciones | 11410 |
| IVA saldo técnico / a favor | 11411 |
| IVA saldo libre | 11412 |
| IIBB retenciones SIRCREB | 11418 |
| IIBB retenciones agentes | 11419 |
| IVA débito 21% | 21401 |
| IVA débito 10,5% | 21402 |
| IIBB a pagar | 21404 |
| TISH a pagar | 21407 |
| TISH gasto | 42401 |
| IIBB gasto | 42405 |
| Proveedores (SQL subdiario) | 21101 |

## SQL Tango (lectura, sin claves)

- Diccionario de empresas: base `Diccionario_039846_001`, tabla `Empresa` (NombreEmpresa → NombreBD).
- Facturas proveedores: vista `V_SubdiarioAsientosIV`.
- Sueldos: tablas `FORMULA`, `CONCEPTO`, `VARIABLE`, `TIPO_CONCEPTO`.
- Instancia típica de la oficina: `TANGOSRV\AXSQLEXPRESS`. Hay ~200 empresas en el diccionario.

## Dónde está la documentación Axoft recolectada

- Escritorio: `C:\Users\recep\Desktop\Tango` (ayudas HTML 26ar: Nexo, Contabilidad, IVA, Procesos generales, Sueldos).
- Red: `\\TANGOSRV\Compartido\CLIENTES\1 - Normativa - Vencimientos\Tango`
- Guías del estudio: asiento automático IVA (por comprobante vs determinación mensual) y modelos DETIVA / DETIIBB / DETTISH.
