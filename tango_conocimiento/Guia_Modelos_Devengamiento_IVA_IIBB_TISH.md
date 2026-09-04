# Modelos de asiento Tango — devengamiento mensual IVA / IIBB / TISH

Parametrización para **Contabilidad > Archivos > Modelos de asientos** + **Generación masiva de asientos**.

Complementa `Guia_Asiento_Automatico_IVA_Tango.md`:

- **Partes 1 a 5 de esa guía**: asiento **por comprobante** (Liquidador de IVA). No es este trabajo.
- **Este documento**: asiento **mensual de determinación / devengamiento**, uno por impuesto, que se dispara a fin de mes.

Fuente: ayudas Axoft 26ar en esta misma carpeta (Modelos de asientos, Variables contables, Generación masiva, Tipos de asiento).

---

## Qué hace cada modelo

| Modelo | Código Tango | Qué deja asentado |
|---|---|---|
| Determinación IVA | `DETIVA` | Netea DF, CF, retenciones y percepciones. Cierra contra **IVA a pagar** o **saldo a favor**. Deja DF/CF en cero si el plan está armado así. |
| Devengamiento IIBB | `DETIIBB` | Dr gasto (impuesto determinado). Aplica retenciones, percepciones, SIRCREB y SAF anterior. Cierra contra **IIBB a pagar** o **SAF nuevo**. |
| Devengamiento TISH | `DETTISH` | Dr gasto tasa. Aplica retenciones TISH, derecho de oficina y SAF. Cierra contra **Tasa a pagar** o **SAF nuevo**. |

Tango **no** calcula la DDJJ. Trae **saldos de cuentas** con `ACUCTA` / `MOVCTA`. Por eso IVA se automatiza bien (DF/CF ya están en el mayor). IIBB y TISH necesitan que el **determinado** tenga una fuente (ver más abajo).

---

## Reglas de Tango que no se pueden ignorar

De **Generación masiva de asientos**:

1. Renglón con importe **0** → se descarta (sirve para el `SI()` del cierre).
2. Importe **negativo** → **rechaza el asiento**. Siempre `ABS(...)`.
3. Asiento sin líneas → rechazo.
4. Si el tipo de asiento entra como **Ingresado**, Debe = Haber es obligatorio.

Consecuencia: **no uses `SI()` en la columna D/H**. Esa columna es D o H fijo. El cierre se arma con **dos renglones** (a pagar / saldo a favor); el que no aplica queda en 0 y Tango lo tira.

---

## Paso 0 — Una sola vez por empresa

### 0.1 Tipos de asiento

**Procesos generales > Datos contables > Tipos de asiento**

| Código | Descripción | Habilitado | Estado inicial | Módulos |
|---|---|---|---|---|
| `DETIVA` | Determinación mensual IVA | Sí | Ingresado | Contabilidad |
| `DETIIBB` | Devengamiento IIBB | Sí | Ingresado | Contabilidad |
| `DETTISH` | Devengamiento TISH | Sí | Ingresado | Contabilidad |

Agrupación opcional: `IMPUESTOS`. Leyenda de encabezado, por ejemplo `Determinación IVA {mes}`.

Si la empresa ya tiene el tipo `IVA` pero **no está habilitado para Contabilidad**, no lo uses: Tango rechaza la importación/generación. Por eso el estudio exporta IVA como `VARIOS`. Estos tipos nuevos van solo a Contabilidad.

### 0.2 Completar la ficha de cuentas

Abrí el Excel `Ficha_Parametrizacion_Asientos_Impuestos.xlsx` (misma carpeta), hoja **Cuentas**. Pegá el código real del plan de esa empresa. Ese código es el que va en las fórmulas (`«DF21»`, etc.).

---

## Paso 1 — Macros (capa de parametrización)

**Contabilidad > Archivos > Variables de contabilidad** → alta, tipo **Macro**, parámetro de salida **número real**.

Usá el botón **Fx** (definición guiada) o pegá la gramática.

Plantilla de saldo a la fecha del proceso:

```
ABS(ACUCTA('Todas', 'CODIGO', 'CODIGO', 'S', EJACT, HASFE, 'H'))
```

- `'Todas'`: tipo de cuenta (prioridad sobre el rango; acá no filtra).
- `'CODIGO'` / `'CODIGO'`: desde / hasta. Una sola cuenta.
- `'S'`: saldo (Debe − Haber). Por eso el `ABS`.
- `EJACT`: ejercicio elegido en Generación masiva.
- `HASFE`: “Hasta fecha” del proceso (último día del mes).
- `'H'`: histórico, **sin** ajustes ni cierres.

Si preferís solo el movimiento del mes (no arrastre):

```
ABS(MOVCTA('Todas', 'CODIGO', 'CODIGO', 'S', EJACT, DESFE, HASFE, 'H'))
```

**Recomendación IVA:** `ACUCTA` + `HASFE`, para que el asiento deje las cuentas de DF/CF en cero aunque haya quedado un resto.

**Recomendación gasto IIBB/TISH (cuenta de resultado):** `MOVCTA` + `DESFE`/`HASFE`, para no tomar el acumulado del ejercicio.

### Macros IVA (ejemplo; reemplazá códigos)

| Código | Fórmula |
|---|---|
| `IVA_DF21` | `ABS(ACUCTA('Todas', '«DF21»', '«DF21»', 'S', EJACT, HASFE, 'H'))` |
| `IVA_DF105` | igual con cuenta DF 10,5 % |
| `IVA_DF27` | igual con cuenta DF 27 % |
| `IVA_CF21` | igual con CF 21 % |
| `IVA_CF105` | igual con CF 10,5 % |
| `IVA_CF27` | igual con CF 27 % |
| `IVA_RET` | igual con Retenciones IVA |
| `IVA_PER` | igual con Percepciones IVA |
| `IVA_STEC` | igual con Saldo técnico período anterior |
| `IVA_SLIB` | igual con Saldo libre disponibilidad |
| `IVA_POS` | `IVA_DF21+IVA_DF105+IVA_DF27-IVA_CF21-IVA_CF105-IVA_CF27-IVA_RET-IVA_PER-IVA_STEC-IVA_SLIB` |

Si una alícuota no existe en el plan, no crees la macro ni el renglón.

Primera prueba: en un mes ya cerrado, compará cada macro contra el mayor. Si `ABS` te da el importe pero el asiento **duplica** el lado (por el signo del saldo), invertí D/H de ese renglón.

---

## Paso 2 — Modelo `DETIVA`

**Contabilidad > Archivos > Modelos de asientos**

### Solapa Principal

- Código: `DETIVA`
- Descripción: Determinación mensual de IVA
- Tipo de asiento: `DETIVA`
- Leyenda: Determinación IVA

### Solapa Cuentas contables

D/H es el lado **habitual** de ese renglón. Importe = nombre de la macro (o la fórmula completa).

| Nro | Cuenta (concepto) | D/H | Fórmula / Importe |
|---|---|---|---|
| 1 | IVA Débito Fiscal 21 % | D | `IVA_DF21` |
| 2 | IVA Débito Fiscal 10,5 % | D | `IVA_DF105` |
| 3 | IVA Débito Fiscal 27 % | D | `IVA_DF27` |
| 4 | IVA Crédito Fiscal 21 % | H | `IVA_CF21` |
| 5 | IVA Crédito Fiscal 10,5 % | H | `IVA_CF105` |
| 6 | IVA Crédito Fiscal 27 % | H | `IVA_CF27` |
| 7 | Retenciones IVA | H | `IVA_RET` |
| 8 | Percepciones IVA | H | `IVA_PER` |
| 9 | Saldo técnico IVA (arrastre) | H | `IVA_STEC` |
| 10 | Saldo libre IVA (arrastre) | H | `IVA_SLIB` |
| 11 | IVA a pagar | H | `SI(IVA_POS>0, IVA_POS, 0)` |
| 12 | Saldo a favor IVA (nuevo período) | D | `SI(IVA_POS<0, ABS(IVA_POS), 0)` |

Ecuación (la misma que usa el estudio):

```
posición = DF − (CF + retenciones + percepciones + saldos técnicos + saldos libres)
posición > 0  →  IVA a pagar (Haber)
posición < 0  →  saldo a favor (Debe)
```

Si las NC de compras/ventas están en **cuentas distintas** de DF/CF, agregá renglones (NC compras al Debe con DF; NC ventas al Haber con CF), y sumalas/restalas en `IVA_POS`.

### Solapa Parametrización

- [x] Habilitado para generación masiva de asientos
- Generación: **Asientos**
- Asientos: **Contable**
- Moneda: corriente
- Clase: **Básico**

---

## Paso 3 — Modelo `DETIIBB`

Misma mecánica. Ecuación del estudio:

```
posición = impuesto determinado − (retenciones + percepciones + SIRCREB + SAF anterior)
posición > 0  →  IIBB a pagar (Haber)
posición < 0  →  SAF nuevo (Debe)
```

### El determinado no está en el mayor (caso típico)

El gasto IIBB **se genera con este asiento**. No hay saldo previo para `ACUCTA`. Opciones:

| Opción | Cuándo usarla | Cómo |
|---|---|---|
| **A — Provisión sobre ventas** | Alícuota simple, una jurisdicción | Macro `IIBB_DET` = `MOVCTA('Todas','«VENTAS_DESDE»','«VENTAS_HASTA»','H',EJACT,DESFE,HASFE,'H')*0.03` (ajustá alícuota y rango de ventas). Es **provisión**, no DDJJ. |
| **B — Carga del modelo a mano** | Determinado sale de ARBA / CM / Excel del estudio | En el modelo, renglón 1 **sin fórmula**. Abrís el modelo desde asientos, completás el determinado, el resto de renglones sí puede llevar fórmula. No sirve para generación masiva de ese renglón. |
| **C — Excel del estudio → Tango** | Querés el determinado exacto de la DDJJ | Seguí generando el asiento en Devengamientos e importá con la plantilla Tango. Este modelo queda para IVA (o para aplicar solo créditos). |

### Si usás opción A (generación masiva)

Macros extra:

| Código | Fórmula |
|---|---|
| `IIBB_DET` | ventas del mes × alícuota (o el importe que definas) |
| `IIBB_RET` | `ABS(ACUCTA('Todas', '«RET_IIBB»', '«RET_IIBB»', 'S', EJACT, HASFE, 'H'))` |
| `IIBB_PER` | percepciones sufridas |
| `IIBB_SIR` | SIRCREB / retenciones bancarias |
| `IIBB_SFA` | saldo a favor período anterior |
| `IIBB_POS` | `IIBB_DET-IIBB_RET-IIBB_PER-IIBB_SIR-IIBB_SFA` |

Renglones:

| Nro | Cuenta | D/H | Fórmula |
|---|---|---|---|
| 1 | IIBB determinado / gasto | D | `IIBB_DET` |
| 2 | Retenciones IIBB sufridas | H | `IIBB_RET` |
| 3 | Percepciones IIBB sufridas | H | `IIBB_PER` |
| 4 | SIRCREB / ret. bancarias | H | `IIBB_SIR` |
| 5 | SAF IIBB período anterior | H | `IIBB_SFA` |
| 6 | IIBB a pagar | H | `SI(IIBB_POS>0, IIBB_POS, 0)` |
| 7 | SAF IIBB nuevo período | D | `SI(IIBB_POS<0, ABS(IIBB_POS), 0)` |

Parametrización: igual que `DETIVA` (habilitado masivo, contable, básico).

Convenio multilateral: duplicá el modelo (`DETCM`) con las cuentas CM. No mezcles ARBA y CM en el mismo asiento.

---

## Paso 4 — Modelo `DETTISH`

Ecuación del estudio:

```
posición = tasa determinada − (retenciones TISH + derecho de oficina + SAF anterior)
```

Misma salvedad que IIBB: la tasa determinada sale de la DDJJ municipal, no del mayor. Opción A = base × alícuota municipal; B = carga manual; C = Excel del estudio.

| Nro | Cuenta | D/H | Fórmula |
|---|---|---|---|
| 1 | Gasto tasa / TISH determinada | D | `TISH_DET` |
| 2 | Retenciones TISH | H | `TISH_RET` |
| 3 | Derecho de oficina | H | `TISH_DOF` |
| 4 | SAF TISH período anterior | H | `TISH_SFA` |
| 5 | Tasa TISH a pagar | H | `SI(TISH_POS>0, TISH_POS, 0)` |
| 6 | SAF TISH nuevo período | D | `SI(TISH_POS<0, ABS(TISH_POS), 0)` |

---

## Paso 5 — Correrlo cada mes

**Contabilidad > Asientos > Generación masiva de asientos**

1. En la grilla, tildá `DETIVA` (y `DETIIBB` / `DETTISH` si están en opción A).
2. **Fecha a generar:** último día del mes.
3. **Concepto:** editá si hace falta (`Determinación IVA 08/2026`).
4. En **Parámetros** (aparece porque hay fórmulas con variables):
   - Ejercicio = el del mes.
   - Períodos = el mes (desde = hasta).
   - Fechas a procesar: **01/mm** a **último día** (llenan `DESFE` y `HASFE`).
   - Asientos: Contables.
5. Aceptar. Live con los asientos; si hay rechazo, Excel con el motivo (casi siempre: no balancea, negativo, o cuenta inexistente).

Orden práctico del cierre:

1. Que estén exportados a Contabilidad todos los asientos de comprobantes del mes (Liquidador / compras / ventas / bancos).
2. Recién ahí generación masiva de `DETIVA` / `DETIIBB` / `DETTISH`.
3. Control: mayor de DF/CF en cero; IVA a pagar = posición de la DDJJ.

Botón **Asientos generados** del modelo: historial. Si hay que reprocesar, eliminá el asiento de ese mes y volvé a generar.

---

## Checklist de la primera empresa

- [ ] Tipos `DETIVA` / `DETIIBB` / `DETTISH` creados y habilitados en Contabilidad
- [ ] Hoja **Cuentas** del Excel completada con el plan real
- [ ] Macros dadas de alta; una prueba de importe vs mayor
- [ ] Tres modelos con D/H, fórmulas y tilde de generación masiva
- [ ] Prueba en un mes ya conocido (no el mes abierto de producción)
- [ ] Asiento balanceado; DF/CF quedan en el saldo esperado
- [ ] Recién después, repetir en el resto de las empresas (cambiar códigos en las macros)

---

## Relación con el Liquidador de IVA

No confundir:

| Pieza | Dónde | Para qué |
|---|---|---|
| Modelo de **ingreso de comprobantes** | Liquidador de IVA | Cada factura: neto, IVA, percepción. |
| Parametrización contable del tipo de comprobante | Liquidador de IVA | Que esa factura genere asiento. |
| Exportación de asientos | Liquidador → Contabilidad | Llevar DF/CF/percepciones al mayor. |
| Modelo `DETIVA` / `DETIIBB` / `DETTISH` | **Contabilidad** | Un asiento de fin de mes que determina / devenga. |

TISH no tiene módulo Tango. Solo existe como modelo de Contabilidad (y/o como asiento exportado desde el estudio).
