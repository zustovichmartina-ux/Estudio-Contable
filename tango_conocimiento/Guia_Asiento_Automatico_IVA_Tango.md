# Guía: asientos automáticos de IVA en Tango

Basada en la documentación oficial de Axoft/Tango disponible en la carpeta `Tango` (módulos Liquidador de IVA, Contabilidad y Procesos generales).

Este documento cubre **dos asientos distintos** que suelen confundirse:

- **Partes 1 a 5**: el asiento de **devengamiento por comprobante** — el débito/crédito fiscal que se genera cada vez que cargás una factura de compra o venta en Liquidador de IVA.
- **Parte 6**: el asiento **mensual de determinación de IVA** — el que netea Débito Fiscal contra Crédito Fiscal (por alícuota) y percepciones, y deja el saldo en "IVA a pagar" (o saldo a favor). Este es un asiento aparte, que se arma en Contabilidad con un modelo de asiento y fórmulas de saldo.

---

## Parte 1 — Configuración previa (se hace una sola vez)

Si tu Tango ya está en producción, es probable que la mayoría de estos pasos ya estén hechos. Revisalos igual antes de la prueba.

### 1. Habilitar las cuentas contables de IVA para el módulo

En **Contabilidad (o Procesos generales) > Archivos > Cuentas**, ubicá (o creá) las cuentas de:
- IVA Débito Fiscal (para ventas)
- IVA Crédito Fiscal (para compras)

En la solapa **Módulos** de cada cuenta, verificá que estén habilitadas para **Liquidador de IVA**. Si no está tildado, el sistema no va a dejar usarlas en los modelos de ingreso.

### 2. Definir el/los Tipo de asiento

En **Procesos generales > Datos contables > Tipos de asiento**, confirmá que existe un tipo de asiento habilitado para el módulo Liquidador de IVA (por ejemplo "IVA COMPRAS" / "IVA VENTAS", o uno genérico). Es un dato obligatorio para el paso siguiente.

### 3. Parámetros contables del Liquidador de IVA

En **Liquidador de IVA > Archivos > Parámetros contables del Liquidador de IVA**, en la solapa Principal:
- Tildá **"Genera asiento en el ingreso de comprobantes"** (así el asiento se contabiliza en el momento de cargar la factura, no después).
- Decidí si activás **"Respeta definición del modelo de ingreso"** (si lo activás, no vas a poder tocar cuentas/importes a mano al cargar el comprobante, solo auxiliares).

### 4. Parametrización contable por Tipo de comprobante

En **Liquidador de IVA > Archivos > Parametrización contable en Liquidador de IVA > Parametrización contable - Tipos de comprobantes**:
- Elegí el tipo de comprobante que vas a usar para la prueba (por ejemplo "Factura A - Compras").
- Tildá **"Genera asiento"**.
- Asigná el **Tipo de asiento** definido en el paso 2 (obligatorio).
- Opcional: elegí una leyenda para el asiento.

### 5. (Opcional) Cuentas específicas por Cliente/Proveedor

Si querés que un proveedor o cliente puntual impacte en una cuenta distinta a la genérica del modelo, cargala en:
- **Parametrización contable - Proveedores** (hasta 2 cuentas)
- **Parametrización contable - Clientes** (hasta 2 cuentas)

Si no configurás nada acá, se usa la cuenta definida en el Modelo de ingreso.

### 6. Revisar el Modelo de ingreso de comprobantes a usar

En **Liquidador de IVA > Archivos > Modelos de ingreso de comprobantes**, abrí el modelo que vas a usar para la prueba (por ejemplo el modelo "A" para responsables inscriptos) y verificá, en su parametrización contable:
- Que cada renglón (neto gravado, IVA, etc.) tenga una **cuenta contable** asignada.
- Que tenga el **D/H** (debe/haber) correcto.
- Que el **Origen** de la cuenta sea el que corresponde (cuenta del modelo, o CC1/CC2/CP1/CP2 si querés que tome la del cliente/proveedor).

Importante: no uses el mismo modelo para compras y ventas — la doc recomienda duplicarlo (botón "Copiar") porque las cuentas de débito/crédito fiscal son distintas.

### 7. Verificar la alícuota de IVA a usar

En **Liquidador de IVA > Archivos > Impuesto al valor agregado > Alícuotas de IVA**, confirmá que existe la alícuota que vas a facturar (21%, 10.5%, etc.).

---

## Parte 2 — Cargar el comprobante de prueba

1. Andá a **Liquidador de IVA > Comprobantes > Registración de comprobantes** y elegí **Nuevo**.
2. Completá:
   - Tipo de comprobante (el parametrizado en el paso 4).
   - Cliente o Proveedor (real u ocasional, código `000000`).
   - Fecha de emisión y Fecha contable.
   - Modelo de ingreso (se propone el habitual del cliente/proveedor; podés cambiarlo con F3).
3. Cargá los importes que pide el modelo (neto gravado, alícuota, etc.). El sistema calcula el IVA solo si el modelo está armado para eso.
4. Al confirmar el alta, si "Genera asiento en el ingreso de comprobantes" está activo, **el asiento se genera automáticamente en ese momento** — ahí ocurre el devengamiento.

### Verificar el asiento generado

- En el mismo comprobante, usá el botón/opción **"Ver asiento"** para ver los renglones (cuenta, debe, haber).
- Confirmá que:
  - Tiene como mínimo 2 renglones.
  - Debe y Haber están balanceados (el sistema no deja grabar un asiento descuadrado).
  - La cuenta de IVA Crédito/Débito Fiscal aparece con el importe correcto según la alícuota.
- En el título del comprobante, el **estado del asiento** debería decir "Generado".

Si el asiento no se generó (quedó "Sin generar"), revisá el paso 3 (parámetro "Genera asiento en el ingreso") o el paso 4 (tildado "Genera asiento" en el tipo de comprobante), o generalo manualmente después con el proceso del punto siguiente.

---

## Parte 3 — Generar el asiento después (si no se generó al cargar)

Si preferís no generar el asiento en el momento de la carga, andá a **Liquidador de IVA > Procesos periódicos > Contabilización > Generación de asientos contables**:
- Elegí el rango de fechas (por defecto, el mes actual).
- En "Comprobantes a procesar" dejá **"Sin generar"**.
- Activá **"Visualiza comprobantes a procesar"** para revisar la lista antes de generar (podés destildar el comprobante de prueba si no lo querés incluir).
- Ejecutá el proceso y confirmá que el comprobante de prueba pasó a estado "Generado".

---

## Parte 4 — Exportar el asiento a Contabilidad

1. Andá a **Liquidador de IVA > Procesos periódicos > Contabilización > Exportación > Exportación de asientos contables**.
2. Como Destino elegí **"Base de datos actual"** (asumiendo que Contabilidad está en la misma base).
3. Tipo de generación: **"Comprobante"** (genera un asiento por comprobante — más fácil de rastrear en una prueba).
4. En "Comprobantes a procesar con asiento" dejá **"Generado"**.
5. Activá **"Visualiza asientos exportados"** para ver el reporte de control al terminar.
6. Ejecutá el proceso. Al finalizar te va a mostrar una grilla con los asientos importados (o rechazados, con el motivo).
7. Hacé doble clic sobre el asiento importado para abrirlo directamente en **Contabilidad**, y confirmá ahí que el asiento quedó registrado con las cuentas y auxiliares correctos.

Si algún asiento no pasa las validaciones de Contabilidad, el sistema anula todo el lote y el comprobante vuelve a estado "Generado" (podés corregir y reintentar). Para consultar o deshacer una exportación ya hecha, están los procesos **"Lotes contables generados"** y **"Anulación de lotes contables generados"**, dentro de la misma carpeta.

---

## Parte 5 — Si más adelante querés automatizarlo sin intervención manual

Una vez que la prueba funcione bien, para que la generación y exportación corran solas (por ejemplo todas las noches o una vez al mes), andá a **Procesos generales > Transferencias > Automatización**:
- **Automatización de generación de asientos**
- **Automatización de exportación de asientos**

Ahí podés programar la frecuencia, el usuario y el período a procesar para que Tango ejecute estos mismos procesos sin que nadie los dispare a mano.

---

## Parte 6 — El asiento mensual de determinación de IVA (neteo a "IVA a pagar")

Este es el asiento tipo el de tu ejemplo: una sola línea por cuenta (IVA Débito Fiscal, IVA Crédito Fiscal por cada alícuota, Percepción IVA) más una línea que balancea contra "IVA a pagar". No lo genera Liquidador de IVA — se arma en **Contabilidad**, usando un **Modelo de asiento** con fórmulas que traen el saldo de cada cuenta, y después se dispara con **Generación masiva de asientos**.

### 6.1 Crear el modelo de asiento

Andá a **Contabilidad > Archivos > Modelos de asientos** y creá un modelo nuevo (por ejemplo `DETIVA`):

- **Código / Descripción**: por ejemplo `DETIVA` — "Determinación mensual de IVA".
- **Tipo de asiento**: elegí o creá uno específico (por ejemplo "IVA DETERMINACIÓN"), para poder filtrarlo después en listados.
- **Leyenda**: opcional, por ejemplo "IVA período {mes/año}".

### 6.2 Cargar los renglones (uno por cuenta)

En la solapa **Cuentas contables** del modelo, agregá un renglón por cada cuenta que aparece en tu asiento de ejemplo:

| Cuenta | D/H | Fórmula (Importe) |
|---|---|---|
| IVA Débito Fiscal | D | saldo de la cuenta (ver 6.3) |
| IVA Crédito Fiscal | H | saldo de la cuenta |
| IVA Crédito Fiscal 10,5% | H | saldo de la cuenta |
| IVA Crédito Fiscal 27% | H | saldo de la cuenta |
| IVA Débito Fiscal (si tenés una segunda cuenta, p. ej. NC) | H | saldo de la cuenta |
| Percepción IVA | H | saldo de la cuenta |
| IVA a pagar | D o H (según el signo) | diferencia (ver 6.4) |

No cargues un importe fijo: en la columna **Fórmula/Importe** hacé clic en el botón **"..."** para abrir el asistente de **Definición guiada de fórmulas** y armar la fórmula con la variable de saldo.

### 6.3 La fórmula para traer el saldo de cada cuenta

Usá la variable **ACUCTA** (saldo acumulado a una fecha) o **MOVCTA** (movimiento entre dos fechas — más intuitivo si estas cuentas se resetean a cero cada mes con este mismo asiento):

```
ACUCTA('Todas', '21401', '21401', 'S', EJACT, HASFE, 'H')
```

- `'21401'` / `'21401'`: rango de cuenta (desde/hasta) — poné el código real de tu cuenta de IVA Débito Fiscal.
- `'S'`: tipo de movimiento = Saldo.
- `EJACT`: ejercicio actual (variable básica del sistema).
- `HASFE`: "hasta fecha" seleccionada al correr el proceso (así cada mes trae el saldo a esa fecha, sin que edites la fórmula).
- `'H'`: tipo de saldo histórico (excluye ajustes y cierres).

Repetí esta fórmula, cambiando el rango de cuenta, para cada línea de Crédito Fiscal y Percepción IVA. Si preferís el movimiento del período en lugar del saldo acumulado, usá `MOVCTA('Todas', '21401', '21401', 'S', EJACT, DESFE, HASFE, 'H')`.

### 6.4 La fórmula del renglón que balancea ("IVA a pagar")

Para que el asiento cierre solo, la línea de "IVA a pagar" tiene que ser la diferencia entre el total de débitos y el total de créditos de las líneas anteriores. Se arma combinando las mismas variables ACUCTA de cada cuenta con una función lógica, por ejemplo:

```
ABS( ACUCTA(...DebitoFiscal...) - ACUCTA(...CreditoFiscal21...) - ACUCTA(...CreditoFiscal10_5...) - ACUCTA(...CreditoFiscal27...) - ACUCTA(...Percepcion...) )
```

Y en la columna D/H podés usar `SI()` para que el sistema decida si va al Debe (saldo a favor) o al Haber (IVA a pagar), según el signo del resultado. Esta parte conviene armarla con tu contador o con soporte de Axoft la primera vez, porque la sintaxis exacta de SI()/IF() dentro del asistente de fórmulas requiere probarla contra tu plan de cuentas real.

### 6.5 Habilitar el modelo para generación masiva

En la solapa **Parametrización** del modelo:
- Tildá **"Habilitado para generación masiva de asientos"**.
- Asientos: **Contable**.
- Moneda: moneda corriente.
- Clase de asiento: **Básico**.

### 6.6 Generarlo todos los meses

Andá a **Contabilidad > Asientos > Generación masiva de asientos**:
1. Seleccioná el modelo `DETIVA` en la grilla.
2. En **Parámetros**, indicá el ejercicio/período y la fecha a procesar (por ejemplo, el último día del mes).
3. Confirmá. El sistema calcula los importes con las fórmulas y genera el asiento; te va a mostrar una consulta con el resultado, y si algo se rechaza (por ejemplo el asiento no balancea), aparece un Excel con el detalle del motivo.

### 6.7 Automatizarlo sin correrlo a mano

Una vez que el modelo esté probado y de confianza, se puede programar igual que los procesos de la Parte 5, para que se dispare solo cada mes sin que nadie entre a ejecutarlo.

Nota: para que las cuentas de IVA Débito/Crédito Fiscal muestren en cada período sólo el movimiento del mes (y no arrastren de meses anteriores), es una práctica común que este mismo asiento las deje en cero — revisá con tu contador si tu plan de cuentas ya está armado así antes de automatizar.

---

## Checklist rápido antes de probar

**Asiento por comprobante (Partes 1 a 5)**
- [ ] Cuentas de IVA Débito/Crédito Fiscal habilitadas para el módulo Liquidador de IVA
- [ ] Tipo de asiento definido y habilitado
- [ ] "Genera asiento en el ingreso de comprobantes" activo (Parámetros contables del Liquidador de IVA)
- [ ] Tipo de comprobante de prueba con "Genera asiento" tildado y tipo de asiento asignado
- [ ] Modelo de ingreso con cuentas y D/H cargados en cada renglón
- [ ] Alícuota de IVA a usar, dada de alta

**Asiento mensual de determinación (Parte 6)**
- [ ] Modelo de asiento `DETIVA` creado con un renglón por cuenta
- [ ] Fórmulas ACUCTA/MOVCTA cargadas y probadas para cada renglón
- [ ] Fórmula del renglón "IVA a pagar" validada con el contador
- [ ] Modelo habilitado para generación masiva de asientos
- [ ] Prueba corrida desde Generación masiva de asientos y asiento resultante balanceado
