# Base operativa del estudio (Marti)

Reglas de trabajo del Estudio Contable / Zaidcorp. Usarlas en el chat de Tango y en procesos de la web.
No incluye claves, TANGO.INI ni contraseñas.

## Cómo se editan fórmulas de Tango Sueldos

- Por sociedad: Archivos → Liquidación → Fórmulas de liquidación → Apertura → Exportar / Importar (Excel nativo).
- Copiar entre sociedades: Herramientas → Administrador → Empresas → Transferir información.
- Editar de a UNA sociedad. ASCII/SICOSS no es el canal de fórmulas.
- Actualizar solo la fórmula NO alcanza: hay que actualizar/refrescar el CONCEPTO que la referencia y recién reliquidar.
- Una misma fórmula (nro) puede estar compartida por varios conceptos: al cambiarla afecta a todos. Si hace falta, crear fórmula nueva.
- Campo Fórmula (azul): si se cambia el número puede crear registro nuevo. Al modificar, conservar ROW_VER en el import.
- Circuito: bajar planilla vacía de Tango → armar FormulaImporte / FormulaCantidad → pegar o importar.

## Sintaxis que el estudio usa

- Condicional: `SI(condicion;verdadero;falso)` (en Tango a veces coma; en exports del estudio suele ir `;`).
- OR: `O` (no escribir OR).
- Códigos de modalidad y legajo van entre comillas: `"099"`, `"1"`.
- `USUELD` = sueldo básico a liquidar. Nunca `USUELO` (error de sintaxis).
- `CTRATO` = modalidad de contratación (texto).
- `LEGAJO` = número de legajo (texto). Comparar `"1"`, `"10"`.
- `DGICOND` = condición SIJP. `"2"` = jubilada.
- `IMPOR` = importe fijo del concepto. `CANTIDAD` = resultado de la fórmula Cantidad.
- Excel de RECEPCION no reconoce IFS: usar SI anidado.

## CTRATO que usa el estudio

- `"099"` = directivo / socio → suele excluirse (importe 0) en aportes y haberes generales.
- `"048"` = modalidad a excluir en varias retenciones/aportes (PIMIPAC y similares).
- `"001"` = jornada / tiempo parcial → 50% cuando el concepto es suma fija por IMPOR.
- `"008"` = jornada completa → 100% del IMPOR.

## Patrones de fórmula (Importe; Cantidad no tocar salvo pedido)

Excluir directivo:

```
SI(CTRATO="099";0;<formula_original>)
```

Excluir 048:

```
SI(CTRATO="048";0;<formula_original>)
```

Excluir 099 o jubilada:

```
SI((CTRATO="099") O (DGICOND="2");0;<formula_original>)
```

50% / 100% sobre IMPOR (va en IMPORTE; Cantidad vacía). El % NO va como 50/100 en Cantidad si Importe ya usa IMPOR:

```
SI(CTRATO="001";IMPOR*0.5;SI(CTRATO="008";IMPOR;0))
```

Solo ciertas modalidades (inclusión), ej. honorario técnico:

```
SI((CTRATO="048") O (CTRATO="099");USUELD/30*CANTIDAD;0)
```

Concepto 1 sueldo básico del estudio: `USUELD/30*CANTIDAD` (no el proporcional).

Si el legajo tiene mal el CTRATO, la fórmula “parece” mal: corregir la modalidad y refrescar conceptos.

## Liquidación

- Separar Primera/Segunda quincena vs Mensual. Al emitir fin de mes: filtrar solo Mensual.
- Remunerativo = TOTHAB. No remunerativo = TOTNR.
- Adelanto 20005 si aparece sin querer: revisar Liquidaciones habilitadas / particulares (no solo el flag Anticipo).
- Un concepto que liquida en **0** no se borra como en Excel. Para que no se calcule: en el concepto, **Liquidaciones habilitadas**, sacarlo de esa liquidación (mensual / quincena). Para que no se imprima en el recibo: marcar **no imprimir si el importe es cero**. Si da 0 por CTRATO 099/048, la fórmula está bien: no hace falta eliminarlo.

## Papeles de bancos / conciliación

1. Solo el saldo inicial del PRIMER mes viene del extracto.
2. Meses siguientes: inicial = final del mes anterior (arrastre).
3. Final = inicial + créditos − débitos.
4. “Según resumen” = saldo del extracto.
5. Diferencia (final − extracto) NO se fuerza a 0: es desvío a revisar.
6. No inventar meses sin extracto: dejar pendiente.
7. Banco Provincia a veces duplica movimientos de fin de mes al inicio del siguiente: quedarse con los del mes que corresponde.
8. Transferencias entre cuentas propias: una sola vez del lado que RECIBE. El lado que SALE no se importa. COELSA = Mercado Pago.
9. CUIT propio no es contraparte.

## Conceptos Bancos (asiento a Tango)

Instructivo: `\\TANGOSRV\Compartido\CLIENTES\zzInstrucciones\Conceptos Bancos 2.xlsx`.

- Verificación: saldo inicial + créditos − débitos = saldo final del extracto.
- Dos bloques: ASIENTO A IMPORTAR A TANGO / NO VA AL ASIENTO DE TANGO.
- IVA discriminado (comisiones/intereses + IVA CF) → bloque NO Tango (ya va en Compras).
- Netear cuentas excepto FCI y transferencias propias (esas en 2 filas).
- Todo con fórmula (SUMIFS), no pegar resultados. Formato número Contabilidad.
- Import Tango: un asiento por banco y mes; Clase Basico; tipo CN; moneda PES; fecha = último día del mes.

## AFIP / monotributo (web y facturación)

- Control: el mes del monto = período facturado (fecha desde/hasta), NUNCA la fecha de emisión.
- Si el PDF no trae "Período Facturado" (venta de bienes), se usa la emisión y se marca como supuesto.
- Recibos se cargan igual que facturas. No se descartan porque el texto cite "por la factura nro X".
- Correlatividad: Facturas, NC y Recibos son series independientes (y por punto de venta).
- NC siempre restan. Recibos entran en positivo.
- Factura en USD: Imp. Total = Importe Dólares × Tipo de Cambio (fórmula, no valor pegado).
- Si el CUIT emisor del PDF no coincide con el cliente, no se carga.
- PDFs nominados: Cliente + número de factura.
- Emitir rápido si la planilla ya está revisada; no reconfirmar CAE por CAE.
- No pedir ni guardar claves fiscales en Excel ni en el chat.

## IVA — Mis Retenciones (web)

- Agrupar por Fecha Ret./Perc., nunca por Fecha Comprobante.
- Crédito = Importe Ret./Perc. (no Importe Total ni Excedente).
- Estado Pendiente no suma hasta Tomada.
