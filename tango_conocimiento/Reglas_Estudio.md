# Reglas del estudio que la web tiene que respetar

Fuente: instructivos del estudio (monotributo, IVA, asientos, conciliaciones).
Si una regla de acá choca con la de un archivo de referencia puntual, manda el de referencia.

## Cómo trabajar

- Español rioplatense, al grano.
- No inventar alícuotas, topes ni valores fiscales. Si pueden haber cambiado, decir que hay que chequear ARCA.
- Fechas siempre `dd/mm/yyyy`. Encabezado de mes: `mmm-yy` (ene-26).
- Importes en Excel: formato Contabilidad. Cero se ve como `-`.
- Un resultado de cálculo va con fórmula, no tipeado. Solo el dato de origen se pega.
- Notas de crédito: si el PDF/export las trae en positivo, restan. Nunca sumarlas como venta.
- Preguntar antes de imputar una cuenta si no está claro. No forzar “la más probable”.

## Monotributo / recategorización (web)

- El mes de cada comprobante es **Período Facturado Desde**, no la Fecha de Emisión.
- Si el PDF no trae período facturado (típico en bienes), se usa la emisión y se marca como supuesto.
- Recibos (Recibo C u otro) se cargan. No se descartan porque digan “por la factura nro X”.
- Facturas, NC y Recibos tienen correlatividad **aparte** (y por punto de venta).
- Factura en USD: Imp. Total = Importe Dólares × Tipo de Cambio (fórmula).
- Recategorización: enero = 01/01–31/12 del año anterior; julio = 01/07 anterior–30/06 en curso.
- El control contra AFIP usa el semestre fijo. La proyección de “cuánto se puede facturar este mes” usa los últimos 12 meses rodantes.
- Si el CUIT emisor del PDF no es el del cliente activo, no se carga.

## IVA — retenciones y percepciones sufridas

- Agrupar por **Fecha Ret./Perc.** del reporte Mis Retenciones. Nunca por Fecha Comprobante.
- El crédito es **Importe Ret./Perc.** No Importe Total ni Importe Excedente.
- Estado **Pendiente** no suma hasta pasarlo a Tomada.

## Asientos a Tango (desde la web)

- Fecha del asiento = último día del mes del período.
- Moneda `PES`. Leyenda de renglones vacía (el concepto va en la cabecera).
- Debe = Haber al centavo. Cuentas madre/rubro no se exportan. `99999` bloquea.
- Devengamientos IVA/IIBB/TISH se exportan como tipo **VARIOS** (en muchas empresas el tipo IVA no está habilitado).
- Conciliación bancaria se exporta como tipo **CN** (instructivo de asientos).
- El código de cuenta es el del plan de **esa** sociedad. No reutilizar el de otro cliente.

## Conciliación bancaria

- Primer mes: saldo inicial del extracto. Meses siguientes: arrastre del final anterior.
- No forzar la diferencia a 0. No inventar meses sin extracto.
- Transferencias entre cuentas propias: una sola vez, del lado que recibe.
- CUIT propio no es contraparte.
