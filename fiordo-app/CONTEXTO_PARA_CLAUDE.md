# Contexto del proyecto (para retomar en otra conversación)

Este archivo resume todo lo decidido hasta ahora. Leelo antes de tocar el código.

## Qué es

App web de **punto de equilibrio para Fiordo Café** (cafetería de especialidad, Las Heras 2808, Mar del Plata),
cliente de un estudio contable. La usa la contadora y el dueño desde la compu y el celular.

- **Publicada en Vercel:** https://estudio-contable-t8lz.vercel.app (se despliega sola al hacer push a `master`).
- **Repositorio:** `zustovichmartina-ux/Estudio-Contable`, carpeta `fiordo-app/` (Vercel → Root Directory = `fiordo-app`).
- **Stack:** Next.js 15 (App Router) + React 19 + TypeScript. Sin Tailwind: CSS propio con tokens en `app/globals.css`.
  Librerías: `xlsx` (SheetJS 0.18.5) para leer/escribir Excel, `jspdf` + `jspdf-autotable` para el PDF,
  `@upstash/redis` para la base compartida.
- La usuaria **no quiere Streamlit**; quiere Vercel. Es contadora, no programadora: explicaciones simples, en castellano rioplatense (vos).

## Límite importante (no cambiar)

El cliente vende parte "en negro". La app calcula el punto de equilibrio **con todas las ventas registradas en Bistrosoft**
para gestión. **No** se construyen funciones para calcular cuánto facturar/declarar ni para separar lo facturado de lo no
facturado con ese fin (sería ayudar a evadir). El reporte de Bistrosoft trae columnas "Comandas" vs "Facturas":
no se usan para eso.

## Cómo calcula

- **Ventas:** todo lo registrado en Bistrosoft, se cobre como se cobre.
- **Gastos fijos** = cargados a mano (pestaña Gastos; se copian a los meses siguientes) + facturas de ARCA de
  proveedores clasificados como "Gasto fijo" (sin IVA).
- **Variables (% de la venta):**
  - Mercadería: compras de ARCA clasificadas "Mercadería" ÷ ventas del mes (si hay); si no, costo de receta por
    producto; si falta, estimado (32 % por defecto). Configurable: auto / compras / recetas.
  - Comisiones: **por medio de pago** sobre lo vendido con cada medio (Efectivo 0 %, Tarjetas 3,5 %, QR 1 %, Online 4 %:
    valores de referencia editables, a confirmar con el cliente).
  - Ingresos brutos: **4 %** (confirmado por la usuaria).
- Margen de contribución = 1 − variables. **Punto de equilibrio = fijos ÷ margen**. Por día = ÷ días que abre
  (días del mes − días cerrados; Fiordo abre todos los días).
- Semáforo diario, proyección, margen de seguridad, análisis por día de semana, aporte por producto.

## Datos que entran (archivos reales del cliente, verificados)

1. **Bistrosoft – Detalle de Transacciones** (xlsx). Encabezado en la fila 4. Cada ticket: fila "Venta"/"Comanda"/
   "Venta (Multipago)" con el total + líneas "- ITEM" y "- COMBO" (precio = total de la línea), "- COMBO ITEM" en $0,
   descuentos "- ITEM DESCUENTO" y "- ITEM DESCUENTO POR PRODUCTO". Filas "CAJA (...)" no son ventas.
   Se suman las líneas; verificado: 1.432 tickets de octubre 2026 cuadran exactos ($ 29.682.875).
   Fechas como texto "31-10-2026"; artículos con SKU entre paréntesis (se quita).
2. **Bistrosoft – Reporte de Ventas por Día** (xlsx): una fila por día con total, ticket promedio, comandas/facturas,
   cantidades y columnas de medios de pago (bajo el título agrupado "Ventas por medios de pago"). No trae productos:
   se guarda como "Ventas sin detalle de productos" y **solo completa días que no tienen detalle** (el Detalle siempre gana).
   Verificado: noviembre 1–22 = $ 19.549.038.
3. **Bistrosoft – Costo de Mercaderías**: vino **vacío** (el cliente no cargó costos en Bistrosoft). Por eso existe la
   "planilla de costos" (Productos → bajar, completar, subir).
4. **ARCA – Mis Comprobantes Recibidos** (xlsx/csv, formato viejo o nuevo): compras sin IVA ni percepciones
   (Responsable Inscripto, configurable), notas de crédito restan, moneda extranjera × tipo de cambio.
   Cada proveedor (por CUIT) se clasifica: mercadería / gasto fijo (con rubro) / no incluir; sugerencia automática
   por nombre (`lib/compras.ts`), se recuerda. **Probado solo con un archivo inventado**: falta probar con uno real.

Los archivos reales del cliente **no están en el repo** (son datos privados). Pedíselos a la usuaria si hacen falta.

## Estructura

```
app/page.tsx            pantalla principal, navegación (#hoy #cargar #productos #gastos #reportes), carga/deshacer
app/api/datos/route.ts  guardado en Upstash Redis (GET/PUT/DELETE); sin variables de entorno responde 501
middleware.ts           contraseña opcional (APP_USER / APP_PASSWORD)
components/             Hoy, Cargar, Productos, Gastos, Reportes, Grafico, Logo, Iconos
lib/motor.ts            cálculo, medios de pago, aviso de WhatsApp, conclusiones
lib/lector.ts           lectura de Excel/CSV, formatos de Bistrosoft, incorporar ventas por día
lib/compras.ts          compras de ARCA y clasificación de proveedores
lib/reportes.ts         PDF mensual, Excel, planilla de costos
lib/almacen.ts          nube (Redis) o navegador (localStorage) + migraciones de configuración
lib/tipos.ts            tipos y valores iniciales
public/                 íconos, manifest, sw.js (app instalable y uso sin conexión)
```

Correr local: `npm install` y `npm run dev` (o doble clic en `ABRIR_Fiordo.bat` en Windows). Probar: `npx next build`.

## Identidad visual (relevada del Instagram @fiordocafe)

Marrón #762303, crema #FFEEDA/#FFEDD9, terracota #A0553A, beige #E4E0C7, texto #3B1A0C.
Títulos Quicksand, textos Montserrat. Logo: círculo marrón con lobo marino entre olas en crema
(en la app hay una versión simplificada en `components/Logo.tsx`; falta el logo original).
Tono: de "vos", cercano, emojis 🤎☕️✨.

## Pendiente

1. **Base de datos compartida:** conectar Upstash for Redis en Vercel (Storage) y redeploy. Hasta eso, cada
   navegador guarda sus propios datos (la usuaria lo vio en el iPhone: "no hay ventas").
2. **Contraseña:** APP_USER / APP_PASSWORD en Vercel.
3. **Datos reales:** gastos fijos (sobre todo sueldos y cargas), compras de ARCA reales, costos de productos,
   tasas reales de comisión (Mercado Pago u otro).
4. **API de Bistrosoft:** existe (bistrosoft.com/api) pero sin documentación pública; la usuaria iba a pedir acceso
   al equipo comercial (ventas@bistrosoft.com). Cuando llegue: conector + tarea diaria (Vercel Cron) que traiga las
   ventas a Redis.
5. Logo original en alta calidad.
6. `xlsx` 0.18.5 tiene avisos de seguridad conocidos (archivos maliciosos); evaluar pasar a la versión oficial de SheetJS.
