# Fiordo · Punto de equilibrio

Aplicación web para seguir, día a día, si Fiordo Café vende lo suficiente para cubrir sus gastos.
Se usa desde la compu y desde el celular. Las ventas se cargan desde el Excel que exporta Bistrosoft.

## Qué hace

| Pestaña | Para qué sirve |
|---|---|
| **Hoy** | Semáforo del último día, avance del mes contra el punto de equilibrio, proyección, margen de seguridad, ventas por día y por día de la semana. |
| **Cargar** | Subir el Excel de ventas. Reconoce directo el **Detalle de Transacciones** de Bistrosoft; para otros archivos detecta la fila de títulos y las columnas, y se pueden corregir antes de guardar. Si un día ya estaba cargado, se reemplaza (no se duplica). |
| **Productos** | Margen y aporte de cada producto. Costo de receta uno por uno, o todos juntos: **Bajar planilla de costos** → completar en Excel → **Subir costos** (acepta también el reporte "Costo de Mercaderías" de Bistrosoft). |
| **Gastos** | Gastos fijos del mes, comisiones, ingresos brutos y días que cierra el local. |
| **Reportes** | Informe mensual en **PDF**, detalle en **Excel** y aviso listo para mandar por **WhatsApp**. |

Punto de equilibrio = gastos fijos ÷ margen de contribución, donde margen de contribución = 1 − (mercadería + comisiones + ingresos brutos).

## Usarla en la compu (Windows)

1. Instalar [Node.js](https://nodejs.org) (versión LTS), una sola vez.
2. Doble clic en **`ABRIR_Fiordo.bat`**. La primera vez instala lo necesario.
3. Se abre Chrome en `http://localhost:3000`.
4. Desde el celular, conectado al mismo WiFi, entrar a la dirección `http://192.168.x.x:3000` que muestra la ventana negra.

Con la terminal: `npm install` y después `npm run dev`.

## Publicarla en Vercel (para usarla desde cualquier lugar)

1. En [vercel.com](https://vercel.com) → **Add New → Project** → importar el repositorio de GitHub.
2. En **Root Directory** elegir `fiordo-app`. Deploy.
3. **Contraseña** (recomendado, son datos del negocio): en *Settings → Environment Variables* agregar
   `APP_USER` y `APP_PASSWORD`. Después volver a desplegar.
4. **Mismos datos en el celular y en la compu**: en el proyecto de Vercel → *Storage* → conectar
   **Upstash Redis** (plan gratis). Vercel agrega solo `KV_REST_API_URL` y `KV_REST_API_TOKEN`.
   Volver a desplegar.

Sin el paso 4, cada navegador guarda sus propios datos (lo que cargues en la compu no aparece en el celular).

## Cómo sacar el Excel de Bistrosoft

En Bistro Web (`webaccess.bistrosoft.com`) exportar el reporte **Detalle de Transacciones** del mes a Excel
(archivo `DetalleDeTransacciones_FIORDO_CAFÉ_-_....xlsx`) y subirlo en la pestaña **Cargar**.

Cómo se lee ese reporte:
- Cada ticket tiene una fila de encabezado (`Venta`, `Comanda`, `Venta (Multipago)`) con el total, y debajo sus líneas.
- Se suman las líneas `- ITEM` y `- COMBO` (el precio ya es el total de la línea) y los descuentos
  (`- ITEM DESCUENTO`, `- ITEM DESCUENTO POR PRODUCTO`). Los `- COMBO ITEM` van en $0 y se ignoran.
- Los movimientos de caja (apertura, retiro, depósito, cierre, ajustes) no son ventas y se ignoran.
- Probado con octubre 2026: 1.432 tickets, todos cuadran contra su encabezado.

En `ejemplo/ventas_ejemplo_noviembre.xlsx` hay un archivo inventado de formato genérico para probar.

## Cambiar el estilo de la marca

Usa la identidad de Fiordo relevada del Instagram (marrón #762303, crema #FFEEDA, terracota #A0553A,
Quicksand para títulos y Montserrat para textos). El logo es una versión simplificada.

- Colores y tipografías: tokens de `:root` en `app/globals.css` (y el bloque de modo oscuro).
- Colores del PDF: `lib/tema.ts`.
- Logo: cuando esté el original, guardarlo en `public/` y usarlo en `components/Logo.tsx` y `public/icon.svg`.

## Estructura

```
app/page.tsx          pantalla principal y navegación
app/api/datos         guardado en la base (Upstash Redis) cuando está conectada
components/           pestañas (Hoy, Cargar, Productos, Gastos, Reportes) y gráfico
lib/motor.ts          cálculo del punto de equilibrio y textos del aviso
lib/lector.ts         lectura de Excel/CSV y detección de columnas
lib/reportes.ts       PDF y Excel
lib/almacen.ts        guardado (nube o navegador)
middleware.ts         contraseña de acceso
```
