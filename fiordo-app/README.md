# Fiordo · Punto de equilibrio

Aplicación web para seguir, día a día, si Fiordo Café vende lo suficiente para cubrir sus gastos.
Se usa desde la compu y desde el celular. Las ventas se cargan desde el Excel que exporta Bistrosoft.

## Qué hace

| Pestaña | Para qué sirve |
|---|---|
| **Hoy** | Semáforo del último día, avance del mes contra el punto de equilibrio, proyección, margen de seguridad, ventas por día y por día de la semana. |
| **Cargar** | Subir el Excel/CSV de ventas. Detecta solo la fila de títulos y las columnas (fecha, artículo, cantidad, importe, medio de pago, ticket); se pueden corregir antes de guardar. Si un día ya estaba cargado, se reemplaza (no se duplica). |
| **Productos** | Margen y aporte de cada producto. Acá se carga el **costo de receta** de cada uno. |
| **Gastos** | Gastos fijos del mes, comisiones, ingresos brutos y días de apertura. |
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

En Bistro Web (`webaccess.bistrosoft.com`) → **Reportes / Ventas** → reporte de **ventas por artículo** o
**detalle de ventas** → rango de fechas del mes → **Exportar a Excel**. Subirlo en la pestaña **Cargar**.

Lo ideal es una fila por artículo vendido con fecha, cantidad e importe (si trae medio de pago y número
de ticket, también se usan). En `ejemplo/ventas_ejemplo_noviembre.xlsx` hay un archivo inventado con ese formato.

## Cambiar el estilo de la marca

El estilo actual es **provisorio**. Para adaptarlo a Fiordo Café:

- Colores y tipografías: tokens de `:root` en `app/globals.css` (y el bloque de modo oscuro).
- Colores del PDF: `lib/tema.ts`.
- Logo: reemplazar `public/icon.svg`.

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
