# Facturación electrónica por Web Service (WSFEv1): pasos en ARCA

## Uso desde la web

La solapa **ARCA → Facturación** (estudiocontablemdp.streamlit.app) usa este mismo emisor.
Los certificados no van en el repo: se cargan en Secrets de Streamlit Cloud (ver el README, sección «Facturación ARCA»).
Los CAE, el ticket WSAA (~12 h) y los datos del emisor se guardan en la base del estudio
(`TURSO_DATABASE_URL` / `TURSO_AUTH_TOKEN` en Secrets). Sin esas claves, en Cloud el disco se borra en cada redeploy.

El instructivo de delegaciones y puntos de venta de abajo sigue vigente. El CLI (`emitir.py`) queda para uso local.

---

**Estudio Zona Güemes · Martina Zustovich · CUIT 23-42284343-4** · 01/10/2026

Referencias: ✅ confirmado en documentación oficial de ARCA · ⚠️ sólo en fuentes de terceros (tutoriales de software de facturación) · ❓ no confirmado, revisar en pantalla.
Manuales oficiales descargados en `docs/` (WSASS_manual, WSASS_como_adherirse, wsaa_obtener_certificado_produccion, wsaa_asociar_certificado_a_wsn_produccion, adminrel.delegarws, manual WSFEv1 v4.7, QRespecificaciones).

Archivos que ya están listos en `/workspace/arca_fe/cert/`:
| Archivo | Para qué |
|---|---|
| `estudiozg.key` | Clave privada (permisos 600). **No se sube a ningún lado, no se manda por mail.** Hacer copia de resguardo en un lugar seguro. Sirve para los dos ambientes. |
| `estudiozg-homo.csr` | Pedido de certificado de **homologación** (CN=estudiozg-homo). Se pega en WSASS. |
| `estudiozg-prod.csr` | Pedido de certificado de **producción** (CN=estudiozg). Se sube en Administración de Certificados Digitales. |

Subject usado (formato oficial ✅ WSASS_manual §4.2 y cert-req-howto): `/C=AR/O=Estudio Zona Guemes/CN=estudiozg[-homo]/serialNumber=CUIT 23422843434` ("CUIT", un espacio y los 11 dígitos, sin guiones). Sin "ü" para evitar problemas de codificación.

---
## (a) Homologación (pruebas) vía WSASS
Requisito: clave fiscal **de persona física**, nivel 2 o más ✅ (WSASS no es delegable).

1. **Adherir WSASS** (una sola vez) ✅ (WSASS_como_adherirse):
   arca.gob.ar → Iniciar sesión con CUIT 23422843434 → **Administrador de Relaciones de Clave Fiscal** → **Adherir servicio** → **ARCA › Servicios interactivos** → elegir **WSASS - Autogestión Certificados Homologación** (nombre exacto del ítem ❓) → **Continuar** → salir y volver a entrar. Aparece "WSASS" en Mis Servicios.
2. **Crear el certificado** ✅ (WSASS_manual §5): WSASS → menú **Nuevo Certificado**:
   - Nombre simbólico del DN: `estudiozg-homo`
   - CUIT: viene cargada (23422843434)
   - Solicitud de certificado PKCS#10: pegar **todo** el texto de `estudiozg-homo.csr` (incluidas las líneas BEGIN/END)
   - **Crear DN y Obtener Certificado** → copiar el texto PEM que devuelve y guardarlo como `cert/estudiozg-homo.crt` (Bloc de notas, sin cambios).
3. **Autorizar wsfe** ✅ (WSASS_manual §6): menú **Crear Autorización a Servicio**:
   - Nombre simbólico del DN: `estudiozg-homo`
   - CUIT representado: **23422843434** (su propia CUIT, para la primera prueba). En homologación se pueden agregar autorizaciones con la CUIT de clientes como representado sin que el cliente haga nada ✅ (§3.3 / FAQ 12.3).
   - Servicio: **wsfe** – Facturación Electrónica (texto exacto de la lista ❓) → **Crear Autorización de Acceso**.
4. Probar: `python emitir.py --excel <archivo>.xlsx --env homo` (los CAE de homologación no tienen validez; el PDF sale con marca "HOMOLOGACIÓN").

## (b) Producción
Requisito: clave fiscal nivel 3 ✅.

1. **Administración de Certificados Digitales** ✅ (wsaa_obtener_certificado_produccion):
   - Si no aparece en la lista de servicios: **Administrador de Relaciones de Clave Fiscal** → **Nueva Relación** → **BUSCAR** → elegir "Administración de Certificados Digitales" → Representante **BUSCAR** → CUIT 23422843434 → **Confirmar** → salir y volver a entrar (si sigue sin aparecer, aceptar con el servicio "Aceptación de Designación").
   - Entrar al servicio → elegir el contribuyente (su CUIT) → **Agregar alias** → alias: `estudiozg` → subir el archivo `estudiozg-prod.csr` → **Agregar alias**.
   - Clic en **Ver** (a la derecha del alias) → descargar el certificado → guardarlo como `cert/estudiozg-prod.crt`.
2. **Asociar el certificado (computador fiscal) al servicio** ✅ (wsaa_asociar_certificado_a_wsn_produccion, adminrel.delegarws):
   **Administrador de Relaciones de Clave Fiscal** → representado: su CUIT → **Nueva Relación** → **BUSCAR** servicio → agrupación **ARCA › WebServices** (rótulo "ARCA" ⚠️; antes decía "AFIP") → **Facturación Electrónica** → Representante **BUSCAR** → en "Computador Fiscal" elegir **estudiozg** → **Confirmar** → revisar → **Confirmar** (constancia F3283/E).
   Esto la habilita para facturar *por su propia CUIT*; para cada cliente se repite en el paso (c)-3.

## (c) Delegación de cada cliente a la CUIT 23-42284343-4
Por cliente, **una vez** (y una relación por servicio) ✅ (adminrel.delegarws §2.3):

1. **El cliente** (o su Administrador de Relaciones, con su propia clave fiscal nivel 3):
   arca.gob.ar → **Administrador de Relaciones de Clave Fiscal** → elegir la persona/empresa que factura → **Nueva Relación** → **BUSCAR** servicio → **ARCA › WebServices › Facturación Electrónica** → Representante **BUSCAR** → en "CUIT/CUIL/CDI del Usuario" cargar **23422843434** (no elegir computador fiscal) → **BUSCAR** → **Confirmar** → aparece la advertencia de que el autorizado debe aceptar → **Confirmar** (constancia F3283/E).
   ⚠️ Si sale "Ud. no cuenta con Computadores Fiscales registrados…", se puede ignorar y continuar (tutoriales de terceros).
2. **Martina acepta**: **Administrador de Relaciones** → opción del menú principal para **designaciones pendientes de aceptación** (es la 4ª opción; nombre visible "Aceptación de Designación" ⚠️/❓) → **Aceptar** la de ese cliente.
3. **Martina asigna el computador fiscal para ese cliente** ✅: **Administrador de Relaciones** → en el desplegable elegir **al cliente** como representado → **Nueva Relación** → servicio **Facturación Electrónica** (sólo aparecen los servicios que el cliente delegó) → Representante: **Computador Fiscal `estudiozg`** (en una subdelegación sólo se permite computador fiscal) → **Confirmar** → **Confirmar**.
   Recién ahí el servicio queda operativo para ese cliente ✅.
- Revocación: el cliente entra a Administrador de Relaciones → **Consultar** → lupa en la fila del estudio → **Revocar** ✅.
- Técnica: el ticket (TA) es del certificado del estudio y sirve para todos los clientes; en cada factura el programa manda la CUIT del cliente en `Auth/Cuit`. Si el cliente no completó (c)-1 a (c)-3, ARCA rechaza con error de autorización.

## (d) Punto de venta para Web Services (en cada cliente)
✅ (guía oficial "¿Cómo doy de alta un Controlador Fiscal…?", mismo servicio) / sistema a elegir ⚠️:
1. Con clave fiscal del cliente (o de Martina si el cliente le delegó el servicio "Administración de puntos de venta y domicilios" ❓): **Administración de puntos de venta y domicilios** → elegir la persona → **A/B/M Puntos de venta** → **Agregar**.
2. Datos: Número (el próximo libre, p. ej. 0005), Nombre de fantasía (opcional), **Sistema**:
   - Monotributista: **"Factura Electrónica - Monotributo - Web Services"** ⚠️
   - Responsable Inscripto (y exento): **"RECE para aplicativo y web services"** ⚠️
   - Domicilio: el que corresponda → **Aceptar**.
3. **No usar el punto de venta de Comprobantes en Línea**: el web service sólo autoriza en puntos de venta de Web Services. Anotar el número en la columna "Punto de venta (WS)" de la planilla.
4. FCE MiPyME: usa el mismo punto de venta WS; además el receptor debe estar obligado y el monto superar el mínimo (eso se consulta en *wsfecred*, no incluido acá) ❓.

---
## Uso del emisor (resumen)
```bash
cd /workspace/arca_fe
.venv/bin/python crear_modelo.py                       # regenera el Excel modelo
.venv/bin/python emitir.py --excel Facturas.xlsx --dry-run [--pdf-borrador]   # valida, NO envía
.venv/bin/python emitir.py --dummy --env homo          # estado de servidores ARCA homologación
.venv/bin/python emitir.py --parametros 23422843434 --env homo   # tablas FEParamGet* (requiere .crt)
.venv/bin/python emitir.py --excel Facturas.xlsx --env homo      # emite en homologación
.venv/bin/python emitir.py --excel Facturas.xlsx --env prod --confirmar-produccion   # PRODUCCIÓN
.venv/bin/python -m pytest -q tests                    # pruebas sin certificado
```
- Certificados esperados: `cert/estudiozg-homo.crt` y `cert/estudiozg-prod.crt` (o `--cert`). Clave: `cert/estudiozg.key`.
- El TA se guarda en `cache/homo/` o `cache/prod/` según ambiente (~12 h). **Usar un solo equipo/carpeta de cache por certificado**: si otro proceso pide un TA vigente, ARCA responde `coe.alreadyAuthenticated` ✅.
- Resultados en la hoja **Resultado** del mismo Excel (se agregan filas; la hoja Facturas no se toca). Copia de seguridad del Excel en `salida/backups/` antes de cada envío. PDFs en `salida/pdf/`, log en `salida/log/`.
- Protecciones: filas `EJEMPLO` nunca van a producción; si un comprobante con el mismo contenido ya figura APROBADO en ese ambiente no se reenvía; ante un corte de red se consulta `FECompUltimoAutorizado`/`FECompConsultar` y se marca **VERIFICAR**.
- Condición IVA receptor: siempre se envía (`CondicionIVAReceptorId`, RG 5616). Las fuentes difieren sobre desde cuándo se rechaza sin el dato (Afip SDK: 01/09/2026; prensa sept-2026: 01/12/2026) ⚠️; el programa la exige siempre.
- **Notas de crédito/débito** (NC/ND A 3/2, B 8/7, C 13/12): en "Tipo comprobante" elegir NC x / ND x y completar el comprobante asociado de una de dos formas: (1) "Cbte asociado ID factura" = ID de otra fila del mismo Excel (se emite primero la factura y la NC usa el número que devolvió ARCA) o de una corrida anterior aprobada (hoja Resultado); (2) "Cbte asociado tipo / PtoVta / Nro / fecha" para un comprobante emitido fuera del Excel. Controles: misma letra (ARCA 10040), mismo emisor, NC ≤ total del original menos NC ya emitidas; al enviar se verifica el original en ARCA con FECompConsultar. NC/ND de FCE no están soportadas (requieren código de anulación).
- Consumidor final sin identificar sólo si el total es < $10.000.000 (RG 5824/2026).

## Pendiente de verificar en pantalla
- Texto exacto de los ítems: WSASS en "Servicios interactivos", servicio "wsfe" en WSASS, rótulo "ARCA › WebServices", nombre de la opción de aceptación de designación, nombres de "Sistema" del punto de venta.
- En homologación, si el punto de venta debe existir (normalmente se acepta cualquiera) ❓.
