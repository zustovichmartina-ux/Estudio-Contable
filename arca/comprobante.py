"""Lectura de la planilla, agrupación, validación y armado del FECAEDetRequest."""
import datetime as dt
import hashlib
import json
import re
from decimal import Decimal, ROUND_HALF_UP

from .codigos import (ALICUOTAS, ASOC_PERMITIDOS, CLASE_DOC, COND_IVA_POR_DESC, COND_IVA_RECEPTOR,
                      CONCEPTOS, DOC_TIPOS, ETIQUETA_CBTE, TIPOS_CBTE, TRANSFERENCIA_FCE)

# (clave interna, encabezado visible, ancho)  -> orden de columnas de la hoja Facturas
COLUMNAS = [
    ("marca", "Marca (EJEMPLO = no se emite)", 14),
    ("id", "ID factura (opcional, agrupa ítems)", 14),
    ("cuit_emisor", "CUIT emisor (cliente)", 15),
    ("razon_emisor", "Razón social emisor", 26),
    ("dom_emisor", "Domicilio comercial emisor", 28),
    ("iibb_emisor", "Ingresos Brutos emisor", 14),
    ("inicio_act", "Inicio actividades emisor", 13),
    ("cond_iva_emisor", "Condición IVA emisor", 22),
    ("pto_vta", "Punto de venta (WS)", 10),
    ("tipo", "Tipo comprobante", 11),
    ("concepto", "Concepto", 18),
    ("fecha", "Fecha comprobante", 12),
    ("cond_venta", "Condición de venta", 16),
    ("doc_tipo", "Doc. tipo receptor", 18),
    ("doc_nro", "Doc. nro receptor", 15),
    ("nombre_rec", "Nombre / Razón social receptor", 28),
    ("dom_rec", "Domicilio receptor", 26),
    ("cond_iva", "Condición IVA receptor (RG 5616)", 30),
    ("detalle", "Detalle / descripción", 36),
    ("neto", "Neto gravado (A/B) · Importe (C)", 15),
    ("alicuota", "Alícuota IVA", 16),
    ("iva", "IVA (fórmula)", 13),
    ("total", "Total (fórmula)", 14),
    ("fsd", "Servicio desde", 12),
    ("fsh", "Servicio hasta", 12),
    ("fvto", "Vto. pago", 12),
    ("cbu", "CBU emisor (sólo FCE)", 24),
    ("fce_transf", "Transferencia FCE", 22),
    ("asoc_id", "Cbte asociado ID factura (mismo Excel)", 16),
    ("asoc_tipo", "Cbte asociado tipo", 11),
    ("asoc_pto", "Cbte asociado PtoVta", 10),
    ("asoc_nro", "Cbte asociado Nro", 11),
    ("asoc_fecha", "Cbte asociado fecha", 12),
]
OPCIONALES_PLANILLA = {"asoc_id", "asoc_tipo", "asoc_pto", "asoc_nro", "asoc_fecha"}  # compatibilidad con planillas viejas
CLAVES = [c[0] for c in COLUMNAS]
COND_IVA_EMISOR = ["IVA Responsable Inscripto", "Responsable Monotributo", "IVA Sujeto Exento"]
COND_VENTA = ["Contado", "Cuenta corriente", "Transferencia bancaria", "Tarjeta de débito",
              "Tarjeta de crédito", "Cheque", "Otra"]

UMBRAL_CF_SIN_IDENTIFICAR = Decimal("10000000")  # RG 5824/2026 (antes RG 5700/2025)
D2 = Decimal("0.01")


def r2(x):
    return Decimal(x).quantize(D2, rounding=ROUND_HALF_UP)


def cuit_valido(c):
    c = re.sub(r"\D", "", str(c or ""))
    if len(c) != 11:
        return False
    pesos = [5, 4, 3, 2, 7, 6, 5, 4, 3, 2]
    s = sum(int(a) * b for a, b in zip(c[:10], pesos))
    dv = 11 - s % 11
    dv = 0 if dv == 11 else (9 if dv == 10 else dv)
    return dv == int(c[10])


def solo_digitos(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return re.sub(r"\D", "", str(v))


def a_fecha(v):
    if v in (None, ""):
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = str(v).strip()
    for f in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y%m%d"):
        try:
            return dt.datetime.strptime(s, f).date()
        except ValueError:
            pass
    raise ValueError(f"fecha inválida: {v!r}")


def a_decimal(v):
    if v in (None, ""):
        return None
    if isinstance(v, (int, float)):
        return Decimal(str(v))
    s = str(v).strip().replace("$", "").replace(" ", "")
    if "," in s:  # formato 1.234,56
        s = s.replace(".", "").replace(",", ".")
    return Decimal(s)


def leer_planilla(wb, hoja="Facturas"):
    """Devuelve lista de dicts (una por fila con datos) con '_fila' = nro de fila Excel."""
    ws = wb[hoja]
    fila_enc = None
    for r in range(1, 15):
        vals = [str(ws.cell(r, c).value or "") for c in range(1, len(COLUMNAS) + 1)]
        if any(v.startswith("CUIT emisor") for v in vals):
            fila_enc = r
            break
    if fila_enc is None:
        raise ValueError("No se encontró la fila de encabezados (columna 'CUIT emisor').")
    enc = {str(ws.cell(fila_enc, c).value or "").strip(): c for c in range(1, ws.max_column + 1)}
    mapa = {}
    for k, h, _ in COLUMNAS:
        if h not in enc:
            if k in OPCIONALES_PLANILLA:
                continue
            raise ValueError(f"Falta la columna '{h}' en la hoja {hoja}.")
        mapa[k] = enc[h]
    filas = []
    for r in range(fila_enc + 1, ws.max_row + 1):
        d = {k: None for k in CLAVES}
        d.update({k: ws.cell(r, c).value for k, c in mapa.items()})
        if all(d[k] in (None, "") for k in ("cuit_emisor", "detalle", "neto", "doc_nro")):
            continue
        d["_fila"] = r
        filas.append(d)
    return filas


def agrupar(filas):
    """Agrupa filas con mismo ID factura (y mismo CUIT emisor) en un comprobante."""
    grupos, orden = {}, []
    for f in filas:
        idf = str(f.get("id") or "").strip()
        clave = (solo_digitos(f.get("cuit_emisor")), idf) if idf else ("fila", f["_fila"])
        if clave not in grupos:
            grupos[clave] = []
            orden.append(clave)
        grupos[clave].append(f)
    return [grupos[k] for k in orden]


CAMPOS_CABECERA = ["marca", "cuit_emisor", "pto_vta", "tipo", "concepto", "fecha", "doc_tipo",
                   "doc_nro", "cond_iva", "fsd", "fsh", "fvto", "cbu", "fce_transf",
                   "asoc_id", "asoc_tipo", "asoc_pto", "asoc_nro", "asoc_fecha"]


class Comprobante:
    def __init__(self, filas, hoy=None):
        self.filas = filas
        self.f0 = filas[0]
        self.hoy = hoy or dt.date.today()
        self.errores, self.avisos = [], []
        self.det = None
        self.items = []
        self.cbte_tipo = self.pto_vta = self.cuit = None
        self.clase = None
        self.es_fce = False
        self.clase_doc = "FC"
        self.asoc = None          # dict con la referencia al comprobante asociado (NC/ND)
        self.asoc_desc = None     # texto para PDF / Resultado
        self.es_ejemplo = any(str(f.get("marca") or "").strip().upper() == "EJEMPLO" for f in filas)

    @property
    def filas_txt(self):
        return ",".join(str(f["_fila"]) for f in self.filas)

    def huella(self):
        datos = [{k: str(f.get(k)) for k in CLAVES if k not in ("iva", "total")} for f in self.filas]
        return hashlib.sha256(json.dumps(datos, sort_keys=True).encode()).hexdigest()[:16]

    def validar(self):
        e, f0 = self.errores, self.f0
        # Consistencia de cabecera entre filas agrupadas
        for f in self.filas[1:]:
            for k in CAMPOS_CABECERA:
                if str(f.get(k) or "") != str(f0.get(k) or ""):
                    e.append(f"Fila {f['_fila']}: '{k}' distinto de la fila {f0['_fila']} (mismo ID factura).")
        self.cuit = solo_digitos(f0.get("cuit_emisor"))
        if not cuit_valido(self.cuit):
            e.append(f"CUIT emisor inválido: {f0.get('cuit_emisor')!r}")
        try:
            self.pto_vta = int(solo_digitos(f0.get("pto_vta")) or 0)
            if not 1 <= self.pto_vta <= 99999:
                raise ValueError
        except ValueError:
            e.append("Punto de venta inválido (1 a 99999; debe ser un PV dado de alta para Web Services).")
        tipo = str(f0.get("tipo") or "").strip()
        if tipo not in TIPOS_CBTE:
            e.append(f"Tipo comprobante inválido: {tipo!r} (usar {', '.join(TIPOS_CBTE)}).")
            return False
        self.cbte_tipo, self.clase, self.es_fce = TIPOS_CBTE[tipo]
        self.clase_doc = CLASE_DOC[self.cbte_tipo]
        self._validar_asociado()
        ce = str(f0.get("cond_iva_emisor") or "").strip()
        if ce == "IVA Responsable Inscripto" and self.clase == "C":
            e.append("Emisor Responsable Inscripto no emite comprobantes C (usar A o B).")
        elif ce in ("Responsable Monotributo", "IVA Sujeto Exento") and self.clase in ("A", "B"):
            e.append(f"Emisor {ce} emite comprobantes C, no {self.clase}.")
        elif not ce:
            self.avisos.append("Sin 'Condición IVA emisor' (se infiere de la clase para el PDF).")
        concepto = CONCEPTOS.get(str(f0.get("concepto") or "").strip())
        if not concepto:
            e.append(f"Concepto inválido: {f0.get('concepto')!r}")
        try:
            fecha = a_fecha(f0.get("fecha")) or self.hoy
            fsd, fsh, fvto = (a_fecha(f0.get(k)) for k in ("fsd", "fsh", "fvto"))
        except ValueError as ex:
            e.append(str(ex))
            return False
        # Ventanas de fecha del manual WSFEv1 (CbteFch)
        delta = (fecha - self.hoy).days
        if self.es_fce:
            if not -5 <= delta <= 1:
                e.append(f"Fecha {fecha:%d/%m/%Y}: para FCE debe estar entre 5 días antes y 1 día después de hoy.")
        elif concepto == 1 and abs(delta) > 5:
            e.append(f"Fecha {fecha:%d/%m/%Y}: para Productos debe estar a ±5 días de hoy.")
        elif concepto in (2, 3) and abs(delta) > 10:
            e.append(f"Fecha {fecha:%d/%m/%Y}: para Servicios debe estar a ±10 días de hoy.")
        if concepto in (2, 3):
            if not (fsd and fsh and fvto):
                e.append("Concepto Servicios: completar Servicio desde, Servicio hasta y Vto. pago.")
            elif fsh < fsd:
                e.append("Servicio hasta es anterior a Servicio desde.")
        if fvto and fvto < fecha:
            e.append("Vto. pago debe ser igual o posterior a la fecha del comprobante.")
        if self.es_fce and not fvto:
            e.append("FCE: Vto. pago es obligatorio.")

        # Receptor
        dt_lbl = str(f0.get("doc_tipo") or "").strip()
        doc_tipo = DOC_TIPOS.get(dt_lbl)
        doc_nro = solo_digitos(f0.get("doc_nro"))
        if doc_tipo is None:
            e.append(f"Doc. tipo receptor inválido: {dt_lbl!r}")
        elif doc_tipo in (80, 86):
            if not cuit_valido(doc_nro):
                e.append(f"{dt_lbl} receptor inválido (dígito verificador): {f0.get('doc_nro')!r}")
        elif doc_tipo == 96:
            if not 6 <= len(doc_nro) <= 8:
                e.append(f"DNI receptor inválido: {f0.get('doc_nro')!r}")
        elif doc_tipo == 99:
            doc_nro = "0"
        if self.clase == "A" and doc_tipo != 80:
            e.append("Comprobante A: el receptor debe identificarse con CUIT.")
        if self.es_fce and doc_tipo != 80:
            e.append("FCE: el receptor debe identificarse con CUIT.")
        cond_lbl = str(f0.get("cond_iva") or "").strip()
        cond = COND_IVA_POR_DESC.get(cond_lbl)
        if cond is None:
            e.append(f"Condición IVA receptor obligatoria (RG 5616) o inválida: {cond_lbl!r}")
        elif self.clase not in COND_IVA_RECEPTOR[cond][1]:
            e.append(f"Condición IVA '{cond_lbl}' no admitida para comprobantes clase {self.clase} "
                     f"(error ARCA 10243).")
        if doc_tipo == 99 and cond not in (None, 5):
            self.avisos.append("Receptor sin identificar: lo habitual es Condición 'Consumidor Final'.")

        # Importes
        bases = {}  # id_iva -> [base, iva]
        imp_neto = imp_iva = imp_opex = imp_conc = Decimal(0)
        for f in self.filas:
            try:
                neto = a_decimal(f.get("neto"))
            except Exception:
                e.append(f"Fila {f['_fila']}: importe inválido {f.get('neto')!r}")
                continue
            if neto is None or neto <= 0:
                e.append(f"Fila {f['_fila']}: el importe debe ser mayor a 0.")
                continue
            neto = r2(neto)
            alic_lbl = str(f.get("alicuota") or "").strip()
            if self.clase == "C":
                if alic_lbl not in ("", "No corresponde (C)"):
                    e.append(f"Fila {f['_fila']}: los comprobantes C no discriminan IVA (alícuota debe ser 'No corresponde (C)').")
                imp_neto += neto
                self.items.append({"detalle": f.get("detalle"), "neto": neto, "alic": "", "iva": Decimal(0), "total": neto})
                continue
            if alic_lbl not in ALICUOTAS or ALICUOTAS[alic_lbl][2] == "c":
                e.append(f"Fila {f['_fila']}: alícuota IVA inválida para clase {self.clase}: {alic_lbl!r}")
                continue
            aid, tasa, clase_imp = ALICUOTAS[alic_lbl]
            iva = r2(neto * Decimal(str(tasa)))
            if clase_imp == "gravado":
                b = bases.setdefault(aid, [Decimal(0), Decimal(0)])
                b[0] += neto
                b[1] += iva
                imp_neto += neto
                imp_iva += iva
            elif clase_imp == "exento":
                imp_opex += neto
            else:
                imp_conc += neto
            self.items.append({"detalle": f.get("detalle"), "neto": neto, "alic": alic_lbl, "iva": iva, "total": neto + iva})
            if not f.get("detalle"):
                self.avisos.append(f"Fila {f['_fila']}: sin detalle.")
        total = imp_neto + imp_iva + imp_opex + imp_conc
        if doc_tipo == 99 and total >= UMBRAL_CF_SIN_IDENTIFICAR:
            e.append(f"Total ${total:,.2f} ≥ $10.000.000: hay que identificar al consumidor final (RG 5824/2026).")
        if not str(f0.get("nombre_rec") or "").strip() and doc_tipo != 99:
            self.avisos.append("Sin nombre del receptor (se usa sólo para el PDF).")

        opcionales = []
        if self.es_fce:
            cbu = solo_digitos(f0.get("cbu"))
            if len(cbu) != 22:
                e.append("FCE: CBU emisor obligatorio de 22 dígitos (opcional 2101).")
            transf = TRANSFERENCIA_FCE.get(str(f0.get("fce_transf") or "").strip())
            if not transf:
                e.append("FCE: elegir Transferencia SCA o ADC (opcional 27).")
            opcionales = [{"Id": "2101", "Valor": cbu}, {"Id": "27", "Valor": transf}]
            self.avisos.append("FCE: verificar que el receptor esté obligado y que el monto supere el mínimo "
                               "(consultable en wsfecred, no en wsfev1).")
        elif f0.get("cbu") or f0.get("fce_transf"):
            self.avisos.append("CBU/Transferencia se ignoran: sólo aplican a FCE.")

        self.det = {
            "Concepto": concepto, "DocTipo": doc_tipo, "DocNro": int(doc_nro or 0),
            "CbteDesde": None, "CbteHasta": None, "CbteFch": fecha.strftime("%Y%m%d"),
            "ImpTotal": float(total), "ImpTotConc": float(imp_conc), "ImpNeto": float(imp_neto),
            "ImpOpEx": float(imp_opex), "ImpTrib": 0.0, "ImpIVA": float(imp_iva),
            "FchServDesde": fsd.strftime("%Y%m%d") if concepto in (2, 3) and fsd else None,
            "FchServHasta": fsh.strftime("%Y%m%d") if concepto in (2, 3) and fsh else None,
            "FchVtoPago": fvto.strftime("%Y%m%d") if (concepto in (2, 3) or self.es_fce) and fvto else None,
            "MonId": "PES", "MonCotiz": 1, "CondicionIVAReceptorId": cond,
            "Iva": [{"Id": k, "BaseImp": float(v[0]), "Importe": float(v[1])} for k, v in sorted(bases.items())]
                   if self.clase != "C" else None,
            "Opcionales": opcionales or None,
        }
        if self.clase in ("A", "B") and imp_neto == 0 and not bases:
            self.det["Iva"] = None
        self.fecha = fecha
        if self.asoc and "tipo" in self.asoc:
            self.verificar_asociado()
        return not e

    def _validar_asociado(self):
        f0, e = self.f0, self.errores
        aid = str(f0.get("asoc_id") or "").strip()
        atipo = str(f0.get("asoc_tipo") or "").strip()
        hay_ext = any(f0.get(k) not in (None, "") for k in ("asoc_tipo", "asoc_pto", "asoc_nro"))
        if self.clase_doc == "FC":
            if aid or hay_ext:
                e.append("Una factura no lleva comprobante asociado (sólo NC/ND).")
            return
        if aid and hay_ext:
            e.append("Comprobante asociado: usar el ID factura del Excel O tipo/PtoVta/Nro, no ambos.")
            return
        if aid:
            self.asoc = {"id": aid}
            return
        if not hay_ext:
            e.append(f"{self.clase_doc} sin comprobante asociado: completar 'Cbte asociado ID factura' "
                     "o 'Cbte asociado tipo/PtoVta/Nro' (validación ARCA 10197).")
            return
        if atipo not in TIPOS_CBTE:
            e.append(f"Cbte asociado tipo inválido: {atipo!r}")
            return
        try:
            pto = int(solo_digitos(f0.get("asoc_pto")) or 0)
            nro = int(solo_digitos(f0.get("asoc_nro")) or 0)
            fch = a_fecha(f0.get("asoc_fecha"))
        except ValueError as ex:
            e.append(f"Cbte asociado: {ex}")
            return
        if not 0 < pto < 99999:
            e.append("Cbte asociado PtoVta inválido (ARCA 10058).")
        if not 0 < nro < 99999999:
            e.append("Cbte asociado Nro inválido (ARCA 10059).")
        self.asoc = {"tipo": TIPOS_CBTE[atipo][0], "pto": pto, "nro": nro, "fecha": fch,
                     "cuit": None, "total": None}

    def verificar_asociado(self, total_original=None, ya_acreditado=Decimal(0)):
        """Reglas de negocio sobre self.asoc ya resuelto (tipo/pto/nro conocidos)."""
        a, e = self.asoc, []
        if not a or "tipo" not in a:
            return True
        if a["tipo"] not in ASOC_PERMITIDOS.get(self.cbte_tipo, set()):
            e.append(f"No se puede asociar un {ETIQUETA_CBTE.get(a['tipo'], a['tipo'])} a un "
                     f"{ETIQUETA_CBTE[self.cbte_tipo]}: debe ser de la misma letra (ARCA 10040).")
        if a.get("cuit") and str(a["cuit"]) != self.cuit:
            e.append("El comprobante asociado debe ser del mismo CUIT emisor.")
        if total_original is not None and self.det and self.clase_doc == "NC":
            disponible = Decimal(str(total_original)) - ya_acreditado
            if Decimal(str(self.det["ImpTotal"])) > disponible:
                e.append(f"La NC (${self.det['ImpTotal']:,.2f}) supera el importe disponible del comprobante "
                         f"asociado (${disponible:,.2f} = total ${float(total_original):,.2f} menos NC previas).")
        if a.get("fecha") and getattr(self, "fecha", None) and a["fecha"] > self.fecha:
            e.append("La fecha del comprobante asociado no puede ser posterior a la de la NC/ND.")
        nro_txt = f"{a['nro']:08d}" if a.get("nro") else "(nro. a asignar)"
        et = ETIQUETA_CBTE.get(a["tipo"], str(a["tipo"]))
        nombre = {"NC": "Nota de Crédito", "ND": "Nota de Débito"}.get(et[:2], "Factura")
        letra = et[-1] if et[-1] in "ABC" else ""
        self.asoc_desc = (f"{nombre} {'FCE ' if et.startswith('FCE') else ''}{letra} (cód. {a['tipo']}) "
                          f"{a['pto']:05d}-{nro_txt}" + (f" del {a['fecha']:%d/%m/%Y}" if a.get("fecha") else ""))
        if self.det is not None:
            ca = {"Tipo": a["tipo"], "PtoVta": a["pto"], "Nro": a.get("nro") or 0, "Cuit": self.cuit}
            if a.get("fecha"):
                ca["CbteFch"] = a["fecha"].strftime("%Y%m%d")
            self.det["CbtesAsoc"] = [ca]
        for x in e:
            if x not in self.errores:
                self.errores.append(x)
        return not e

    def clave_asoc(self):
        a = self.asoc or {}
        return f"{a.get('tipo')}-{a.get('pto')}-{a.get('nro')}" if a.get("nro") else None

    def con_numero(self, nro):
        d = dict(self.det)
        d["CbteDesde"] = d["CbteHasta"] = nro
        return d
