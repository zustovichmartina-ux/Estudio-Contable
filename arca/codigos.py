"""Tablas de códigos ARCA (manual WSFEv1 v4.7, 01/09/2026, y métodos FEParamGet*).

Verificar periódicamente contra FEParamGetTiposCbte / FEParamGetTiposDoc /
FEParamGetTiposIva / FEParamGetCondicionIvaReceptor (se pueden consultar con
`emitir.py --parametros`)."""

URLS = {
    "homo": {
        "wsaa": "https://wsaahomo.afip.gov.ar/ws/services/LoginCms",
        "wsfe": "https://wswhomo.afip.gov.ar/wsfev1/service.asmx",
    },
    "prod": {
        "wsaa": "https://wsaa.afip.gov.ar/ws/services/LoginCms",
        "wsfe": "https://servicios1.afip.gov.ar/wsfev1/service.asmx",
    },
}

# Etiqueta de la planilla -> (código ARCA, clase, es_fce)
TIPOS_CBTE = {
    "A": (1, "A", False),
    "B": (6, "B", False),
    "C": (11, "C", False),
    "FCE A": (201, "A", True),
    "FCE B": (206, "B", True),
    "FCE C": (211, "C", True),
    "NC A": (3, "A", False),
    "ND A": (2, "A", False),
    "NC B": (8, "B", False),
    "ND B": (7, "B", False),
    "NC C": (13, "C", False),
    "ND C": (12, "C", False),
}
ETIQUETA_CBTE = {v[0]: k for k, v in TIPOS_CBTE.items()}
# Clase de documento: FC factura, NC nota de crédito, ND nota de débito
CLASE_DOC = {1: "FC", 6: "FC", 11: "FC", 201: "FC", 206: "FC", 211: "FC",
             3: "NC", 8: "NC", 13: "NC", 2: "ND", 7: "ND", 12: "ND"}
# Comprobantes que se pueden asociar (manual WSFEv1 v4.7, validación 10040)
ASOC_PERMITIDOS = {
    2: {1, 2, 3, 4, 5, 34, 39, 60, 63, 88, 991}, 3: {1, 2, 3, 4, 5, 34, 39, 60, 63, 88, 991},
    7: {6, 7, 8, 9, 10, 35, 40, 61, 64, 88, 991}, 8: {6, 7, 8, 9, 10, 35, 40, 61, 64, 88, 991},
    12: {11, 12, 13, 15}, 13: {11, 12, 13, 15},
}
NOMBRE_CBTE = {1: "FACTURA", 6: "FACTURA", 11: "FACTURA",
               201: "FACTURA DE CRÉDITO ELECTRÓNICA MiPyMEs (FCE)",
               206: "FACTURA DE CRÉDITO ELECTRÓNICA MiPyMEs (FCE)",
               211: "FACTURA DE CRÉDITO ELECTRÓNICA MiPyMEs (FCE)",
               3: "NOTA DE CRÉDITO", 8: "NOTA DE CRÉDITO", 13: "NOTA DE CRÉDITO",
               2: "NOTA DE DÉBITO", 7: "NOTA DE DÉBITO", 12: "NOTA DE DÉBITO"}

CONCEPTOS = {"Productos": 1, "Servicios": 2, "Productos y Servicios": 3}

# Etiqueta -> código FEParamGetTiposDoc
DOC_TIPOS = {
    "CUIT": 80,
    "CUIL": 86,
    "CDI": 87,
    "DNI": 96,
    "Pasaporte": 94,
    "Consumidor Final (sin identificar)": 99,
}
NOMBRE_DOC = {v: k for k, v in DOC_TIPOS.items()}
NOMBRE_DOC[99] = "Sin identificar"

# Condición frente al IVA del receptor (RG 5616) - Anexo manual v4.7
# código: (descripción, clases permitidas)
COND_IVA_RECEPTOR = {
    1: ("IVA Responsable Inscripto", {"A", "C"}),
    4: ("IVA Sujeto Exento", {"B", "C"}),
    5: ("Consumidor Final", {"B", "C"}),
    6: ("Responsable Monotributo", {"A", "C"}),
    7: ("Sujeto No Categorizado", {"B", "C"}),
    8: ("Proveedor del Exterior", {"B", "C"}),
    9: ("Cliente del Exterior", {"B", "C"}),
    10: ("IVA Liberado – Ley N° 19.640", {"B", "C"}),
    13: ("Monotributista Social", {"A", "C"}),
    15: ("IVA No Alcanzado", {"B", "C"}),
    16: ("Monotributo Trabajador Independiente Promovido", {"A", "C"}),
}
COND_IVA_POR_DESC = {d: k for k, (d, _) in COND_IVA_RECEPTOR.items()}

# Alícuotas: etiqueta -> (id AlicIva, tasa, tipo)
#  tipo: "gravado" (va a ImpNeto + array Iva), "exento" (ImpOpEx), "no_gravado" (ImpTotConc), "c" (comprobante C)
ALICUOTAS = {
    "21%": (5, 0.21, "gravado"),
    "10,5%": (4, 0.105, "gravado"),
    "27%": (6, 0.27, "gravado"),
    "0%": (3, 0.0, "gravado"),
    "Exento": (None, 0.0, "exento"),
    "No gravado": (None, 0.0, "no_gravado"),
    "No corresponde (C)": (None, 0.0, "c"),
}

TRANSFERENCIA_FCE = {"SCA - Transferencia al Sistema de Circulación Abierta": "SCA",
                     "ADC - Agente de Depósito Colectivo": "ADC"}

QR_URL = "https://www.arca.gob.ar/fe/qr/"
