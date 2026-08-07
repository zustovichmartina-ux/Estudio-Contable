#!/usr/bin/env python
"""
Smoke funcional del Estudio Contable — offline, no toca la sesión Streamlit.

Qué controla:
  1) La web responde (GET :8501) — solo lectura
  2) Import de app/procesador
  3) Extracción IVA (balance Oftalmología)
  4) Extracción IIBB línea a línea + partida doble
  5) NC no se filtran como ruido; pie PAYWAY no entra al asiento

Uso:
  python smoke_estudio.py
  powershell -File correr_smoke_estudio.ps1

Resultado: logs/smoke_ultimo.json (+ exit code 0/1).
"""
from __future__ import annotations

import json
import sys
import traceback
import urllib.error
import urllib.request
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

LOGS = ROOT / "logs"
RESULT_JSON = LOGS / "smoke_ultimo.json"
RESULT_TXT = LOGS / "smoke_ultimo.txt"
BALANCE = ROOT / "Copia de OFTALMOLOGIA RELE Balance 2026.xlsx"
URL_WEB = "http://127.0.0.1:8501/"


@dataclass
class Check:
    nombre: str
    ok: bool
    detalle: str = ""


@dataclass
class SmokeReport:
    ok: bool
    fecha: str
    checks: list[Check] = field(default_factory=list)

    def add(self, nombre: str, ok: bool, detalle: str = "") -> None:
        self.checks.append(Check(nombre, ok, detalle))
        if not ok:
            self.ok = False


def _check_web(report: SmokeReport) -> None:
    try:
        with urllib.request.urlopen(URL_WEB, timeout=8) as resp:
            code = getattr(resp, "status", None) or resp.getcode()
        report.add("web_http", 200 <= int(code) < 400, f"status={code}")
    except (urllib.error.URLError, TimeoutError, OSError) as exc:
        report.add("web_http", False, f"no responde: {exc}")


def _check_imports(report: SmokeReport) -> None:
    try:
        import procesador as proc  # noqa: F401
        report.add("import_procesador", True)
    except Exception as exc:
        report.add("import_procesador", False, f"{type(exc).__name__}: {exc}")
        return
    try:
        import py_compile

        py_compile.compile(str(ROOT / "app.py"), doraise=True)
        report.add("compile_app", True, "app.py sin errores de sintaxis")
    except Exception as exc:
        report.add("compile_app", False, f"{type(exc).__name__}: {exc}")


def _check_nc_ruido(report: SmokeReport) -> None:
    import procesador as proc

    casos_ok = (
        "Nota de crédito compras",
        "Nota de credito ventas 21%",
        "NC compras",
    )
    for c in casos_ok:
        if proc._es_fila_ruido_balance(c):
            report.add("nc_no_ruido", False, f"filtró como ruido: {c}")
            return
    if not proc._es_fila_ruido_balance("Nota: ver detalle"):
        report.add("nc_no_ruido", False, "nota de pie debería ser ruido")
        return
    if not proc._es_fila_ruido_balance("PAYWAY 06/2024"):
        report.add("nc_no_ruido", False, "PAYWAY debería ser ruido")
        return
    tipo_nc = proc.inferir_tipo_movimiento_desde_concepto(
        "Nota de crédito compras", monto_original=-100.0, impuesto="IVA",
    )
    if tipo_nc != "Debe":
        report.add("nc_no_ruido", False, f"NC compras tipo={tipo_nc} (esperado Debe)")
        return
    report.add("nc_no_ruido", True, "NC proyectables; pie/PAYWAY filtrados")


def _check_iva_extract(report: SmokeReport) -> None:
    import procesador as proc

    if not BALANCE.is_file():
        report.add("iva_extract", False, f"falta balance: {BALANCE.name}")
        return
    r = proc.extraer_filas_universales_balance_por_periodo_con_errores(
        str(BALANCE), "IVA", "04/2025",
    )
    if r.error:
        report.add("iva_extract", False, f"{r.error_tipo}: {r.error}")
        return
    if len(r.filas or []) < 3:
        report.add("iva_extract", False, f"pocas filas: {len(r.filas or [])}")
        return
    report.add("iva_extract", True, f"{len(r.filas)} filas en 04/2025")


def _check_iibb_lineas(report: SmokeReport) -> None:
    import unicodedata

    import procesador as proc

    def _norm(f: dict) -> str:
        s = str(f.get("descripcion") or f.get("concepto_raw") or "").lower()
        return "".join(
            c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn"
        )

    if not BALANCE.is_file():
        report.add("iibb_lineas", False, f"falta balance: {BALANCE.name}")
        return

    r = proc.extraer_filas_universales_balance_por_periodo_con_errores(
        str(BALANCE), "IIBB", "04/2025",
    )
    if r.error:
        report.add("iibb_lineas", False, f"{r.error_tipo}: {r.error}")
        return
    filas = r.filas or []
    if len(filas) < 5:
        report.add("iibb_lineas", False, f"esperaba >=5 lineas, hay {len(filas)}")
        return

    if any("payway" in _norm(f) for f in filas):
        report.add("iibb_lineas", False, "PAYWAY entro al asiento 04/2025")
        return
    if any("oftalmolog" in _norm(f) or "devengamiento de" in _norm(f) for f in filas):
        report.add("iibb_lineas", False, "entro titulo/razon social al asiento")
        return

    total_d = round(sum(float(f.get("debe") or 0) for f in filas), 2)
    total_h = round(sum(float(f.get("haber") or 0) for f in filas), 2)
    if abs(total_d - total_h) > 0.05:
        report.add(
            "iibb_lineas",
            False,
            f"no balancea Debe={total_d} Haber={total_h} dif={round(total_d - total_h, 2)}",
        )
        return

    ret = next(
        (
            f for f in filas
            if "retencion" in _norm(f) and "bancar" not in _norm(f)
        ),
        None,
    )
    if ret and str(ret.get("codigo")) == "11450":
        report.add("iibb_lineas", False, "retenciones tomo 11450 en vez de 11418")
        return

    # Conceptos del Excel aunque esten en $0
    tiene_saldo = any("saldo a favor" in _norm(f) for f in filas)
    tiene_redondeo = any("redondeo" in _norm(f) for f in filas)
    tiene_impuesto = any("impuesto" in _norm(f) for f in filas)
    tiene_ret = any("retencion" in _norm(f) for f in filas)
    tiene_perc = any("percep" in _norm(f) for f in filas)
    tiene_pagar = any("pagar" in _norm(f) for f in filas)
    if not (tiene_impuesto and tiene_ret and tiene_perc and tiene_pagar and tiene_saldo and tiene_redondeo):
        report.add(
            "iibb_lineas",
            False,
            "faltan conceptos del Excel (impuesto/ret/perc/pagar/saldo/redondeo)",
        )
        return

    r6 = proc.extraer_filas_universales_balance_por_periodo_con_errores(
        str(BALANCE), "IIBB", "06/2025",
    )
    if r6.error:
        report.add("iibb_lineas", False, f"06/2025 {r6.error_tipo}: {r6.error}")
        return
    if any("payway" in _norm(f) for f in (r6.filas or [])):
        report.add("iibb_lineas", False, "ruido PAYWAY en 06/2025")
        return

    report.add(
        "iibb_lineas",
        True,
        f"04/2025 {len(filas)} lineas (con saldo/redondeo); 06/2025 sin PAYWAY",
    )


def run() -> SmokeReport:
    report = SmokeReport(ok=True, fecha=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    _check_web(report)
    _check_imports(report)
    if any(c.nombre == "import_procesador" and c.ok for c in report.checks):
        try:
            _check_nc_ruido(report)
        except Exception as exc:
            report.add("nc_no_ruido", False, f"{type(exc).__name__}: {exc}")
        try:
            _check_iva_extract(report)
        except Exception as exc:
            report.add("iva_extract", False, traceback.format_exc(limit=3))
        try:
            _check_iibb_lineas(report)
        except Exception as exc:
            report.add("iibb_lineas", False, traceback.format_exc(limit=3))
    return report


def _persist(report: SmokeReport) -> None:
    LOGS.mkdir(parents=True, exist_ok=True)
    payload = {
        "ok": report.ok,
        "fecha": report.fecha,
        "checks": [asdict(c) for c in report.checks],
    }
    RESULT_JSON.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    lines = [
        f"SMOKE {'PASS' if report.ok else 'FAIL'}  {report.fecha}",
        "-" * 48,
    ]
    for c in report.checks:
        mark = "OK " if c.ok else "FAIL"
        extra = f" — {c.detalle}" if c.detalle else ""
        lines.append(f"[{mark}] {c.nombre}{extra}")
    RESULT_TXT.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> int:
    report = run()
    _persist(report)
    print(RESULT_TXT.read_text(encoding="utf-8"), end="")
    return 0 if report.ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
