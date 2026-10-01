#!/usr/bin/env python3
"""Cola de rutinas para el asistente externo.

Usa la misma base que la web (database.obtener_conexion):
Turso si existen TURSO_DATABASE_URL y TURSO_AUTH_TOKEN; si no, el SQLite local.

Ejemplos (desde la raíz del repo):

    python scripts/cola_rutinas.py listar --estado PENDIENTE
    python scripts/cola_rutinas.py tomar 12
    python scripts/cola_rutinas.py terminar 12 --estado OK --resultado "Listo" --archivos "C:\\ruta\\salida.xlsx"
    python scripts/cola_rutinas.py terminar 12 --estado OK --resultado "Listo" --preview-json preview.json
    python scripts/cola_rutinas.py ficha --json
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from rutinas import (  # noqa: E402
    ESTADOS,
    ESTADOS_CIERRE,
    ErrorRutina,
    exportar_ficha,
    listar_pedidos,
    normalizar_preview,
    tomar_pedido,
    terminar_pedido,
)


def _imprimir(payload: object, *, ok: bool = True) -> int:
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0 if ok else 1


def _pedido_json(pedido: dict) -> dict:
    return {
        "id": pedido["id"],
        "rutina": pedido["rutina"],
        "nombre": pedido.get("nombre") or pedido["rutina"],
        "parametros": pedido.get("parametros"),
        "solicitado_por": pedido.get("solicitado_por"),
        "creado_en": pedido.get("creado_en"),
        "estado": pedido.get("estado"),
        "tomado_en": pedido.get("tomado_en"),
        "terminado_en": pedido.get("terminado_en"),
        "resultado": pedido.get("resultado"),
        "archivos": pedido.get("archivos"),
        "preview": pedido.get("preview"),
    }


def _leer_preview(ruta: str) -> dict:
    path = Path(ruta)
    if not path.is_file():
        raise ErrorRutina(f"No está el archivo de preview: {path}")
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ErrorRutina(f"El preview no es JSON válido: {exc}") from exc
    return normalizar_preview(data)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Cola de rutinas del estudio. La web encola; este script toma y cierra."
    )
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_ficha = sub.add_parser("ficha", help="Exporta la ficha de clientes de Proyecciones")
    p_ficha.add_argument("--json", action="store_true", required=True, help="Imprime la ficha en JSON")

    p_listar = sub.add_parser("listar", help="Lista pedidos en JSON")
    p_listar.add_argument("--estado", required=True, choices=ESTADOS)
    p_listar.add_argument("--limite", type=int, default=200)

    p_tomar = sub.add_parser("tomar", help="Pasa un PENDIENTE a EN_CURSO")
    p_tomar.add_argument("id", type=int)

    p_terminar = sub.add_parser("terminar", help="Cierra un EN_CURSO en OK o ERROR")
    p_terminar.add_argument("id", type=int)
    p_terminar.add_argument("--estado", required=True, choices=ESTADOS_CIERRE)
    p_terminar.add_argument("--resultado", required=True)
    p_terminar.add_argument("--archivos", default=None)
    p_terminar.add_argument(
        "--preview-json",
        default=None,
        help="Archivo JSON {resumen_md, tablas:[{titulo, columnas, filas}], archivos:[rutas]}",
    )

    args = parser.parse_args(argv)
    try:
        if args.cmd == "ficha":
            return _imprimir(exportar_ficha())
        if args.cmd == "listar":
            pedidos = listar_pedidos(estado=args.estado, limite=args.limite)
            return _imprimir([_pedido_json(p) for p in pedidos])
        if args.cmd == "tomar":
            pedido = tomar_pedido(args.id)
            return _imprimir({"ok": True, "pedido": _pedido_json(pedido)})
        preview = _leer_preview(args.preview_json) if args.preview_json else None
        pedido = terminar_pedido(
            args.id, args.estado, args.resultado, args.archivos, preview=preview
        )
        return _imprimir({"ok": True, "pedido": _pedido_json(pedido)})
    except ErrorRutina as exc:
        return _imprimir({"ok": False, "error": str(exc)}, ok=False)


if __name__ == "__main__":
    sys.exit(main())
