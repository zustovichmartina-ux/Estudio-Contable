"""Lecturas de las pantallas, con cache de Streamlit.

El número de generación sube al escribir, así un guardado se ve al toque.
El TTL cubre lo que escribió otro proceso. `bd` entra en la clave para no
mezclar dos archivos en el mismo proceso (los tests cambian `DB_PATH`).

Las funciones directas viven en cada módulo. Este archivo las envuelve: los
imports de esos módulos quedan adentro de la función cacheada cuando arman
un ciclo con este archivo.
"""
from __future__ import annotations

import streamlit as st

import database

_TTL_MAESTRO = 30
_TTL_CORTO = 5


def _bd() -> str:
    return str(database.DB_PATH)


@st.cache_data(ttl=_TTL_MAESTRO, show_spinner=False)
def clientes(generacion: int, bd: str) -> list[dict]:
    del generacion, bd
    return database._listar_clientes_directo()


@st.cache_data(ttl=_TTL_MAESTRO, show_spinner=False)
def cliente(generacion: int, bd: str, cliente_id: int) -> dict | None:
    del generacion, bd
    return database._obtener_cliente_directo(cliente_id)


@st.cache_data(ttl=_TTL_MAESTRO, show_spinner=False)
def semilla(generacion: int, bd: str, clave: str) -> str | None:
    del generacion, bd
    return database.leer_semilla(clave)


@st.cache_data(ttl=_TTL_MAESTRO, show_spinner=False)
def usuarios(generacion: int, bd: str, solo_activos: bool) -> list[dict]:
    del generacion, bd
    # auth_oficina importa este módulo: el import queda en la llamada.
    import auth_oficina

    return auth_oficina._listar_usuarios_directo(solo_activos)


@st.cache_data(ttl=_TTL_CORTO, show_spinner=False)
def pedido_abierto(generacion: int, bd: str, codigo: str) -> dict | None:
    del generacion, bd
    import rutinas

    return rutinas._pedido_abierto_directo(codigo)


@st.cache_data(ttl=_TTL_CORTO, show_spinner=False)
def ultimo_pedido(generacion: int, bd: str, codigo: str) -> dict | None:
    del generacion, bd
    import rutinas

    return rutinas._ultimo_pedido_directo(codigo)


@st.cache_data(ttl=_TTL_CORTO, show_spinner=False)
def pedidos(generacion: int, bd: str, estado: str, tope: int) -> list[dict]:
    del generacion, bd
    import rutinas

    return rutinas._listar_pedidos_directo(estado or None, tope)


@st.cache_data(ttl=_TTL_CORTO, show_spinner=False)
def emisiones(generacion: int, bd: str, cuit: str, ambiente: str, limite: int) -> list[dict]:
    del generacion, bd
    import arca.persistencia as arca_persistencia

    return arca_persistencia._listar_emisiones_directo(cuit, ambiente, limite)


@st.cache_data(ttl=_TTL_CORTO, show_spinner=False)
def emisores(generacion: int, bd: str) -> list[dict]:
    del generacion, bd
    import arca.persistencia as arca_persistencia

    return arca_persistencia._listar_emisores_directo()


def clientes_de_pantalla() -> list[dict]:
    return clientes(database.generacion_datos(), _bd())


def cliente_de_pantalla(cliente_id: int) -> dict | None:
    return cliente(database.generacion_datos(), _bd(), int(cliente_id))


def semilla_de_pantalla(clave: str) -> str | None:
    return semilla(database.generacion_datos(), _bd(), clave)


def usuarios_de_pantalla(solo_activos: bool) -> list[dict]:
    return usuarios(database.generacion_datos(), _bd(), bool(solo_activos))


def pedido_abierto_de_pantalla(codigo: str) -> dict | None:
    return pedido_abierto(database.generacion_datos(), _bd(), codigo)


def ultimo_pedido_de_pantalla(codigo: str) -> dict | None:
    return ultimo_pedido(database.generacion_datos(), _bd(), codigo)


def pedidos_de_pantalla(estado: str, tope: int) -> list[dict]:
    return pedidos(database.generacion_datos(), _bd(), estado, tope)


def emisiones_de_pantalla(cuit: str, ambiente: str, limite: int) -> list[dict]:
    return emisiones(database.generacion_datos(), _bd(), cuit, ambiente, limite)


def emisores_de_pantalla() -> list[dict]:
    return emisores(database.generacion_datos(), _bd())
