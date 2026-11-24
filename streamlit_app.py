"""Interfaz de generación de cargues AGP. Ejecutar con Streamlit."""

from __future__ import annotations

import hashlib
import logging

import streamlit as st

from cargue_agp import (
    ENCABEZADOS, MAX_ARCHIVO_MB, OPCIONES_BT, OPCIONES_MT, ErrorCargue,
    analizar_excel, cargar_catalogo, construir_ubicacion, filas_cargue,
    generar_excel, normalizar_tam,
)

VERSION = "1.0.0"
ESTILOS = """
<style>
.block-container {max-width:1180px;padding-top:2.5rem;padding-bottom:2.5rem;}
[data-testid="stSidebar"] {background:#eaf0f2;}
.agp-hero {background:#142f43;border-radius:18px;border-bottom:5px solid #4ec4a5;
  padding:30px 34px;margin-bottom:26px;color:#fff;}
.agp-hero .eyebrow {color:#8de0c8;font-size:12px;letter-spacing:2px;font-weight:700;}
.agp-hero h1 {color:#fff;font-size:36px;letter-spacing:-.8px;line-height:1.15;
  margin:10px 0 14px;padding:0;}
.agp-hero p {color:#dbe7ee;margin:0;max-width:750px;font-size:16px;}
.agp-section {font-size:20px;font-weight:650;margin:0 0 5px;color:#182f43;}
.agp-section span {display:inline-flex;background:#e5f4ee;color:#12675b;
  width:28px;height:28px;align-items:center;justify-content:center;border-radius:8px;
  font-size:14px;margin-right:9px;}
.agp-empty {padding:28px 22px;border:1px dashed #acbdc5;border-radius:12px;
  text-align:center;background:#f0f5f5;color:#526976;margin-top:12px;}
[data-testid="stMetric"] {background:#fff;border:1px solid #dfe7eb;
  padding:16px 18px;border-radius:12px;}
.agp-footer {border-top:1px solid #d4dfe4;margin-top:32px;padding-top:14px;
  color:#576d7a;font-size:12px;}
@media(max-width:640px) {
 .agp-hero {padding:24px 20px;} .agp-hero h1 {font-size:28px;}
 .block-container {padding-top:3.5rem;}
}
</style>
"""


def limpiar_resultado() -> None:
    st.session_state.pop("resultado", None)
    st.session_state.pop("error_generacion", None)


def cambiar_archivo() -> None:
    limpiar_resultado()
    st.session_state.pop("lectura", None)


def mayusculas(campo: str) -> None:
    st.session_state[campo] = st.session_state[campo].strip().upper()
    limpiar_resultado()


def cambiar_nivel() -> None:
    st.session_state.pop("sufijo", None)
    limpiar_resultado()


def seccion(numero: int, titulo: str) -> None:
    st.markdown(f'<div class="agp-section"><span>{numero}</span>{titulo}</div>',
                unsafe_allow_html=True)


def main() -> None:
    st.set_page_config(page_title="Cargue de valorización · AGP", page_icon="📄", layout="wide")
    st.markdown(ESTILOS, unsafe_allow_html=True)
    st.markdown(
        '<div class="agp-hero"><div class="eyebrow">DIRECCIÓN AGP</div>'
        '<h1>Cargue de valorización</h1>'
        '<p>Consolide los materiales y prestaciones de su PMS, excluya los materiales '
        'de aportación y descargue el archivo para cargar en AGP.</p></div>', unsafe_allow_html=True)
    try:
        catalogo = cargar_catalogo()
    except ErrorCargue as exc:
        st.error(str(exc))
        st.stop()

    with st.sidebar:
        st.markdown("### Cargue AGP")
        st.caption("Generador de archivos de valorización")
        st.divider()
        st.markdown("**Cómo preparar el cargue**")
        st.markdown("1. Cargue el PMS.\n2. Diligencie ubicación técnica y TAM.\n3. Revise y descargue el Excel.")
        st.divider()
        st.markdown("**Catálogo de aportación**")
        st.write(f"{len(catalogo.codigos)} materiales para excluir automáticamente.")
        st.caption(f"Versión del catálogo: {catalogo.version[:12]}")
        with st.expander("Consultar códigos"):
            st.dataframe([{"CODIGO": c} for c in sorted(catalogo.codigos)],
                         hide_index=True, width="stretch", height=250)
        with st.expander("Formatos admitidos"):
            st.markdown("**RG11:** hoja `RG11_budget`.\n\n"
                        "**E1 consolidado:** `RG10_MO_MT`.\n\n"
                        "**E1 anterior:** `RG10_SER` y `RG36_MAT`.\n\n"
                        "Si coexisten, se usa RG11 primero y luego E1 consolidado.")
        st.caption(f"Aplicación {VERSION}")

    izquierda, derecha = st.columns([1.05, 1], gap="large")
    with izquierda:
        with st.container(border=True):
            seccion(1, "Archivo de origen")
            st.caption("Cargue el RG11 o E1. El nombre del archivo puede ser cualquiera.")
            archivo = st.file_uploader("Archivo PMS", type=["xlsx", "xlsm"], key="pms",
                                       on_change=cambiar_archivo,
                                       help=f"Máximo {MAX_ARCHIVO_MB} MB. No admite Excel .xls.")
            st.caption("Las cantidades repetidas se suman por código. El catálogo se aplica automáticamente.")

    with derecha:
        with st.container(border=True):
            seccion(2, "Datos del cargue")
            nivel = st.radio("Nivel de tensión", ["MT", "BT"], horizontal=True, key="nivel",
                             format_func=lambda v: "Media tensión (MT)" if v == "MT" else "Baja tensión (BT)",
                             on_change=cambiar_nivel)
            base = st.text_input("Código de línea" if nivel == "MT" else "Código de emplazamiento AGP",
                                 placeholder="Ej.: CD1000861" if nivel == "MT" else "Ej.: CD1B166",
                                 key="codigo_base", on_change=mayusculas, args=("codigo_base",),
                                 help="Diligencie el código base, sin la estructura seleccionable.")
            opciones = OPCIONES_MT if nivel == "MT" else OPCIONES_BT
            sufijo = st.selectbox("Estructura de ubicación", opciones, index=None,
                                  placeholder="Seleccione una estructura", key="sufijo",
                                  on_change=limpiar_resultado)
            tam = st.text_input("TAM", placeholder="Ej.: QSR", key="tam",
                                on_change=mayusculas, args=("tam",),
                                help="Se convierte automáticamente a mayúsculas.")
            if base and sufijo:
                try:
                    st.caption("Ubicación técnica que se usará en todas las filas")
                    st.code(construir_ubicacion(nivel, base, sufijo), language=None)
                except ErrorCargue as exc:
                    st.warning(str(exc))
            st.caption("PEL/TRATA: 001 · CIT: XK0 · EQUIPO y CONJUNTO: vacíos")

    resumen = None
    if archivo is not None:
        contenido = archivo.getvalue()
        firma = hashlib.sha256(contenido + catalogo.version.encode()).hexdigest()
        lectura = st.session_state.get("lectura")
        if lectura is None or lectura["firma"] != firma:
            limpiar_resultado()
            try:
                with st.spinner("Leyendo códigos, consolidando cantidades y cruzando aportación…"):
                    resumen = analizar_excel(contenido, catalogo)
                lectura = {"firma": firma, "resumen": resumen, "error": None}
            except ErrorCargue as exc:
                lectura = {"firma": firma, "resumen": None, "error": exc}
            st.session_state.lectura = lectura
        resumen = lectura["resumen"]
        if lectura["error"]:
            error = lectura["error"]
            st.error(str(error))
            if error.incidencias:
                st.dataframe(error.incidencias, hide_index=True, width="stretch")

    st.write("")
    seccion(3, "Revisión y descarga")
    if resumen is None:
        st.markdown('<div class="agp-empty">Cargue un PMS para revisar los códigos que '
                    'se incluirán en el archivo.</div>', unsafe_allow_html=True)
    else:
        a, b, c, d = st.columns(4)
        a.metric("Filas de origen", resumen.filas_leidas)
        b.metric("Códigos unificados", resumen.codigos_unificados)
        c.metric("Códigos excluidos", len({r.codigo for r in resumen.excluidos}))
        d.metric("Códigos para cargar", len(resumen.incluidos))
        st.caption(f"Formato detectado: {resumen.formato}. "
                   "Las cantidades cero se conservan; las cantidades decimales no se redondean.")
        with st.expander("Revisar materiales excluidos por aportación", expanded=False):
            if resumen.excluidos:
                st.dataframe([{"MATERIAL": r.codigo, "CANTIDAD": float(r.cantidad), "MOTIVO": r.motivo}
                              for r in resumen.excluidos], hide_index=True, width="stretch")
            else:
                st.write("No se encontraron coincidencias de aportación.")
        if not resumen.incluidos:
            st.warning("Todos los códigos fueron excluidos por aportación. No hay filas para cargar.")
        else:
            pendiente = None
            try:
                construir_ubicacion(nivel, base, sufijo)
                normalizar_tam(tam)
            except ErrorCargue as exc:
                pendiente = str(exc)
            if pendiente:
                st.info("Complete el código base, la estructura de ubicación y el TAM para generar el archivo.")
            if st.button("Generar archivo de cargue", type="primary", disabled=pendiente is not None,
                         width="stretch", icon=":material/file_save:"):
                limpiar_resultado()
                try:
                    contenido_excel = generar_excel(resumen, nivel, base, sufijo, tam)
                    filas = filas_cargue(resumen, nivel, base, sufijo, tam)
                    st.session_state.resultado = {"contenido": contenido_excel, "filas": filas}
                except ErrorCargue as exc:
                    st.session_state.error_generacion = str(exc)
                except Exception:
                    logging.exception("Error al generar el cargue AGP")
                    st.session_state.error_generacion = "No se pudo generar el Excel. Revise los datos e intente nuevamente."
            if st.session_state.get("error_generacion"):
                st.error(st.session_state.error_generacion)
            resultado = st.session_state.get("resultado")
            if resultado:
                st.success(f"Archivo generado con {len(resultado['filas'])} códigos.")
                vista = [dict(zip(ENCABEZADOS, fila)) for fila in resultado["filas"]]
                for fila in vista:
                    fila["MATERIAL"] = str(fila["MATERIAL"])
                    fila["EQUIPO"] = ""
                    fila["CONJUNTO"] = ""
                st.dataframe(vista,
                             hide_index=True, width="stretch")
                st.caption("El Excel descargado contiene una sola hoja, los ocho encabezados y los datos, sin formato de tabla.")
                st.download_button("Descargar CARGUE_AGP.xlsx", resultado["contenido"],
                                   file_name="CARGUE_AGP.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   on_click="ignore", type="primary", width="stretch",
                                   icon=":material/download:")
    st.markdown('<div class="agp-footer">Dirección AGP · Cargue de materiales y prestaciones</div>',
                unsafe_allow_html=True)


if __name__ == "__main__":
    main()
