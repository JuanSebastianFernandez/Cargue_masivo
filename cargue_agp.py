"""Lectura PMS y generación del cargue AGP, independiente de Streamlit.

Los encabezados, normalización y prioridad de formatos siguen Valida_mat_pres.
El RG11 conserva el signo original. E1 aplica sus reglas de movimiento.
"""

from __future__ import annotations

import csv
import hashlib
import io
import math
import re
import unicodedata
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from zipfile import BadZipFile, ZIP_DEFLATED, ZipFile

import openpyxl
from openpyxl.styles import Alignment, Font


ENCABEZADOS = (
    "PEL/TRATA", "CIT", "ubicación tecnica", "TAM", "EQUIPO", "CONJUNTO",
    "MATERIAL", "CANTIDAD",
)
OPCIONES_MT = (
    "AEN", "AEN-APO", "AEN-CUC", "AEN-SEC", "AEN-REC", "AEN-CCO", "AEN-PAR",
    "CVA", "CVA-APO", "CVA-CUC", "CVA-SEC", "CVA-REC", "CVA-CCO", "CVA-PAR",
    "CVI", "CVI-CNL", "CST",
)
# Se conserva LBTAEN00-APO exactamente como lo solicitó el usuario.
OPCIONES_BT = ("LBT-CVA00", "LBT-CVA00-APO", "LBT-AEN00", "LBTAEN00-APO")
RUTA_CATALOGO = Path(__file__).parent / "catalogo" / "materiales_aportacion.csv"
MAX_ARCHIVO_MB = 60
MAX_XML_BYTES = 256 * 1024 * 1024
MAX_FILAS = 50_000
MAX_COLUMNAS = 200
MAX_CELDAS = 500_000
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


class ErrorCargue(ValueError):
    """Error operativo con incidencias opcionales para mostrar en la interfaz."""

    def __init__(self, mensaje: str, incidencias: list[dict] | None = None):
        super().__init__(mensaje)
        self.incidencias = incidencias or []


@dataclass(frozen=True)
class Catalogo:
    codigos: frozenset[str]
    version: str


@dataclass(frozen=True)
class Registro:
    codigo: str
    cantidad: Decimal
    tipo: str
    hoja: str
    fila: int
    motivo_exclusion: str = ""


@dataclass(frozen=True)
class Consolidado:
    codigo: str
    cantidad: Decimal
    motivo: str = ""


@dataclass(frozen=True)
class Resumen:
    formato: str
    filas_leidas: int
    codigos_unificados: int
    incluidos: tuple[Consolidado, ...]
    excluidos: tuple[Consolidado, ...]


def normalizar_encabezado(valor: object) -> str:
    texto = "" if valor is None else str(valor)
    texto = "".join(c for c in unicodedata.normalize("NFKD", texto)
                    if not unicodedata.combining(c))
    return re.sub(r"[^A-Z0-9]+", " ", texto.upper()).strip()


def normalizar_codigo(valor: object) -> str:
    if valor is None or isinstance(valor, bool):
        return ""
    if isinstance(valor, float):
        if not math.isfinite(valor):
            return ""
        if valor.is_integer():
            return str(int(valor))
    texto = re.sub(r"\s+", "", str(valor)).upper()
    if re.fullmatch(r"\d+\.0+", texto):
        texto = texto.split(".")[0]
    return texto


def normalizar_cantidad(valor: object) -> Decimal:
    """Admite coma decimal, ambos separadores y negativos con signo final."""
    if valor is None or isinstance(valor, bool):
        raise ValueError("Cantidad vacía o no numérica.")
    texto = re.sub(r"\s+", "", str(valor))
    if texto.endswith("-") and texto.count("-") == 1:
        texto = "-" + texto[:-1]
    if "," in texto and "." in texto:
        texto = (texto.replace(".", "").replace(",", ".")
                 if texto.rfind(",") > texto.rfind(".") else texto.replace(",", ""))
    elif "," in texto:
        texto = texto.replace(",", ".")
    elif texto.count(".") > 1:
        partes = texto.split(".")
        texto = "".join(partes[:-1]) + "." + partes[-1]
    try:
        numero = Decimal(texto)
    except InvalidOperation as exc:
        raise ValueError("Cantidad vacía o no numérica.") from exc
    if not numero.is_finite() or not math.isfinite(float(numero)):
        raise ValueError("La cantidad debe ser un número finito.")
    return numero


def cargar_catalogo(ruta: Path = RUTA_CATALOGO) -> Catalogo:
    try:
        contenido = ruta.read_bytes()
        lector = csv.DictReader(io.StringIO(contenido.decode("utf-8-sig")))
        if lector.fieldnames != ["CODIGO"]:
            raise ValueError("Debe tener una única columna llamada CODIGO.")
        codigos: set[str] = set()
        for numero, fila in enumerate(lector, 2):
            codigo = normalizar_codigo(fila.get("CODIGO"))
            if None in fila or not re.fullmatch(r"[0-9]+", codigo):
                raise ValueError(f"Código inválido en la fila {numero}.")
            codigos.add(codigo)
        if not codigos:
            raise ValueError("El catálogo está vacío.")
        return Catalogo(frozenset(codigos), hashlib.sha256(contenido).hexdigest())
    except (OSError, UnicodeError, csv.Error, ValueError) as exc:
        raise ErrorCargue(f"No se pudo leer el catálogo de aportación: {exc}") from exc


def construir_ubicacion(nivel: str, codigo_base: str, sufijo: str) -> str:
    opciones = {"MT": OPCIONES_MT, "BT": OPCIONES_BT}
    if nivel not in opciones or sufijo not in opciones[nivel]:
        raise ErrorCargue("Seleccione un nivel de tensión y una estructura válidos.")
    base = codigo_base.strip().upper()
    if not re.fullmatch(r"[A-Z0-9]+", base):
        raise ErrorCargue("Diligencie solo el código de línea o emplazamiento, sin espacios ni sufijo.")
    ubicacion = f"{base}-{sufijo}"
    if len(ubicacion) > 32767:
        raise ErrorCargue("La ubicación técnica supera el tamaño permitido por Excel.")
    return ubicacion


def normalizar_tam(tam: str) -> str:
    texto = tam.strip().upper()
    if not texto:
        raise ErrorCargue("Diligencie el TAM.")
    if len(texto) > 32767 or re.search(r"[\x00-\x1f]", texto):
        raise ErrorCargue("El TAM contiene caracteres o una longitud no permitidos por Excel.")
    return texto


def _preparar_excel(contenido: bytes) -> bytes:
    """Comprueba el archivo y sanea nombres de impresión sin tocar las celdas."""
    if not contenido or len(contenido) > MAX_ARCHIVO_MB * 1024 * 1024:
        raise ErrorCargue(f"Cargue un Excel no vacío de hasta {MAX_ARCHIVO_MB} MB.")
    try:
        with ZipFile(io.BytesIO(contenido)) as archivo:
            infos = archivo.infolist()
            if len(infos) > 2048 or sum(i.file_size for i in infos) > MAX_XML_BYTES:
                raise ErrorCargue("El Excel supera el límite de contenido descomprimido (256 MB).")
            if "xl/workbook.xml" not in archivo.namelist():
                raise ErrorCargue("El archivo no es un libro Excel .xlsx o .xlsm válido.")
            raiz = ET.fromstring(archivo.read("xl/workbook.xml"))
            nombres = raiz.find(f"{{{NS}}}definedNames")
            cambio = False
            if nombres is not None:
                for nombre in list(nombres):
                    texto = nombre.text or ""
                    if any(error in texto for error in ("#REF!", "#N/A")) or (
                        "Excel_BuiltIn_Print" in nombre.get("name", "")
                    ):
                        nombres.remove(nombre)
                        cambio = True
            if not cambio:
                return contenido
            salida = io.BytesIO()
            with ZipFile(salida, "w", ZIP_DEFLATED) as limpio:
                for info in infos:
                    datos = (ET.tostring(raiz, encoding="utf-8", xml_declaration=True)
                             if info.filename == "xl/workbook.xml" else archivo.read(info.filename))
                    limpio.writestr(info, datos)
            return salida.getvalue()
    except ErrorCargue:
        raise
    except (BadZipFile, KeyError, ET.ParseError, RuntimeError, OSError) as exc:
        raise ErrorCargue("No se pudo abrir el Excel. Compruebe que no esté dañado o protegido.") from exc


def identificar_formato(hojas: list[str]) -> tuple[str, tuple[str, ...]]:
    # Una única fuente por libro: evita sumar dos versiones de la misma información.
    if "RG11_budget" in hojas:
        return "RG11", ("RG11_budget",)
    if "RG10_MO_MT" in hojas:
        return "E1 consolidado", ("RG10_MO_MT",)
    if {"RG10_SER", "RG36_MAT"}.issubset(hojas):
        return "E1 anterior", ("RG10_SER", "RG36_MAT")
    raise ErrorCargue("No se reconoce el PMS. Debe contener RG11_budget, RG10_MO_MT "
                     "o las hojas RG10_SER y RG36_MAT. El nombre del archivo es libre.")


def _buscar_encabezados(ws, aliases: dict[str, tuple[str, ...]]) -> tuple[int, dict[str, int]]:
    if ((ws.max_row or 0) > MAX_FILAS or (ws.max_column or 0) > MAX_COLUMNAS
            or (ws.max_row or 0) * (ws.max_column or 0) > MAX_CELDAS):
        raise ErrorCargue(f"La hoja {ws.title} supera los límites de lectura: "
                         "50.000 filas, 200 columnas o 500.000 celdas.")
    for numero, fila in enumerate(ws.iter_rows(max_row=40, max_col=MAX_COLUMNAS,
                                              values_only=True), 1):
        nombres = [normalizar_encabezado(v) for v in fila]
        indices = {campo: next((nombres.index(normalizar_encabezado(a)) for a in opciones
                               if normalizar_encabezado(a) in nombres), None)
                   for campo, opciones in aliases.items()}
        if all(i is not None for i in indices.values()):
            return numero, indices
    raise ErrorCargue(f"No se encontraron los encabezados en {ws.title}: "
                     + ", ".join(opciones[0] for opciones in aliases.values()))


def _leer_hoja(ws, formulas, formato: str, incidencias: list[dict]) -> list[Registro]:
    if formato == "RG11":
        aliases = {"codigo": ("cod_prest/mat", "codigo", "cod prest mat"),
                   "cantidad": ("quantity", "cantidad")}
    elif ws.title == "RG10_MO_MT":
        aliases = {"codigo": ("PRESTACION MATERIAL", "CODIGO PRESTACION MATERIAL"),
                   "cantidad": ("CANTIDAD", "CANTIDAD TOTAL"),
                   "tipo": ("MO MT",), "movimiento": ("MOVIMIENTO", "TIPO DE MOVIMIENTO")}
    elif ws.title == "RG10_SER":
        aliases = {"codigo": ("PRESTACION",), "cantidad": ("CANTIDAD",)}
    else:
        aliases = {"codigo": ("CODIGO MATERIAL E4E",), "cantidad": ("CANTIDAD",),
                   "movimiento": ("TIPO DE MOVIMIENTO",)}
    inicio, indices = _buscar_encabezados(ws, aliases)
    max_col = max(indices.values()) + 1
    registros: list[Registro] = []
    # Leer también las fórmulas permite detectar resultados sin caché, incluso
    # cuando código y cantidad aparecen vacíos en data_only=True.
    valores = ws.iter_rows(min_row=inicio + 1, max_col=max_col, values_only=True)
    expresiones = formulas.iter_rows(min_row=inicio + 1, max_col=max_col, values_only=True)
    for numero, (fila, fila_formula) in enumerate(zip(valores, expresiones), inicio + 1):
        datos = {campo: fila[i] for campo, i in indices.items()}
        if all(datos[c] in (None, "") and fila_formula[indices[c]] in (None, "")
               for c in ("codigo", "cantidad")):
            continue
        try:
            for campo, indice in indices.items():
                if isinstance(fila_formula[indice], str) and fila_formula[indice].startswith("="):
                    if fila[indice] is None or (isinstance(fila[indice], str)
                                               and fila[indice].startswith("#")):
                        raise ValueError(f"{campo}: fórmula sin resultado válido. "
                                         "Recalcule y guarde el archivo en Excel.")
            codigo = normalizar_codigo(datos["codigo"])
            if not re.fullmatch(r"[A-Z0-9]+", codigo):
                raise ValueError("Código vacío o inválido.")
            cantidad = normalizar_cantidad(datos["cantidad"])
            tipo = "MATERIAL" if codigo.isdigit() else "PRESTACION"
            if ws.title == "RG10_SER":
                tipo = "PRESTACION"
            elif ws.title == "RG36_MAT":
                tipo = "MATERIAL"
            elif ws.title == "RG10_MO_MT":
                tipo = normalizar_encabezado(datos["tipo"])
                if tipo not in {"PRESTACION", "MATERIAL"}:
                    raise ValueError("MO_MT debe indicar PRESTACION o MATERIAL.")
            movimiento = normalizar_encabezado(datos.get("movimiento"))
            motivo = ""
            if tipo == "MATERIAL" and formato != "RG11":
                if "MATERIAL DE APORTACION" in movimiento:
                    motivo = "Material de aportación según movimiento E1"
                retiro = ("RETIRADO" in movimiento if ws.title == "RG36_MAT"
                          else movimiento in {"RETIRADO", "RETIRADA", "RETIRO", "CHATARRA"})
                if retiro:
                    cantidad = -abs(cantidad)
            registros.append(Registro(codigo, cantidad, tipo, ws.title, numero, motivo))
        except ValueError as exc:
            incidencias.append({"HOJA": ws.title, "FILA": numero,
                                "CODIGO": str(datos["codigo"] or ""), "DETALLE": str(exc)})
    return registros


def _orden_codigo(codigo: str) -> tuple:
    return (0, int(codigo), codigo) if codigo.isdigit() else (1, codigo, codigo)


def analizar_excel(contenido: bytes, catalogo: Catalogo) -> Resumen:
    if not catalogo.codigos:
        raise ErrorCargue("El catálogo de aportación no puede estar vacío.")
    limpio = _preparar_excel(contenido)
    wb = formulas = None
    try:
        wb = openpyxl.load_workbook(io.BytesIO(limpio), read_only=True,
                                   data_only=True, keep_links=False)
        formulas = openpyxl.load_workbook(io.BytesIO(limpio), read_only=True,
                                         data_only=False, keep_links=False)
        formato, hojas = identificar_formato(wb.sheetnames)
        registros: list[Registro] = []
        incidencias: list[dict] = []
        for hoja in hojas:
            registros.extend(_leer_hoja(wb[hoja], formulas[hoja], formato, incidencias))
        if incidencias:
            raise ErrorCargue(f"Se encontraron {len(incidencias)} filas con errores. "
                             "Corríjalas en el PMS y vuelva a cargarlo.", incidencias)
        if not registros:
            raise ErrorCargue("El PMS no contiene materiales o prestaciones para procesar.")
        incluidos: dict[str, Decimal] = {}
        excluidos: dict[tuple[str, str], Decimal] = {}
        for registro in registros:
            motivo = registro.motivo_exclusion
            # El maestro solicitado se cruza en todos los formatos admitidos.
            if registro.tipo == "MATERIAL" and registro.codigo in catalogo.codigos:
                motivo = "Coincide con el catálogo de materiales de aportación"
            if motivo:
                clave = registro.codigo, motivo
                excluidos[clave] = excluidos.get(clave, Decimal(0)) + registro.cantidad
            else:
                incluidos[registro.codigo] = incluidos.get(registro.codigo, Decimal(0)) + registro.cantidad
        return Resumen(
            formato, len(registros), len({r.codigo for r in registros}),
            tuple(Consolidado(c, incluidos[c]) for c in sorted(incluidos, key=_orden_codigo)),
            tuple(Consolidado(c, q, m) for (c, m), q in
                  sorted(excluidos.items(), key=lambda x: (_orden_codigo(x[0][0]), x[0][1]))),
        )
    except ErrorCargue:
        raise
    except Exception as exc:
        raise ErrorCargue("No se pudo leer el contenido del PMS. Abra el archivo en Excel, "
                         "recalcule y guárdelo como .xlsx antes de reintentar.") from exc
    finally:
        if wb is not None:
            wb.close()
        if formulas is not None:
            formulas.close()


def _codigo_excel(codigo: str) -> str | int:
    # Igual al ejemplo: materiales numéricos; prestaciones alfanuméricas.
    # Los ceros iniciales y códigos largos deben conservarse como texto.
    if codigo.isdigit() and len(codigo) <= 15 and not codigo.startswith("0"):
        return int(codigo)
    return codigo


def filas_cargue(resumen: Resumen, nivel: str, codigo_base: str, sufijo: str,
                 tam: str) -> list[tuple]:
    ubicacion = construir_ubicacion(nivel, codigo_base, sufijo)
    tam = normalizar_tam(tam)
    if not resumen.incluidos:
        raise ErrorCargue("Todos los códigos fueron excluidos por aportación. No hay filas para cargar.")
    return [("001", "XK0", ubicacion, tam, None, None, _codigo_excel(r.codigo),
             int(r.cantidad) if r.cantidad == r.cantidad.to_integral_value() else float(r.cantidad))
            for r in resumen.incluidos]


def generar_excel(resumen: Resumen, nivel: str, codigo_base: str, sufijo: str,
                   tam: str) -> bytes:
    """Una sola hoja ordinaria: ocho encabezados, sin tablas ni totales."""
    filas = filas_cargue(resumen, nivel, codigo_base, sufijo, tam)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CARGUE_AGP"
    ws.append(ENCABEZADOS)
    for fila in filas:
        ws.append(fila)
    for fila in ws.iter_rows():
        for celda in fila:
            celda.font = Font(name="Calibri", size=11, bold=celda.row == 1)
            celda.alignment = Alignment(vertical="center")
            # TAM puede contener símbolos; siempre se escribe como texto, nunca fórmula.
            if isinstance(celda.value, str):
                celda.data_type = "s"
    for letra, ancho in zip("ABCDEFGH", (16, 12, 34, 18, 17, 17, 20, 18)):
        ws.column_dimensions[letra].width = ancho
    ws.row_dimensions[1].height = 24
    ws.freeze_panes = "A2"
    for fila in ws.iter_rows(min_row=2):
        fila[0].number_format = "@"
        fila[7].number_format = "General"
    salida = io.BytesIO()
    wb.save(salida)
    wb.close()
    return salida.getvalue()
