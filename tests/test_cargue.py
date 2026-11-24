"""Pruebas de reglas de negocio y estructura del archivo de cargue."""

import io
import unittest
from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile

import openpyxl

from cargue_agp import (
    Catalogo, ENCABEZADOS, OPCIONES_BT, OPCIONES_MT, ErrorCargue, analizar_excel,
    cargar_catalogo, construir_ubicacion, generar_excel, normalizar_cantidad,
)


def libro(hojas):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for nombre, filas in hojas.items():
        ws = wb.create_sheet(nombre)
        for fila in filas:
            ws.append(fila)
    salida = io.BytesIO()
    wb.save(salida)
    wb.close()
    return salida.getvalue()


def rg11(*filas):
    return libro({"RG11_budget": [("Título",), (), ("cod_prest/mat", "quantity"), *filas]})


CATALOGO = Catalogo(frozenset({"251363", "274311"}), "prueba")


class ReglasCargue(unittest.TestCase):
    def test_consolida_normaliza_y_excluye(self):
        resumen = analizar_excel(rg11(
            (" oc2137 ", "1,25"), ("OC2137", "2,5"), (300524, 2),
            ("300524.0", 4), (251363, 16), (251363, 4),
        ), CATALOGO)
        self.assertEqual(resumen.filas_leidas, 6)
        self.assertEqual(resumen.codigos_unificados, 3)
        self.assertEqual([(r.codigo, r.cantidad) for r in resumen.incluidos],
                         [("300524", Decimal(6)), ("OC2137", Decimal("3.75"))])
        self.assertEqual(resumen.excluidos[0].cantidad, Decimal(20))

    def test_rg11_no_cambia_signo_por_retiro(self):
        datos = libro({"RG11_budget": [("cod_prest/mat", "quantity", "tipo trabajo"),
                                      (970161, 3, "Retiro")]})
        self.assertEqual(analizar_excel(datos, CATALOGO).incluidos[0].cantidad, 3)

    def test_decimales_cero_y_negativos_no_se_pierden(self):
        resumen = analizar_excel(rg11((1, "0,1"), (1, "0,2"), (2, 0), (3, "2-")), CATALOGO)
        self.assertEqual([r.cantidad for r in resumen.incluidos],
                         [Decimal("0.3"), Decimal(0), Decimal(-2)])
        for texto, numero in [("1.234,56", "1234.56"), ("1,234.56", "1234.56"),
                              (" 1\u00a0234,5 ", "1234.5")]:
            self.assertEqual(normalizar_cantidad(texto), Decimal(numero))

    def test_todas_las_ubicaciones_autorizadas(self):
        self.assertEqual(len(OPCIONES_MT), 17)
        self.assertEqual(len(OPCIONES_BT), 4)
        for nivel, opciones, base in [("MT", OPCIONES_MT, "cd1000861"),
                                     ("BT", OPCIONES_BT, "cd1b166")]:
            for sufijo in opciones:
                self.assertEqual(construir_ubicacion(nivel, base, sufijo), f"{base.upper()}-{sufijo}")
        for args in [("MT", "", "AEN"), ("BT", "CD1 B166", "LBT-CVA00"),
                     ("MT", "CD1000861-AEN", "AEN"), ("BT", "CD1B166", "AEN")]:
            with self.assertRaises(ErrorCargue):
                construir_ubicacion(*args)

    def test_excel_plano_y_tipos_de_celda(self):
        r = analizar_excel(rg11((300524, 3), ("OC2137", "1,25"), ("001234", 2)), CATALOGO)
        contenido = generar_excel(r, "BT", "cd1b166", "LBT-CVA00-APO", "qsr")
        wb = openpyxl.load_workbook(io.BytesIO(contenido))
        ws = wb.active
        self.assertEqual(wb.sheetnames, ["CARGUE_AGP"])
        self.assertEqual(tuple(c.value for c in ws[1]), ENCABEZADOS)
        self.assertEqual((ws.max_row, ws.max_column), (4, 8))
        self.assertEqual(len(ws.tables), 0)
        self.assertIsNone(ws.auto_filter.ref)
        filas = list(ws.values)[1:]
        for fila in filas:
            self.assertEqual(fila[:6], ("001", "XK0", "CD1B166-LBT-CVA00-APO", "QSR", None, None))
        self.assertEqual([f[6] for f in filas], ["001234", 300524, "OC2137"])
        self.assertEqual(ws["A2"].data_type, "s")
        self.assertEqual(ws["H4"].value, 1.25)
        self.assertEqual(ws["H4"].data_type, "n")
        wb.close()
        with ZipFile(io.BytesIO(contenido)) as z:
            self.assertFalse(any(n.startswith("xl/tables/") for n in z.namelist()))

    def test_tam_obligatorio_y_se_escribe_como_texto(self):
        r = analizar_excel(rg11((1, 1)), CATALOGO)
        with self.assertRaises(ErrorCargue):
            generar_excel(r, "MT", "CD1", "AEN", "  ")
        wb = openpyxl.load_workbook(io.BytesIO(generar_excel(r, "MT", "CD1", "AEN", "=tam")))
        self.assertEqual(wb.active["D2"].value, "=TAM")
        self.assertEqual(wb.active["D2"].data_type, "s")
        wb.close()

    def test_no_exporta_resultados_parciales_con_datos_invalidos(self):
        for fila in [(None, 2), (1, "texto"), ("#N/A", 3), (1, None), (1, True)]:
            with self.subTest(fila=fila), self.assertRaises(ErrorCargue) as exc:
                analizar_excel(rg11((300524, 1), fila), CATALOGO)
            self.assertTrue(exc.exception.incidencias)

    def test_formulas_sin_cache_no_se_omiten(self):
        for fila in [(1, "=1+1"), ("=1+1", "=1+1")]:
            with self.assertRaises(ErrorCargue) as exc:
                analizar_excel(rg11(fila), CATALOGO)
            self.assertIn("Recalcule", str(exc.exception.incidencias))

    def test_libro_invalido_vacio_y_sin_formato(self):
        for contenido in [b"", b"No es un Excel", libro({"Hoja1": [(1, 2)]}), rg11()]:
            with self.assertRaises(ErrorCargue):
                analizar_excel(contenido, CATALOGO)

    def test_solo_aportacion_impide_descarga(self):
        r = analizar_excel(rg11((251363, 3)), CATALOGO)
        self.assertFalse(r.incluidos)
        with self.assertRaises(ErrorCargue):
            generar_excel(r, "MT", "CD1", "AEN", "QSR")

    def test_e1_consolidado_movimientos_y_prioridad(self):
        datos = libro({
            "RG10_MO_MT": [("PRESTACION/MATERIAL", "CANTIDAD", "MO_MT", "MOVIMIENTO"),
                           (300524, 5, "MATERIAL", "Instalado"),
                           (300524, 2, "MATERIAL", "Retirado"),
                           (999, 1, "MATERIAL", "Material de aportación"),
                           (274311, 8, "MATERIAL", "Instalado"),
                           ("OC2137", 3, "PRESTACION", "Retiro")],
            "RG10_SER": [("PRESTACION", "CANTIDAD"), ("OC2137", 999)],
            "RG36_MAT": [("CODIGO MATERIAL E4E", "CANTIDAD", "TIPO DE MOVIMIENTO"),
                         (300524, 999, "Instalado")],
        })
        r = analizar_excel(datos, CATALOGO)
        self.assertEqual(r.formato, "E1 consolidado")
        self.assertEqual([(x.codigo, x.cantidad) for x in r.incluidos], [("300524", 3), ("OC2137", 3)])
        self.assertEqual({x.codigo for x in r.excluidos}, {"999", "274311"})
        # RG11 gana si ambos formatos coexisten.
        datos = libro({"RG11_budget": [("cod_prest/mat", "quantity"), (1, 2)],
                       "RG10_MO_MT": [("IGNORADO",)]})
        self.assertEqual(analizar_excel(datos, CATALOGO).formato, "RG11")

    def test_e1_anterior(self):
        datos = libro({"RG10_SER": [("PRESTACION", "CANTIDAD"), ("OC2137", 2)],
                       "RG36_MAT": [("CODIGO MATERIAL E4E", "CANTIDAD", "TIPO DE MOVIMIENTO"),
                                    (300524, 4, "Material instalado"),
                                    (300524, 1, "Material retirado"),
                                    (999, 6, "Material de aportación")]})
        r = analizar_excel(datos, CATALOGO)
        self.assertEqual(r.formato, "E1 anterior")
        self.assertEqual([(x.codigo, x.cantidad) for x in r.incluidos], [("300524", 3), ("OC2137", 2)])

    def test_catalogo_autonomo_y_validado(self):
        self.assertEqual(len(cargar_catalogo().codigos), 103)
        with TemporaryDirectory() as temporal:
            ruta = Path(temporal) / "catalogo.csv"
            for texto in ["OTRO\n123\n", "CODIGO\nABC\n", "CODIGO\n", "CODIGO\n123,456\n"]:
                ruta.write_text(texto, encoding="utf-8")
                with self.assertRaises(ErrorCargue):
                    cargar_catalogo(ruta)


if __name__ == "__main__":
    unittest.main()
