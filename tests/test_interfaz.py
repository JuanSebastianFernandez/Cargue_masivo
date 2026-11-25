"""Pruebas de la sesión y descarga con AppTest de Streamlit."""

import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from streamlit.testing.v1 import AppTest

from test_cargue import rg11

APP = Path(__file__).resolve().parents[1] / "streamlit_app.py"


class InterfazCargue(unittest.TestCase):
    def test_inicio_y_cambio_a_bt(self):
        app = AppTest.from_file(str(APP)).run(timeout=20)
        self.assertFalse(app.exception)
        self.assertEqual(len(app.text_input), 2)
        self.assertEqual(len(app.selectbox[0].options), 17)
        app.radio[0].set_value("BT").run()
        self.assertFalse(app.exception)
        self.assertEqual(len(app.selectbox[0].options), 4)
        self.assertEqual(app.text_input[0].label, "Código de emplazamiento AGP")

    def test_generacion_y_cambios_invalidan_descarga(self):
        archivo = SimpleNamespace(getvalue=lambda: rg11((300524, 2), ("OC2137", "1,5")))
        with patch("streamlit.file_uploader", return_value=archivo):
            app = AppTest.from_file(str(APP)).run(timeout=20)
            self.assertFalse(app.exception)
            self.assertTrue(app.button(key="generar_cargue").disabled)
            app.text_input(key="codigo_base").set_value("cd1000861").run()
            app.selectbox(key="sufijo").select("AEN-APO").run()
            app.text_input(key="tam").set_value("qsr").run()
            self.assertEqual(app.text_input(key="tam").value, "QSR")
            self.assertFalse(app.button(key="generar_cargue").disabled)
            app.button(key="generar_cargue").click().run()
            self.assertFalse(app.exception)
            self.assertEqual(len(app.get("download_button")), 1)
            self.assertEqual(app.session_state["resultado"]["filas"][0][:4],
                             ("001", "XK0", "CD1000861-AEN-APO", "QSR"))
            vista = next(elemento.value for elemento in app.dataframe
                         if "EQUIPO" in elemento.value.columns)
            self.assertTrue((vista["EQUIPO"] == "").all())
            self.assertTrue((vista["CONJUNTO"] == "").all())
            app.text_input(key="tam").set_value("otro").run()
            self.assertFalse(app.exception)
            self.assertEqual(len(app.get("download_button")), 0)
            app.button(key="generar_cargue").click().run()
            app.radio[0].set_value("BT").run()
            self.assertEqual(len(app.get("download_button")), 0)
            self.assertTrue(app.button(key="generar_cargue").disabled)

    def test_error_de_datos_bloquea_generacion(self):
        archivo = SimpleNamespace(getvalue=lambda: rg11((300524, "inválido")))
        with patch("streamlit.file_uploader", return_value=archivo):
            app = AppTest.from_file(str(APP)).run(timeout=20)
            self.assertFalse(app.exception)
            self.assertTrue(app.error)
            self.assertEqual(len(app.get("download_button")), 0)
            self.assertNotIn("generar_cargue", [boton.key for boton in app.button])
            self.assertIn("nuevo_cargue", [boton.key for boton in app.button])

    def test_nuevo_cargue_limpia_sesion_y_admite_otro_archivo(self):
        archivos = {"pms_0": SimpleNamespace(getvalue=lambda: rg11((300524, 2)))}
        with patch("streamlit.file_uploader", side_effect=lambda *a, **kw: archivos.get(kw["key"])) as selector:
            app = AppTest.from_file(str(APP)).run(timeout=20)
            app.radio[0].set_value("BT").run()
            app.text_input(key="codigo_base").set_value("CD1B166").run()
            app.selectbox(key="sufijo").select("LBT-CVA00-APO").run()
            app.text_input(key="tam").set_value("QSR").run()
            app.button(key="generar_cargue").click().run()
            self.assertFalse(app.exception)
            self.assertEqual(len(app.get("download_button")), 1)

            app.button(key="nuevo_cargue_resultado").click().run()
            self.assertFalse(app.exception)
            self.assertEqual(selector.call_args.kwargs["key"], "pms_1")
            self.assertEqual(app.radio[0].value, "MT")
            self.assertEqual(app.text_input(key="codigo_base").value, "")
            self.assertEqual(app.text_input(key="tam").value, "")
            self.assertIsNone(app.selectbox(key="sufijo").value)
            self.assertEqual(len(app.get("download_button")), 0)
            for clave in ("lectura", "resultado", "error_generacion"):
                self.assertNotIn(clave, app.session_state)

            # El segundo archivo se procesa desde cero y reemplaza los códigos del primero.
            archivos["pms_1"] = SimpleNamespace(getvalue=lambda: rg11(("OC2137", 5)))
            app.run()
            app.text_input(key="codigo_base").set_value("CD2000").run()
            app.selectbox(key="sufijo").select("AEN").run()
            app.text_input(key="tam").set_value("OTRO").run()
            app.button(key="generar_cargue").click().run()
            self.assertFalse(app.exception)
            self.assertEqual(app.session_state["resultado"]["filas"],
                             [("001", "XK0", "CD2000-AEN", "OTRO", None, None, "OC2137", 5)])

            # El botón de la barra lateral realiza el mismo reinicio.
            app.button(key="nuevo_cargue").click().run()
            self.assertFalse(app.exception)
            self.assertEqual(selector.call_args.kwargs["key"], "pms_2")
            self.assertEqual(len(app.get("download_button")), 0)

    def test_nuevo_cargue_recupera_una_lectura_con_error(self):
        archivo = SimpleNamespace(getvalue=lambda: rg11((300524, "inválido")))
        with patch("streamlit.file_uploader", side_effect=lambda *a, **kw:
                   archivo if kw["key"] == "pms_0" else None):
            app = AppTest.from_file(str(APP)).run(timeout=20)
            self.assertTrue(app.error)
            app.session_state["error_generacion"] = "Error anterior"
            app.button(key="nuevo_cargue").click().run()
            self.assertFalse(app.exception)
            self.assertFalse(app.error)
            self.assertNotIn("lectura", app.session_state)
            self.assertNotIn("error_generacion", app.session_state)


if __name__ == "__main__":
    unittest.main()
