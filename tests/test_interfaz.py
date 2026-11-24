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
            self.assertTrue(app.button[0].disabled)
            app.text_input(key="codigo_base").set_value("cd1000861").run()
            app.selectbox(key="sufijo").select("AEN-APO").run()
            app.text_input(key="tam").set_value("qsr").run()
            self.assertEqual(app.text_input(key="tam").value, "QSR")
            self.assertFalse(app.button[0].disabled)
            app.button[0].click().run()
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
            app.button[0].click().run()
            app.radio[0].set_value("BT").run()
            self.assertEqual(len(app.get("download_button")), 0)
            self.assertTrue(app.button[0].disabled)

    def test_error_de_datos_bloquea_generacion(self):
        archivo = SimpleNamespace(getvalue=lambda: rg11((300524, "inválido")))
        with patch("streamlit.file_uploader", return_value=archivo):
            app = AppTest.from_file(str(APP)).run(timeout=20)
            self.assertFalse(app.exception)
            self.assertTrue(app.error)
            self.assertEqual(len(app.get("download_button")), 0)
            self.assertEqual(len(app.button), 0)


if __name__ == "__main__":
    unittest.main()
