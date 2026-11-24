"""Compara el cargue generado con un Excel de referencia, celda por celda."""

import argparse
import io
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from cargue_agp import analizar_excel, cargar_catalogo, generar_excel


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("rg11", type=Path)
    parser.add_argument("esperado", type=Path)
    parser.add_argument("--salida", type=Path)
    args = parser.parse_args()
    resumen = analizar_excel(args.rg11.read_bytes(), cargar_catalogo())
    contenido = generar_excel(resumen, "MT", "CD1000861", "AEN-APO", "QSR")
    wb = openpyxl.load_workbook(io.BytesIO(contenido))
    esperado = openpyxl.load_workbook(args.esperado, read_only=True, data_only=True)
    try:
        actual, referencia = list(wb.active.values), list(esperado.active.values)
        if actual != referencia:
            raise SystemExit("ERROR: los valores o encabezados no coinciden con el ejemplo.")
        if len(wb.worksheets) != 1 or wb.active.max_column != 8 or wb.active.tables:
            raise SystemExit("ERROR: estructura de salida incorrecta.")
        print(f"OK: {len(actual) - 1} filas coinciden celda por celda; una hoja y ninguna tabla Excel.")
        print(f"Origen: {resumen.filas_leidas} filas. Excluidos: {len(resumen.excluidos)} códigos.")
    finally:
        wb.close()
        esperado.close()
    if args.salida:
        args.salida.parent.mkdir(parents=True, exist_ok=True)
        args.salida.write_bytes(contenido)
        print(f"Archivo guardado: {args.salida.resolve()}")


if __name__ == "__main__":
    main()
