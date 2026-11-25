# Cargue de valorización AGP

Aplicación Streamlit para convertir un PMS RG11 o E1 en `CARGUE_AGP.xlsx`.
Incluye el catálogo de 103 materiales de aportación suministrado para este proyecto.
Funciona sin depender de las carpetas de `Valida_mat_pres` ni de los archivos de ejemplo.

## Ejecutar en Windows

Requiere Python 3.11 o posterior. Desde esta carpeta:

```powershell
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
.venv\Scripts\python -m streamlit run streamlit_app.py
```

Abra la dirección local que muestre Streamlit. No es necesario activar el entorno de PowerShell.
También puede ejecutar `iniciar.ps1` después de instalar las dependencias.

## Uso

1. Cargue el PMS `.xlsx` o `.xlsm`. Su nombre es libre.
2. Seleccione MT o BT, diligencie el código de línea o de emplazamiento AGP y seleccione la estructura.
3. Diligencie el TAM. La aplicación lo transforma a mayúsculas.
4. Revise los códigos consolidados y los materiales excluidos.
5. Pulse **Generar archivo de cargue**, revise la vista previa y descargue el Excel.
6. Pulse **Nuevo cargue**, en la barra lateral o debajo de la descarga, para procesar otro PMS sin recargar la página.

**Nuevo cargue** limpia el archivo, la lectura, los errores, la ubicación técnica, el TAM
y el resultado anterior. El nivel de tensión vuelve a MT y el catálogo de aportación permanece disponible.

Un cambio de archivo, catálogo, tensión, código base, estructura o TAM retira la descarga anterior.
Los archivos se procesan en memoria y cada sesión conserva su propio resultado.

## Estructura de salida

Una hoja ordinaria llamada `CARGUE_AGP`, sin objeto Tabla de Excel, sin filtros, totales,
portadas ni hojas auxiliares. Los encabezados están en A1:H1 y los datos empiezan en A2:

```text
PEL/TRATA | CIT | ubicación tecnica | TAM | EQUIPO | CONJUNTO | MATERIAL | CANTIDAD
```

- `PEL/TRATA` siempre es el texto `001`, conservando los ceros.
- `CIT` siempre es `XK0`.
- `ubicación tecnica` une el código base del profesional y la estructura seleccionada con un guion.
- `TAM` es obligatorio y se escribe en mayúsculas.
- `EQUIPO` y `CONJUNTO` quedan vacíos.
- `MATERIAL` incluye materiales y prestaciones, con cantidades sumadas por código.
- `CANTIDAD` es numérica. Se conservan decimales, ceros y signos; no se aplica redondeo de negocio.

Los códigos numéricos se ordenan primero y luego los alfanuméricos, como en el ejemplo.
Se conservan como texto los códigos con ceros iniciales o más de 15 dígitos.

### Ubicaciones admitidas

**MT:** AEN, AEN-APO, AEN-CUC, AEN-SEC, AEN-REC, AEN-CCO, AEN-PAR,
CVA, CVA-APO, CVA-CUC, CVA-SEC, CVA-REC, CVA-CCO, CVA-PAR, CVI, CVI-CNL, CST.

**BT:** LBT-CVA00, LBT-CVA00-APO, LBT-AEN00, LBTAEN00-APO.
La última opción se conserva literalmente como se solicitó, sin insertar un guion entre LBT y AEN00.

## Lectura y exclusión

Se adaptaron del proyecto `Valida_mat_pres` la detección por hojas y encabezados,
la normalización de códigos y cantidades y las reglas de movimientos de E1.
El motor de este proyecto es autónomo.

| Prioridad | Formato | Campos de origen |
| --- | --- | --- |
| 1 | RG11, hoja `RG11_budget` | `cod_prest/mat` y `quantity` o `cantidad` |
| 2 | E1 consolidado, hoja `RG10_MO_MT` | `PRESTACION/MATERIAL`, `CANTIDAD`, `MO_MT` y `MOVIMIENTO` |
| 3 | E1 anterior, hojas `RG10_SER` y `RG36_MAT` | `PRESTACION`, `CODIGO MATERIAL E4E`, `CANTIDAD` y `TIPO DE MOVIMIENTO` |

Los encabezados se buscan en las primeras 40 filas; no se depende de letras de columna.
Si coexisten varios formatos se procesa solo el de mayor prioridad.

Los materiales que coincidan con `catalogo/materiales_aportacion.csv` se omiten en todos
los formatos. E1 también excluye los marcados como Material de aportación en el movimiento.
En E1 se resta el material retirado según las reglas del proyecto de referencia.
En RG11 se conserva el signo de `quantity`, incluso si otra columna indica Retiro,
para reproducir el cargue de ejemplo. El tipo de trabajo no cambia la cantidad de RG11.

Las filas con cantidades inválidas, código faltante o fórmulas sin resultado guardado
bloquean la generación y se muestran con hoja y fila. Para fórmulas, recalcule el PMS
en Excel y guárdelo antes de cargarlo. Los errores de fórmulas ajenos a las columnas
que se procesan no intervienen en el cargue.

### Actualizar aportación

Edite `catalogo/materiales_aportacion.csv` conservando una única columna `CODIGO`,
un código numérico por fila y codificación UTF-8. La interfaz muestra el número de códigos
y la huella de la versión. Si el catálogo no se puede leer, la generación se detiene.

## Comprobaciones

```powershell
.venv\Scripts\python -m unittest discover -s tests -v
```

Para contrastar con los archivos de referencia, sin incluirlos en el repositorio:

```powershell
.venv\Scripts\python scripts/verificar_ejemplo.py "RUTA\EJemplo_RG11.xlsx" "RUTA\CARGUE_EJEMPLO.xlsx" --salida outputs/CARGUE_AGP.xlsx
```

Con el RG11 suministrado y ubicación `CD1000861-AEN-APO`, TAM `QSR`, se verificaron
30 filas de origen, 10 códigos excluidos y 20 filas finales coincidentes celda por celda
con el ejemplo, incluyendo encabezados y tipos de datos.

## Límites

Hasta 60 MB por archivo y 256 MB descomprimidos, con un máximo de 2.048 elementos
internos. Cada hoja procesada puede tener hasta 50.000 filas, 200 columnas y 500.000
posiciones. Las hojas auxiliares no se procesan. No se admiten archivos `.xls` antiguos.
Estos límites previenen cargas excesivas; la capacidad simultánea depende del servidor.

Para desplegar en Streamlit use `streamlit_app.py`, `requirements.txt`, `cargue_agp.py`,
`.streamlit/config.toml` y `catalogo/materiales_aportacion.csv`. No se requieren secretos.
La aplicación no incorpora autenticación.

Referencias técnicas: [descargas de Streamlit](https://docs.streamlit.io/develop/api-reference/widgets/st.download_button)
y [pruebas de interfaz con AppTest](https://docs.streamlit.io/develop/api-reference/app-testing).
