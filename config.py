"""
config.py - Configuración centralizada del sistema de automatización archivística.
"""
from pathlib import Path

# ─── Rutas de entrada ────────────────────────────────────────────────────────
# Cambia estas rutas para apuntar a tus archivos reales antes de ejecutar.
BASE_DIR    = Path(__file__).parent
EXCEL_PATH  = BASE_DIR / "copia_cata_DiegoAguilar.xlsx"  # Archivo Excel .xlsx de entrada
PDF_PATH    = BASE_DIR / "documento.pdf"                  # PDF escaneado de entrada

# ─── Carpetas de salida ───────────────────────────────────────────────────────
OUTPUT_DIR  = BASE_DIR / "output"            # Raíz donde se crean las carpetas
LOGS_DIR    = BASE_DIR / "logs"

PROCESS_LOG    = LOGS_DIR / "process.log"
PENDIENTES_CSV = LOGS_DIR / "pendientes.csv"

# ─── Parámetros del Excel ─────────────────────────────────────────────────────
# Estructura del Excel copia_cata_DiegoAguilar.xlsx:
#   Fila 4:  Seccion: XVI          → Siglo (nivel 2 de la jerarquía)
#   Fila 7:  Código del fondo: N7  → Número de acervo (nivel 1 de la jerarquía)
#   Fila 8:  Encabezados de columnas
#   Fila 9:  Sub-encabezados (FECHA INICIAL / FECHA FINAL)
#   Filas 10+: Marcadores de sección (Protocolo Nº X, Registro Nº X) + datos
SKIP_ROWS = 7   # skiprows para pandas (salta filas 1-7, usa fila 8 como header)

# Filas del Excel donde se leen los metadatos globales del fondo
META_ROW_SIGLO   = 4   # «Seccion: XVI»
META_ROW_ACERVO  = 7   # «Código del fondo: N7»

# Nombres de columna tal como aparecen en la fila 8 del Excel (después de strip).
COL_REGISTRO   = "N° DE REGISTRO"
COL_ESCRIBANO  = "ESCRIBANO/\nNOTARIO"    # la celda tiene salto de línea
COL_PROTOCOLO  = "N° DE PROT."
COL_FOLIOS     = "N° DE FOLIOS"
COL_LUGAR      = "DATA TÓPICA (Lugar)"
COL_FECHA_INI  = "FECHA INICIAL"           # sub-columna de DATA CRÓNICA
COL_TITULO     = "TÍTULO/\nESCRITURA"      # título libre
COL_TITULO_EST = "Titulo estandar"         # columna I – se usa para la carpeta
COL_INT1       = "INTERESADO 1"
COL_INT2       = "INTERESADO 2"
COL_OBS        = "OBSERVACIONES"

# ─── Comportamiento ───────────────────────────────────────────────────────────
# DRY_RUN = True  →  simula el proceso sin escribir ningún archivo en disco.
# Útil para validar los datos antes de una ejecución real.
DRY_RUN = False

# Detección de saltos de secuencia entre registros consecutivos.
# True  →  registros con salto de folio se marcan como pendientes y NO se procesan.
# False →  se emite solo un WARNING en el log pero el registro sí se procesa.
CHECK_SEQUENCE_GAPS = True
