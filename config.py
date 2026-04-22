"""
config.py - Configuración centralizada del sistema de automatización archivística.
"""
from pathlib import Path

# ─── Rutas de entrada ────────────────────────────────────────────────────────
# Cambia estas rutas para apuntar a tus archivos reales antes de ejecutar.
BASE_DIR    = Path(__file__).parent
EXCEL_PATH  = BASE_DIR / "inventario.xlsx"   # Archivo Excel .xlsx de entrada
PDF_PATH    = BASE_DIR / "documento.pdf"     # PDF escaneado de entrada

# ─── Carpetas de salida ───────────────────────────────────────────────────────
OUTPUT_DIR  = BASE_DIR / "output"            # Raíz donde se crean las carpetas
LOGS_DIR    = BASE_DIR / "logs"

PROCESS_LOG    = LOGS_DIR / "process.log"
PENDIENTES_CSV = LOGS_DIR / "pendientes.csv"

# ─── Parámetros del Excel ─────────────────────────────────────────────────────
# Las primeras 8 filas del Excel son encabezados institucionales;
# la fila 8 (índice 7) es el encabezado real de columnas.
SKIP_ROWS = 7   # skiprows para pandas (0-indexed: salta filas 0-7, usa fila 8 como header)

# Nombres de columna esperados (después de normalizar).
# Si tu Excel usa nombres distintos, ajusta aquí.
COL_REGISTRO  = "N° DE REGISTRO"
COL_ESCRIBANO = "ESCRIBANO/NOTARIO"
COL_PROTOCOLO = "N° DE PROT."
COL_FOLIOS    = "N° DE FOLIOS"
COL_LUGAR     = "DATA TOPICA(Lugar)"
COL_FECHA_INI = "FECHA INICIAL"    # sub-columna dentro de DATA CRONICA
COL_TITULO    = "TITULO/ESCRITURA"
COL_INT1      = "INTERESADO 1"
COL_INT2      = "INTERESADO 2"
COL_OBS       = "OBSERVACIONES"

# ─── Comportamiento ───────────────────────────────────────────────────────────
# DRY_RUN = True  →  simula el proceso sin escribir ningún archivo en disco.
# Útil para validar los datos antes de una ejecución real.
DRY_RUN = False

# Detección de saltos de secuencia entre registros consecutivos.
# True  →  registros con salto de folio se marcan como pendientes y NO se procesan.
# False →  se emite solo un WARNING en el log pero el registro sí se procesa.
CHECK_SEQUENCE_GAPS = True
