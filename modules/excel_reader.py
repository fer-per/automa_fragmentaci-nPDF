"""
excel_reader.py - Lectura y normalización del archivo Excel (.xlsx).

Maneja:
- Encabezados institucionales (skiprows)
- Columnas con sub-encabezados (FECHA INICIAL / FECHA FINAL dentro de DATA CRONICA)
- Filas completamente vacías
"""
import logging
from pathlib import Path
from typing import Optional
import pandas as pd

logger = logging.getLogger(__name__)


def _normalize_col(name: str) -> str:
    """Normaliza un nombre de columna: strip de espacios y minúsculas para comparación."""
    return str(name).strip()


def load_excel(
    excel_path: Path,
    skip_rows: int = 7,
    col_fecha_ini: str = "FECHA INICIAL",
) -> pd.DataFrame:
    """
    Carga el Excel y devuelve un DataFrame limpio y listo para procesar.

    El índice del DataFrame corresponde al número de fila REAL del Excel
    (1-indexed, tal como lo ve el usuario en la hoja de cálculo).
    Esto garantiza que la GUI muestre los mismos números de fila que el Excel.

    Estructura del Excel:
      - Filas 1..skip_rows: encabezados institucionales (se omiten)
      - Fila skip_rows+1: encabezado real de columnas
      - Filas skip_rows+2 en adelante: datos
      - "DATA CRONICA" tiene dos sub-columnas: FECHA INICIAL y FECHA FINAL

    Args:
        excel_path:   Ruta al archivo .xlsx
        skip_rows:    Número de filas a saltar (0-indexed) antes del header.
        col_fecha_ini: Nombre a asignar a la columna de fecha inicial.

    Returns:
        DataFrame con columnas normalizadas, filas vacías eliminadas,
        y el índice = número de fila real en el Excel (1-indexed).
    """
    path = Path(excel_path)
    if not path.exists():
        raise FileNotFoundError(f"Archivo Excel no encontrado: {path}")

    logger.info(f"Cargando Excel: {path}")

    df = pd.read_excel(
        path,
        skiprows=skip_rows,
        header=0,
        engine="openpyxl",
        dtype=str,           # Todo como string; evita conversiones numéricas inesperadas
    )

    # Normalizar nombres de columna (strip)
    df.columns = [_normalize_col(c) for c in df.columns]

    # Eliminar columnas completamente sin nombre (Unnamed: X)
    unnamed = [c for c in df.columns if str(c).startswith("Unnamed:")]
    if unnamed:
        df.drop(columns=unnamed, inplace=True)

    # El Excel tiene dos sub-columnas bajo "DATA CRONICA".
    # Tras el drop de Unnamed, DATA CRONICA puede quedar como una sola columna
    # con FECHA INICIAL en su posición y FECHA FINAL perdida, o puede duplicarse.
    # Renombramos si la columna "DATA CRONICA" duplicada aparece como "DATA CRONICA.1"
    rename_map = {}
    if "DATA CRONICA" in df.columns and "DATA CRONICA.1" in df.columns:
        rename_map["DATA CRONICA"] = col_fecha_ini
        rename_map["DATA CRONICA.1"] = "FECHA FINAL"
    elif "DATA CRONICA" in df.columns:
        rename_map["DATA CRONICA"] = col_fecha_ini
    df.rename(columns=rename_map, inplace=True)

    # Asignar el número de fila REAL del Excel como índice.
    # skiprows filas se omitieron + 1 fila de header = skip_rows + 1 filas antes de datos.
    # La primera fila de datos en el Excel es skip_rows + 2 (1-indexed).
    # El DataFrame index 0 corresponde a esa fila.
    df.index = df.index + skip_rows + 2  # ahora df.index = fila real del Excel

    # Eliminar filas completamente vacías (conservando el índice real)
    df.dropna(how="all", inplace=True)

    # Reemplazar 'nan' string por vacío
    df = df.fillna("")

    logger.info(f"Excel cargado: {len(df)} registros (filas {df.index.min()}–{df.index.max()} del Excel)")
    return df
