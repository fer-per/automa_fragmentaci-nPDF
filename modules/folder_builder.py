"""
folder_builder.py - Construcción de la estructura de carpetas y nombres de archivo.

Estructura generada:
    output/
    └── [Escribano]/
        └── [Año]/
            └── Protocolo N° [X]/
                └── Registro N° [X]/
                    └── [Tipo de Escritura]/
                        └── [N- Mes]/
                            └── [Interesado 1]/
                                └── [Interesado 2].pdf

Ejemplo:
    output/PORTUGAL, Cesar/1567/Protocolo N° 1/Registro N° 5/Obligacion/5- Mayo/Diego de Serna/Antonio de Barco.pdf
"""
import re
import logging
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)

# Caracteres inválidos en nombres de archivo/carpeta en Windows
_INVALID_CHARS_RE = re.compile(r'[\\/:*?"<>|]')

# Mapa de número de mes a nombre (para fechas en formato dd/mm/yyyy)
_MONTH_NAMES = {
    1: "1- Enero", 2: "2- Febrero", 3: "3- Marzo", 4: "4- Abril",
    5: "5- Mayo", 6: "6- Junio", 7: "7- Julio", 8: "8- Agosto",
    9: "9- Septiembre", 10: "10- Octubre", 11: "11- Noviembre", 12: "12- Diciembre",
}


def sanitize(name: str, fallback: str = "Sin_Nombre") -> str:
    """
    Elimina caracteres inválidos para nombres de carpeta/archivo en Windows.
    También recorta espacios y puntos al final (inválido en Windows).
    """
    if not name or not str(name).strip():
        return fallback
    cleaned = _INVALID_CHARS_RE.sub("_", str(name).strip())
    cleaned = cleaned.rstrip(". ")
    return cleaned or fallback


def _parse_date(fecha_str: str) -> tuple[Optional[int], Optional[int]]:
    """
    Intenta parsear una fecha en formato 'd/m/yyyy' o 'dd/mm/yyyy'.
    Retorna (año, mes) o (None, None) si no es parseable.
    """
    try:
        parts = str(fecha_str).strip().split("/")
        if len(parts) == 3:
            day, month, year = int(parts[0]), int(parts[1]), int(parts[2])
            return year, month
    except (ValueError, IndexError):
        pass
    return None, None


def build_output_path(
    output_dir: Path,
    escribano: str,
    protocolo: str,
    registro: str,
    titulo: str,
    fecha_ini: str,
    interesado1: str,
    interesado2: str,
    dry_run: bool = False,
) -> Path:
    """
    Construye y crea la ruta completa de destino para el PDF de un registro.

    Jerarquía: Escribano / Año / Protocolo N° X / Registro N° X / Tipo / Mes / Int1 / Int2.pdf

    Usa colisión de nombre: si el archivo ya existe, agrega sufijo _2, _3, etc.

    Returns:
        Path completo al archivo .pdf de destino.
    """
    year, month = _parse_date(fecha_ini)
    year_str  = str(year) if year else "Sin_Año"
    month_str = _MONTH_NAMES.get(month, "Sin_Mes") if month else "Sin_Mes"
    prot_str  = sanitize(protocolo, "Sin_Protocolo")
    reg_str   = sanitize(registro, "Sin_Registro")
    folder_path = (
        output_dir
        / sanitize(escribano)
        / year_str
        / f"Protocolo N° {prot_str}"
        / f"Registro N° {reg_str}"
        / sanitize(titulo, "Sin_Tipo")
        / month_str
        / sanitize(interesado1, "Sin_Interesado1")
    )

    base_name = sanitize(interesado2, "Sin_Interesado2")
    dest_path = folder_path / f"{base_name}.pdf"

    # Resolución de colisiones
    if dest_path.exists() and not dry_run:
        counter = 2
        while True:
            candidate = folder_path / f"{base_name}_{counter}.pdf"
            if not candidate.exists():
                dest_path = candidate
                logger.debug(f"Colisión resuelta → {dest_path.name}")
                break
            counter += 1

    if not dry_run:
        folder_path.mkdir(parents=True, exist_ok=True)

    return dest_path
