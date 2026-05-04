"""
folder_builder.py - Construcción de la estructura de carpetas y nombres de archivo.

Jerarquía generada (11 niveles):
    output/
    └── ACERVO DOCUMENTAL NUMERO 7/          ← fila 7 del Excel (ej. «Código del fondo: N7»)
        └── SIGLO XVI/                        ← fila 4 del Excel (ej. «Seccion: XVI»)
            └── FONDO DOCUMENTAL/             ← literal fijo
                └── DIEGO DE AGUILAR/         ← Escribano/Notario
                    └── 1586/                 ← Año (extraído de fecha inicial)
                        └── PROTOCOLO 16/     ← N° de protocolo
                            └── REGISTRO 1/  ← N° de registro
                                └── PODER/   ← Titulo estandar (columna I)
                                    └── 1. ENERO/          ← Mes (formato «N. NOMBRE»)
                                        └── Interesado 1/  ← Solo el PRIMER nombre completo
                                            └── Interesado 2.pdf  ← Solo el PRIMER nombre completo

Regla de nombres de interesado:
    Si el campo contiene múltiples nombres separados por coma (ej. "juan perez, miguel serna"),
    solo se toma el PRIMER nombre completo ("juan perez") para crear la carpeta/archivo.
    El resto de nombres se ignoran en la construcción de rutas.

Ejemplo real:
    output/ACERVO DOCUMENTAL NUMERO 7/SIGLO XVI/FONDO DOCUMENTAL/
    DIEGO DE AGUILAR/1586/PROTOCOLO 16/REGISTRO 1/PODER/1. ENERO/
    Ramirianez de Sarabia/Antonio de Valderrama.pdf
"""
import re
import logging
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)

# Caracteres invalidos en nombres de archivo/carpeta en Windows
_INVALID_CHARS_RE = re.compile(r'[\\/:*?"<>|]')

# Tabla de valores romanos (de mayor a menor para el algoritmo greedy)
_ROMAN_TABLE = [
    ('M', 1000), ('CM', 900), ('D', 500), ('CD', 400),
    ('C', 100),  ('XC', 90),  ('L', 50),  ('XL', 40),
    ('X', 10),   ('IX', 9),   ('V', 5),   ('IV', 4),  ('I', 1),
]


def _roman_to_arabic(s: str) -> str:
    """
    Convierte un numero romano a arabigo como string.
    Si el valor no es un numero romano valido, devuelve el string original.
    Ejemplos: 'XVI' -> '16', 'IV' -> '4', 'abc' -> 'abc'
    """
    text = s.strip().upper()
    if not text:
        return s
    result = 0
    i = 0
    for numeral, value in _ROMAN_TABLE:
        while text[i:i + len(numeral)] == numeral:
            result += value
            i += len(numeral)
    # Solo es romano valido si consumimos todos los caracteres
    if i == len(text) and result > 0:
        return str(result)
    return s  # no es romano, devolver original

# Mapa de número de mes a nombre en formato «N. NOMBRE» (punto + mayúsculas)
_MONTH_NAMES = {
    1:  "1. ENERO",
    2:  "2. FEBRERO",
    3:  "3. MARZO",
    4:  "4. ABRIL",
    5:  "5. MAYO",
    6:  "6. JUNIO",
    7:  "7. JULIO",
    8:  "8. AGOSTO",
    9:  "9. SEPTIEMBRE",
    10: "10. OCTUBRE",
    11: "11. NOVIEMBRE",
    12: "12. DICIEMBRE",
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


def _first_full_name(raw: str) -> str:
    """
    Extrae el PRIMER nombre completo de un campo que puede contener varios nombres
    separados por coma.

    Regla:
        "juan perez, miguel serna"  →  "juan perez"
        "juan perez"                →  "juan perez"
        ""                          →  ""

    No se altera capitalización ni se elimina ningún carácter; solo se corta
    en la primera coma (si existe) y se hace strip del resultado.
    """
    if not raw or not str(raw).strip():
        return ""
    # Dividir por la primera coma
    parts = str(raw).split(",", 1)
    return parts[0].strip()


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
    acervo_num: str,
    siglo: str,
    escribano: str,
    protocolo: str,
    registro: str,
    titulo_est: str,
    fecha_ini: str,
    interesado1: str,
    interesado2: str,
    dry_run: bool = False,
) -> Path:
    """
    Construye y crea la ruta completa de destino para el PDF de un registro.

    Para interesado1 e interesado2 se utiliza SOLO el primer nombre completo
    (hasta la primera coma) cuando el campo contiene varios nombres separados
    por coma.

    Jerarquía de 11 niveles:
      1. ACERVO DOCUMENTAL NUMERO {acervo_num}
      2. SIGLO {siglo}
      3. FONDO DOCUMENTAL   (literal fijo)
      4. {escribano}
      5. {año}              (extraído de fecha_ini)
      6. PROTOCOLO {protocolo}
      7. REGISTRO {registro}
      8. {titulo_est}       (Titulo estandar, columna I)
      9. {N}. {MES}         (mes extraído de fecha_ini)
     10. {interesado1}      (carpeta — solo primer nombre completo)
     11. {interesado2}.pdf  (nombre del archivo — solo primer nombre completo)

    En caso de colisión de nombre, agrega sufijo _2, _3, etc.

    Returns:
        Path completo al archivo .pdf de destino.
    """
    year, month = _parse_date(fecha_ini)
    year_str  = str(year)                               if year  else "Sin_Año"
    month_str = _MONTH_NAMES.get(month, "Sin_Mes")      if month else "Sin_Mes"

    prot_str = sanitize(str(protocolo).strip(), "Sin_Protocolo")
    reg_str  = sanitize(str(registro).strip(),  "Sin_Registro")

    # ── Extraer solo el primer nombre completo para interesados ──
    int1_first = _first_full_name(interesado1)
    int2_first = _first_full_name(interesado2)

    if int1_first != interesado1.strip():
        logger.debug(
            f"Interesado1 recortado: '{interesado1.strip()}' -> '{int1_first}'"
        )
    if int2_first != interesado2.strip():
        logger.debug(
            f"Interesado2 recortado: '{interesado2.strip()}' -> '{int2_first}'"
        )

    # Convertir siglo de romano a arabigo (ej. XVI -> 16)
    siglo_display = _roman_to_arabic(sanitize(siglo, "Sin_Siglo"))

    folder_path = (
        output_dir
        / f"ACERVO DOCUMENTAL NUMERO {sanitize(acervo_num, 'Sin_Acervo')}"
        / f"SIGLO {siglo_display}"
        / "FONDO DOCUMENTAL"
        / sanitize(escribano, "Sin_Escribano")
        / year_str
        / f"PROTOCOLO {prot_str}"
        / f"REGISTRO {reg_str}"
        / sanitize(titulo_est, "Sin_Titulo")
        / month_str
        / sanitize(int1_first, "Sin_Interesado1")
    )

    base_name = sanitize(int2_first, "Sin_Interesado2")
    dest_path = folder_path / f"{base_name}.pdf"

    # Resolución de colisiones de nombre
    if dest_path.exists() and not dry_run:
        counter = 2
        while True:
            candidate = folder_path / f"{base_name}_{counter}.pdf"
            if not candidate.exists():
                dest_path = candidate
                logger.debug(f"Colision resuelta -> {dest_path.name}")
                break
            counter += 1

    if not dry_run:
        folder_path.mkdir(parents=True, exist_ok=True)

    return dest_path
