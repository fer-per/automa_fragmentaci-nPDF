"""
folio_parser.py - Convierte notación de folios r/v a números de página PDF.

Cada hoja física tiene dos caras:
    recto (r) → página PDF = 2*N - 1
    vuelta (v) → página PDF = 2*N

Casos soportados:
    "1r"       → [1]
    "1v"       → [2]
    "1r-1v"    → [1, 2]
    "4r-6v"    → [7, 8, 9, 10, 11, 12]
    "7v-8r"    → [14, 15]
    "8v-12r"   → [16, 17, 18, 19, 20, 21, 22, 23]
    "13r-15v"  → [25, 26, 27, 28, 29, 30]
"""
import re
from typing import Optional

# Regex que captura: folio_inicio, cara_inicio y opcionalmente folio_fin, cara_fin
_FOLIO_RE = re.compile(
    r'^\s*(\d+)(r|v)(?:\s*-\s*(\d+)(r|v))?\s*$',
    re.IGNORECASE
)


def _folio_to_page(folio: int, cara: str) -> int:
    """Convierte un número de folio y cara a número de página PDF (1-indexed)."""
    cara = cara.lower()
    if cara == 'r':
        return 2 * folio - 1
    else:  # 'v'
        return 2 * folio


def parse_folio_range(texto: str) -> tuple[Optional[list[int]], Optional[str]]:
    """
    Parsea un rango de folios en notación r/v y devuelve la lista de páginas PDF.

    Returns:
        (pages, None)         si el formato es válido
        (None, error_msg)     si el formato es inválido o incoherente
    """
    if not texto or not str(texto).strip():
        return None, "Campo de folios vacío"

    texto = str(texto).strip()
    m = _FOLIO_RE.match(texto)
    if not m:
        return None, f"Formato de folios inválido: '{texto}'"

    f_ini = int(m.group(1))
    c_ini = m.group(2).lower()
    f_fin = int(m.group(3)) if m.group(3) else None
    c_fin = m.group(4).lower() if m.group(4) else None

    page_ini = _folio_to_page(f_ini, c_ini)

    # Caso: folio único (sin guión)
    if f_fin is None:
        return [page_ini], None

    page_fin = _folio_to_page(f_fin, c_fin)

    # Validar coherencia del rango
    if page_fin < page_ini:
        return None, (
            f"Rango incoherente: '{texto}' "
            f"(página inicio={page_ini} > página fin={page_fin})"
        )

    pages = list(range(page_ini, page_fin + 1))
    return pages, None


def last_page_of_range(texto: str) -> Optional[int]:
    """
    Devuelve la ÚLTIMA página PDF de un rango de folios.
    Útil para la detección de saltos de secuencia.
    Retorna None si el texto no es parseable.
    """
    pages, err = parse_folio_range(texto)
    if err or not pages:
        return None
    return pages[-1]


def first_page_of_range(texto: str) -> Optional[int]:
    """
    Devuelve la PRIMERA página PDF de un rango de folios.
    Útil para la detección de saltos de secuencia.
    Retorna None si el texto no es parseable.
    """
    pages, err = parse_folio_range(texto)
    if err or not pages:
        return None
    return pages[0]
