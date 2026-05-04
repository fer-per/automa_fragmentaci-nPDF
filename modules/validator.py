"""
validator.py - Validación de registros antes del procesamiento.

Detecta:
- Folios vacíos
- Formato de folio inválido
- Páginas fuera del rango del PDF
- Saltos de secuencia entre registros consecutivos

Soporta offset de inicio del PDF:
  folio_inicio_excel  →  número de folio donde inicia el protocolo en el Excel
  pdf_page_inicio     →  número de página PDF donde empieza la primera imagen
"""
from typing import Optional
from modules.folio_parser import parse_folio_range, first_page_of_range, last_page_of_range


def validate_record(
    row: dict,
    col_folios: str,
    total_pdf_pages: int,
    prev_last_page: Optional[int] = None,
    check_gaps: bool = True,
    folio_inicio_excel: int = 1,
    pdf_page_inicio: int = 1,
    segments: list = None,
) -> tuple[bool, str]:
    """
    Valida un registro del Excel antes de procesarlo.

    Args:
        row:                Diccionario con los datos del registro (una fila del DataFrame).
        col_folios:         Nombre de la columna que contiene el rango de folios.
        total_pdf_pages:    Total de páginas del PDF fuente.
        prev_last_page:     Última página del registro anterior (para detección de saltos).
        check_gaps:         Si True, activa la detección de saltos de secuencia.
        folio_inicio_excel: Número de folio donde comienza el protocolo en el Excel.
        pdf_page_inicio:    Número de página PDF donde empieza la primera imagen del protocolo.

    Returns:
        (True, "")            si el registro es válido
        (False, error_msg)    si hay algún problema
    """
    folio_val = row.get(col_folios)

    # 1. Folios vacíos / nulos
    if folio_val is None or str(folio_val).strip() == '' or str(folio_val).strip().lower() == 'nan':
        return False, "Folios vacíos o nulos"

    # 2. Formato inválido + cálculo de páginas con offset
    pages, parse_err = parse_folio_range(
        str(folio_val),
        folio_inicio_excel=folio_inicio_excel,
        pdf_page_inicio=pdf_page_inicio,
        segments=segments,
    )
    if parse_err:
        return False, parse_err

    # 3. Páginas fuera del rango del PDF
    max_page = max(pages)
    min_page = min(pages)
    if max_page > total_pdf_pages:
        return False, (
            f"Folio '{folio_val}' requiere página {max_page} "
            f"pero el PDF solo tiene {total_pdf_pages} páginas"
        )
    if min_page < 1:
        return False, (
            f"Folio '{folio_val}' resulta en número de página < 1 "
            f"(revisa el folio de inicio del Excel y la página de inicio del PDF)"
        )

    # 4. Detección de salto de secuencia
    if check_gaps and prev_last_page is not None:
        current_first = first_page_of_range(
            str(folio_val),
            folio_inicio_excel=folio_inicio_excel,
            pdf_page_inicio=pdf_page_inicio,
            segments=segments,
        )
        if current_first is not None:
            expected_next = prev_last_page + 1
            if current_first != expected_next:
                missing_pages = list(range(expected_next, current_first))
                return False, (
                    f"Salto de secuencia detectado: "
                    f"se esperaba página {expected_next} "
                    f"pero se encontró página {current_first} "
                    f"(folios '{folio_val}'). "
                    f"Páginas faltantes: {missing_pages}"
                )

    return True, ""
