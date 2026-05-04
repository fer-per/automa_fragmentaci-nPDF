# -*- coding: utf-8 -*-
"""
folio_parser.py - Convierte notacion de folios r/v a numeros de pagina PDF.

Cada hoja fisica tiene dos caras:
    recto (r)  -> pagina PDF = 2*N - 1  (relativo al folio 1)
    vuelta (v) -> pagina PDF = 2*N

Soporte de offset:
    Si el PDF no empieza en el folio 1 sino en el folio F_inicio,
    se resta (F_inicio - 1) hojas al calculo.

    pdf_page = folio_page_absoluto - folio_page_de_F_inicio + page_pdf_inicio

Soporte de multiples segmentos:
    Cuando el PDF tiene saltos (algunos folios no estan escaneados),
    se pueden definir varios segmentos (page_abs_inicio, pdf_page_inicio)
    para mapear correctamente cada seccion del protocolo.

Casos soportados (sin offset):
    "1r"       -> [1]
    "1v"       -> [2]
    "1r-1v"    -> [1, 2]
    "4r-6v"    -> [7, 8, 9, 10, 11, 12]

Con offset (folio_inicio_excel=30, pdf_page_inicio=1):
    "30r"      -> [1]   (primera imagen del PDF)
    "30v"      -> [2]
    "31r-31v"  -> [3, 4]
"""
import re
from typing import Optional

# Regex: captura folio_inicio, cara_inicio y opcionalmente folio_fin, cara_fin
_FOLIO_RE = re.compile(
    r'^\s*(\d+)(r|v)(?:\s*-\s*(\d+)(r|v))?\s*$',
    re.IGNORECASE
)


def _folio_to_page_abs(folio: int, cara: str) -> int:
    # Convierte folio+cara a numero de pagina ABSOLUTO (sin offset).
    cara = cara.lower()
    if cara == 'r':
        return 2 * folio - 1
    else:
        return 2 * folio


# Tipo de segmento interno: (page_abs_inicio, pdf_page_inicio)
# page_abs_inicio: pagina absoluta del folio que inicia el segmento (ej. "401r" -> 801)
# Esto permite comparar con la pagina absoluta del folio que se quiere resolver.
Segments = list  # list[tuple[int, int]]


def folio_text_to_page_abs(texto: str):
    """
    Convierte una cadena de folio con cara ('401r', '53v', etc.) a su pagina absoluta.
    Retorna int si es valido, None si el formato no es reconocido.
    """
    texto = str(texto).strip()
    m = _FOLIO_RE.match(texto)
    if not m:
        return None
    f = int(m.group(1))
    c = m.group(2).lower()
    return _folio_to_page_abs(f, c)


def _find_segment(page_abs: int, segments: list) -> tuple:
    """
    Dado la pagina absoluta de un folio, devuelve el segmento aplicable
    (page_abs_inicio, pdf_page_inicio).
    El segmento aplicable es el de mayor page_abs_inicio que sea <= page_abs.
    """
    best = segments[0] if segments else (1, 1)
    for seg in sorted(segments, key=lambda x: x[0]):
        if page_abs >= seg[0]:
            best = seg
        else:
            break
    return best


def _apply_offset(page_abs: int, folio_inicio_excel: int, pdf_page_inicio: int) -> int:
    # Transforma una pagina absoluta a la pagina real dentro del PDF usando offset simple.
    first_abs = _folio_to_page_abs(folio_inicio_excel, 'r')
    offset = first_abs - pdf_page_inicio
    return page_abs - offset


def folio_to_page(folio: int, cara: str,
                  folio_inicio_excel: int = 1,
                  pdf_page_inicio: int = 1,
                  segments: list = None) -> int:
    """
    Convierte un numero de folio y cara a numero de pagina PDF (1-indexed).

    Si se proporcionan 'segments' (lista de (page_abs_inicio, pdf_page_inicio)),
    se busca el segmento aplicable por pagina absoluta y se aplica su offset.
    De lo contrario usa folio_inicio_excel y pdf_page_inicio como fallback.
    """
    page_abs = _folio_to_page_abs(folio, cara)
    if segments:
        seg_page_abs_start, seg_pdf_start = _find_segment(page_abs, segments)
        return page_abs - (seg_page_abs_start - seg_pdf_start)
    return _apply_offset(page_abs, folio_inicio_excel, pdf_page_inicio)


def parse_folio_range(
    texto: str,
    folio_inicio_excel: int = 1,
    pdf_page_inicio: int = 1,
    segments: list = None,
) -> tuple:
    """
    Parsea un rango de folios en notacion r/v y devuelve la lista de paginas PDF.

    Calcula la pagina PDF de CADA folio individualmente usando el segmento correcto.
    Esto garantiza que:
      - Los registros que cruzan fronteras de segmento se mapeen correctamente.
      - La lista resultante no contiene duplicados.
      - Las paginas siempre estan en orden creciente dentro del rango.

    Args:
        texto:              Texto de folios, ej. "30r-31v".
        folio_inicio_excel: Numero de folio donde comienza el protocolo (fallback si no hay segments).
        pdf_page_inicio:    Pagina PDF de inicio (fallback si no hay segments).
        segments:           Lista de (page_abs_ini, pdf_pag_ini) para mapeo multi-segmento.

    Returns:
        (pages, None)     si el formato es valido
        (None, error_msg) si el formato es invalido o incoherente
    """
    if not texto or not str(texto).strip():
        return None, "Campo de folios vacio"

    texto = str(texto).strip()
    m = _FOLIO_RE.match(texto)
    if not m:
        return None, f"Formato de folios invalido: '{texto}'"

    f_ini = int(m.group(1))
    c_ini = m.group(2).lower()
    f_fin = int(m.group(3)) if m.group(3) else None
    c_fin = m.group(4).lower() if m.group(4) else None

    # Caso de folio simple (sin rango)
    if f_fin is None:
        page_ini = folio_to_page(f_ini, c_ini, folio_inicio_excel, pdf_page_inicio, segments)
        return [page_ini], None

    abs_ini = _folio_to_page_abs(f_ini, c_ini)
    abs_fin = _folio_to_page_abs(f_fin, c_fin)

    if abs_fin < abs_ini:
        return None, (
            f"Rango incoherente: '{texto}' "
            f"(folio fin {f_fin}{c_fin} < folio inicio {f_ini}{c_ini})"
        )

    # Computar la pagina PDF de cada folio usando su segmento correcto.
    # Esto maneja correctamente rangos que cruzan fronteras de segmento.
    pages = []
    seen  = set()
    for abs_page in range(abs_ini, abs_fin + 1):
        if segments:
            seg_abs, seg_pdf = _find_segment(abs_page, segments)
            pdf_page = abs_page - (seg_abs - seg_pdf)
        else:
            first_abs = _folio_to_page_abs(folio_inicio_excel, 'r')
            pdf_page  = abs_page - (first_abs - pdf_page_inicio)

        if pdf_page < 1:
            continue  # pagina fuera de rango (offset incorrecto)

        # Evitar duplicados en caso de saltos de segmento dentro del rango
        if pdf_page not in seen:
            pages.append(pdf_page)
            seen.add(pdf_page)

    if not pages:
        return None, f"El rango '{texto}' no produce paginas PDF validas (revisa el offset/segmentos)"

    # Ordenar ascendentemente para garantizar secuencia correcta en el PDF de salida.
    # Esto maneja rangos que cruzan fronteras de segmento (las paginas pueden ser
    # no-monoticas antes del ordenamiento si el segmento reinicia a un numero menor).
    pages.sort()

    return pages, None


def adjust_pages_for_missing_folios(
    pages: list,
    ignored_folio_pages: set,
    active_seg=None,
    segments: list = None,
) -> list:
    """
    Ajusta los numeros de pagina PDF por folios fisicamente ausentes del escaneo.

    El shift SOLO se aplica a las paginas ignoradas que pertenecen al mismo segmento
    que el registro actual (active_seg). Las paginas de segmentos nuevos ya tienen
    una posicion fisica definida explicitamente y no se deben desplazar.

    Logica de pertenencia: un folio ignorado con valor nominal 'ign' pertenece al
    segmento (seg_abs, seg_pdf) si:
        abs_equivalente = ign + (seg_abs - seg_pdf)
        seg_abs <= abs_equivalente < siguiente_seg_abs
    """
    if not ignored_folio_pages:
        return list(pages)

    # Determinar que paginas ignoradas son relevantes para el segmento activo
    if active_seg is not None and segments and len(segments) > 1:
        seg_abs, seg_pdf = active_seg
        sorted_segs = sorted(segments, key=lambda x: x[0])
        try:
            idx = sorted_segs.index(active_seg)
        except ValueError:
            idx = 0
        next_seg_abs = sorted_segs[idx + 1][0] if idx + 1 < len(sorted_segs) else float('inf')

        relevant_ignored = set()
        for ign in ignored_folio_pages:
            implied_abs = ign + (seg_abs - seg_pdf)
            if seg_abs <= implied_abs < next_seg_abs:
                relevant_ignored.add(ign)
    else:
        relevant_ignored = set(ignored_folio_pages)

    if not relevant_ignored:
        return list(pages)   # Ningun folio ignorado afecta este segmento

    ignored_sorted = sorted(relevant_ignored)
    result = []
    for nominal_page in pages:
        if nominal_page in relevant_ignored:
            continue  # Hoja ausente del escaneo
        shift = sum(1 for ign in ignored_sorted if ign < nominal_page)
        physical_page = nominal_page - shift
        if physical_page >= 1:
            result.append(physical_page)

    return result



def last_page_of_range(
    texto: str,
    folio_inicio_excel: int = 1,
    pdf_page_inicio: int = 1,
    segments: list = None,
) -> Optional[int]:
    # Devuelve la ULTIMA pagina PDF de un rango de folios.
    pages, err = parse_folio_range(texto, folio_inicio_excel, pdf_page_inicio, segments)
    if err or not pages:
        return None
    return pages[-1]


def first_page_of_range(
    texto: str,
    folio_inicio_excel: int = 1,
    pdf_page_inicio: int = 1,
    segments: list = None,
) -> Optional[int]:
    # Devuelve la PRIMERA pagina PDF de un rango de folios.
    pages, err = parse_folio_range(texto, folio_inicio_excel, pdf_page_inicio, segments)
    if err or not pages:
        return None
    return pages[0]


def analyze_folio_sequence(folio_list: list, indices: list = None) -> dict:
    """
    Analiza la sucesion de una lista de cadenas de folios del Excel.

    Detecta:
      - Entradas con formato invalido.
      - Saltos de secuencia entre registros consecutivos.
      - Solapamientos entre registros.

    Args:
        folio_list: Lista de cadenas tal como aparecen en el Excel.
        indices:    Lista opcional de identificadores de fila (ej. numeros de fila del Excel).
                    Si se proporciona, se muestra en el reporte de cada problema.

    Returns:
        Dict con 'ok', 'gaps', 'overlaps', 'invalid', 'summary'.
    """
    gaps            = []
    overlaps        = []
    invalid_entries = []

    prev_last_folio = None
    prev_last_cara  = None
    prev_texto      = None
    prev_row_id     = None

    for idx, texto in enumerate(folio_list):
        row_id  = indices[idx] if (indices and idx < len(indices)) else (idx + 1)
        texto_s = str(texto).strip() if texto else ""
        if not texto_s or texto_s.lower() == "nan":
            continue

        m = _FOLIO_RE.match(texto_s)
        if not m:
            invalid_entries.append({
                "index":  idx,
                "row_id": row_id,
                "texto":  texto_s,
                "detalle": f"Fila {row_id}: Formato invalido -> '{texto_s}'",
            })
            prev_last_folio = None
            prev_last_cara  = None
            prev_texto      = texto_s
            prev_row_id     = row_id
            continue

        f_ini = int(m.group(1))
        c_ini = m.group(2).lower()
        f_fin = int(m.group(3)) if m.group(3) else f_ini
        c_fin = m.group(4).lower() if m.group(4) else c_ini

        page_ini_abs = _folio_to_page_abs(f_ini, c_ini)
        page_fin_abs = _folio_to_page_abs(f_fin, c_fin)

        if page_fin_abs < page_ini_abs:
            invalid_entries.append({
                "index":  idx,
                "row_id": row_id,
                "texto":  texto_s,
                "detalle": f"Fila {row_id}: Rango incoherente '{texto_s}' (fin < inicio)",
            })
            prev_last_folio = None
            prev_last_cara  = None
            prev_texto      = texto_s
            prev_row_id     = row_id
            continue

        if prev_last_folio is not None:
            expected_page = _folio_to_page_abs(prev_last_folio, prev_last_cara) + 1
            actual_page   = page_ini_abs

            if actual_page > expected_page:
                gaps.append({
                    "index":       idx,
                    "row_id":      row_id,
                    "prev_row_id": prev_row_id,
                    "prev_texto":  prev_texto,
                    "curr_texto":  texto_s,
                    "detalle": f"Fila {row_id}: Salto de '{prev_texto}' a '{texto_s}'",
                })
            elif actual_page < expected_page:
                overlaps.append({
                    "index":       idx,
                    "row_id":      row_id,
                    "prev_row_id": prev_row_id,
                    "prev_texto":  prev_texto,
                    "curr_texto":  texto_s,
                    "detalle": f"Fila {row_id}: Solapamiento de '{prev_texto}' con '{texto_s}'",
                })

        prev_last_folio = f_fin
        prev_last_cara  = c_fin
        prev_texto      = texto_s
        prev_row_id     = row_id

    ok = not gaps and not overlaps and not invalid_entries

    # Construir resumen
    lines = []
    if ok:
        lines.append("[OK]  Sucesion de folios correcta. No se detectaron saltos ni solapamientos.")
    else:
        if invalid_entries:
            lines.append(f"[!!]  {len(invalid_entries)} entrada(s) con formato invalido:")
            for e in invalid_entries:
                lines.append(f"    . {e['detalle']}")
        if gaps:
            lines.append(f"[!!]  {len(gaps)} salto(s) detectado(s):")
            for g in gaps:
                lines.append(f"    . {g['detalle']}")
        if overlaps:
            lines.append(f"[!!]  {len(overlaps)} solapamiento(s) detectado(s):")
            for o in overlaps:
                lines.append(f"    . {o['detalle']}")

    return {
        "ok":       ok,
        "gaps":     gaps,
        "overlaps": overlaps,
        "invalid":  invalid_entries,
        "summary":  "\n".join(lines),
    }


def _page_abs_to_folio_label(page_abs: int) -> str:
    # Convierte una pagina absoluta (1-based) a etiqueta de folio: Nr o Nv.
    folio = (page_abs + 1) // 2
    cara  = "r" if page_abs % 2 == 1 else "v"
    return f"{folio}{cara}"


def _pages_to_folio_labels(pages: list) -> str:
    # Convierte lista de paginas absolutas a cadena legible, ej. 3r, 3v, 4r.
    if not pages:
        return ""
    labels = [_page_abs_to_folio_label(p) for p in pages]
    return ", ".join(labels)
