"""
data_cronica_analyzer.py - Analiza la columna DATA CRONICA (FECHA INICIAL y FECHA FINAL).

Verifica:
  1. Formato correcto: d/m/yyyy o dd/mm/yyyy.
  2. Coherencia interna: fecha_final >= fecha_inicial por registro.
  3. Progresion cronologica: las fechas iniciales deben ser no decrecientes
     (cada registro no deberia tener una fecha anterior a la del registro previo).
  4. Muestra exactamente que registros rompen el formato o la progresion.
"""
import re
from typing import Optional

# Patron de fecha valida: d/m/yyyy (acepta 1 o 2 digitos en dia/mes)
_DATE_RE = re.compile(r'^(\d{1,2})/(\d{1,2})/(\d{4})$')


def _parse_date(valor: str) -> Optional[tuple]:
    """
    Parsea una fecha en formato d/m/yyyy.
    Returns (dia, mes, anio) o None si no es parseable.
    """
    v = str(valor).strip() if valor else ""
    if not v or v.lower() == "nan":
        return None
    m = _DATE_RE.match(v)
    if not m:
        return None
    dia, mes, anio = int(m.group(1)), int(m.group(2)), int(m.group(3))
    # Validar rangos basicos
    if not (1 <= mes <= 12 and 1 <= dia <= 31 and anio >= 1000):
        return None
    return (anio, mes, dia)  # ordenar por (anio, mes, dia)


def analyze_data_cronica(
    registros: list,
    col_fecha_ini: str,
    col_fecha_fin: str = None,
    col_registro: str = None,
) -> dict:
    """
    Analiza la DATA CRONICA de una lista de registros.

    Args:
        registros:     Lista de dicts (filas del DataFrame).
        col_fecha_ini: Nombre de la columna FECHA INICIAL.
        col_fecha_fin: Nombre de la columna FECHA FINAL (opcional).
        col_registro:  Nombre de la columna N de REGISTRO (para reportes).

    Returns:
        Dict con:
          'ok'            : bool
          'invalid_fmt'   : list  - registros con formato de fecha invalido
          'incoherent'    : list  - registros donde fecha_fin < fecha_ini
          'non_progress'  : list  - registros donde la fecha va hacia atras
          'summary'       : str
          'total'         : int
          'ok_count'      : int
    """
    invalid_fmt   = []
    incoherent    = []
    non_progress  = []
    ok_count      = 0
    total         = 0

    prev_fecha_ini_parsed = None
    prev_reg_id           = None

    for idx, row in enumerate(registros):
        val_ini = row.get(col_fecha_ini, "")
        val_fin = row.get(col_fecha_fin, "") if col_fecha_fin else None
        reg_id  = str(row.get(col_registro, idx + 1)).strip() if col_registro else str(idx + 1)
        total  += 1

        parsed_ini = _parse_date(val_ini)
        parsed_fin = _parse_date(val_fin) if val_fin is not None else None

        # 1. Formato invalido en FECHA INICIAL
        ini_str = str(val_ini).strip() if val_ini else ""
        fin_str = str(val_fin).strip() if val_fin else ""

        if ini_str and ini_str.lower() != "nan" and parsed_ini is None:
            invalid_fmt.append({
                "index": idx,
                "reg_id": reg_id,
                "campo": "FECHA INICIAL",
                "valor": ini_str,
                "detalle": f"Reg. {reg_id}: FECHA INICIAL con formato invalido -> '{ini_str}' (esperado d/m/yyyy)",
            })

        # 2. Formato invalido en FECHA FINAL
        if col_fecha_fin and fin_str and fin_str.lower() != "nan" and parsed_fin is None:
            invalid_fmt.append({
                "index": idx,
                "reg_id": reg_id,
                "campo": "FECHA FINAL",
                "valor": fin_str,
                "detalle": f"Reg. {reg_id}: FECHA FINAL con formato invalido -> '{fin_str}' (esperado d/m/yyyy)",
            })

        # 3. Coherencia interna: FECHA FINAL >= FECHA INICIAL
        if parsed_ini and parsed_fin and parsed_fin < parsed_ini:
            incoherent.append({
                "index": idx,
                "reg_id": reg_id,
                "fecha_ini": ini_str,
                "fecha_fin": fin_str,
                "detalle": (
                    f"Reg. {reg_id}: FECHA FINAL ({fin_str}) es anterior a "
                    f"FECHA INICIAL ({ini_str})"
                ),
            })

        # 4. Progresion cronologica entre registros consecutivos
        if parsed_ini and prev_fecha_ini_parsed is not None:
            if parsed_ini < prev_fecha_ini_parsed:
                non_progress.append({
                    "index": idx,
                    "reg_id": reg_id,
                    "prev_reg_id": prev_reg_id,
                    "fecha_actual": ini_str,
                    "fecha_anterior": _tuple_to_str(prev_fecha_ini_parsed),
                    "detalle": (
                        f"Reg. {reg_id}: FECHA INICIAL ({ini_str}) es anterior a "
                        f"la del reg. {prev_reg_id} ({_tuple_to_str(prev_fecha_ini_parsed)}) "
                        f"-> regresion cronologica"
                    ),
                })

        if parsed_ini:
            prev_fecha_ini_parsed = parsed_ini
            prev_reg_id           = reg_id
            ok_count += 1

    ok = not invalid_fmt and not incoherent and not non_progress

    # Construir resumen
    lines = []
    if ok:
        lines.append(f"[OK]  DATA CRONICA correcta en los {total} registros.")
    else:
        lines.append(f"[>>]  Total registros analizados: {total}")
        lines.append(f"[OK]  Con fecha valida y progresiva: {ok_count}")
        if invalid_fmt:
            lines.append(f"[ERR] Formato invalido ({len(invalid_fmt)}):")
            for e in invalid_fmt[:20]:
                lines.append(f"    . {e['detalle']}")
            if len(invalid_fmt) > 20:
                lines.append(f"    ... y {len(invalid_fmt) - 20} mas")
        if incoherent:
            lines.append(f"[!!]  Fecha final anterior a fecha inicial ({len(incoherent)}):")
            for e in incoherent[:20]:
                lines.append(f"    . {e['detalle']}")
            if len(incoherent) > 20:
                lines.append(f"    ... y {len(incoherent) - 20} mas")
        if non_progress:
            lines.append(f"[!!]  Regresion cronologica ({len(non_progress)}):")
            for e in non_progress[:20]:
                lines.append(f"    . {e['detalle']}")
            if len(non_progress) > 20:
                lines.append(f"    ... y {len(non_progress) - 20} mas")

    return {
        "ok":           ok,
        "invalid_fmt":  invalid_fmt,
        "incoherent":   incoherent,
        "non_progress": non_progress,
        "summary":      "\n".join(lines),
        "total":        total,
        "ok_count":     ok_count,
    }


def _tuple_to_str(t: tuple) -> str:
    """Convierte (anio, mes, dia) a formato d/m/yyyy."""
    return f"{t[2]}/{t[1]}/{t[0]}"
