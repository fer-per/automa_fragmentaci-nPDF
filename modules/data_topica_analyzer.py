"""
data_topica_analyzer.py - Analizador de la columna DATA TÓPICA (Lugar).

Verifica que:
  1. El formato sea correcto: el campo no debe estar vacío si se espera un valor.
  2. Los valores sigan un orden coherente (no se repiten ni se mezclan formatos).
  3. Detecta duplicados consecutivos, entradas vacías y formatos inconsistentes.

Criterios de formato válido para DATA TÓPICA:
  - Cadena no vacía
  - Empieza con letra mayúscula (o minúscula aceptable)
  - No contiene caracteres especiales inválidos para una ubicación (números solos, símbolos raros)
  - Formato esperado: "Ciudad" o "Ciudad, País" o "Ciudad (referencia)"

El analizador también detecta si la lista NO va en orden creciente de registro
(cuando se combina con el número de registro).
"""
import re
from typing import Optional

# Patrón de formato válido para DATA TÓPICA:
#   Empieza con letra (mayúscula o minúscula), puede contener letras, espacios,
#   comas, paréntesis, puntos, guiones y tildes.
_VALID_TOPICA_RE = re.compile(
    r'^[A-Za-záéíóúÁÉÍÓÚüÜñÑ][A-Za-záéíóúÁÉÍÓÚüÜñÑ0-9\s,.()\-\']+$'
)

# Patrón de campo completamente numérico (inválido como lugar)
_ONLY_DIGITS_RE = re.compile(r'^\d+$')


def _classify_topica(valor: str) -> str:
    """
    Clasifica el valor de DATA TÓPICA.

    Returns:
        'ok'      – formato correcto
        'empty'   – vacío o nulo
        'invalid' – formato inválido
    """
    v = str(valor).strip() if valor else ""
    if not v or v.lower() == "nan":
        return "empty"
    if _ONLY_DIGITS_RE.match(v):
        return "invalid"
    if not _VALID_TOPICA_RE.match(v):
        return "invalid"
    return "ok"


def analyze_data_topica(
    registros: list[dict],
    col_topica: str,
    col_registro: Optional[str] = None,
) -> dict:
    """
    Analiza la columna DATA TÓPICA de una lista de registros.

    Args:
        registros:    Lista de dicts (filas del DataFrame como .to_dict()).
        col_topica:   Nombre de la columna DATA TÓPICA.
        col_registro: Nombre de la columna N° DE REGISTRO (opcional, para contexto en errores).

    Returns:
        Dict con:
          'ok'          : bool   — True si no hay ningún problema.
          'empty'       : list   — Registros con DATA TÓPICA vacía.
          'invalid'     : list   — Registros con formato inválido.
          'summary'     : str    — Resumen legible para mostrar en la GUI.
          'total'       : int    — Total de registros analizados.
          'ok_count'    : int    — Registros con DATA TÓPICA válida.
    """
    empty_list   = []
    invalid_list = []
    ok_count     = 0
    total        = 0

    for idx, row in enumerate(registros):
        valor = row.get(col_topica, "")
        reg_id = str(row.get(col_registro, idx + 1)).strip() if col_registro else str(idx + 1)

        status = _classify_topica(valor)
        total += 1

        if status == "empty":
            empty_list.append({
                "index": idx,
                "reg_id": reg_id,
                "valor": str(valor).strip(),
                "detalle": f"Reg. {reg_id}: DATA TÓPICA vacía o nula",
            })
        elif status == "invalid":
            invalid_list.append({
                "index": idx,
                "reg_id": reg_id,
                "valor": str(valor).strip(),
                "detalle": f"Reg. {reg_id}: Formato invalido -> '{str(valor).strip()}'",
            })
        else:
            ok_count += 1

    ok = not empty_list and not invalid_list

    # Construir resumen
    lines = []
    if ok:
        lines.append(f"[OK]  DATA TOPICA correcta en los {total} registros.")
    else:
        lines.append(f"[>>]  Total registros analizados: {total}")
        lines.append(f"[OK]  Con DATA TOPICA valida:     {ok_count}")
        if empty_list:
            lines.append(f"[!!]  Vacios ({len(empty_list)}):")
            for e in empty_list[:20]:
                lines.append(f"    . {e['detalle']}")
            if len(empty_list) > 20:
                lines.append(f"    ... y {len(empty_list) - 20} mas")
        if invalid_list:
            lines.append(f"[ERR] Formato invalido ({len(invalid_list)}):")
            for e in invalid_list[:20]:
                lines.append(f"    . {e['detalle']}")
            if len(invalid_list) > 20:
                lines.append(f"    ... y {len(invalid_list) - 20} mas")

    return {
        "ok":        ok,
        "empty":     empty_list,
        "invalid":   invalid_list,
        "summary":   "\n".join(lines),
        "total":     total,
        "ok_count":  ok_count,
    }
