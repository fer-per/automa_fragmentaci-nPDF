import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, r'c:\Users\Alonso\Videos\automa_fragmentaci-nPDF')

from modules.folio_parser import (
    parse_folio_range, analyze_folio_sequence,
    folio_text_to_page_abs, _folio_to_page_abs, _find_segment,
)

# ── folio_text_to_page_abs ────────────────────────────────────────────────────
assert folio_text_to_page_abs("401r") == 801,  "FALLO 401r abs"
assert folio_text_to_page_abs("53v")  == 106,  "FALLO 53v abs"
assert folio_text_to_page_abs("1r")   == 1,    "FALLO 1r abs"
assert folio_text_to_page_abs("abc")  is None, "FALLO abc no es folio"
print("OK  folio_text_to_page_abs: 401r=801, 53v=106, 1r=1, abc=None")

# ── Segmentos basados en page_abs ─────────────────────────────────────────────
# Primera seg: 30r (abs=59) = pag1   Segunda: 401r (abs=801) = pag635
segs = [(_folio_to_page_abs(30, 'r'), 1), (_folio_to_page_abs(401, 'r'), 635)]
# = [(59, 1), (801, 635)]

p, e = parse_folio_range('30r', segments=segs)
assert p == [1] and e is None, f"FALLO 30r: {p}"
print(f"OK  30r -> pag {p[0]}")

p2, e2 = parse_folio_range('346v', segments=segs)
# abs(346v)=692, seg(59,1): offset=59-1=58, pdf=692-58=634
assert p2 == [634], f"FALLO 346v: {p2}"
print(f"OK  346v -> pag {p2[0]}")

p3, e3 = parse_folio_range('401r', segments=segs)
# abs(401r)=801, seg(801,635): offset=801-635=166, pdf=801-166=635
assert p3 == [635], f"FALLO 401r: {p3}"
print(f"OK  401r -> pag {p3[0]}")

p4, e4 = parse_folio_range('403v', segments=segs)
# abs(403v)=806, seg(801,635): pdf=806-166=640
assert p4 == [640], f"FALLO 403v: {p4}"
print(f"OK  403v -> pag {p4[0]}")

# Segmento ingresado como "401r" (como lo haría la GUI)
page_abs_401r = folio_text_to_page_abs("401r")  # 801
segs2 = [(_folio_to_page_abs(1,'r'), 1), (page_abs_401r, 635)]
p5, e5 = parse_folio_range('401r', segments=segs2)
assert p5 == [635], f"FALLO segs2 401r: {p5}"
print(f"OK  401r via folio_text_to_page_abs -> pag {p5[0]}")

# ── analyze_folio_sequence con índices de fila ────────────────────────────────
folios  = ['1r-1v', '2r-2v', '4r-4v']
indices = [10, 11, 15]   # numeros de fila del Excel
r = analyze_folio_sequence(folios, indices=indices)
assert not r['ok'], "FALLO: deberia haber gap"
assert r['gaps'][0]['row_id'] == 15, f"FALLO row_id: {r['gaps'][0]}"
# Detalle NO debe incluir lista de hojas faltantes, solo la fila y los folios
detalle = r['gaps'][0]['detalle']
assert '3r' not in detalle and '3v' not in detalle, f"FALLO: detalle no debe listar hojas: {detalle}"
assert 'Fila 15' in detalle, f"FALLO: debe indicar fila: {detalle}"
print(f"OK  gap con fila: '{detalle}'")

# Sucesion correcta
r2 = analyze_folio_sequence(['1r-1v','2r-2v','3r-3v'], indices=[10,11,12])
assert r2['ok'], "FALLO sucesion correcta"
print("OK  sucesion correcta confirmada")

# ── Numeros romanos ───────────────────────────────────────────────────────────
from modules.folder_builder import _roman_to_arabic
assert _roman_to_arabic("XVI") == "16"
assert _roman_to_arabic("abc") == "abc"
print("OK  roman_to_arabic: XVI->16, abc->abc")

# ── main.py importa ───────────────────────────────────────────────────────────
import main
print("OK  main.py importa sin errores")

print("\n=== TODOS LOS TESTS PASARON ===")
