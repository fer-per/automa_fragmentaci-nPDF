"""
generate_test_pdf.py - Genera un PDF sintético de prueba con N páginas numeradas.

Uso:
    python generate_test_pdf.py                  # genera 384 páginas (192 hojas r+v)
    python generate_test_pdf.py --hojas 100      # genera 200 páginas (100 hojas × 2)
    python generate_test_pdf.py --pages 50       # genera 50 páginas directamente
    python generate_test_pdf.py --output mi.pdf  # nombre de salida personalizado

Requiere: reportlab
    pip install reportlab
"""
import argparse
from pathlib import Path

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas


def generate_test_pdf(output_path: Path, num_pages: int) -> None:
    """
    Genera un PDF de prueba con `num_pages` páginas.
    Cada página muestra su número (1-indexed) y el folio equivalente.
    """
    c = canvas.Canvas(str(output_path), pagesize=A4)
    width, height = A4

    for page_num in range(1, num_pages + 1):
        # Calcular folio y cara equivalentes
        folio = (page_num + 1) // 2
        cara = "r" if page_num % 2 != 0 else "v"

        # Fondo color suave alternado
        if page_num % 2 == 0:
            c.setFillColorRGB(0.95, 0.97, 1.0)
        else:
            c.setFillColorRGB(1.0, 0.98, 0.95)
        c.rect(0, 0, width, height, fill=1, stroke=0)

        # Número de página grande al centro
        c.setFillColorRGB(0.2, 0.2, 0.4)
        c.setFont("Helvetica-Bold", 72)
        c.drawCentredString(width / 2, height / 2 + 1 * cm, f"Página {page_num}")

        # Folio equivalente
        c.setFont("Helvetica", 36)
        c.setFillColorRGB(0.5, 0.3, 0.1)
        c.drawCentredString(width / 2, height / 2 - 2 * cm, f"Folio {folio}{cara}")

        # Borde decorativo
        c.setStrokeColorRGB(0.6, 0.6, 0.8)
        c.setLineWidth(2)
        c.rect(1 * cm, 1 * cm, width - 2 * cm, height - 2 * cm)

        # Texto pequeño con metadatos de prueba
        c.setFont("Helvetica", 10)
        c.setFillColorRGB(0.4, 0.4, 0.4)
        c.drawString(1.5 * cm, 1.5 * cm, f"PDF de prueba sintético │ Página {page_num}/{num_pages}")
        c.drawRightString(width - 1.5 * cm, 1.5 * cm, "fracmen_auto – Sistema Archivístico")

        c.showPage()

    c.save()
    print(f"✓ PDF de prueba generado: {output_path} ({num_pages} páginas)")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Genera un PDF sintético de prueba para el sistema archivístico."
    )
    group = parser.add_mutually_exclusive_group()
    group.add_argument(
        "--hojas", type=int, default=None,
        help="Número de hojas (folios) a generar. Cada hoja = 2 páginas (r + v)."
    )
    group.add_argument(
        "--pages", type=int, default=None,
        help="Número de páginas a generar directamente."
    )
    parser.add_argument(
        "--output", type=str, default="documento.pdf",
        help="Nombre del archivo PDF de salida (default: documento.pdf)"
    )
    args = parser.parse_args()

    # Determinar número de páginas
    if args.pages is not None:
        num_pages = args.pages
    elif args.hojas is not None:
        num_pages = args.hojas * 2
    else:
        # Default: 192 hojas = 384 páginas
        num_pages = 384

    print(f"Generando {num_pages} páginas ({num_pages // 2} hojas r+v)...")
    output_path = Path(__file__).parent / args.output
    generate_test_pdf(output_path, num_pages)
