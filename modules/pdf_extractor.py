"""
pdf_extractor.py - Extracción de páginas de un PDF fuente y escritura del PDF resultante.

Usa pypdf (sucesor moderno de PyPDF2).
El PdfReader se abre UNA SOLA VEZ fuera de este módulo y se reutiliza
para todos los registros (evita I/O repetido en PDFs grandes).
"""
import logging
from pathlib import Path
from typing import Union

from pypdf import PdfReader, PdfWriter

logger = logging.getLogger(__name__)


def open_pdf(pdf_path: Union[str, Path]) -> PdfReader:
    """
    Abre el PDF fuente y retorna el PdfReader.
    Llama a esta función UNA SOLA VEZ y reutiliza el objeto en todo el proceso.

    Raises:
        FileNotFoundError: si la ruta no existe.
    """
    path = Path(pdf_path)
    if not path.exists():
        raise FileNotFoundError(f"PDF no encontrado: {path}")
    logger.info(f"Abriendo PDF fuente: {path} ({path.stat().st_size // 1024} KB)")
    reader = PdfReader(str(path))
    logger.info(f"PDF fuente tiene {len(reader.pages)} páginas")
    return reader


def extract_pages(
    reader: PdfReader,
    page_numbers: list[int],
    dest_path: Path,
    dry_run: bool = False,
) -> bool:
    """
    Extrae las páginas indicadas del PdfReader y escribe un nuevo PDF en dest_path.

    Args:
        reader:       PdfReader ya abierto del PDF fuente.
        page_numbers: Lista de números de página (1-indexed) a extraer.
        dest_path:    Ruta completa donde se escribirá el nuevo PDF.
        dry_run:      Si True, no escribe nada en disco.

    Returns:
        True si la extracción fue exitosa, False en caso de error.
    """
    total_pages = len(reader.pages)
    writer = PdfWriter()

    for page_num in page_numbers:
        idx = page_num - 1  # pypdf usa índice 0-based
        if idx < 0 or idx >= total_pages:
            logger.error(
                f"Página {page_num} fuera de rango (PDF tiene {total_pages} páginas). "
                f"Saltando extracción."
            )
            return False
        writer.add_page(reader.pages[idx])

    if dry_run:
        logger.info(f"[DRY RUN] Se escribiría: {dest_path} ({len(page_numbers)} páginas)")
        return True

    dest_path.parent.mkdir(parents=True, exist_ok=True)
    with open(dest_path, "wb") as f:
        writer.write(f)

    logger.debug(f"PDF escrito: {dest_path} ({len(page_numbers)} páginas)")
    return True
