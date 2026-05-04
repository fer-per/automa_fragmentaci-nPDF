"""
main.py - Punto de entrada principal del sistema de automatización archivística.

Orquesta todos los módulos:
  1. Carga el Excel y el PDF fuente
  2. Filtra las filas según rango de filas indicado (protocolo)
  3. Valida cada registro (folios, rango, secuencia)
  4. Extrae páginas y escribe el PDF de salida
  5. Registra errores en logs/pendientes.csv
  6. Muestra barra de progreso y resumen final

Uso básico:
    python main.py                                         # usa config.py
    python main.py --preview                               # muestra filas del Excel y sale
    python main.py --excel inv.xlsx --pdf doc.pdf          # archivos específicos
    python main.py --row-start 1 --row-end 50              # solo filas 1-50 del Excel
    python main.py --dry-run                               # simula sin escribir archivos
    python main.py --no-gap-check                          # sin detección de saltos

Flujo típico para un Excel grande con múltiples protocolos:
    1. python main.py --excel big.xlsx --preview
       → ver listado de filas con sus datos
    2. python main.py --excel big.xlsx --pdf prot1.pdf --row-start 1 --row-end 120
       → procesar solo el protocolo 1
    3. python main.py --excel big.xlsx --pdf prot2.pdf --row-start 121 --row-end 300
       → procesar el protocolo 2, etc.
"""
import argparse
import csv
import logging
import sys
import time
from pathlib import Path

from tqdm import tqdm

import config
from modules.excel_reader import load_excel, load_excel_metadata
from modules.folio_parser import parse_folio_range, last_page_of_range
from modules.pdf_extractor import open_pdf, extract_pages
from modules.folder_builder import build_output_path
from modules.validator import validate_record


# ─── Configuración del logging ────────────────────────────────────────────────

def setup_logging(logs_dir: Path) -> None:
    logs_dir.mkdir(parents=True, exist_ok=True)
    log_format = "%(asctime)s [%(levelname)s] %(message)s"
    logging.basicConfig(
        level=logging.DEBUG,
        format=log_format,
        handlers=[
            logging.FileHandler(logs_dir / "process.log", encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )


# ─── Escritor de pendientes.csv ───────────────────────────────────────────────

class PendientesWriter:
    """Escribe registros problemáticos en pendientes.csv de forma incremental."""

    FIELDNAMES = [
        "N° DE REGISTRO", "ESCRIBANO/NOTARIO", "N° DE FOLIOS",
        "TITULO/ESCRITURA", "INTERESADO 1", "INTERESADO 2",
        "MOTIVO",
    ]

    def __init__(self, csv_path: Path, dry_run: bool = False):
        self.dry_run = dry_run
        self._path = csv_path
        self._file = None
        self._writer = None
        if not dry_run:
            csv_path.parent.mkdir(parents=True, exist_ok=True)
            self._file = open(csv_path, "a", newline="", encoding="utf-8-sig")
            write_header = csv_path.stat().st_size == 0 if csv_path.exists() else True
            self._writer = csv.DictWriter(self._file, fieldnames=self.FIELDNAMES, extrasaction="ignore")
            if write_header:
                self._writer.writeheader()

    def write(self, row: dict, motivo: str) -> None:
        row_out = {k: row.get(k, "") for k in self.FIELDNAMES}
        row_out["MOTIVO"] = motivo
        if self.dry_run:
            logging.getLogger(__name__).info(f"[DRY RUN] Pendiente: Reg {row_out.get('N° DE REGISTRO','?')} - {motivo}")
        else:
            self._writer.writerow(row_out)
            self._file.flush()

    def close(self) -> None:
        if self._file:
            self._file.close()


# ─── Modo Preview ─────────────────────────────────────────────────────────────

def run_preview(excel_path: Path, skip_rows: int) -> None:
    """
    Imprime un listado numerado de todas las filas del Excel.
    La columna FILA muestra el número de fila real del Excel (1-indexed).
    Útil para decidir el rango --row-start / --row-end antes de procesar.
    """
    meta = load_excel_metadata(
        excel_path,
        meta_row_siglo=config.META_ROW_SIGLO,
        meta_row_acervo=config.META_ROW_ACERVO,
    )
    df = load_excel(excel_path, skip_rows=skip_rows)
    total = len(df)

    print(f"\nExcel: {excel_path}  |  Total filas de datos: {total}")
    print(f"Fondo: ACERVO DOCUMENTAL NUMERO {meta['acervo_num']}  |  SIGLO {meta['siglo']}")
    print(f"Filas del Excel: {df.index.min()} a {df.index.max()}")
    print("-" * 120)
    print(f"{'FILA':>5}  {'REG':>5}  {'ESCRIBANO':<20}  {'PROT':>5}  {'FOLIOS':<10}  {'TITULO EST.':<20}  {'INTERESADO 1'}")
    print("-" * 120)

    for fila_excel, row in df.iterrows():
        reg        = str(row.get(config.COL_REGISTRO,   "")).strip()
        escribano  = str(row.get(config.COL_ESCRIBANO,  "")).replace("\n", " ").strip()[:20]
        protocolo  = str(row.get(config.COL_PROTOCOLO,  "")).strip()
        folios     = str(row.get(config.COL_FOLIOS,     "")).strip()
        titulo_est = str(row.get(config.COL_TITULO_EST, "")).strip()[:20]
        int1       = str(row.get(config.COL_INT1,       "")).strip()

        # Marcar filas sin folio
        marker = "  " if folios else " *"
        print(f"{fila_excel:>5}{marker} {reg:>5}  {escribano:<20}  {protocolo:>5}  {folios:<10}  {titulo_est:<20}  {int1}")

    print("-" * 120)
    print(f"  Total: {total} filas  |  (*) = sin folios")
    print(f"\nEjemplo de uso:")
    print(f"  python main.py --excel \"{excel_path}\" --pdf mi_doc.pdf --row-start {df.index.min()} --row-end {df.index.max()}\n")


# ─── Proceso principal ────────────────────────────────────────────────────────

def run(
    excel_path: Path,
    pdf_path: Path,
    row_start: int,
    row_end: int,
    dry_run: bool,
    check_gaps: bool,
    folio_inicio_excel: int = 1,
    pdf_page_inicio: int = 1,
) -> None:
    """
    Procesa las filas del Excel usando las filas reales del archivo.
    Los metadatos del fondo (acervo_num, siglo) se leen al inicio desde las filas de cabecera.

    Args:
        excel_path:         Ruta al archivo Excel .xlsx
        pdf_path:           Ruta al PDF fuente de este protocolo
        row_start:          Fila real del Excel donde iniciar (0 = desde la primera fila de datos)
        row_end:            Fila real del Excel donde terminar (inclusive). 0 = hasta el final.
        dry_run:            Si True, no escribe archivos en disco
        check_gaps:         Si True, detecta saltos de secuencia entre folios
        folio_inicio_excel: Numero de folio con que comienza el protocolo en el Excel (ej. 30).
        pdf_page_inicio:    Pagina real del PDF donde empieza la primera imagen del protocolo.
    """
    logger = logging.getLogger(__name__)
    start_time = time.time()

    logger.info("=" * 60)
    logger.info("INICIO DEL PROCESO ARCHIVISTICO")
    logger.info(f"  Excel:     {excel_path}")
    logger.info(f"  PDF:       {pdf_path}")
    logger.info(f"  Filas:     {row_start} -> {'fin' if row_end == 0 else row_end}")
    logger.info(f"  Salida:    {config.OUTPUT_DIR}")
    logger.info(f"  DRY RUN:   {dry_run}")
    logger.info(f"  Saltos:    {'ON' if check_gaps else 'OFF'}")
    logger.info(f"  Folio ini Excel: {folio_inicio_excel}")
    logger.info(f"  Pag ini PDF:     {pdf_page_inicio}")
    logger.info("=" * 60)

    # 1a. Leer metadatos del fondo (siglo y número de acervo) desde las filas de cabecera
    meta = load_excel_metadata(
        excel_path,
        meta_row_siglo=config.META_ROW_SIGLO,
        meta_row_acervo=config.META_ROW_ACERVO,
    )
    acervo_num = meta["acervo_num"]
    siglo      = meta["siglo"]
    logger.info(f"  Fondo:     ACERVO DOCUMENTAL NUMERO {acervo_num}  |  SIGLO {siglo}")

    # 1b. Cargar Excel completo (el índice del DataFrame = fila real del Excel)
    df = load_excel(excel_path, skip_rows=config.SKIP_ROWS)

    # 2. Aplicar filtro de rango usando filas reales del Excel
    if row_start > 0:
        df = df[df.index >= row_start]
    if row_end > 0:
        df = df[df.index <= row_end]

    total = len(df)
    if total == 0:
        logger.error(f"No hay filas en el rango {row_start}–{row_end}. Abortando.")
        sys.exit(1)

    logger.info(f"Filas seleccionadas: {total} (filas Excel {df.index.min()} a {df.index.max()})")

    # 3. Abrir PDF fuente (una sola vez)
    reader = open_pdf(pdf_path)
    total_pdf_pages = len(reader.pages)

    # 4. Preparar escritor de pendientes
    pendientes = PendientesWriter(config.PENDIENTES_CSV, dry_run=dry_run)

    processed = 0
    skipped   = 0
    prev_last_page = None

    # 5. Iterar registros con barra de progreso
    for fila_excel, row in tqdm(df.iterrows(), total=total, desc="Procesando", unit="reg"):
        row_dict = row.to_dict()

        is_valid, error_msg = validate_record(
            row=row_dict,
            col_folios=config.COL_FOLIOS,
            total_pdf_pages=total_pdf_pages,
            prev_last_page=prev_last_page,
            check_gaps=check_gaps,
            folio_inicio_excel=folio_inicio_excel,
            pdf_page_inicio=pdf_page_inicio,
        )

        reg_id = row_dict.get(config.COL_REGISTRO, f"fila_{fila_excel}")

        if not is_valid:
            nivel = logging.WARNING if "Salto" in error_msg else logging.INFO
            logger.log(nivel, f"Fila {fila_excel} / Reg {reg_id}: OMITIDO - {error_msg}")
            pendientes.write(row_dict, motivo=error_msg)
            skipped += 1
            continue

        # Parsear folios
        folio_str = str(row_dict.get(config.COL_FOLIOS, "")).strip()
        pages, _ = parse_folio_range(
            folio_str,
            folio_inicio_excel=folio_inicio_excel,
            pdf_page_inicio=pdf_page_inicio,
        )

        # Construir ruta de destino (jerarquía de 11 niveles)
        dest_path = build_output_path(
            output_dir=config.OUTPUT_DIR,
            acervo_num=acervo_num,
            siglo=siglo,
            escribano=str(row_dict.get(config.COL_ESCRIBANO, "")).replace("\n", " "),
            protocolo=str(row_dict.get(config.COL_PROTOCOLO, "")),
            registro=str(row_dict.get(config.COL_REGISTRO, "")),
            titulo_est=str(row_dict.get(config.COL_TITULO_EST, "")),
            fecha_ini=str(row_dict.get(config.COL_FECHA_INI, "")),
            interesado1=str(row_dict.get(config.COL_INT1, "")),
            interesado2=str(row_dict.get(config.COL_INT2, "")),
            dry_run=dry_run,
        )

        # Extraer y escribir páginas
        success = extract_pages(
            reader=reader,
            page_numbers=pages,
            dest_path=dest_path,
            dry_run=dry_run,
        )

        if success:
            logger.info(
                f"Fila {fila_excel} / Reg {reg_id}: OK - {dest_path.relative_to(config.BASE_DIR)} "
                f"({len(pages)} pag.: {pages})"
            )
            processed += 1
            prev_last_page = last_page_of_range(
                folio_str,
                folio_inicio_excel=folio_inicio_excel,
                pdf_page_inicio=pdf_page_inicio,
            )
        else:
            logger.error(f"Fila {fila_excel} / Reg {reg_id}: ERROR al extraer paginas - {folio_str}")
            pendientes.write(row_dict, motivo="Error al extraer paginas del PDF")
            skipped += 1

    pendientes.close()

    # 6. Resumen final
    elapsed = time.time() - start_time
    logger.info("=" * 60)
    logger.info("RESUMEN DEL PROCESO")
    logger.info(f"  Filas seleccionadas: {total}")
    logger.info(f"  Procesados:          {processed}")
    logger.info(f"  Pendientes:          {skipped}")
    logger.info(f"  Tiempo:              {elapsed:.2f} segundos")
    if not dry_run:
        logger.info(f"  Pendientes CSV:      {config.PENDIENTES_CSV}")
        logger.info(f"  Log completo:        {config.PROCESS_LOG}")
    logger.info("=" * 60)

    if skipped > 0:
        print(f"\n[ATENCION] {skipped} registro(s) no procesados. Revisa: {config.PENDIENTES_CSV}")


# ─── Punto de entrada ─────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description=(
            "Sistema de automatizacion archivistica.\n"
            "Extrae PDFs individuales por protocolo desde un PDF escaneado + inventario Excel.\n\n"
            "Flujo tipico para Excel grande con multiples protocolos:\n"
            "  1. Previsualizar filas:  python main.py --excel big.xlsx --preview\n"
            "  2. Procesar protocolo 1: python main.py --excel big.xlsx --pdf prot1.pdf --row-start 1 --row-end 120\n"
            "  3. Procesar protocolo 2: python main.py --excel big.xlsx --pdf prot2.pdf --row-start 121 --row-end 300"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    # ── Archivos de entrada ──
    parser.add_argument(
        "--excel", type=str, default=None,
        metavar="RUTA",
        help=f"Ruta al archivo Excel .xlsx (default: {config.EXCEL_PATH})",
    )
    parser.add_argument(
        "--pdf", type=str, default=None,
        metavar="RUTA",
        help=f"Ruta al PDF fuente del protocolo (default: {config.PDF_PATH})",
    )

    # ── Rango de filas ──
    parser.add_argument(
        "--row-start", type=int, default=1,
        metavar="N",
        help="Primera fila de datos a procesar, 1-indexed (default: 1 = inicio)",
    )
    parser.add_argument(
        "--row-end", type=int, default=0,
        metavar="N",
        help="Ultima fila de datos a procesar, inclusiva (default: 0 = hasta el final)",
    )

    # ── Modos especiales ──
    parser.add_argument(
        "--preview", action="store_true",
        help="Muestra listado de filas del Excel y sale sin procesar nada.",
    )
    parser.add_argument(
        "--dry-run", action="store_true", default=config.DRY_RUN,
        help="Simula el proceso sin escribir archivos en disco.",
    )
    parser.add_argument(
        "--no-gap-check", action="store_true", default=False,
        help="Desactiva la deteccion de saltos de secuencia entre folios.",
    )
    parser.add_argument(
        "--folio-inicio", type=int, default=1,
        metavar="N",
        help=(
            "Numero de folio con el que COMIENZA el protocolo en el Excel "
            "(default: 1). Ej.: si el Excel inicia en 30r, usa --folio-inicio 30"
        ),
    )
    parser.add_argument(
        "--pdf-page-inicio", type=int, default=1,
        metavar="N",
        help=(
            "Pagina real del PDF donde empieza la primera imagen del protocolo "
            "(default: 1). Ej.: si hay 2 portadas antes, usa --pdf-page-inicio 3"
        ),
    )

    args = parser.parse_args()

    # Resolver rutas: argumento CLI tiene prioridad sobre config.py
    excel_path = Path(args.excel) if args.excel else config.EXCEL_PATH
    pdf_path   = Path(args.pdf)   if args.pdf   else config.PDF_PATH

    # ── Modo preview ──
    if args.preview:
        try:
            run_preview(excel_path, skip_rows=config.SKIP_ROWS)
        except FileNotFoundError as e:
            print(f"[ERROR] Archivo no encontrado: {e}")
            sys.exit(1)
        sys.exit(0)

    # ── Proceso completo ──
    setup_logging(config.LOGS_DIR)

    try:
        run(
            excel_path=excel_path,
            pdf_path=pdf_path,
            row_start=args.row_start,
            row_end=args.row_end,
            dry_run=args.dry_run,
            check_gaps=not args.no_gap_check,
            folio_inicio_excel=args.folio_inicio,
            pdf_page_inicio=args.pdf_page_inicio,
        )
    except FileNotFoundError as e:
        logging.critical(f"Archivo no encontrado: {e}")
        sys.exit(1)
    except Exception as e:
        logging.critical(f"Error critico inesperado: {e}", exc_info=True)
        sys.exit(1)
