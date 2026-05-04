"""
gui.py - Interfaz gráfica del Sistema de Automatización Archivística.

"""
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import logging
import sys
import os
import time
import queue

# ── Imports del proyecto ──
import config
from modules.excel_reader import load_excel, load_excel_metadata
from modules.folio_parser import (
    parse_folio_range, last_page_of_range, analyze_folio_sequence,
    folio_text_to_page_abs, _folio_to_page_abs, _find_segment,
    adjust_pages_for_missing_folios,
)
from modules.pdf_extractor import open_pdf, extract_pages
from modules.folder_builder import build_output_path
from modules.validator import validate_record
from modules.data_topica_analyzer import analyze_data_topica
from modules.data_cronica_analyzer import analyze_data_cronica


# ─── Colores y estilos ───────────────────────────────────────────────────────
BG           = "#1a1b2e"
BG_CARD      = "#252640"
BG_INPUT     = "#2f3052"
BG_HOVER     = "#363860"
FG           = "#e8e8f0"
FG_DIM       = "#9090b0"
FG_MUTED     = "#6868a0"
ACCENT       = "#6c63ff"
ACCENT_LIGHT = "#8b83ff"
ACCENT_HOVER = "#9d96ff"
SUCCESS      = "#4ade80"
SUCCESS_DIM  = "#22c55e"
WARNING      = "#fbbf24"
ERROR        = "#f87171"
BORDER       = "#3a3b60"

FONT         = ("Segoe UI", 11)
FONT_BOLD    = ("Segoe UI", 11, "bold")
FONT_TITLE   = ("Segoe UI", 20, "bold")
FONT_SUBTITLE= ("Segoe UI", 12)
FONT_SMALL   = ("Segoe UI", 9)
FONT_STEP    = ("Segoe UI", 13, "bold")
FONT_LOG     = ("Consolas", 9)
FONT_BTN_BIG = ("Segoe UI", 15, "bold")


# ─── Tooltip helper ──────────────────────────────────────────────────────────
class ToolTip:
    """Tooltip emergente para widgets. Aparece al pasar el mouse."""
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip_window = None
        widget.bind("<Enter>", self._show)
        widget.bind("<Leave>", self._hide)

    def _show(self, event=None):
        if self.tip_window:
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(
            tw, text=self.text, justify="left",
            background="#3a3a5c", foreground="#e0e0f0",
            relief="solid", borderwidth=1,
            font=("Segoe UI", 10), padx=10, pady=6,
        )
        label.pack()

    def _hide(self, event=None):
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None


# ─── Logger handler que envía mensajes a la GUI ──────────────────────────────
class QueueHandler(logging.Handler):
    """Envía log records a una queue para que la GUI los muestre."""
    def __init__(self, log_queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record):
        self.log_queue.put(self.format(record))


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Fracmentación de Archivos — Archivos Historicos")
        self.configure(bg=BG)
        self.minsize(920, 720)
        self.geometry("980x780")

        # Try setting icon (optional, won't crash if missing)
        try:
            self.iconbitmap(default="")
        except Exception:
            pass

        # Estado
        self.excel_path = None
        self.pdf_path   = None
        self.output_dir = config.OUTPUT_DIR
        self.df         = None
        self.processing = False
        self.cancel_flag = False

        # Offset de inicio (folio Excel vs. página PDF)
        self.folio_inicio_excel = 1   # número de folio con que inicia el protocolo en el Excel
        self.pdf_page_inicio    = 1   # página del PDF donde empieza la primera imagen del protocolo

        # Cola para logs en tiempo real
        self.log_queue = queue.Queue()

        # Anti-flicker: track last known sizes to avoid unnecessary relayouts
        self._last_canvas_width = 0
        self._last_content_height = 0
        self._resize_after_id = None

        self._build_ui()
        self._center_window()
        self._poll_log_queue()

    # ─── Centrar ventana ──────────────────────────────────────────────────
    def _center_window(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    # ─── Construir la interfaz ────────────────────────────────────────────
    def _build_ui(self):
        # Scrollable main frame — uses a stable approach to avoid flicker
        self.canvas = tk.Canvas(self, bg=BG, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.main = tk.Frame(self.canvas, bg=BG)

        # Only update scrollregion when the content's ACTUAL size changes
        self.main.bind("<Configure>", self._on_content_configure)
        self._canvas_window = self.canvas.create_window((0, 0), window=self.main, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # Responsive width — debounced to avoid relayout storms during scroll
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        # Mousewheel scroll
        self.bind_all("<MouseWheel>",
                       lambda e: self.canvas.yview_scroll(-1*(e.delta//120), "units"))

        pad = {"padx": 24, "pady": 4}

        # ── Header ──
        header = tk.Frame(self.main, bg=BG)
        header.pack(fill="x", padx=24, pady=(20, 0))

        tk.Label(header, text="📄  Automatización de Fracmentación de Archivos Historicos",
                 font=FONT_TITLE, bg=BG, fg=FG).pack(anchor="w")
        tk.Label(header, text="Separa en documentos individuales por número de folios.",
                 font=FONT_SUBTITLE, bg=BG, fg=FG_DIM, wraplength=800, anchor="w", justify="left"
                 ).pack(anchor="w", pady=(2, 0))

        # Separador
        tk.Frame(self.main, bg=BORDER, height=1).pack(fill="x", padx=24, pady=(15, 10))

        # ══════════════════════════════════════════════════════════════════
        # PASO 1: Seleccionar archivos
        # ══════════════════════════════════════════════════════════════════
        self._step_header("①", "Selecciona tus archivos",
                          "Elige el Excel con el inventario y el PDF escaneado.")

        files_card = self._make_card()

        # ── Excel ──
        row_e = tk.Frame(files_card, bg=BG_CARD)
        row_e.pack(fill="x", padx=16, pady=(14, 6))

        self.icon_excel = tk.Label(row_e, text="⬜", font=("Segoe UI", 14),
                                   bg=BG_CARD, fg=FG_DIM, width=2)
        self.icon_excel.pack(side="left")

        lbl_frame_e = tk.Frame(row_e, bg=BG_CARD)
        lbl_frame_e.pack(side="left", fill="x", expand=True, padx=(6, 8))
        tk.Label(lbl_frame_e, text="Archivo Excel (.xlsx)",
                 font=FONT_BOLD, bg=BG_CARD, fg=FG, anchor="w").pack(anchor="w")
        self.lbl_excel = tk.Label(lbl_frame_e, text="No seleccionado",
                                  font=FONT_SMALL, bg=BG_CARD, fg=FG_DIM, anchor="w")
        self.lbl_excel.pack(anchor="w")

        btn_e = self._make_button(row_e, "Seleccionar…", self._select_excel, width=14)
        btn_e.pack(side="right")
        ToolTip(btn_e, "Haz clic para buscar tu archivo Excel\ncon el inventario de documentos.")

        # ── PDF ──
        row_p = tk.Frame(files_card, bg=BG_CARD)
        row_p.pack(fill="x", padx=16, pady=(6, 14))

        self.icon_pdf = tk.Label(row_p, text="⬜", font=("Segoe UI", 14),
                                  bg=BG_CARD, fg=FG_DIM, width=2)
        self.icon_pdf.pack(side="left")

        lbl_frame_p = tk.Frame(row_p, bg=BG_CARD)
        lbl_frame_p.pack(side="left", fill="x", expand=True, padx=(6, 8))
        tk.Label(lbl_frame_p, text="Archivo PDF protocolo escaneado",
                 font=FONT_BOLD, bg=BG_CARD, fg=FG, anchor="w").pack(anchor="w")
        self.lbl_pdf = tk.Label(lbl_frame_p, text="No seleccionado",
                                 font=FONT_SMALL, bg=BG_CARD, fg=FG_DIM, anchor="w")
        self.lbl_pdf.pack(anchor="w")

        btn_p = self._make_button(row_p, "Seleccionar…", self._select_pdf, width=14)
        btn_p.pack(side="right")
        ToolTip(btn_p, "Haz clic para buscar el PDF escaneado\nque contiene los documentos originales.")

        # ══════════════════════════════════════════════════════════════════
        # PASO 2: Vista previa y rango
        # ══════════════════════════════════════════════════════════════════
        self._step_header("②", "Revisa los datos y elige el rango",
                          "La tabla muestra las filas del Excel. Puedes procesar todas o solo un rango.")

        preview_card = self._make_card()

        # Tabla de vista previa
        table_container = tk.Frame(preview_card, bg=BG_CARD)
        table_container.pack(fill="x", padx=16, pady=(14, 8))

        cols = ("fila", "reg", "escribano", "prot", "folios", "titulo", "interesado1")
        col_names = {
            "fila": "Fila", "reg": "Reg.", "escribano": "Escribano",
            "prot": "Prot.", "folios": "Folios", "titulo": "Título",
            "interesado1": "Interesado 1",
        }

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Dark.Treeview",
                         background=BG_INPUT, foreground=FG,
                         fieldbackground=BG_INPUT,
                         font=FONT_SMALL, rowheight=26, borderwidth=0)
        style.configure("Dark.Treeview.Heading",
                         background="#2c2d50", foreground=ACCENT_LIGHT,
                         font=("Segoe UI", 10, "bold"),
                         borderwidth=0, relief="flat")
        style.map("Dark.Treeview",
                   background=[("selected", ACCENT)],
                   foreground=[("selected", "#ffffff")])

        self.tree = ttk.Treeview(
            table_container, columns=cols, show="headings",
            height=7, style="Dark.Treeview",
        )
        for col_id in cols:
            w = 50 if col_id in ("fila", "reg", "prot") else 85 if col_id == "folios" else 150
            self.tree.heading(col_id, text=col_names[col_id])
            anchor = "center" if col_id in ("fila", "reg", "prot", "folios") else "w"
            self.tree.column(col_id, width=w, minwidth=40, anchor=anchor)

        tree_scroll = ttk.Scrollbar(table_container, orient="vertical",
                                     command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        self.tree.pack(side="left", fill="x", expand=True)
        tree_scroll.pack(side="right", fill="y")

        # Info de la tabla
        self.lbl_table_info = tk.Label(
            preview_card,
            text="  ↑  Carga un archivo Excel para ver las filas aquí",
            font=FONT_SMALL, bg=BG_CARD, fg=FG_MUTED, anchor="w",
        )
        self.lbl_table_info.pack(fill="x", padx=16, pady=(0, 8))

        # ── Rango de filas ──
        range_frame = tk.Frame(preview_card, bg=BG_CARD)
        range_frame.pack(padx=16, pady=(4, 14))

        tk.Label(range_frame, text="Inicia desde fila:", font=FONT_BOLD,
                 bg=BG_CARD, fg=FG).pack(side="left", padx=(0, 5))
        self.spin_start = tk.Spinbox(
            range_frame, from_=1, to=99999, width=7, font=FONT,
            bg=BG_INPUT, fg=FG, buttonbackground=BG_CARD,
            insertbackground=FG, relief="flat", justify="center",
        )
        self.spin_start.pack(side="left", padx=(0, 20))
        ToolTip(self.spin_start, "Primera fila del Excel que quieres procesar.\n"
                "Normalmente déjalo en 1.")

        tk.Label(range_frame, text="Termina en fila:", font=FONT_BOLD,
                 bg=BG_CARD, fg=FG).pack(side="left", padx=(0, 5))
        self.spin_end = tk.Spinbox(
            range_frame, from_=0, to=99999, width=7, font=FONT,
            bg=BG_INPUT, fg=FG, buttonbackground=BG_CARD,
            insertbackground=FG, relief="flat", justify="center",
        )
        self.spin_end.delete(0, "end")
        self.spin_end.insert(0, "0")
        self.spin_end.pack(side="left", padx=(0, 10))
        ToolTip(self.spin_end, "Última fila que quieres procesar.\nDeja en 0 para procesar todas.")

        hint_0 = tk.Label(range_frame, text="", font=FONT_SMALL,
                          bg=BG_CARD, fg=FG_MUTED)
        hint_0.pack(side="left")

        # ══════════════════════════════════════════════════════════════════
        # PASO 2b: Configuración de inicio de folios / PDF
        # ══════════════════════════════════════════════════════════════════
        self._step_header("②b", "Configuración de inicio",
                          "Indica desde qué folio empieza el protocolo y desde qué página del PDF.")

        offset_card = self._make_card()
        offset_row = tk.Frame(offset_card, bg=BG_CARD)
        offset_row.pack(padx=16, pady=(14, 10))

        tk.Label(offset_row, text="Folio inicio del Excel:",
                 font=FONT_BOLD, bg=BG_CARD, fg=FG).pack(side="left", padx=(0, 5))
        self.spin_folio_ini = tk.Spinbox(
            offset_row, from_=1, to=99999, width=7, font=FONT,
            bg=BG_INPUT, fg=FG, buttonbackground=BG_CARD,
            insertbackground=FG, relief="flat", justify="center",
            command=self._on_offset_change,
        )
        self.spin_folio_ini.delete(0, "end")
        self.spin_folio_ini.insert(0, "1")
        self.spin_folio_ini.pack(side="left", padx=(0, 20))
        self.spin_folio_ini.bind("<FocusOut>", lambda e: self._on_offset_change())
        ToolTip(self.spin_folio_ini,
                "Número de folio con el que EMPIEZA el protocolo en el Excel.\n"
                "Ej.: si el Excel inicia en '30r', escribe 30.\n"
                "Si empieza en '1r', déjalo en 1 (valor por defecto).")

        tk.Label(offset_row, text="Página PDF de inicio:",
                 font=FONT_BOLD, bg=BG_CARD, fg=FG).pack(side="left", padx=(0, 5))
        self.spin_pdf_ini = tk.Spinbox(
            offset_row, from_=1, to=99999, width=7, font=FONT,
            bg=BG_INPUT, fg=FG, buttonbackground=BG_CARD,
            insertbackground=FG, relief="flat", justify="center",
            command=self._on_offset_change,
        )
        self.spin_pdf_ini.delete(0, "end")
        self.spin_pdf_ini.insert(0, "1")
        self.spin_pdf_ini.pack(side="left", padx=(0, 10))
        self.spin_pdf_ini.bind("<FocusOut>", lambda e: self._on_offset_change())
        ToolTip(self.spin_pdf_ini,
                "Número de página REAL del PDF donde empieza la primera imagen del protocolo.\n"
                "Ej.: si hay 2 portadas antes, escribe 3.\n"
                "Si no hay portadas, déjalo en 1 (valor por defecto).")

        self.lbl_offset_info = tk.Label(
            offset_card,
            text="  -> La primera imagen del PDF corresponde al folio 1r del protocolo.",
            font=FONT_SMALL, bg=BG_CARD, fg=FG_MUTED, anchor="w",
        )
        self.lbl_offset_info.pack(fill="x", padx=16, pady=(0, 6))

        # ── Segmentos adicionales (para saltos en el PDF) ─────────────────
        seg_lbl = tk.Label(
            offset_card,
            text="  Segmentos adicionales (cuando el PDF no contiene los folios del salto):",
            font=FONT_SMALL, bg=BG_CARD, fg=FG_DIM, anchor="w",
        )
        seg_lbl.pack(fill="x", padx=16, pady=(2, 4))

        seg_frame = tk.Frame(offset_card, bg=BG_CARD)
        seg_frame.pack(fill="x", padx=16, pady=(0, 4))

        # Treeview de segmentos
        self.tree_segs = ttk.Treeview(
            seg_frame, columns=("folio", "pdf_pag"), show="headings",
            height=3, style="Dark.Treeview",
        )
        self.tree_segs.heading("folio",   text="Folio inicio Excel")
        self.tree_segs.heading("pdf_pag", text="Pagina PDF inicio")
        self.tree_segs.column("folio",   width=160, anchor="center")
        self.tree_segs.column("pdf_pag", width=160, anchor="center")
        self.tree_segs.pack(side="left", fill="x", expand=True)

        seg_btn_frame = tk.Frame(seg_frame, bg=BG_CARD)
        seg_btn_frame.pack(side="left", padx=(8, 0))

        # Entradas para nuevo segmento
        ent_row = tk.Frame(offset_card, bg=BG_CARD)
        ent_row.pack(fill="x", padx=16, pady=(2, 8))

        tk.Label(ent_row, text="Folio (ej: 401r, 53v):", font=FONT_SMALL,
                 bg=BG_CARD, fg=FG).pack(side="left", padx=(0, 3))
        self.ent_seg_folio = tk.Entry(
            ent_row, width=9, font=FONT, bg=BG_INPUT, fg=FG,
            insertbackground=FG, relief="flat", justify="center",
        )
        self.ent_seg_folio.insert(0, "401r")
        self.ent_seg_folio.pack(side="left", padx=(0, 10))

        tk.Label(ent_row, text="Pagina PDF:", font=FONT_SMALL,
                 bg=BG_CARD, fg=FG).pack(side="left", padx=(0, 3))
        self.ent_seg_pag = tk.Entry(
            ent_row, width=7, font=FONT, bg=BG_INPUT, fg=FG,
            insertbackground=FG, relief="flat", justify="center",
        )
        self.ent_seg_pag.insert(0, "635")
        self.ent_seg_pag.pack(side="left", padx=(0, 10))

        btn_add_seg = self._make_button(ent_row, "+ Agregar", self._add_segment, width=10)
        btn_add_seg.pack(side="left", padx=(0, 6))
        ToolTip(btn_add_seg,
                "Agrega un segmento: indica desde que folio del Excel\n"
                "empieza la siguiente seccion en el PDF.\n"
                "Util cuando el PDF no contiene los folios del salto.")

        btn_del_seg = self._make_button(ent_row, "- Eliminar", self._del_segment, width=10)
        btn_del_seg.configure(bg="#4a4a6a")
        btn_del_seg.pack(side="left")
        ToolTip(btn_del_seg, "Elimina el segmento seleccionado en la tabla.")

        # ── Páginas a ignorar ────────────────────────────────────────────
        tk.Frame(offset_card, bg=BORDER, height=1).pack(fill="x", padx=16, pady=(6, 0))

        ign_lbl = tk.Label(
            offset_card,
            text="  Páginas PDF a ignorar (hojas que NO deben incluirse en ningún documento):",
            font=FONT_SMALL, bg=BG_CARD, fg=FG_DIM, anchor="w",
        )
        ign_lbl.pack(fill="x", padx=16, pady=(6, 4))

        ign_frame = tk.Frame(offset_card, bg=BG_CARD)
        ign_frame.pack(fill="x", padx=16, pady=(0, 4))

        self.tree_ignore = ttk.Treeview(
            ign_frame, columns=("entrada", "paginas"), show="headings",
            height=3, style="Dark.Treeview",
        )
        self.tree_ignore.heading("entrada", text="Entrada")
        self.tree_ignore.heading("paginas", text="Páginas PDF ignoradas")
        self.tree_ignore.column("entrada", width=130, anchor="center")
        self.tree_ignore.column("paginas", width=340, anchor="w")
        self.tree_ignore.pack(side="left", fill="x", expand=True)

        ign_scroll = ttk.Scrollbar(ign_frame, orient="vertical",
                                    command=self.tree_ignore.yview)
        self.tree_ignore.configure(yscrollcommand=ign_scroll.set)
        ign_scroll.pack(side="right", fill="y")

        ign_ent_row = tk.Frame(offset_card, bg=BG_CARD)
        ign_ent_row.pack(fill="x", padx=16, pady=(2, 10))

        tk.Label(ign_ent_row, text="Pág. o rango (ej: 5  ó  10-15):",
                 font=FONT_SMALL, bg=BG_CARD, fg=FG).pack(side="left", padx=(0, 4))
        self.ent_ignore_pag = tk.Entry(
            ign_ent_row, width=10, font=FONT, bg=BG_INPUT, fg=FG,
            insertbackground=FG, relief="flat", justify="center",
        )
        self.ent_ignore_pag.insert(0, "5")
        self.ent_ignore_pag.pack(side="left", padx=(0, 10))
        ToolTip(self.ent_ignore_pag,
                "Escribe una pagina simple (ej: 5)\n"
                "o un rango (ej: 10-15) para ignorar varias a la vez.")

        btn_add_ign = self._make_button(
            ign_ent_row, "+ Ignorar", self._add_ignored_page, width=10)
        btn_add_ign.configure(bg="#7c3d2a")
        btn_add_ign.bind("<Enter>", lambda e: btn_add_ign.configure(bg="#a0532f"))
        btn_add_ign.bind("<Leave>", lambda e: btn_add_ign.configure(bg="#7c3d2a"))
        btn_add_ign.pack(side="left", padx=(0, 6))
        ToolTip(btn_add_ign,
                "Agrega la pagina o rango indicado a la lista de exclusion.\n"
                "Esas paginas seran omitidas en todos los documentos generados.")

        btn_del_ign = self._make_button(
            ign_ent_row, "- Quitar", self._del_ignored_page, width=10)
        btn_del_ign.configure(bg="#4a4a6a")
        btn_del_ign.pack(side="left")
        ToolTip(btn_del_ign, "Quita la entrada seleccionada de la lista de ignorados.")

        # ── Folios a ignorar (solo afecta el conteo, NO el PDF extraído) ──
        tk.Frame(offset_card, bg="#3a2a10", height=1).pack(fill="x", padx=16, pady=(2, 0))

        ign_folio_lbl = tk.Label(
            offset_card,
            text="  Folios a ignorar — solo excluye del conteo de folios (el PDF se extrae completo):",
            font=FONT_SMALL, bg=BG_CARD, fg=FG_DIM, anchor="w",
        )
        ign_folio_lbl.pack(fill="x", padx=16, pady=(6, 2))

        ign_folio_tree_frame = tk.Frame(offset_card, bg=BG_CARD)
        ign_folio_tree_frame.pack(fill="x", padx=16, pady=(0, 4))

        self.tree_ignore_folios = ttk.Treeview(
            ign_folio_tree_frame, columns=("entrada", "detalle"), show="headings",
            height=3, style="Dark.Treeview",
        )
        self.tree_ignore_folios.heading("entrada", text="Folio(s) ignorado(s)")
        self.tree_ignore_folios.heading("detalle", text="Posiciones excluidas del conteo")
        self.tree_ignore_folios.column("entrada", width=130, anchor="center")
        self.tree_ignore_folios.column("detalle", width=340, anchor="w")
        self.tree_ignore_folios.pack(side="left", fill="x", expand=True)

        ign_folio_scroll2 = ttk.Scrollbar(ign_folio_tree_frame, orient="vertical",
                                           command=self.tree_ignore_folios.yview)
        self.tree_ignore_folios.configure(yscrollcommand=ign_folio_scroll2.set)
        ign_folio_scroll2.pack(side="right", fill="y")

        ign_folio_row = tk.Frame(offset_card, bg=BG_CARD)
        ign_folio_row.pack(fill="x", padx=16, pady=(2, 4))

        tk.Label(ign_folio_row, text="Folio (ej: 40r  ó  40r-41v):",
                 font=FONT_SMALL, bg=BG_CARD, fg=FG).pack(side="left", padx=(0, 4))
        self.ent_ignore_folio = tk.Entry(
            ign_folio_row, width=12, font=FONT, bg=BG_INPUT, fg=FG,
            insertbackground=FG, relief="flat", justify="center",
        )
        self.ent_ignore_folio.insert(0, "40r")
        self.ent_ignore_folio.pack(side="left", padx=(0, 10))
        ToolTip(self.ent_ignore_folio,
                "Escribe el folio como aparece en el Excel: 40r, 40v, 40r-41v.\n"
                "Esas posiciones se excluirán del conteo de folios.\n"
                "El PDF se extrae completo sin omitir ninguna página.")

        btn_add_ign_folio = self._make_button(
            ign_folio_row, "+ Ignorar Folio", self._add_ignored_folio, width=14)
        btn_add_ign_folio.configure(bg="#5a3a00")
        btn_add_ign_folio.bind("<Enter>", lambda e: btn_add_ign_folio.configure(bg="#7a5200"))
        btn_add_ign_folio.bind("<Leave>", lambda e: btn_add_ign_folio.configure(bg="#5a3a00"))
        btn_add_ign_folio.pack(side="left", padx=(0, 6))
        ToolTip(btn_add_ign_folio,
                "Agrega el rango de folios a la lista de exclusión de conteo.\n"
                "Solo afecta cuántas páginas se asignan al registro.\n"
                "El PDF se extrae completo (sin saltar páginas).")

        btn_del_ign_folio = self._make_button(
            ign_folio_row, "- Quitar", self._del_ignored_folio, width=10)
        btn_del_ign_folio.configure(bg="#4a4a6a")
        btn_del_ign_folio.pack(side="left")
        ToolTip(btn_del_ign_folio, "Quita la entrada seleccionada de la lista de folios ignorados.")

        self.lbl_ignore_folio_preview = tk.Label(
            offset_card,
            text="  →  Escribe un folio y pulsa el botón para agregarlo al conteo ignorado.",
            font=FONT_SMALL, bg=BG_CARD, fg=FG_MUTED, anchor="w",
        )
        self.lbl_ignore_folio_preview.pack(fill="x", padx=16, pady=(0, 2))

        self.lbl_ignore_folio_count = tk.Label(
            offset_card,
            text="  Sin folios ignorados en el conteo.",
            font=FONT_SMALL, bg=BG_CARD, fg=FG_MUTED, anchor="w",
        )
        self.lbl_ignore_folio_count.pack(fill="x", padx=16, pady=(0, 8))

        # ══════════════════════════════════════════════════════════════════
        # PASO 2c: Analizadores
        # ══════════════════════════════════════════════════════════════════
        self._step_header("②c", "Analizadores",
                          "Verifica la sucesión de folios y la DATA TÓPICA antes de fragmentar.")

        analyzer_card = self._make_card()
        btn_row_an = tk.Frame(analyzer_card, bg=BG_CARD)
        btn_row_an.pack(padx=16, pady=(14, 6))

        btn_analyze_folios = self._make_button(
            btn_row_an, "Analizar Folios", self._analyze_folios, width=18)
        btn_analyze_folios.pack(side="left", padx=(0, 8))
        ToolTip(btn_analyze_folios,
                "Verifica que la sucesion de folios del Excel sea continua.\n"
                "Muestra que hojas o caras faltan si hay saltos.")

        btn_analyze_topica = self._make_button(
            btn_row_an, "Analizar TOPICA", self._analyze_data_topica, width=18)
        btn_analyze_topica.pack(side="left", padx=(0, 8))
        ToolTip(btn_analyze_topica,
                "Verifica que la columna DATA TOPICA tenga formato correcto.")

        btn_analyze_cronica = self._make_button(
            btn_row_an, "Analizar CRONICA", self._analyze_data_cronica, width=18)
        btn_analyze_cronica.pack(side="left", padx=(0, 8))
        ToolTip(btn_analyze_cronica,
                "Verifica el formato de fechas (d/m/yyyy) y que las fechas\n"
                "vayan en orden cronologico progresivo.")

        btn_check_coverage = self._make_button(
            btn_row_an, "Verificar Cobertura PDF", self._analyze_pdf_coverage, width=22)
        btn_check_coverage.configure(bg="#2d6a4f")
        btn_check_coverage.bind("<Enter>", lambda e: btn_check_coverage.configure(bg="#40916c"))
        btn_check_coverage.bind("<Leave>", lambda e: btn_check_coverage.configure(bg="#2d6a4f"))
        btn_check_coverage.pack(side="left", padx=(0, 8))
        ToolTip(btn_check_coverage,
                "Compara la ultima pagina PDF que usaria el rango configurado\n"
                "con el total real de paginas del PDF.\n"
                "Indica si el PDF esta bien alineado o si sobran/faltan hojas.\n"
                "Se actualiza automaticamente al cambiar segmentos o ignorados.")

        btn_report = self._make_button(
            btn_row_an, "\U0001f4cb Reporte", self._generate_fragmentation_report, width=14)
        btn_report.configure(bg="#3a3a7a")
        btn_report.bind("<Enter>", lambda e: btn_report.configure(bg="#5050aa"))
        btn_report.bind("<Leave>", lambda e: btn_report.configure(bg="#3a3a7a"))
        btn_report.pack(side="left")
        ToolTip(btn_report,
                "Genera un reporte detallado del mapeo folio \u2192 p\u00e1ginas PDF.\n"
                "Muestra columnas del Excel y resalta registros afectados\n"
                "por segmentos, p\u00e1ginas ignoradas o folios ignorados.")

        # Panel de resultados de análisis (texto expandible)
        self.txt_analyzer = tk.Text(
            analyzer_card, height=7, bg=BG_INPUT, fg=FG_DIM,
            font=FONT_LOG, relief="flat", wrap="word",
            insertbackground=FG, state="disabled",
            padx=12, pady=10,
        )
        an_scroll = ttk.Scrollbar(analyzer_card, orient="vertical",
                                   command=self.txt_analyzer.yview)
        self.txt_analyzer.configure(yscrollcommand=an_scroll.set)
        self.txt_analyzer.pack(side="left", fill="both", expand=True, padx=(16, 0), pady=(6, 14))
        an_scroll.pack(side="right", fill="y", padx=(0, 16), pady=(6, 14))

        self.txt_analyzer.tag_configure("ok",   foreground=SUCCESS)
        self.txt_analyzer.tag_configure("warn", foreground=WARNING)
        self.txt_analyzer.tag_configure("err",  foreground=ERROR)
        self.txt_analyzer.tag_configure("info", foreground=FG_DIM)

        # ══════════════════════════════════════════════════════════════════
        # PASO 3: Carpeta de salida
        # ══════════════════════════════════════════════════════════════════
        self._step_header("③", "Carpeta de salida",
                          "Elige dónde se guardarán los PDFs generados.")

        output_card = self._make_card()
        row_out = tk.Frame(output_card, bg=BG_CARD)
        row_out.pack(fill="x", padx=16, pady=14)

        tk.Label(row_out, text="📁", font=("Segoe UI", 14),
                 bg=BG_CARD, fg=FG_DIM).pack(side="left")

        self.lbl_output = tk.Label(
            row_out, text=str(self.output_dir),
            font=FONT, bg=BG_INPUT, fg=FG,
            relief="flat", padx=10, pady=6, anchor="w",
        )
        self.lbl_output.pack(side="left", fill="x", expand=True, padx=(8, 8))

        btn_out = self._make_button(row_out, "Cambiar…", self._select_output_dir, width=12)
        btn_out.pack(side="right")
        ToolTip(btn_out, "Cambia la carpeta donde se guardan\nlos documentos separados.")

        # ══════════════════════════════════════════════════════════════════
        # PASO 4: Opciones y Procesar
        # ══════════════════════════════════════════════════════════════════
        self._step_header("④", "¡Procesar!",
                          "Revisa que todo esté listo y presiona el botón.")

        action_card = self._make_card()

        # Opciones
        opts_row = tk.Frame(action_card, bg=BG_CARD)
        opts_row.pack(padx=16, pady=(14, 8))

        self.var_no_gap = tk.BooleanVar(value=True)
        cb = tk.Checkbutton(
            opts_row, text="  Ignorar saltos de secuencia entre folios",
            variable=self.var_no_gap, font=FONT, bg=BG_CARD, fg=FG,
            selectcolor=BG_INPUT, activebackground=BG_CARD, activeforeground=FG,
            cursor="hand2",
        )
        cb.pack(side="left")
        ToolTip(cb, "Si está marcado, el sistema NO detendrá un registro\n"
                "solo porque hay un salto numérico entre folios.\n"
                "Recomendado: dejarlo marcado.")

        # Estado de "listo para procesar"
        self.lbl_ready = tk.Label(
            action_card, text="", font=FONT_SMALL, bg=BG_CARD, fg=FG_DIM,
        )
        self.lbl_ready.pack(pady=(0, 4))
        self._update_ready_label()

        # Botones: Procesar y Cancelar
        btn_row = tk.Frame(action_card, bg=BG_CARD)
        btn_row.pack(pady=(2, 6))

        self.btn_process = tk.Button(
            btn_row, text="▶   PROCESAR", font=FONT_BTN_BIG,
            bg=ACCENT, fg="#ffffff",
            activebackground=ACCENT_HOVER, activeforeground="#ffffff",
            relief="flat", cursor="hand2", padx=36, pady=10,
            command=self._start_processing,
        )
        self.btn_process.pack(side="left", padx=(0, 10))
        ToolTip(self.btn_process, "Inicia el proceso de extracción de documentos.\n"
                "Asegúrate de haber seleccionado ambos archivos.")

        self.btn_cancel = tk.Button(
            btn_row, text="✕  Cancelar", font=FONT_BOLD,
            bg="#4a4a6a", fg=FG, activebackground="#5a5a7a", activeforeground=FG,
            relief="flat", cursor="hand2", padx=16, pady=10,
            command=self._cancel_processing, state="disabled",
        )
        self.btn_cancel.pack(side="left")

        # Barra de progreso
        style.configure("Accent.Horizontal.TProgressbar",
                         troughcolor=BG_INPUT, background=ACCENT, thickness=22)
        self.progress = ttk.Progressbar(
            action_card, mode="determinate", length=500,
            style="Accent.Horizontal.TProgressbar",
        )
        self.progress.pack(pady=(6, 4))

        self.lbl_progress = tk.Label(
            action_card, text="", font=FONT, bg=BG_CARD, fg=FG_DIM,
        )
        self.lbl_progress.pack(pady=(0, 14))

        # ══════════════════════════════════════════════════════════════════
        # Registro de actividad (log en tiempo real)
        # ══════════════════════════════════════════════════════════════════
        log_header = tk.Frame(self.main, bg=BG)
        log_header.pack(fill="x", padx=24, pady=(12, 4))
        tk.Label(log_header, text="Registro de actividad",
                 font=("Segoe UI", 11, "bold"), bg=BG, fg=FG_DIM, anchor="w").pack(side="left")

        log_card = self._make_card()
        self.txt_log = tk.Text(
            log_card, height=8, bg=BG_INPUT, fg=FG_DIM,
            font=FONT_LOG, relief="flat", wrap="word",
            insertbackground=FG, state="disabled",
            padx=12, pady=10,
        )
        log_scroll = ttk.Scrollbar(log_card, orient="vertical",
                                    command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=log_scroll.set)
        self.txt_log.pack(side="left", fill="both", expand=True, padx=(16, 0), pady=14)
        log_scroll.pack(side="right", fill="y", padx=(0, 16), pady=14)

        # Tag para colores en el log
        self.txt_log.tag_configure("ok", foreground=SUCCESS)
        self.txt_log.tag_configure("warn", foreground=WARNING)
        self.txt_log.tag_configure("err", foreground=ERROR)
        self.txt_log.tag_configure("info", foreground=FG_DIM)

        # ══════════════════════════════════════════════════════════════════
        # Panel de resultados (oculto inicialmente)
        # ══════════════════════════════════════════════════════════════════
        self.results_frame = tk.Frame(self.main, bg=BG_CARD)

        self.lbl_result_title = tk.Label(
            self.results_frame, text="",
            font=("Segoe UI", 15, "bold"), bg=BG_CARD, fg=SUCCESS,
        )
        self.lbl_result_title.pack(pady=(16, 6))

        self.lbl_result_body = tk.Label(
            self.results_frame, text="",
            font=FONT, bg=BG_CARD, fg=FG, justify="left",
        )
        self.lbl_result_body.pack(padx=16, pady=(0, 6))

        self.btn_open_output = tk.Button(
            self.results_frame, text="📂  Abrir carpeta de salida",
            font=FONT_BOLD, bg="#3b3b5a", fg=FG, relief="flat",
            cursor="hand2", padx=20, pady=8,
            command=self._open_output_folder,
        )
        self.btn_open_output.pack(pady=(4, 16))
        ToolTip(self.btn_open_output, "Abre en el Explorador de Windows la carpeta\n"
                "donde se guardaron los PDFs.")

        # Espaciado final
        tk.Frame(self.main, bg=BG, height=20).pack()

    # ─── Scroll & resize stability ─────────────────────────────────────────
    def _on_content_configure(self, event):
        """Called when the inner frame changes size. Only updates scrollregion
        if the content height actually changed, preventing feedback loops."""
        new_height = self.main.winfo_reqheight()
        if new_height != self._last_content_height:
            self._last_content_height = new_height
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        """Called when the canvas (viewport) is resized. Debounced so that
        dragging the scrollbar does NOT cause continuous relayouts."""
        new_width = event.width
        if new_width != self._last_canvas_width:
            self._last_canvas_width = new_width
            # Cancel any pending resize to debounce rapid events
            if self._resize_after_id is not None:
                self.after_cancel(self._resize_after_id)
            self._resize_after_id = self.after(30, self._apply_canvas_width, new_width)

    def _apply_canvas_width(self, width):
        """Actually apply the new width to the inner frame."""
        self._resize_after_id = None
        self.canvas.itemconfig(self._canvas_window, width=width)

    # ─── Helpers de UI ────────────────────────────────────────────────────
    def _step_header(self, number, title, subtitle):
        frame = tk.Frame(self.main, bg=BG)
        frame.pack(fill="x", padx=24, pady=(14, 4))

        top = tk.Frame(frame, bg=BG)
        top.pack(fill="x")

        tk.Label(top, text=number, font=("Segoe UI", 16),
                 bg=BG, fg=ACCENT).pack(side="left", padx=(0, 8))
        tk.Label(top, text=title, font=FONT_STEP,
                 bg=BG, fg=FG).pack(side="left")

        tk.Label(frame, text=subtitle, font=FONT_SMALL,
                 bg=BG, fg=FG_DIM, anchor="w").pack(anchor="w", padx=(34, 0))

    def _make_card(self):
        """Crea un frame tipo tarjeta con borde sutil."""
        outer = tk.Frame(self.main, bg=BORDER, padx=1, pady=1)
        outer.pack(fill="x", padx=24, pady=(0, 2))
        inner = tk.Frame(outer, bg=BG_CARD)
        inner.pack(fill="both", expand=True)
        return inner

    def _make_button(self, parent, text, command, width=None):
        btn = tk.Button(
            parent, text=text, font=FONT_BOLD,
            bg=ACCENT, fg="#ffffff",
            activebackground=ACCENT_HOVER, activeforeground="#ffffff",
            relief="flat", cursor="hand2", padx=10, pady=5,
            command=command,
        )
        if width:
            btn.configure(width=width)
        # Hover effect
        btn.bind("<Enter>", lambda e: btn.configure(bg=ACCENT_HOVER))
        btn.bind("<Leave>", lambda e: btn.configure(bg=ACCENT))
        return btn

    def _update_ready_label(self):
        """Muestra si el sistema está listo para procesar."""
        excel_ok = self.excel_path is not None
        pdf_ok   = self.pdf_path is not None

        if excel_ok and pdf_ok:
            self.lbl_ready.config(
                text="✓  Todo listo. Puedes presionar PROCESAR.", fg=SUCCESS_DIM,
            )
        else:
            missing = []
            if not excel_ok:
                missing.append("el Excel")
            if not pdf_ok:
                missing.append("el PDF")
            self.lbl_ready.config(
                text=f"⚠  Falta seleccionar: {' y '.join(missing)}", fg=WARNING,
            )

    # ─── Seleccionar archivos ─────────────────────────────────────────────
    def _select_excel(self):
        path = filedialog.askopenfilename(
            title="Selecciona el archivo Excel con el inventario",
            filetypes=[
                ("Archivos Excel", "*.xlsx *.xls"),
                ("Todos los archivos", "*.*"),
            ],
        )
        if path:
            self.excel_path = Path(path)
            self.lbl_excel.config(text=self.excel_path.name, fg=FG)
            self.icon_excel.config(text="✅", fg=SUCCESS)
            self._log("Archivo Excel seleccionado: " + self.excel_path.name, "ok")
            self._load_preview()
            self._update_ready_label()

    def _select_pdf(self):
        path = filedialog.askopenfilename(
            title="Selecciona el PDF escaneado con los documentos",
            filetypes=[
                ("Archivos PDF", "*.pdf"),
                ("Todos los archivos", "*.*"),
            ],
        )
        if path:
            self.pdf_path = Path(path)
            self.lbl_pdf.config(text=self.pdf_path.name, fg=FG)
            self.icon_pdf.config(text="✅", fg=SUCCESS)
            self._log("Archivo PDF seleccionado: " + self.pdf_path.name, "ok")
            self._update_ready_label()

    def _select_output_dir(self):
        path = filedialog.askdirectory(
            title="Elige la carpeta donde se guardarán los documentos",
            initialdir=str(self.output_dir),
        )
        if path:
            self.output_dir = Path(path)
            self.lbl_output.config(text=str(self.output_dir))
            self._log("Carpeta de salida cambiada a: " + str(self.output_dir), "info")

    # ─── Cargar vista previa del Excel ────────────────────────────────────
    def _load_preview(self):
        try:
            self.df = load_excel(self.excel_path, skip_rows=config.SKIP_ROWS)
        except Exception as e:
            messagebox.showerror(
                "No se pudo leer el Excel",
                f"Ocurrió un error al abrir el archivo:\n\n{e}\n\n"
                "Verifica que sea un archivo .xlsx válido.",
            )
            self._log(f"Error al leer Excel: {e}", "err")
            return

        # Limpiar tabla
        for item in self.tree.get_children():
            self.tree.delete(item)

        total = len(self.df)
        for fila_excel, row in self.df.iterrows():
            reg       = str(row.get(config.COL_REGISTRO, "")).strip()
            escribano = str(row.get(config.COL_ESCRIBANO, "")).strip()[:25]
            prot      = str(row.get(config.COL_PROTOCOLO, "")).strip()
            folios    = str(row.get(config.COL_FOLIOS, "")).strip()
            titulo    = str(row.get(config.COL_TITULO_EST, "")).strip()[:25]
            int1      = str(row.get(config.COL_INT1, "")).strip()[:25]

            # Limpiar "nan"
            if folios.lower() == "nan":
                folios = ""
            if reg.lower() == "nan":
                reg = ""

            tag = "sin_folio" if not folios else ""
            self.tree.insert("", "end",
                             values=(fila_excel, reg, escribano, prot, folios, titulo, int1),
                             tags=(tag,))

        self.tree.tag_configure("sin_folio", foreground=WARNING)

        # Actualizar spinboxes con filas reales del Excel
        first_row = self.df.index.min()
        last_row  = self.df.index.max()
        self.spin_start.delete(0, "end")
        self.spin_start.insert(0, str(first_row))
        self.spin_end.delete(0, "end")
        self.spin_end.insert(0, str(last_row))
        self.spin_start.config(from_=first_row, to=last_row)
        self.spin_end.config(from_=first_row, to=last_row)

        sin_folio = sum(
            1 for _, r in self.df.iterrows()
            if not str(r.get(config.COL_FOLIOS, "")).strip()
               or str(r.get(config.COL_FOLIOS, "")).strip().lower() == "nan"
        )
        self.lbl_table_info.config(
            text=f"  {total} filas encontradas  ·  {sin_folio} sin folios (en amarillo)",
        )
        self._log(f"Excel cargado: {total} filas, {sin_folio} sin folios.", "info")

    # ─── Log en tiempo real ───────────────────────────────────────────────
    def _log(self, message, tag="info"):
        """Agrega un mensaje al registro de actividad visible."""
        timestamp = time.strftime("%H:%M:%S")
        self.txt_log.config(state="normal")
        self.txt_log.insert("end", f"[{timestamp}]  {message}\n", tag)
        self.txt_log.see("end")
        self.txt_log.config(state="disabled")

    def _poll_log_queue(self):
        """Revisa la cola de logs y los muestra en la GUI."""
        while not self.log_queue.empty():
            try:
                msg = self.log_queue.get_nowait()
                tag = "info"
                if "ERROR" in msg or "CRITICO" in msg:
                    tag = "err"
                elif "WARNING" in msg or "OMITIDO" in msg:
                    tag = "warn"
                elif "OK" in msg:
                    tag = "ok"
                self._log(msg, tag)
            except queue.Empty:
                break
        self.after(200, self._poll_log_queue)

    # ─── Procesar ─────────────────────────────────────────────────────────
    def _start_processing(self):
        if self.processing:
            return

        # Validaciones claras
        if not self.excel_path:
            messagebox.showwarning(
                "Falta el Excel",
                "No has seleccionado un archivo Excel.\n\n"
                "Ve al Paso ① y haz clic en «Seleccionar…» junto a 'Archivo Excel'.",
            )
            return

        if not self.pdf_path:
            messagebox.showwarning(
                "Falta el PDF",
                "No has seleccionado un archivo PDF.\n\n"
                "Ve al Paso ① y haz clic en «Seleccionar…» junto a 'Archivo PDF'.",
            )
            return

        if self.df is None or len(self.df) == 0:
            messagebox.showwarning(
                "Excel vacío",
                "El archivo Excel no contiene datos para procesar.\n"
                "Verifica que sea el archivo correcto.",
            )
            return

        self.processing = True
        self.cancel_flag = False
        self.btn_process.config(state="disabled", text="Procesando…", bg="#555577")
        self.btn_cancel.config(state="normal")
        self.results_frame.pack_forget()
        self.progress["value"] = 0
        self.lbl_progress.config(text="Preparando…")
        self._log("— Proceso iniciado —", "ok")

        thread = threading.Thread(target=self._run_process, daemon=True)
        thread.start()

    def _cancel_processing(self):
        if self.processing:
            self.cancel_flag = True
            self._log("Cancelación solicitada. Esperando que termine la fila actual…", "warn")
            self.btn_cancel.config(state="disabled", text="Cancelando…")

    def _on_offset_change(self):
        """Actualiza los valores de offset y refresca la etiqueta descriptiva."""
        try:
            fi = int(self.spin_folio_ini.get())
            pp = int(self.spin_pdf_ini.get())
            self.folio_inicio_excel = max(1, fi)
            self.pdf_page_inicio    = max(1, pp)
        except ValueError:
            return
        portadas = self.pdf_page_inicio - 1
        texto = (
            f"  Folio {self.folio_inicio_excel}r = pagina {self.pdf_page_inicio} del PDF"
            + (f"  ({portadas} pag. introductorias)" if portadas > 0 else "")
        )
        self.lbl_offset_info.config(text=texto)

    def _add_segment(self):
        """Agrega un segmento adicional a la tabla usando notacion de folio (401r, 53v, etc.)."""
        folio_text = self.ent_seg_folio.get().strip()
        page_abs   = folio_text_to_page_abs(folio_text)
        if page_abs is None:
            messagebox.showwarning(
                "Formato invalido",
                f"'{folio_text}' no es un folio valido.\n"
                "Usa el formato: numero + r o v  (ej: 401r, 53v, 30r)"
            )
            return
        try:
            pag = int(self.ent_seg_pag.get().strip())
            if pag < 1:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Valor invalido", "Pagina PDF debe ser un numero entero positivo.")
            return
        # Mostramos el folio en la tabla pero guardamos page_abs internamente
        self.tree_segs.insert("", "end", values=(folio_text, pag), tags=(str(page_abs),))
        self._auto_refresh_coverage()

    def _del_segment(self):
        """Elimina el segmento seleccionado en la tabla."""
        selected = self.tree_segs.selection()
        if not selected:
            messagebox.showinfo("Sin seleccion", "Selecciona un segmento en la tabla para eliminarlo.")
            return
        for item in selected:
            self.tree_segs.delete(item)
        self._auto_refresh_coverage()

    # ─── Páginas a ignorar ────────────────────────────────────────────────
    def _parse_ignore_input(self, texto: str):
        """
        Parsea la entrada del usuario para páginas ignoradas.
        Acepta:
          - Número simple:  "5"   → {5}
          - Rango:          "10-15" → {10, 11, 12, 13, 14, 15}
        Retorna (set_de_paginas, etiqueta_display) o (None, msg_error).
        """
        texto = texto.strip()
        import re as _re
        m_range = _re.match(r'^(\d+)\s*[-–]\s*(\d+)$', texto)
        m_single = _re.match(r'^(\d+)$', texto)
        if m_range:
            ini = int(m_range.group(1))
            fin = int(m_range.group(2))
            if fin < ini:
                return None, f"Rango incoherente: {ini} > {fin}"
            if ini < 1:
                return None, "El numero de pagina debe ser >= 1"
            return set(range(ini, fin + 1)), f"{ini}-{fin}"
        elif m_single:
            n = int(m_single.group(1))
            if n < 1:
                return None, "El numero de pagina debe ser >= 1"
            return {n}, str(n)
        else:
            return None, f"Formato no reconocido: '{texto}'\nUsa un numero (ej: 5) o rango (ej: 10-15)"

    def _add_ignored_page(self):
        """Agrega una pagina o rango de paginas a la lista de ignorados."""
        texto = self.ent_ignore_pag.get().strip()
        paginas, etiqueta_o_error = self._parse_ignore_input(texto)
        if paginas is None:
            messagebox.showwarning("Entrada invalida", etiqueta_o_error)
            return
        # Mostrar en la tabla
        paginas_sorted = sorted(paginas)
        detalle = ", ".join(str(p) for p in paginas_sorted[:20])
        if len(paginas_sorted) > 20:
            detalle += f" … ({len(paginas_sorted)} paginas)"
        self.tree_ignore.insert("", "end",
                                values=(etiqueta_o_error, detalle),
                                tags=("|".join(str(p) for p in paginas_sorted),))
        self._refresh_ignore_label()
        self._auto_refresh_coverage()

    def _del_ignored_page(self):
        """Elimina la entrada de ignorados seleccionada (pagina o folio)."""
        selected = self.tree_ignore.selection()
        if not selected:
            messagebox.showinfo("Sin seleccion",
                                "Selecciona una entrada en la tabla de ignorados para eliminarla.")
            return
        for item in selected:
            self.tree_ignore.delete(item)
        self._refresh_ignore_label()
        self._auto_refresh_coverage()

    def _add_ignored_folio(self):
        """
        Agrega un folio o rango de folios a la lista de exclusión de conteo.
        Solo afecta cuántas páginas se asignan al registro; el PDF se extrae completo.
        El usuario escribe: 40r, 40v, 40r-41v, etc.
        """
        texto = self.ent_ignore_folio.get().strip()
        if not texto:
            messagebox.showwarning("Entrada vacía", "Escribe un folio o rango de folios.")
            return

        segments = self._get_segments()
        pages, err = parse_folio_range(texto, segments=segments)
        if err or not pages:
            msg = err or "No se pudo calcular páginas para ese folio."
            self.lbl_ignore_folio_preview.config(text=f"  →  Error: {msg}", fg=ERROR)
            messagebox.showwarning("Folio inválido", msg)
            return

        paginas_sorted = sorted(set(pages))
        detalle = ", ".join(str(p) for p in paginas_sorted[:20])
        if len(paginas_sorted) > 20:
            detalle += f" … ({len(paginas_sorted)} posiciones)"

        self.lbl_ignore_folio_preview.config(
            text=f"  →  '{texto}': {len(paginas_sorted)} posición(es) excluidas del conteo (PDF intacto).",
            fg=WARNING)

        self.tree_ignore_folios.insert("", "end",
                                       values=(texto, detalle),
                                       tags=("|".join(str(p) for p in paginas_sorted),))
        self._refresh_ignore_folio_label()
        self._auto_refresh_coverage()

    def _refresh_ignore_label(self):
        """Actualiza el contador de paginas ignoradas."""
        total = len(self._get_ignored_pages())
        if total == 0:
            self.lbl_ignore_count.config(
                text="  Sin páginas ignoradas.", fg=FG_MUTED)
        else:
            self.lbl_ignore_count.config(
                text=f"  {total} página(s) PDF serán ignoradas en la fragmentación.",
                fg=WARNING)

    def _get_ignored_pages(self) -> set:
        """Devuelve el set completo de paginas PDF a ignorar, uniendo todas las entradas."""
        result = set()
        for item in self.tree_ignore.get_children():
            tags = self.tree_ignore.item(item, "tags")
            if tags:
                try:
                    for p_str in tags[0].split("|"):
                        if p_str:
                            result.add(int(p_str))
                except (ValueError, IndexError):
                    pass
        return result

    def _del_ignored_folio(self):
        """Elimina la entrada seleccionada de la lista de folios ignorados en el conteo."""
        selected = self.tree_ignore_folios.selection()
        if not selected:
            messagebox.showinfo("Sin selección",
                                "Selecciona un folio en la tabla para eliminarlo.")
            return
        for item in selected:
            self.tree_ignore_folios.delete(item)
        self._refresh_ignore_folio_label()
        self._auto_refresh_coverage()

    def _get_ignored_folio_pages(self) -> set:
        """
        Devuelve el set de posiciones de página (mapeadas con segmentos) que
        corresponden a folios ausentes del protocolo.
        Estas posiciones se excluyen del conteo de páginas por registro;
        el PDF se extrae sin omitir ninguna página.
        """
        result = set()
        for item in self.tree_ignore_folios.get_children():
            tags = self.tree_ignore_folios.item(item, "tags")
            if tags:
                try:
                    for p_str in tags[0].split("|"):
                        if p_str:
                            result.add(int(p_str))
                except (ValueError, IndexError):
                    pass
        return result

    def _refresh_ignore_folio_label(self):
        """Actualiza el contador de folios ignorados en el conteo."""
        total = len(self._get_ignored_folio_pages())
        if total == 0:
            self.lbl_ignore_folio_count.config(
                text="  Sin folios ignorados en el conteo.", fg=FG_MUTED)
        else:
            self.lbl_ignore_folio_count.config(
                text=f"  {total} posición(es) de folio excluidas del conteo (PDF se extrae completo).",
                fg=WARNING)

    def _auto_refresh_coverage(self):
        """Re-ejecuta el analisis de cobertura PDF si hay Excel y PDF cargados."""
        if self.df is not None and self.pdf_path is not None:
            self._analyze_pdf_coverage()

    def _generate_fragmentation_report(self):
        """
        Genera una ventana con el reporte detallado de fragmentacion:
          - Configuracion activa (segmentos, paginas ignoradas, folios ignorados)
          - Columnas del Excel usadas
          - Mapeo fila -> folio -> paginas PDF con flags de modificacion
        """
        import time as _time

        if self.df is None:
            messagebox.showwarning("Sin datos", "Primero carga un archivo Excel.")
            return

        segments           = self._get_segments()
        ignored_pages      = self._get_ignored_pages()
        ignored_folio_pgs  = self._get_ignored_folio_pages()

        # ── Recopilar entradas de segmentos con etiquetas ──────────────────
        seg_labels = []
        for item in self.tree_segs.get_children():
            vals = self.tree_segs.item(item, "values")
            seg_labels.append(f"Folio {vals[0]} → Pág. PDF {vals[1]}")

        ign_page_labels = []
        for item in self.tree_ignore.get_children():
            vals = self.tree_ignore.item(item, "values")
            ign_page_labels.append(f"{vals[0]}  →  {vals[1]}")

        ign_folio_labels = []
        for item in self.tree_ignore_folios.get_children():
            vals = self.tree_ignore_folios.item(item, "values")
            ign_folio_labels.append(f"{vals[0]}  →  {vals[1]}")

        # ── Construir texto del reporte ────────────────────────────────────
        SEP  = "═" * 68
        SEP2 = "─" * 68
        lines = []
        ts = _time.strftime("%Y-%m-%d  %H:%M:%S")

        lines.append(SEP)
        lines.append("  REPORTE DE FRAGMENTACIÓN")
        lines.append(f"  Generado: {ts}")
        if self.excel_path:
            lines.append(f"  Excel   : {self.excel_path.name}")
        if self.pdf_path:
            lines.append(f"  PDF     : {self.pdf_path.name}")
        lines.append(SEP)

        # Configuración activa
        lines.append("")
        lines.append("  ── CONFIGURACIÓN ACTIVA ──────────────────────────────────────")
        lines.append("")
        lines.append(f"  Folio inicio Excel : {self.folio_inicio_excel}r   →  Pág. PDF inicio: {self.pdf_page_inicio}")
        lines.append("")

        if seg_labels:
            lines.append("  [SALTO PDF]  Segmentos adicionales:")
            for s in seg_labels:
                lines.append(f"    • {s}")
        else:
            lines.append("  [SALTO PDF]  Sin segmentos adicionales.")
        lines.append("")

        if ign_page_labels:
            lines.append("  [PAG-IGN]  Páginas PDF ignoradas (excluidas del output):")
            for s in ign_page_labels:
                lines.append(f"    • {s}")
        else:
            lines.append("  [PAG-IGN]  Sin páginas PDF ignoradas.")
        lines.append("")

        if ign_folio_labels:
            lines.append("  [FOL-IGN]  Folios ignorados en el conteo (PDF intacto):")
            for s in ign_folio_labels:
                lines.append(f"    • {s}")
        else:
            lines.append("  [FOL-IGN]  Sin folios ignorados en el conteo.")
        lines.append("")

        # Columnas del Excel
        import config as _cfg
        lines.append(SEP2)
        lines.append("  COLUMNAS DEL EXCEL UTILIZADAS")
        lines.append(SEP2)
        col_map = [
            ("Folios",        _cfg.COL_FOLIOS),
            ("Registro",      _cfg.COL_REGISTRO),
            ("Escribano",     _cfg.COL_ESCRIBANO.replace("\n", " ")),
            ("Protocolo",     _cfg.COL_PROTOCOLO),
            ("Lugar (Tóp.)",  _cfg.COL_LUGAR),
            ("Fecha inicial", _cfg.COL_FECHA_INI),
            ("Título est.",   _cfg.COL_TITULO_EST),
            ("Interesado 1",  _cfg.COL_INT1),
            ("Interesado 2",  _cfg.COL_INT2),
            ("Observaciones", _cfg.COL_OBS),
        ]
        for label, col in col_map:
            lines.append(f"  {label:<15} : {col}")
        lines.append("")

        # Mapeo folio → páginas PDF
        lines.append(SEP2)
        lines.append("  MAPEO  FILA → FOLIO(S) → PÁGINAS PDF")
        lines.append(SEP2)
        hdr = f"  {'Fila':<6} {'Reg.':<6} {'Folio(s)':<14} {'Págs. PDF (inicio-fin)':<28} {'Flags'}"
        lines.append(hdr)
        lines.append("  " + "-" * 66)

        try:
            row_start = int(self.spin_start.get())
            row_end   = int(self.spin_end.get())
        except ValueError:
            row_start, row_end = 0, 0

        df_slice = self.df.copy()
        if row_start > 0:
            df_slice = df_slice[df_slice.index >= row_start]
        if row_end > 0:
            df_slice = df_slice[df_slice.index <= row_end]

        prev_active_seg = segments[0] if segments else None
        prev_pdf_last   = None
        prev_fila       = None

        for fila_excel, row in df_slice.iterrows():
            folio_str = str(row.get(_cfg.COL_FOLIOS, "")).strip()
            reg       = str(row.get(_cfg.COL_REGISTRO, "")).strip()
            if not folio_str or folio_str.lower() == "nan":
                lines.append(f"  {str(fila_excel):<6} {reg:<6} {'(sin folio)':<14} {'—':<28}")
                prev_fila = fila_excel
                continue

            pages, err = parse_folio_range(folio_str, segments=segments)
            if err or not pages:
                lines.append(f"  {str(fila_excel):<6} {reg:<6} {folio_str:<14} ERROR: {err}")
                prev_fila = fila_excel
                continue

            # Ajustar por folios físicamente ausentes (shift solo en el segmento activo)
            nominal_pages = pages[:]
            if ignored_folio_pgs:
                # Calcular active_seg para este registro
                _first_txt = folio_str.split('-')[0].strip()
                _abs_ini   = folio_text_to_page_abs(_first_txt)
                _act_seg   = _find_segment(_abs_ini, segments) if _abs_ini and len(segments) > 1 else (segments[0] if segments else None)
                pages = adjust_pages_for_missing_folios(
                    pages, ignored_folio_pgs, active_seg=_act_seg, segments=segments
                )
            else:
                pages = list(pages)

            if not pages:
                lines.append(f"  {str(fila_excel):<6} {reg:<6} {folio_str:<14} (sin páginas efectivas)")
                prev_fila = fila_excel
                continue

            p_min, p_max = pages[0], pages[-1]
            sub_lines    = []   # líneas de detalle con excepciones

            # ── Cambio de segmento ─────────────────────────────────────────
            seg_changed = False
            if len(segments) > 1:
                first_folio_txt = folio_str.split('-')[0].strip()
                abs_ini = folio_text_to_page_abs(first_folio_txt)
                if abs_ini is not None:
                    active_seg = _find_segment(abs_ini, segments)
                    if active_seg != prev_active_seg:
                        seg_changed = True
                        sep = (
                            f"  {'':6} {'':6}  ↳ CAMBIO DE SEGMENTO"
                            + (f" (fila {prev_fila} terminó en pág. {prev_pdf_last})" if prev_pdf_last else "")
                            + f"  →  nuevo segmento inicia en pág.PDF {active_seg[1]}"
                        )
                        lines.append(sep)
                    prev_active_seg = active_seg

            # ── Sucesión con el registro anterior (mismo segmento) ────────────
            if prev_pdf_last is not None and not seg_changed:
                expected = prev_pdf_last + 1
                if p_min > expected:
                    sub_lines.append(
                        f"             ⚠ BRECHA: {p_min - expected} pág(s) sin asignar "
                        f"(págs. {expected}–{p_min-1}) tras fila {prev_fila}"
                    )
                elif p_min < expected:
                    sub_lines.append(
                        f"             ⚠ SOLAPAMIENTO: {expected - p_min} pág(s) "
                        f"compartidas con fila {prev_fila} (desde pág. {p_min})"
                    )

            # ── Páginas PDF ignoradas en este rango ───────────────────────
            pag_ign = sorted(p for p in nominal_pages if p in ignored_pages) if ignored_pages else []
            if pag_ign:
                sub_lines.append(
                    f"             ↳ [PAG-IGN] Excluidas del output PDF: "
                    f"{pag_ign[:10]}{'…' if len(pag_ign)>10 else ''}"
                )

            # ── Folios ignorados ──────────────────────────────
            fol_nom_ign = sorted(p for p in nominal_pages if p in ignored_folio_pgs) if ignored_folio_pgs else []
            if fol_nom_ign:
                n_missing = len(fol_nom_ign)
                shift_val = sum(1 for ign in fol_nom_ign if ign <= p_max)
                sub_lines.append(
                    f"             ↳ [FOL-IGN] {n_missing} folio(s) ausentes del escaneo: "
                    f"{fol_nom_ign[:10]}{'…' if n_missing>10 else ''} "
                    f"(págs. posteriores desplazadas -{shift_val})"
                )

            # ── Descripción de páginas ──────────────────────────────
            n_nominal  = len(nominal_pages)
            n_physical = len(pages)
            if fol_nom_ign:
                pag_desc = f"{p_min}–{p_max}  ({n_nominal} nom. → {n_physical} físicas)"
            else:
                pag_desc = f"{p_min}–{p_max}  ({len(pages)} págs.)"

            inline = ""
            if seg_changed:       inline = "[NUEVO SEG.]"
            elif pag_ign:         inline = f"[{len(pag_ign)} pag.ignorada(s)]"
            elif fol_nom_ign:     inline = f"[{len(fol_nom_ign)} folio(s) ausentes]"
            elif sub_lines:       inline = "[VER DETALLE ↓]"

            lines.append(f"  {str(fila_excel):<6} {reg:<6} {folio_str:<14} {pag_desc:<34} {inline}")
            lines.extend(sub_lines)

            prev_pdf_last = p_max
            prev_fila     = fila_excel

        lines.append("")
        lines.append(SEP)
        lines.append("  FIN DEL REPORTE")
        lines.append(SEP)
        report_text = "\n".join(lines)

        # ── Ventana de reporte ─────────────────────────────────────────────
        win = tk.Toplevel(self)
        win.title("Reporte de Fragmentación")
        win.configure(bg=BG)
        win.geometry("920x620")
        win.minsize(700, 400)

        top_bar = tk.Frame(win, bg=BG)
        top_bar.pack(fill="x", padx=16, pady=(12, 4))
        tk.Label(top_bar, text="\U0001f4cb  Reporte de Fragmentación",
                 font=FONT_STEP, bg=BG, fg=FG).pack(side="left")

        def _save_report():
            from tkinter import filedialog as _fd
            path = _fd.asksaveasfilename(
                title="Guardar reporte",
                defaultextension=".txt",
                filetypes=[("Texto", "*.txt"), ("Todos", "*.*")],
                initialfile=f"reporte_fragmentacion_{_time.strftime('%Y%m%d_%H%M%S')}.txt",
            )
            if path:
                try:
                    with open(path, "w", encoding="utf-8") as f:
                        f.write(report_text)
                    messagebox.showinfo("Guardado", f"Reporte guardado en:\n{path}")
                except Exception as exc:
                    messagebox.showerror("Error", f"No se pudo guardar:\n{exc}")

        btn_save = tk.Button(
            top_bar, text="\U0001f4be  Guardar .txt", font=FONT_BOLD,
            bg="#3a3a7a", fg=FG, activebackground="#5050aa", activeforeground=FG,
            relief="flat", cursor="hand2", padx=12, pady=4,
            command=_save_report,
        )
        btn_save.pack(side="right")

        txt = tk.Text(
            win, bg=BG_INPUT, fg=FG, font=FONT_LOG,
            relief="flat", wrap="none", padx=12, pady=10,
            insertbackground=FG,
        )
        sb_y = ttk.Scrollbar(win, orient="vertical",   command=txt.yview)
        sb_x = ttk.Scrollbar(win, orient="horizontal", command=txt.xview)
        txt.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)

        sb_y.pack(side="right",  fill="y")
        sb_x.pack(side="bottom", fill="x")
        txt.pack(fill="both", expand=True, padx=(16, 0), pady=(4, 0))

        # Tags para resaltar
        txt.tag_configure("header",   foreground=ACCENT_LIGHT, font=("Consolas", 9, "bold"))
        txt.tag_configure("seg_flag", foreground="#ffd700")
        txt.tag_configure("pag_flag", foreground=ERROR)
        txt.tag_configure("fol_flag", foreground=WARNING)
        txt.tag_configure("col_key",  foreground=FG_DIM)
        txt.tag_configure("normal",   foreground=FG)

        for line in lines:
            if line.startswith("═") or line.startswith("─"):
                txt.insert("end", line + "\n", "header")
            elif line.strip().startswith("REPORTE") or line.strip().startswith("FIN DEL") \
                    or line.strip().startswith("CONFIGURACIÓN") \
                    or line.strip().startswith("COLUMNAS") \
                    or line.strip().startswith("MAPEO"):
                txt.insert("end", line + "\n", "header")
            elif "[SALTO" in line or "[SALTO PDF]" in line:
                txt.insert("end", line + "\n", "seg_flag")
            elif "[PAG-IGN" in line:
                txt.insert("end", line + "\n", "pag_flag")
            elif "[FOL-IGN" in line:
                txt.insert("end", line + "\n", "fol_flag")
            elif " : " in line and "Folio" not in line and "PDF" not in line \
                    and line.strip().startswith(tuple(c[0] for c in col_map)):
                txt.insert("end", line + "\n", "col_key")
            else:
                txt.insert("end", line + "\n", "normal")

        txt.config(state="disabled")

    def _get_segments(self) -> list:
        """
        Devuelve la lista de segmentos como (page_abs_inicio, pdf_page_inicio),
        ordenada de menor a mayor page_abs.
        El primer segmento siempre proviene de los spinboxes principales.
        """
        # Primer segmento: folio_inicio_excel (siempre recto) -> pdf_page_inicio
        first_page_abs = _folio_to_page_abs(self.folio_inicio_excel, 'r')
        segs = [(first_page_abs, self.pdf_page_inicio)]
        for item in self.tree_segs.get_children():
            tags = self.tree_segs.item(item, "tags")
            vals = self.tree_segs.item(item, "values")
            try:
                # El tag almacena el page_abs calculado al insertar
                page_abs = int(tags[0]) if tags else folio_text_to_page_abs(str(vals[0]))
                pdf_pag  = int(vals[1])
                if page_abs and pdf_pag > 0:
                    segs.append((page_abs, pdf_pag))
            except (ValueError, IndexError, TypeError):
                pass
        return sorted(segs, key=lambda x: x[0])

    # ─── Analizador de sucesión de folios ─────────────────────────────────
    def _analyze_folios(self):
        """Ejecuta el analisis de sucesion de folios y muestra el resultado con numeros de fila."""
        if self.df is None:
            messagebox.showwarning("Sin datos", "Primero carga un archivo Excel.")
            return

        folio_list = []
        row_indices = []
        for fila_excel, row in self.df.iterrows():
            folio_list.append(str(row.get(config.COL_FOLIOS, "")).strip())
            row_indices.append(fila_excel)

        result = analyze_folio_sequence(folio_list, indices=row_indices)
        self._show_analyzer_result("ANALISIS DE SUCESION DE FOLIOS", result["summary"], result["ok"])

    # ─── Analizador de DATA TÓPICA ─────────────────────────────────────────
    def _analyze_data_topica(self):
        """Ejecuta el análisis de DATA TÓPICA y muestra el resultado."""
        if self.df is None:
            messagebox.showwarning("Sin datos", "Primero carga un archivo Excel.")
            return

        registros = [row.to_dict() for _, row in self.df.iterrows()]
        result = analyze_data_topica(
            registros,
            col_topica=config.COL_LUGAR,
            col_registro=config.COL_REGISTRO,
        )
        self._show_analyzer_result("ANALISIS DATA TOPICA", result["summary"], result["ok"])

    def _analyze_data_cronica(self):
        """Ejecuta el analisis de DATA CRONICA y muestra el resultado."""
        if self.df is None:
            messagebox.showwarning("Sin datos", "Primero carga un archivo Excel.")
            return
        registros = [row.to_dict() for _, row in self.df.iterrows()]
        result = analyze_data_cronica(
            registros,
            col_fecha_ini=config.COL_FECHA_INI,
            col_fecha_fin="FECHA FINAL",
            col_registro=config.COL_REGISTRO,
        )
        self._show_analyzer_result("ANALISIS DATA CRONICA", result["summary"], result["ok"])

    # ─── Verificador de cobertura PDF ──────────────────────────────────────
    def _analyze_pdf_coverage(self):
        """
        Verifica si la ultima hoja del PDF coincide exactamente con el
        final del rango configurado (filas manuales + segmentos adicionales).

        Logica:
          1. Requiere PDF cargado para conocer total_pdf_pages.
          2. Toma el rango de filas (spin_start / spin_end).
          3. Toma los segmentos (spinboxes principales + segmentos adicionales).
          4. Para cada fila del Excel en el rango, convierte sus folios a paginas PDF.
          5. La mayor pagina PDF usada = ultima pagina consumida.
          6. Compara con total_pdf_pages del PDF.
        """
        lines = []

        # ── Validaciones previas ──────────────────────────────────────────
        if self.df is None:
            lines.append("[ERR]  No hay Excel cargado. Carga primero un archivo Excel.")
            self._show_analyzer_result("VERIFICACION DE COBERTURA PDF", "\n".join(lines), False)
            return

        if self.pdf_path is None:
            lines.append("[ERR]  No hay PDF seleccionado. Selecciona primero el archivo PDF.")
            self._show_analyzer_result("VERIFICACION DE COBERTURA PDF", "\n".join(lines), False)
            return

        # ── Leer total de paginas del PDF ─────────────────────────────────
        try:
            reader = open_pdf(self.pdf_path)
            total_pdf_pages = len(reader.pages)
        except Exception as exc:
            lines.append(f"[ERR]  No se pudo abrir el PDF: {exc}")
            self._show_analyzer_result("VERIFICACION DE COBERTURA PDF", "\n".join(lines), False)
            return

        # ── Obtener rango de filas ────────────────────────────────────────
        try:
            row_start = int(self.spin_start.get())
            row_end   = int(self.spin_end.get())
        except ValueError:
            lines.append("[ERR]  Rango de filas no valido (revisa los spinboxes).")
            self._show_analyzer_result("VERIFICACION DE COBERTURA PDF", "\n".join(lines), False)
            return

        # ── Filtrar el DataFrame al rango seleccionado ────────────────────
        df_slice = self.df.copy()
        if row_start > 0:
            df_slice = df_slice[df_slice.index >= row_start]
        if row_end > 0:
            df_slice = df_slice[df_slice.index <= row_end]

        if len(df_slice) == 0:
            lines.append("[ERR]  El rango de filas seleccionado no contiene datos.")
            self._show_analyzer_result("VERIFICACION DE COBERTURA PDF", "\n".join(lines), False)
            return

        # ── Obtener segmentos de offset ───────────────────────────────────
        segments = self._get_segments()

        # ── Calcular la mayor pagina PDF utilizada ────────────────────────
        max_page_used = 0
        last_valid_folio = ""
        last_valid_fila  = None
        invalid_count    = 0

        for fila_excel, row in df_slice.iterrows():
            folio_str = str(row.get(config.COL_FOLIOS, "")).strip()
            if not folio_str or folio_str.lower() == "nan":
                continue
            pages, err = parse_folio_range(folio_str, segments=segments)
            if err or not pages:
                invalid_count += 1
                continue
            ultimo = pages[-1]
            if ultimo > max_page_used:
                max_page_used    = ultimo
                last_valid_folio = folio_str
                last_valid_fila  = fila_excel

        # ── Calcular paginas del rango de segmentos adicionales ───────────
        # Los segmentos adicionales redefinen offsets pero no por si solos
        # consumen paginas; solo lo hacen si hay folios en ese tramo.
        # max_page_used ya contempla todos los folios del rango.

        # ── Construir informe ─────────────────────────────────────────────
        rango_desc = (
            f"Filas Excel {df_slice.index.min()} – {df_slice.index.max()}"
            + (" (todas las filas)" if row_end == 0 else "")
        )
        seg_count = len(segments)

        lines.append(f"  PDF seleccionado   : {self.pdf_path.name}")
        lines.append(f"  Total paginas PDF  : {total_pdf_pages}")
        lines.append(f"  Rango de filas     : {rango_desc}")
        lines.append(f"  Segmentos offset   : {seg_count}  {segments}")
        lines.append(f"  Registros en rango : {len(df_slice)}")
        if invalid_count:
            lines.append(f"  [!!]  Folios invalidos ignorados: {invalid_count}")
        lines.append("")

        if max_page_used == 0:
            lines.append("[ERR]  No se encontro ningun folio valido en el rango. "
                         "Verifica el Excel y los segmentos.")
            self._show_analyzer_result("VERIFICACION DE COBERTURA PDF", "\n".join(lines), False)
            return

        lines.append(f"  Ultima pagina PDF usada por el rango: {max_page_used}")
        lines.append(f"    (corresponde al folio '{last_valid_folio}', fila Excel {last_valid_fila})")
        lines.append(f"  Ultima pagina real del PDF           : {total_pdf_pages}")
        lines.append("")

        diferencia = total_pdf_pages - max_page_used

        if diferencia == 0:
            lines.append("[OK]  PERFECTO: La ultima hoja del PDF coincide exactamente con el")
            lines.append("      final del rango. El PDF se separara sin que sobre ninguna hoja.")
            ok = True
        elif diferencia > 0:
            lines.append(f"[!!]  ADVERTENCIA: Sobrarian {diferencia} pagina(s) del PDF sin asignar.")
            lines.append(f"      El rango solo cubre hasta la pagina {max_page_used},")
            lines.append(f"      pero el PDF tiene {total_pdf_pages} paginas en total.")
            lines.append("")
            lines.append("      Posibles causas:")
            lines.append("        - El rango de filas no llega hasta el ultimo registro del protocolo.")
            lines.append("        - Faltan registros en el Excel para las ultimas hojas del PDF.")
            lines.append("        - El offset (folio inicio / pagina PDF) no esta bien configurado.")
            lines.append("        - Faltan segmentos adicionales para cubrir los folios restantes.")
            ok = False
        else:  # diferencia < 0
            lines.append(f"[ERR]  ERROR: El rango requiere {-diferencia} pagina(s) MAS de las que")
            lines.append(f"      tiene el PDF ({total_pdf_pages} paginas disponibles,")
            lines.append(f"      pero el rango alcanza la pagina {max_page_used}).")
            lines.append("")
            lines.append("      Posibles causas:")
            lines.append("        - El offset (folio inicio / pagina PDF) es incorrecto.")
            lines.append("        - El rango de filas excede lo que este PDF contiene.")
            lines.append("        - Algun segmento adicional apunta a una pagina inexistente.")
            ok = False

        self._show_analyzer_result("VERIFICACION DE COBERTURA PDF", "\n".join(lines), ok)

    def _show_analyzer_result(self, titulo: str, summary: str, ok: bool):
        """Escribe el resultado de un analisis en el panel de analizadores."""
        self.txt_analyzer.config(state="normal")
        self.txt_analyzer.delete("1.0", "end")
        ts = time.strftime("%H:%M:%S")
        sep = "-" * 60
        header = f"[{ts}] {titulo}\n{sep}\n"
        self.txt_analyzer.insert("end", header, "info")
        # Colorear linea a linea segun prefijo ASCII
        for line in summary.splitlines():
            if line.startswith("[OK]"):
                tag = "ok"
            elif line.startswith("[ERR]"):
                tag = "err"
            elif line.startswith("[!!]"):
                tag = "warn"
            else:
                tag = "info"
            self.txt_analyzer.insert("end", line + "\n", tag)
        self.txt_analyzer.see("end")
        self.txt_analyzer.config(state="disabled")

    def _run_process(self):
        try:
            row_start  = int(self.spin_start.get())
            row_end    = int(self.spin_end.get())
            check_gaps = not self.var_no_gap.get()
            output_dir = self.output_dir
            segments             = self._get_segments()        # lista de (folio_ini, pdf_pag_ini)
            ignored_pages        = self._get_ignored_pages()   # set de pags PDF a omitir del output
            ignored_folio_pages  = self._get_ignored_folio_pages()  # folios ausentes: solo ajustan conteo

            if ignored_pages:
                self._log_thread(
                    f"Páginas PDF ignoradas ({len(ignored_pages)}): "
                    + ", ".join(str(p) for p in sorted(ignored_pages))
                )
            if ignored_folio_pages:
                self._log_thread(
                    f"Folios ausentes excluidos del conteo ({len(ignored_folio_pages)} posiciones): "
                    + ", ".join(str(p) for p in sorted(ignored_folio_pages))
                )

            # Filtrar DataFrame usando filas reales del Excel
            df_slice = self.df
            if row_start > 0:
                df_slice = df_slice[df_slice.index >= row_start]
            if row_end > 0:
                df_slice = df_slice[df_slice.index <= row_end]
            total = len(df_slice)

            if total == 0:
                self.after(0, lambda: messagebox.showwarning(
                    "Sin datos",
                    "El rango de filas seleccionado no contiene datos.\n"
                    "Ajusta los valores de 'Desde fila' y 'Hasta fila'.",
                ))
                self._reset_button()
                return

            # Abrir PDF
            self.after(0, lambda: self.lbl_progress.config(text="Abriendo PDF…"))
            self._log_thread("Abriendo archivo PDF…")
            reader = open_pdf(self.pdf_path)
            total_pdf_pages = len(reader.pages)
            self._log_thread(f"PDF abierto: {total_pdf_pages} paginas.")
            self._log_thread(
                f"Segmentos de offset: {segments}"
            )

            # Leer metadatos del fondo (acervo y siglo) UNA SOLA VEZ antes del bucle
            meta = load_excel_metadata(
                self.excel_path,
                meta_row_siglo=config.META_ROW_SIGLO,
                meta_row_acervo=config.META_ROW_ACERVO,
            )
            self._log_thread(f"Fondo: ACERVO DOCUMENTAL NUMERO {meta['acervo_num']}  |  SIGLO {meta['siglo']}")

            processed      = 0
            skipped        = 0
            errors         = []
            prev_last_page = None
            prev_seg_proc  = segments[0] if segments else None
            start_time     = time.time()

            for i, (fila_excel, row) in enumerate(df_slice.iterrows()):
                # Revisar cancelación
                if self.cancel_flag:
                    self._log_thread(f"Proceso cancelado por el usuario en la fila {fila_excel}.")
                    break

                row_dict = row.to_dict()
                reg_id   = row_dict.get(config.COL_REGISTRO, f"fila_{fila_excel}")

                # Actualizar progreso
                pct = int((i + 1) / total * 100)
                msg = f"Fila {fila_excel}  ·  Reg. {reg_id}  ({i+1} de {total})"
                self.after(0, lambda p=pct, m=msg: self._update_progress(p, m))

                # Detectar cambio de segmento para omitir el gap-check en la transición
                folio_str_raw = str(row_dict.get(config.COL_FOLIOS, "")).strip()
                check_prev    = prev_last_page
                curr_seg_proc = prev_seg_proc
                if folio_str_raw and folio_str_raw.lower() != 'nan' and len(segments) > 1:
                    _first_txt = folio_str_raw.split('-')[0].strip()
                    _abs_pg    = folio_text_to_page_abs(_first_txt)
                    if _abs_pg is not None:
                        curr_seg_proc = _find_segment(_abs_pg, segments)
                        if curr_seg_proc != prev_seg_proc:
                            check_prev = None  # no aplicar gap-check al cruzar segmento
                            self._log_thread(
                                f"  [SALTO] Fila {fila_excel}: nuevo segmento → pág.PDF {curr_seg_proc[1]}"
                            )

                # Validar
                is_valid, error_msg = validate_record(
                    row=row_dict, col_folios=config.COL_FOLIOS,
                    total_pdf_pages=total_pdf_pages,
                    prev_last_page=check_prev, check_gaps=check_gaps,
                    segments=segments,
                )

                if not is_valid:
                    skipped += 1
                    errors.append(f"Fila {fila_excel}: {error_msg}")
                    self._log_thread(f"Fila {fila_excel} omitida: {error_msg}")
                    continue

                folio_str = folio_str_raw
                pages, _ = parse_folio_range(folio_str, segments=segments)

                # Ajustar páginas por folios físicamente ausentes del escaneo.
                # El shift solo afecta al segmento activo (curr_seg_proc).
                if ignored_folio_pages and pages:
                    pages = adjust_pages_for_missing_folios(
                        pages, ignored_folio_pages,
                        active_seg=curr_seg_proc, segments=segments,
                    )

                dest_path = build_output_path(
                    output_dir=output_dir,
                    acervo_num=meta["acervo_num"],
                    siglo=meta["siglo"],
                    escribano=str(row_dict.get(config.COL_ESCRIBANO, "")).replace("\n", " "),
                    protocolo=str(row_dict.get(config.COL_PROTOCOLO, "")),
                    registro=str(row_dict.get(config.COL_REGISTRO, "")),
                    titulo_est=str(row_dict.get(config.COL_TITULO_EST, "")),
                    fecha_ini=str(row_dict.get(config.COL_FECHA_INI, "")),
                    interesado1=str(row_dict.get(config.COL_INT1, "")),
                    interesado2=str(row_dict.get(config.COL_INT2, "")),
                )

                success = extract_pages(
                    reader=reader,
                    page_numbers=pages,
                    dest_path=dest_path,
                    ignored_pages=ignored_pages,
                )

                if success:
                    processed += 1
                    # Usar la última página efectiva (después de folios ignorados)
                    prev_last_page = pages[-1] if pages else last_page_of_range(folio_str, segments=segments)
                    prev_seg_proc  = curr_seg_proc
                    self._log_thread(f"OK Fila {fila_excel} -> {dest_path.name}")
                else:
                    skipped += 1
                    errors.append(f"Fila {fila_excel}: Error al extraer páginas")
                    self._log_thread(f"✗  Fila {fila_excel}: Error al extraer páginas")

            elapsed = time.time() - start_time

            cancelled = self.cancel_flag
            self.after(0, lambda: self._show_results(
                total, processed, skipped, elapsed, errors, cancelled))

        except FileNotFoundError as e:
            self.after(0, lambda: messagebox.showerror(
                "Archivo no encontrado",
                f"No se encontró el archivo:\n\n{e}\n\n"
                "Verifica que no se haya movido o eliminado.",
            ))
            self._log_thread(f"ERROR: Archivo no encontrado - {e}")
        except Exception as e:
            self.after(0, lambda: messagebox.showerror(
                "Error inesperado",
                f"Ocurrió un error durante el proceso:\n\n{e}\n\n"
                "Revisa el registro de actividad para más detalles.",
            ))
            self._log_thread(f"ERROR CRÍTICO: {e}")
        finally:
            self.after(0, self._reset_button)

    def _log_thread(self, message):
        """Envía un mensaje al log desde un hilo secundario (thread-safe)."""
        tag = "info"
        if message.startswith("✓"):
            tag = "ok"
        elif message.startswith("✗") or "ERROR" in message:
            tag = "err"
        elif "omitida" in message or "cancelad" in message.lower():
            tag = "warn"
        self.after(0, lambda m=message, t=tag: self._log(m, t))

    def _update_progress(self, pct, msg):
        self.progress["value"] = pct
        self.lbl_progress.config(text=msg)

    def _reset_button(self):
        self.processing = False
        self.cancel_flag = False
        self.btn_process.config(state="normal", text="▶   PROCESAR", bg=ACCENT)
        self.btn_cancel.config(state="disabled", text="✕  Cancelar")

    def _show_results(self, total, processed, skipped, elapsed, errors, cancelled=False):
        self.progress["value"] = 100 if not cancelled else self.progress["value"]

        if cancelled:
            self.lbl_progress.config(text="Cancelado por el usuario")
            self.lbl_result_title.config(
                text=f"Proceso cancelado  ·  {processed} PDFs generados antes de cancelar",
                fg=WARNING,
            )
        elif skipped == 0:
            self.lbl_progress.config(text="¡Completado!")
            self.lbl_result_title.config(
                text=f"✅  ¡Todos los {processed} registros se procesaron correctamente!",
                fg=SUCCESS,
            )
        else:
            self.lbl_progress.config(text="Completado con omisiones")
            self.lbl_result_title.config(
                text=f"Proceso terminado  ·  {processed} exitosos, {skipped} omitidos",
                fg=WARNING,
            )

        body = (
            f"Total de filas:          {total}\n"
            f"PDFs generados:          {processed}\n"
            f"No procesados:           {skipped}\n"
            f"Tiempo:                  {elapsed:.1f} segundos\n"
            f"Carpeta de salida:       {self.output_dir}"
        )
        if errors:
            body += f"\n\nDetalles de omisiones (primeros 10):\n"
            for e in errors[:10]:
                body += f"  · {e}\n"
            if len(errors) > 10:
                body += f"  … y {len(errors)-10} más"

        self.lbl_result_body.config(text=body)
        self.results_frame.pack(fill="x", padx=24, pady=(8, 16))

        self._log("— Proceso finalizado —", "ok")
        self._log(f"  Generados: {processed} | Omitidos: {skipped} | Tiempo: {elapsed:.1f}s", "info")

    def _open_output_folder(self):
        output = self.output_dir
        output.mkdir(parents=True, exist_ok=True)
        os.startfile(str(output))


if __name__ == "__main__":
    app = App()
    app.mainloop()
