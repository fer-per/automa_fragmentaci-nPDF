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
from modules.excel_reader import load_excel
from modules.folio_parser import parse_folio_range, last_page_of_range
from modules.pdf_extractor import open_pdf, extract_pages
from modules.folder_builder import build_output_path
from modules.validator import validate_record


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
            titulo    = str(row.get(config.COL_TITULO, "")).strip()[:25]
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

    def _run_process(self):
        try:
            row_start = int(self.spin_start.get())
            row_end   = int(self.spin_end.get())
            check_gaps = not self.var_no_gap.get()
            output_dir = self.output_dir

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
            self._log_thread(f"PDF abierto: {total_pdf_pages} páginas.")

            processed = 0
            skipped   = 0
            errors    = []
            prev_last_page = None
            start_time = time.time()

            for i, (fila_excel, row) in enumerate(df_slice.iterrows()):
                # Revisar cancelación
                if self.cancel_flag:
                    self._log_thread(f"Proceso cancelado por el usuario en la fila {fila_excel}.")
                    break

                row_dict = row.to_dict()
                reg_id = row_dict.get(config.COL_REGISTRO, f"fila_{fila_excel}")

                # Actualizar progreso
                pct = int((i + 1) / total * 100)
                msg = f"Fila {fila_excel}  ·  Reg. {reg_id}  ({i+1} de {total})"
                self.after(0, lambda p=pct, m=msg: self._update_progress(p, m))

                # Validar
                is_valid, error_msg = validate_record(
                    row=row_dict, col_folios=config.COL_FOLIOS,
                    total_pdf_pages=total_pdf_pages,
                    prev_last_page=prev_last_page, check_gaps=check_gaps,
                )

                if not is_valid:
                    skipped += 1
                    errors.append(f"Fila {fila_excel}: {error_msg}")
                    self._log_thread(f"Fila {fila_excel} omitida: {error_msg}")
                    continue

                folio_str = str(row_dict.get(config.COL_FOLIOS, "")).strip()
                pages, _ = parse_folio_range(folio_str)

                dest_path = build_output_path(
                    output_dir=output_dir,
                    escribano=str(row_dict.get(config.COL_ESCRIBANO, "")),
                    protocolo=str(row_dict.get(config.COL_PROTOCOLO, "")),
                    registro=str(row_dict.get(config.COL_REGISTRO, "")),
                    titulo=str(row_dict.get(config.COL_TITULO, "")),
                    fecha_ini=str(row_dict.get(config.COL_FECHA_INI, "")),
                    interesado1=str(row_dict.get(config.COL_INT1, "")),
                    interesado2=str(row_dict.get(config.COL_INT2, "")),
                )

                success = extract_pages(reader=reader, page_numbers=pages,
                                        dest_path=dest_path)

                if success:
                    processed += 1
                    prev_last_page = last_page_of_range(folio_str)
                    self._log_thread(f"✓  Fila {fila_excel} → {dest_path.name}")
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
