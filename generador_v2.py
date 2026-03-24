"""
generador_v2.py  v3.0
Generador de Evaluaciones AMR — Gemini integrado
"""

import os, sys, json, threading, random, copy
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk

from config import get_api_key, set_api_key
from gemini_client import GeminiClient
from file_parser import extraer_texto
from generador_app import Generador, LOGO_PATH, AMR_PATH

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# ── Tema ──────────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

# ── Paleta Apple ──────────────────────────────────────────────────────────────
BG       = "#F2F2F7"   # system background
PANEL    = "#FFFFFF"   # card / surface
BORDER   = "#E5E5EA"   # separator / subtle border
BORDER2  = "#C7C7CC"   # input border (slightly stronger)
MUTED    = "#8E8E93"   # secondary label
TEXT     = "#1C1C1E"   # primary label
TEXT2    = "#3A3A3C"   # secondary text (slightly lighter)
ACCENT   = "#007AFF"   # iOS blue
ACCENT_D = "#0062CC"
GREEN    = "#34C759"   # iOS green
GREEN_D  = "#28A745"
PURPLE   = "#AF52DE"   # iOS purple
PURPLE_D = "#8944B8"
RED      = "#FF3B30"   # iOS red
CHAT_BG  = "#F9F9FB"
BUB_U    = "#007AFF"   # user bubble
BUB_IA   = "#FFFFFF"   # assistant bubble

FONT = "Segoe UI"

# ── Cursos ────────────────────────────────────────────────────────────────────
CURSOS = [
    "1° Básico", "2° Básico", "3° Básico", "4° Básico",
    "5° Básico", "6° Básico", "7° Básico", "8° Básico",
    "1° Medio",  "2° Medio",  "3° Medio",  "4° Medio",
]

# ── Clave de respuestas ───────────────────────────────────────────────────────
def generar_clave_docx(fichas: list, path: str):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Cm(2)
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("CLAVE DE RESPUESTAS")
    r.bold = True
    r.font.size = Pt(14)
    doc.add_paragraph()
    for ficha in fichas:
        p = doc.add_paragraph()
        p.add_run(f"Ficha N° {ficha.get('numero', '')}").bold = True
        t = doc.add_table(rows=1, cols=2)
        t.style = "Table Grid"
        t.alignment = WD_TABLE_ALIGNMENT.LEFT
        for cell, txt in zip(t.rows[0].cells, ["Pregunta", "Respuesta"]):
            cell.text = txt
            cell.paragraphs[0].runs[0].bold = True
        for preg in ficha.get("preguntas", []):
            row = t.add_row().cells
            row[0].text = str(preg.get("numero", ""))
            row[1].text = preg.get("respuesta_correcta", "")
        doc.add_paragraph()
    doc.save(path)


# ════════════════════════════════════════════════════════════════════════════
#  App
# ════════════════════════════════════════════════════════════════════════════
class App(ctk.CTk):

    def __init__(self):
        super().__init__()
        self.title("Generador de Evaluaciones  ·  AMR")
        self.geometry("1120x740")
        self.minsize(920, 640)
        self.configure(fg_color=BG)
        self.tk.call("tk", "scaling", 1.0)
        try:
            # Evita el flash negro en Windows al redimensionar/maximizar
            self.wm_attributes("-transparentcolor", "")
        except Exception:
            pass

        try:
            icon_path = os.path.join(
                getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__))),
                "icono.ico"
            )
            if os.path.isfile(icon_path):
                self.iconbitmap(icon_path)
        except Exception:
            pass

        self._gemini: GeminiClient | None = None
        self._fichas: list | None         = None
        self._texto_base                  = ""
        self._archivo_nombre              = ""
        self._generando                   = False
        self._ph_activo                   = True
        self._chat_w                      = 600  # se actualiza en resize

        self._resize_job = None
        self._verificar_api_key()
        self._build_ui()
        self.bind("<Configure>", self._on_resize)
        self._burbuja(
            "Hola. Describe la evaluación que necesitas y la genero al instante.\n\n"
            "Ejemplo: «Prueba de comprensión sobre el agua para 3° básico, 4 preguntas»",
            "ia"
        )

    # ── API Key ───────────────────────────────────────────────────────────────
    def _verificar_api_key(self):
        key = get_api_key()
        if key:
            self._gemini = GeminiClient(key)
        else:
            self._pedir_api_key()

    def _pedir_api_key(self):
        dlg = ctk.CTkToplevel(self)
        dlg.title("API Key de Gemini")
        dlg.geometry("480x330")
        dlg.resizable(False, False)
        dlg.configure(fg_color=PANEL)
        dlg.grab_set()

        ctk.CTkLabel(dlg, text="Configura tu API Key",
                     text_color=TEXT,
                     font=ctk.CTkFont(FONT, 18, "bold")).pack(
                         padx=32, pady=(32, 4), anchor="w")
        ctk.CTkLabel(dlg, text="Consíguela gratis en aistudio.google.com",
                     text_color=MUTED,
                     font=ctk.CTkFont(FONT, 11)).pack(padx=32, anchor="w")

        entry = ctk.CTkEntry(dlg, width=416,
                             show="•",
                             placeholder_text="Pega tu API Key aquí",
                             font=ctk.CTkFont("Consolas", 11),
                             fg_color=BG, text_color=TEXT,
                             border_color=BORDER2, border_width=1,
                             height=44, corner_radius=10)
        entry.pack(padx=32, pady=(24, 6))

        lbl_err = ctk.CTkLabel(dlg, text="", text_color=RED,
                               font=ctk.CTkFont(FONT, 10))
        lbl_err.pack(padx=32, anchor="w")

        def guardar():
            key = entry.get().strip()
            if not key:
                lbl_err.configure(text="Pega tu API Key primero.")
                return
            lbl_err.configure(text="Verificando…", text_color=MUTED)
            dlg.update_idletasks()
            cli = GeminiClient(key)
            ok, err = cli.verificar_api_key()
            if ok:
                set_api_key(key)
                self._gemini = cli
                dlg.destroy()
            else:
                lbl_err.configure(
                    text=f"Key inválida: {err[:65]}", text_color=RED)

        ctk.CTkButton(dlg, text="Guardar y continuar",
                      fg_color=ACCENT, hover_color=ACCENT_D,
                      text_color="white",
                      font=ctk.CTkFont(FONT, 12, "bold"),
                      height=44, corner_radius=10,
                      command=guardar).pack(
                          padx=32, pady=(16, 0), fill="x")
        self.wait_window(dlg)

    # ── Build UI ──────────────────────────────────────────────────────────────
    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self._build_header()
        self._build_main()

    def _build_header(self):
        hdr = ctk.CTkFrame(self, fg_color=PANEL, corner_radius=0, height=60)
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.grid_propagate(False)
        hdr.grid_columnconfigure(1, weight=1)

        if AMR_PATH:
            try:
                from PIL import Image
                img = ctk.CTkImage(Image.open(AMR_PATH).convert("RGBA"),
                                   size=(34, 34))
                ctk.CTkLabel(hdr, image=img, text="",
                             fg_color=PANEL, corner_radius=8,
                             width=42, height=42).grid(
                                 row=0, column=0, padx=(18, 0), pady=9)
            except Exception:
                pass

        info = ctk.CTkFrame(hdr, fg_color="transparent")
        info.grid(row=0, column=1, sticky="w", padx=14)
        ctk.CTkLabel(info, text="Generador de Evaluaciones",
                     text_color=TEXT,
                     font=ctk.CTkFont(FONT, 14, "bold")).pack(anchor="w")
        ctk.CTkLabel(info, text="Con inteligencia artificial  ·  AMR  ·  by Levil",
                     text_color=MUTED,
                     font=ctk.CTkFont(FONT, 10)).pack(anchor="w")

        ctk.CTkButton(hdr, text="API Key",
                      fg_color="transparent", hover_color=BG,
                      text_color=MUTED,
                      font=ctk.CTkFont(FONT, 10),
                      width=80, height=32, corner_radius=8,
                      command=self._cambiar_api_key).grid(
                          row=0, column=2, padx=(0, 18))

        # Separador inferior
        ctk.CTkFrame(self, fg_color=BORDER, height=1,
                     corner_radius=0).grid(row=0, column=0, sticky="sew")

    def _build_main(self):
        main = ctk.CTkFrame(self, fg_color=BG, corner_radius=0)
        main.grid(row=1, column=0, sticky="nsew", padx=16, pady=16)
        main.grid_columnconfigure(0, weight=0, minsize=288)
        main.grid_columnconfigure(1, weight=1)
        main.grid_rowconfigure(0, weight=1)
        self._build_panel_izq(main)
        self._build_panel_der(main)

    # ── Panel izquierdo ───────────────────────────────────────────────────────
    def _build_panel_izq(self, parent):
        left = ctk.CTkScrollableFrame(parent, fg_color=PANEL,
                                      corner_radius=14,
                                      border_color=BORDER, border_width=1,
                                      scrollbar_button_color=BORDER,
                                      scrollbar_button_hover_color=MUTED)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left.grid_columnconfigure(0, weight=1)
        self._left_canvas = left._parent_canvas
        left.bind("<Enter>", lambda e: self._set_scroll("left"))
        left.bind("<Leave>", lambda e: self._set_scroll(None))

        row = 0

        # ── Pedido ────────────────────────────────────────────────────────────
        row = self._sec(left, "PEDIDO", row)

        self._txt = ctk.CTkTextbox(left, height=110, wrap="word",
                                   font=ctk.CTkFont(FONT, 10),
                                   fg_color=BG, text_color=MUTED,
                                   border_color=BORDER, border_width=1,
                                   corner_radius=10)
        self._txt.grid(row=row, column=0, sticky="ew", padx=14, pady=(6, 0))
        self._txt.insert("1.0",
            "Ej: Prueba de comprensión sobre la Independencia "
            "de Chile para 5° básico con 4 preguntas")
        self._txt.bind("<FocusIn>",  self._limpiar_ph)
        self._txt.bind("<FocusOut>", self._restaurar_ph)
        self._txt.bind("<Button-3>", self._menu_contextual)
        row += 1

        fr_txt = ctk.CTkFrame(left, fg_color="transparent")
        fr_txt.grid(row=row, column=0, sticky="ew", padx=14, pady=(5, 0))
        fr_txt.grid_columnconfigure(1, weight=1)
        row += 1

        ctk.CTkButton(fr_txt, text="Limpiar",
                      fg_color="transparent", hover_color=BG,
                      text_color=MUTED, font=ctk.CTkFont(FONT, 9),
                      width=58, height=24, corner_radius=6,
                      command=self._limpiar_campo).grid(row=0, column=0)

        self._btn_arch = ctk.CTkButton(fr_txt, text="+ Adjuntar texto base",
                                       fg_color="transparent", hover_color=BG,
                                       text_color=ACCENT,
                                       font=ctk.CTkFont(FONT, 9, "bold"),
                                       height=24, corner_radius=6,
                                       command=self._subir_archivo)
        self._btn_arch.grid(row=0, column=1, sticky="e")

        self._lbl_arch = ctk.CTkLabel(left, text="",
                                      text_color=GREEN,
                                      font=ctk.CTkFont(FONT, 8))
        self._lbl_arch.grid(row=row, column=0, sticky="w", padx=14)
        row += 1

        row = self._sep(left, row)

        # ── Opciones ──────────────────────────────────────────────────────────
        row = self._sec(left, "OPCIONES", row)

        opc = ctk.CTkFrame(left, fg_color="transparent")
        opc.grid(row=row, column=0, sticky="ew", padx=14, pady=(6, 0))
        opc.grid_columnconfigure(1, weight=1)
        row += 1

        opciones = [
            ("Curso",               CURSOS,                              "3° Básico"),
            ("Preguntas por ficha", [str(n) for n in range(2, 7)],       "4"),
            ("Cantidad de fichas",  [str(n) for n in range(1, 21)],      "1"),
            ("Tipo de documento",   ["Con encabezado", "Sin encabezado"], "Con encabezado"),
        ]
        self._combos = []
        for i, (lbl, vals, default) in enumerate(opciones):
            ctk.CTkLabel(opc, text=lbl, text_color=MUTED,
                         font=ctk.CTkFont(FONT, 10)).grid(
                             row=i, column=0, sticky="w", pady=5)
            cb = ctk.CTkComboBox(opc, values=vals,
                                 font=ctk.CTkFont(FONT, 10),
                                 fg_color=BG, text_color=TEXT,
                                 border_color=BORDER, border_width=1,
                                 button_color=BORDER,
                                 button_hover_color=MUTED,
                                 dropdown_fg_color=PANEL,
                                 dropdown_text_color=TEXT,
                                 height=34, corner_radius=8,
                                 state="readonly")
            cb.set(default)
            cb.grid(row=i, column=1, sticky="ew", padx=(10, 0), pady=5)
            self._combos.append(cb)

        self._cb_curso, self._cb_preg, self._cb_fich, self._cb_tipo = self._combos

        row = self._sep(left, row)

        # ── Salida ────────────────────────────────────────────────────────────
        row = self._sec(left, "SALIDA WORD", row)

        self._seg_salida = ctk.CTkSegmentedButton(
            left,
            values=["Un solo Word", "Word por ficha"],
            font=ctk.CTkFont(FONT, 10),
            fg_color=BG,
            selected_color=ACCENT,
            selected_hover_color=ACCENT_D,
            unselected_color=BG,
            unselected_hover_color=BORDER,
            text_color=TEXT,
            text_color_disabled=MUTED,
            corner_radius=8,
            height=34,
        )
        self._seg_salida.set("Un solo Word")
        self._seg_salida.grid(row=row, column=0, sticky="ew",
                              padx=14, pady=(6, 0))
        row += 1

        lbl_nota = ctk.CTkLabel(left,
                                text="Word por ficha: máx. 10 · Un solo Word: sin límite",
                                text_color=MUTED,
                                font=ctk.CTkFont(FONT, 8))
        lbl_nota.grid(row=row, column=0, sticky="w", padx=14, pady=(3, 0))
        row += 1

        row = self._sep(left, row)

        # ── Extras ────────────────────────────────────────────────────────────
        row = self._sec(left, "EXTRAS", row)

        self._var_clave   = tk.BooleanVar()
        self._var_version = tk.BooleanVar()

        chk_frame = ctk.CTkFrame(left, fg_color="transparent")
        chk_frame.grid(row=row, column=0, sticky="ew", padx=14, pady=(6, 0))
        row += 1

        for var, texto in [
            (self._var_clave,   "Generar clave de respuestas"),
            (self._var_version, "Generar Versión A y Versión B"),
        ]:
            ctk.CTkCheckBox(chk_frame, text=texto, variable=var,
                            font=ctk.CTkFont(FONT, 10),
                            text_color=TEXT2,
                            fg_color=ACCENT,
                            hover_color=ACCENT_D,
                            border_color=BORDER2,
                            checkmark_color="white",
                            corner_radius=5).pack(anchor="w", pady=4)

        row = self._sep(left, row)

        # ── Botones ───────────────────────────────────────────────────────────
        btns = ctk.CTkFrame(left, fg_color="transparent")
        btns.grid(row=row, column=0, sticky="ew", padx=14, pady=(0, 16))
        btns.grid_columnconfigure((0, 1), weight=1)
        row += 1

        self._btn_gen = ctk.CTkButton(btns, text="Generar",
                                      fg_color=GREEN, hover_color=GREEN_D,
                                      text_color="white",
                                      font=ctk.CTkFont(FONT, 13, "bold"),
                                      height=48, corner_radius=12,
                                      command=self._generar)
        self._btn_gen.grid(row=0, column=0, columnspan=2,
                           sticky="ew", pady=(0, 8))

        self._btn_dl = ctk.CTkButton(btns, text="Descargar Word",
                                     fg_color=ACCENT, hover_color=ACCENT_D,
                                     text_color="white",
                                     font=ctk.CTkFont(FONT, 11, "bold"),
                                     height=40, corner_radius=10,
                                     state="disabled",
                                     command=self._descargar)
        self._btn_dl.grid(row=1, column=0, sticky="ew", padx=(0, 4))

        ctk.CTkButton(btns, text="Nueva conv.",
                      fg_color=PURPLE, hover_color=PURPLE_D,
                      text_color="white",
                      font=ctk.CTkFont(FONT, 11, "bold"),
                      height=40, corner_radius=10,
                      command=self._nueva_conv).grid(
                          row=1, column=1, sticky="ew", padx=(4, 0))

    # ── Helpers de layout ─────────────────────────────────────────────────────
    def _sec(self, parent, texto, row):
        ctk.CTkLabel(parent, text=texto,
                     text_color=MUTED,
                     font=ctk.CTkFont(FONT, 8, "bold")).grid(
                         row=row, column=0, sticky="w",
                         padx=14, pady=(16, 0))
        return row + 1

    def _sep(self, parent, row):
        ctk.CTkFrame(parent, fg_color=BORDER,
                     height=1, corner_radius=0).grid(
                         row=row, column=0, sticky="ew",
                         padx=14, pady=(12, 0))
        return row + 1

    # ── Panel derecho (chat) ──────────────────────────────────────────────────
    def _build_panel_der(self, parent):
        right = ctk.CTkFrame(parent, fg_color=PANEL,
                             corner_radius=14,
                             border_color=BORDER, border_width=1)
        right.grid(row=0, column=1, sticky="nsew")
        right.grid_columnconfigure(0, weight=1)
        right.grid_rowconfigure(1, weight=1)

        hdr = ctk.CTkFrame(right, fg_color=PANEL,
                           corner_radius=0, height=44)
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.grid_propagate(False)
        ctk.CTkLabel(hdr, text="CONVERSACIÓN",
                     text_color=MUTED,
                     font=ctk.CTkFont(FONT, 8, "bold")).pack(
                         side="left", padx=18, pady=12)
        ctk.CTkFrame(right, fg_color=BORDER,
                     height=1, corner_radius=0).grid(
                         row=0, column=0, sticky="sew")

        self._chat_area = ctk.CTkScrollableFrame(
            right, fg_color=CHAT_BG, corner_radius=0,
            scrollbar_button_color=BORDER,
            scrollbar_button_hover_color=MUTED)
        self._chat_area.grid(row=1, column=0, sticky="nsew")
        self._chat_area.grid_columnconfigure(0, weight=1)
        self._chat_area.bind("<Enter>", lambda e: self._set_scroll("chat"))
        self._chat_area.bind("<Leave>", lambda e: self._set_scroll(None))

    # ── Burbujas de chat ──────────────────────────────────────────────────────
    def _on_resize(self, event):
        if event.widget is not self:
            return
        if self._resize_job:
            self.after_cancel(self._resize_job)
        self._resize_job = self.after(120, self._aplicar_resize)

    def _aplicar_resize(self):
        self._resize_job = None
        w = self._chat_area.winfo_width()
        if w > 100:
            self._chat_w = w

    def _burbuja(self, texto: str, rol: str):
        es_user = (rol == "user")
        row     = len(self._chat_area.winfo_children())

        # Burbujas ocupan máximo 72% del ancho del chat
        ancho = max(260, int(self._chat_w * 0.72))

        outer = ctk.CTkFrame(self._chat_area, fg_color="transparent")
        outer.grid(row=row, column=0, sticky="ew", padx=14, pady=5)
        outer.grid_columnconfigure(0, weight=1)

        if es_user:
            bub = ctk.CTkFrame(outer, fg_color=BUB_U, corner_radius=16)
            bub.grid(row=0, column=0, sticky="e", padx=(int(self._chat_w * 0.22), 0))
            ctk.CTkLabel(bub, text=texto,
                         text_color="white",
                         wraplength=ancho, justify="left",
                         font=ctk.CTkFont(FONT, 10),
                         padx=14, pady=10).pack()
        else:
            bub = ctk.CTkFrame(outer, fg_color=BUB_IA,
                               border_color=BORDER, border_width=1,
                               corner_radius=16)
            bub.grid(row=0, column=0, sticky="w", padx=(0, int(self._chat_w * 0.22)))
            ctk.CTkLabel(bub, text=texto,
                         text_color=TEXT2,
                         wraplength=ancho, justify="left",
                         font=ctk.CTkFont(FONT, 10),
                         padx=14, pady=10).pack()

        self.update_idletasks()
        self._chat_area._parent_canvas.yview_moveto(1.0)

    # ── Menú contextual (clic derecho) ───────────────────────────────────────
    def _menu_contextual(self, event):
        menu = tk.Menu(self, tearoff=0,
                       bg=PANEL, fg=TEXT,
                       activebackground=ACCENT, activeforeground="white",
                       font=(FONT, 10),
                       relief="flat", bd=0)
        menu.add_command(label="Cortar",
                         command=lambda: self._txt.event_generate("<<Cut>>"))
        menu.add_command(label="Copiar",
                         command=lambda: self._txt.event_generate("<<Copy>>"))
        menu.add_command(label="Pegar",
                         command=lambda: self._txt.event_generate("<<Paste>>"))
        menu.add_separator()
        menu.add_command(label="Seleccionar todo",
                         command=lambda: self._txt.tag_add("sel", "1.0", "end"))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    # ── Control de scroll ────────────────────────────────────────────────────
    def _set_scroll(self, target: str | None):
        self.unbind_all("<MouseWheel>")
        if target == "chat":
            self.bind_all("<MouseWheel>", self._do_scroll_chat)
        elif target == "left":
            self.bind_all("<MouseWheel>", self._do_scroll_left)

    def _do_scroll_chat(self, event):
        self._chat_area._parent_canvas.yview_scroll(
            int(-1 * (event.delta / 120)) * 3, "units")
        return "break"

    def _do_scroll_left(self, event):
        self._left_canvas.yview_scroll(
            int(-1 * (event.delta / 120)) * 3, "units")
        return "break"

    # ── Placeholder ───────────────────────────────────────────────────────────
    _PH = ("Ej: Prueba de comprensión sobre la Independencia "
           "de Chile para 5° básico con 4 preguntas")

    def _limpiar_campo(self):
        self._txt.delete("1.0", "end")
        self._txt.configure(text_color=TEXT)
        self._ph_activo = False

    def _limpiar_ph(self, _=None):
        if self._ph_activo:
            self._txt.delete("1.0", "end")
            self._txt.configure(text_color=TEXT)
            self._ph_activo = False

    def _restaurar_ph(self, _=None):
        if not self._txt.get("1.0", "end").strip():
            self._txt.configure(text_color=MUTED)
            self._txt.insert("1.0", self._PH)
            self._ph_activo = True

    # ── Adjuntar archivo ──────────────────────────────────────────────────────
    def _subir_archivo(self):
        path = filedialog.askopenfilename(
            title="Seleccionar archivo",
            filetypes=[("Soportados", "*.pdf *.docx *.txt"),
                       ("PDF", "*.pdf"), ("Word", "*.docx"), ("Texto", "*.txt")])
        if not path:
            return
        try:
            self._texto_base     = extraer_texto(path)
            self._archivo_nombre = os.path.basename(path)
            self._lbl_arch.configure(text=f"  {self._archivo_nombre}")
        except Exception as e:
            messagebox.showerror("Error al leer archivo", str(e))

    # ── Generar ───────────────────────────────────────────────────────────────
    def _generar(self):
        if self._generando:
            return
        if not self._gemini:
            messagebox.showerror("Sin API Key", "Configura tu API Key primero.")
            return

        prompt = "" if self._ph_activo else self._txt.get("1.0", "end").strip()
        if not prompt:
            messagebox.showwarning("Pedido vacío",
                                   "Escribe lo que quieres generar.")
            return

        n_fichas    = int(self._cb_fich.get())
        n_preguntas = int(self._cb_preg.get())
        curso       = self._cb_curso.get()

        if self._seg_salida.get() == "Word por ficha" and n_fichas > 10:
            messagebox.showwarning("Límite excedido",
                "En modo «Word por ficha» el máximo son 10 fichas.\n"
                "Reduce la cantidad o cambia a «Un solo Word».")
            return

        self._generando = True
        self._btn_gen.configure(state="disabled", text="Generando…",
                                fg_color=BORDER, text_color=MUTED)
        self._btn_dl.configure(state="disabled")

        meta = f"{n_fichas} ficha(s)  ·  {n_preguntas} preguntas  ·  {curso}"
        if self._archivo_nombre:
            meta += f"  ·  {self._archivo_nombre}"
        self._burbuja(f"{prompt}\n\n{meta}", "user")

        threading.Thread(
            target=self._worker,
            args=(prompt, n_fichas, n_preguntas, curso),
            daemon=True
        ).start()

    def _worker(self, prompt, n_fichas, n_preguntas, curso):
        try:
            fichas = self._gemini.generar(
                prompt=prompt, n_fichas=n_fichas,
                n_preguntas=n_preguntas, curso=curso,
                texto_base=self._texto_base)
            self.after(0, self._on_ok, fichas)
        except Exception as e:
            self.after(0, self._on_err, str(e))

    def _on_ok(self, fichas):
        self._fichas    = fichas
        self._generando = False
        self._btn_gen.configure(state="normal", text="Generar",
                                fg_color=GREEN, text_color="white")
        self._btn_dl.configure(state="normal", fg_color=ACCENT)
        total = sum(len(f.get("preguntas", [])) for f in fichas)
        self._burbuja(
            f"Listo. Generé {len(fichas)} ficha(s) con {total} pregunta(s) en total.\n"
            "Pulsa «Descargar Word» cuando quieras.", "ia")

    def _on_err(self, err):
        self._generando = False
        self._btn_gen.configure(state="normal", text="Generar",
                                fg_color=GREEN, text_color="white")
        self._burbuja(f"Ocurrió un error:\n{err}", "ia")

    # ── Descargar ─────────────────────────────────────────────────────────────
    def _descargar(self):
        if not self._fichas:
            return

        ts     = datetime.now().strftime("%Y%m%d_%H%M%S")
        layout = "doble" if self._cb_tipo.get() == "Con encabezado" else "normal"
        modo   = self._seg_salida.get()
        hechos = []

        try:
            if modo == "Word por ficha":
                fichas_out = self._fichas[:10]  # máximo 10
                folder = filedialog.askdirectory(
                    title="Seleccionar carpeta donde guardar las fichas")
                if not folder:
                    return

                for i, ficha in enumerate(fichas_out, 1):
                    if self._var_version.get():
                        # Versión A
                        path_a = os.path.join(folder,
                                              f"ficha{i}_versionA_{ts}.docx")
                        Generador().generar([ficha], path_a, layout)
                        hechos.append(f"Ficha {i} — Versión A")
                        # Versión B (preguntas mezcladas)
                        fb = copy.deepcopy(ficha)
                        ps = fb.get("preguntas", [])
                        random.shuffle(ps)
                        for j, p in enumerate(ps, 1):
                            p["numero"] = j
                        path_b = os.path.join(folder,
                                              f"ficha{i}_versionB_{ts}.docx")
                        Generador().generar([fb], path_b, layout)
                        hechos.append(f"Ficha {i} — Versión B")
                    else:
                        path = os.path.join(folder, f"ficha{i}_{ts}.docx")
                        Generador().generar([ficha], path, layout)
                        hechos.append(f"Ficha {i}")

                if self._var_clave.get():
                    path_clave = os.path.join(folder, f"clave_{ts}.docx")
                    generar_clave_docx(fichas_out, path_clave)
                    hechos.append("Clave de respuestas")

            else:  # Un solo Word
                out = filedialog.asksaveasfilename(
                    title="Guardar evaluación",
                    initialfile=f"evaluacion_{ts}.docx",
                    defaultextension=".docx",
                    filetypes=[("Word", "*.docx")])
                if not out:
                    return
                base = out[:-5]

                if self._var_version.get():
                    Generador().generar(self._fichas,
                                        f"{base}_versionA.docx", layout)
                    hechos.append("Versión A")
                    fb = copy.deepcopy(self._fichas)
                    for f in fb:
                        ps = f.get("preguntas", [])
                        random.shuffle(ps)
                        for i, p in enumerate(ps, 1):
                            p["numero"] = i
                    Generador().generar(fb, f"{base}_versionB.docx", layout)
                    hechos.append("Versión B")
                else:
                    Generador().generar(self._fichas, out, layout)
                    hechos.append("Evaluación")

                if self._var_clave.get():
                    generar_clave_docx(self._fichas, f"{base}_respuestas.docx")
                    hechos.append("Clave de respuestas")

        except Exception as e:
            messagebox.showerror("Error al guardar", str(e))
            return

        msg = "Guardado:\n" + "\n".join(f"  · {h}" for h in hechos)
        self._burbuja(msg, "ia")
        messagebox.showinfo("Listo", msg)

    # ── Otras acciones ────────────────────────────────────────────────────────
    def _nueva_conv(self):
        if not messagebox.askyesno("Nueva conversación",
                                   "¿Quieres empezar una conversación nueva?\n"
                                   "Se borrará el historial actual."):
            return
        if self._gemini:
            self._gemini.limpiar_historial()
        self._fichas = None
        self._texto_base = self._archivo_nombre = ""
        self._lbl_arch.configure(text="")
        self._btn_dl.configure(state="disabled")
        for w in self._chat_area.winfo_children():
            w.destroy()
        self._burbuja("Conversación nueva. ¿En qué puedo ayudarte?", "ia")

    def _cambiar_api_key(self):
        self._pedir_api_key()


# ════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    App().mainloop()
