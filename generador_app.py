"""
generador_app.py  v3.0
Generador de Evaluaciones — AMR
"""

import json, os, sys
from io import BytesIO
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── Rutas de logos ───────────────────────────────────────────────────────
if getattr(sys, 'frozen', False):
    _dir = sys._MEIPASS
else:
    _dir = os.path.dirname(os.path.abspath(__file__))

LOGO_PATH = None
for _c in [os.path.join(_dir, "logo_gabriela_mistral.png"),
           os.path.join(_dir, "logo.png")]:
    if os.path.isfile(_c):
        LOGO_PATH = _c; break

AMR_PATH = None
for _c in [os.path.join(_dir, "amr logo.png"),
           os.path.join(_dir, "amr.png"),
           os.path.join(_dir, "amr.jpg")]:
    if os.path.isfile(_c):
        AMR_PATH = _c; break

PAGE_W = 18.0  # cm útiles (A4, márgenes 1.5 + 1.5)

# ── Formatos JSON para la IA ─────────────────────────────────────────────
FORMATO_DOBLE = """\
Genera las fichas/evaluaciones que necesites usando EXACTAMENTE este formato JSON.
Devuelve SOLO el JSON, sin texto adicional ni bloques de código.
Cada objeto de la lista es una ficha distinta (cada una tendrá su propio encabezado y ocupará su propia página).

[
  {
    "numero": 1,
    "instruccion": "Lee atentamente el siguiente texto:",
    "pasaje": "Texto del pasaje de la primera ficha.",
    "preguntas": [
      {
        "numero": 1,
        "enunciado": "¿Pregunta de comprensión?",
        "alternativas": [
          "A) Primera opción.",
          "B) Segunda opción.",
          "C) Tercera opción.",
          "D) Cuarta opción."
        ]
      }
    ]
  },
  {
    "numero": 2,
    "instruccion": "Lee atentamente el siguiente texto:",
    "pasaje": "Texto del pasaje de la segunda ficha.",
    "preguntas": [
      {
        "numero": 1,
        "enunciado": "¿Pregunta de comprensión?",
        "alternativas": [
          "A) Primera opción.",
          "B) Segunda opción.",
          "C) Tercera opción.",
          "D) Cuarta opción."
        ]
      }
    ]
  }
]
"""

FORMATO_NORMAL = """\
Genera una evaluación usando EXACTAMENTE este formato JSON.
Devuelve SOLO el JSON, sin texto adicional ni bloques de código.
La evaluación ocupa una página completa (puede tener muchas preguntas).

[
  {
    "numero": 1,
    "instruccion": "Lee atentamente el siguiente texto:",
    "pasaje": "Escribe aquí el texto del pasaje entre comillas.",
    "preguntas": [
      {
        "numero": 1,
        "enunciado": "¿Cuál es la pregunta de comprensión?",
        "alternativas": [
          "A) Primera opción.",
          "B) Segunda opción.",
          "C) Tercera opción.",
          "D) Cuarta opción."
        ]
      },
      {
        "numero": 2,
        "enunciado": "¿Segunda pregunta?",
        "alternativas": [
          "A) Primera opción.",
          "B) Segunda opción.",
          "C) Tercera opción.",
          "D) Cuarta opción."
        ]
      }
    ]
  }
]
"""


# ── Helpers XML ──────────────────────────────────────────────────────────

def _tbl_borders(table, sides=(), color="000000", sz=12):
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr"); tbl.insert(0, tblPr)
    for old in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(old)
    bdr = OxmlElement("w:tblBorders")
    for side in ("top","left","bottom","right","insideH","insideV"):
        el = OxmlElement(f"w:{side}")
        if side in sides:
            el.set(qn("w:val"),   "single")
            el.set(qn("w:sz"),    str(sz))
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), color)
        else:
            el.set(qn("w:val"),   "none")
            el.set(qn("w:sz"),    "0")
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), "auto")
        bdr.append(el)
    tblPr.append(bdr)


def _cell_borders(cell, sides=(), color="000000", sz=8):
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:tcBorders")):
        tcPr.remove(old)
    bdr = OxmlElement("w:tcBorders")
    for side in ("top","left","bottom","right","insideH","insideV"):
        el = OxmlElement(f"w:{side}")
        if side in sides:
            el.set(qn("w:val"),   "single")
            el.set(qn("w:sz"),    str(sz))
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), color)
        else:
            el.set(qn("w:val"),   "none")
            el.set(qn("w:sz"),    "0")
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), "auto")
        bdr.append(el)
    tcPr.append(bdr)


def _col_w(table, col_idx, width_cm):
    for row in table.rows:
        tcPr = row.cells[col_idx]._tc.get_or_add_tcPr()
        for old in tcPr.findall(qn("w:tcW")):
            tcPr.remove(old)
        w = OxmlElement("w:tcW")
        w.set(qn("w:w"),    str(int(width_cm * 567)))
        w.set(qn("w:type"), "dxa")
        tcPr.append(w)


def _tbl_w(table, width_cm):
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr"); tbl.insert(0, tblPr)
    for old in tblPr.findall(qn("w:tblW")):
        tblPr.remove(old)
    w = OxmlElement("w:tblW")
    w.set(qn("w:w"),    str(int(width_cm * 567)))
    w.set(qn("w:type"), "dxa")
    tblPr.append(w)


def _cell_margin(cell, top=0, start=80, bottom=0, end=80):
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:tcMar")):
        tcPr.remove(old)
    mar = OxmlElement("w:tcMar")
    for side, val in [("top",top),("start",start),("bottom",bottom),("end",end)]:
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"),    str(val))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tcPr.append(mar)


def _move_into_cell(table, cell):
    """Mueve una tabla al interior de la celda, antes del último <w:p>."""
    tbl_xml = table._tbl
    tbl_xml.getparent().remove(tbl_xml)
    tc = cell._tc
    children = list(tc)
    last_p = max((i for i, ch in enumerate(children)
                  if ch.tag == qn("w:p")), default=None)
    if last_p is None:
        tc.append(tbl_xml)
    else:
        tc.insert(last_p, tbl_xml)


def _p0(p):
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)


def _set_tbl_no_spacing(table):
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr"); tbl.insert(0, tblPr)
    for old in tblPr.findall(qn("w:tblCellSpacing")):
        tblPr.remove(old)
    cs = OxmlElement("w:tblCellSpacing")
    cs.set(qn("w:w"), "0"); cs.set(qn("w:type"), "dxa")
    tblPr.append(cs)


# ════════════════════════════════════════════════════════════════════════
#  Generador
# ════════════════════════════════════════════════════════════════════════

class Generador:

    def generar(self, fichas: list, output_path: str, layout: str = "doble",
                imagen_bytes: bytes | None = None):
        doc = Document()
        sec = doc.sections[0]
        sec.top_margin    = Cm(1.5)
        sec.bottom_margin = Cm(1.2)
        sec.left_margin   = Cm(1.5)
        sec.right_margin  = Cm(1.5)
        ns = doc.styles["Normal"]
        ns.font.name = "Calibri"
        ns.font.size = Pt(10)
        ns.paragraph_format.space_before = Pt(0)
        ns.paragraph_format.space_after  = Pt(0)

        if layout == "normal":
            self._add_pagina_normal(doc, fichas, imagen_bytes)
        else:  # doble → con encabezado, 1 ficha por página
            for i, ficha in enumerate(fichas):
                if i > 0:
                    p = doc.add_paragraph(); _p0(p)
                    p.add_run().add_break(WD_BREAK.PAGE)
                self._add_evaluacion(doc, ficha, imagen_bytes)

        doc.save(output_path)

    def _add_pagina_normal(self, doc, fichas, imagen_bytes=None):
        """Modo normal: una sola cabecera al inicio, todo el contenido fluye continuo."""
        if not fichas:
            return
        w = PAGE_W - 0.4
        numero = fichas[0].get("numero", "")

        # ── CABECERA ÚNICA ────────────────────────────────────────────────
        hdr_outer = doc.add_table(rows=1, cols=1)
        hdr_outer.alignment = WD_TABLE_ALIGNMENT.CENTER
        _tbl_w(hdr_outer, PAGE_W)
        _set_tbl_no_spacing(hdr_outer)
        _tbl_borders(hdr_outer,
                     sides=("top", "left", "right"),
                     color="000000", sz=10)

        hdr_cell = hdr_outer.rows[0].cells[0]
        hdr_cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        _cell_borders(hdr_cell, sides=())
        _cell_margin(hdr_cell, top=60, start=110, bottom=0, end=110)

        hdr = doc.add_table(rows=1, cols=3)
        _tbl_w(hdr, w)
        _set_tbl_no_spacing(hdr)
        _tbl_borders(hdr, sides=())
        hdr.rows[0].height = Cm(1.2)

        c0, c1, c2 = hdr.rows[0].cells
        _col_w(hdr, 0, 2.8); _col_w(hdr, 2, 2.2)
        for c in (c0, c1, c2):
            _cell_borders(c, sides=())
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        p_logo = c0.paragraphs[0]; _p0(p_logo)
        p_logo.alignment = WD_ALIGN_PARAGRAPH.LEFT
        if LOGO_PATH:
            p_logo.add_run().add_picture(LOGO_PATH, height=Cm(1.1))
        else:
            r = p_logo.add_run("Colegio Cristiano")
            r.font.size = Pt(7); r.font.bold = True

        p_tit = c1.paragraphs[0]; _p0(p_tit)
        p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
        titulo = fichas[0].get("titulo", f"FICHA DE COMPRENSIÓN LECTORA N° {numero}")
        r = p_tit.add_run(titulo)
        r.font.bold = True; r.font.size = Pt(12)

        p_amr = c2.paragraphs[0]; _p0(p_amr)
        p_amr.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if AMR_PATH:
            p_amr.add_run().add_picture(AMR_PATH, height=Cm(1.1))
        else:
            r = p_amr.add_run("AMR")
            r.font.bold = True; r.font.size = Pt(18)

        _move_into_cell(hdr, hdr_cell)

        datos = doc.add_table(rows=1, cols=6)
        _tbl_w(datos, w)
        _set_tbl_no_spacing(datos)
        _tbl_borders(datos,
                     sides=("top","left","bottom","right","insideH","insideV"),
                     color="000000", sz=8)
        datos.rows[0].height = Cm(0.65)

        specs = [
            ("Nombre", True,  1.9),
            ("",       False, 7.5),
            ("Curso",  True,  1.4),
            ("",       False, 1.8),
            ("Fecha",  True,  1.4),
            ("",       False, 3.0),
        ]
        for idx, (lbl, bold, cw) in enumerate(specs):
            c = datos.rows[0].cells[idx]
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            _col_w(datos, idx, cw)
            _cell_margin(c, top=0, start=80, bottom=0, end=40)
            p = c.paragraphs[0]; _p0(p)
            if lbl:
                r = p.add_run(lbl)
                r.font.bold = bold; r.font.size = Pt(10)

        _move_into_cell(datos, hdr_cell)

        # ── BLOQUE DE CONTENIDO CONTINUO (una fila por pregunta) ─────────
        content_outer = doc.add_table(rows=0, cols=1)
        content_outer.alignment = WD_TABLE_ALIGNMENT.CENTER
        _tbl_w(content_outer, PAGE_W)
        _set_tbl_no_spacing(content_outer)
        _tbl_borders(content_outer,
                     sides=("top", "left", "bottom", "right"),
                     color="000000", sz=10)

        def _add_content_row(can_split, top_m=0, bot_m=0):
            row = content_outer.add_row()
            trPr = row._tr.get_or_add_trPr()
            cs = OxmlElement("w:cantSplit")
            cs.set(qn("w:val"), "0" if can_split else "1")
            trPr.append(cs)
            c = row.cells[0]
            c.vertical_alignment = WD_ALIGN_VERTICAL.TOP
            _cell_borders(c, sides=())
            _cell_margin(c, top=top_m, start=110, bottom=bot_m, end=110)
            return c

        for i, ficha in enumerate(fichas):
            instruccion = ficha.get("instruccion", "")
            pasaje      = ficha.get("pasaje", "")
            preguntas   = ficha.get("preguntas", [])

            # Instrucción + pasaje en la misma fila (puede dividirse si el pasaje es largo)
            c = _add_content_row(can_split=True, top_m=60 if i == 0 else 120)
            p = c.paragraphs[0]; _p0(p)
            p.paragraph_format.space_before   = Pt(0)
            p.paragraph_format.space_after    = Pt(3)
            p.paragraph_format.keep_with_next = True
            r = p.add_run(instruccion)
            r.font.size = Pt(9.5); r.font.italic = True

            es_poema = '\n' in pasaje
            # Para poemas: fuente y espaciado compacto para evitar orphans
            tam_pasaje  = Pt(8.5) if es_poema else Pt(10)
            esp_despues = Pt(2)   if es_poema else Pt(6)
            esp_preg    = Pt(2)   if es_poema else Pt(5)
            tam_preg    = Pt(9)   if es_poema else Pt(10)

            p = c.add_paragraph(); _p0(p)
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after  = esp_despues
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if es_poema else WD_ALIGN_PARAGRAPH.JUSTIFY
            for j, linea in enumerate(pasaje.split('\n') if es_poema else [pasaje]):
                if j > 0:
                    p.add_run().add_break()
                r = p.add_run(linea)
                r.font.size = tam_pasaje; r.font.bold = True

            # Imagen adjunta (entre pasaje y preguntas)
            if imagen_bytes:
                pi = c.add_paragraph(); _p0(pi)
                pi.paragraph_format.space_before = Pt(6)
                pi.paragraph_format.space_after  = Pt(6)
                pi.alignment = WD_ALIGN_PARAGRAPH.CENTER
                pi.add_run().add_picture(BytesIO(imagen_bytes), width=Cm(14))

            # Banco de palabras (tipo completar)
            banco = ficha.get("banco_palabras", [])
            if banco and any(pr.get("tipo", "seleccion_multiple") == "completar"
                             for pr in preguntas):
                cb = _add_content_row(can_split=False)
                self._render_banco(cb, banco, tam_preg, first_p=cb.paragraphs[0])

            # Una fila por pregunta
            for preg in preguntas:
                c = _add_content_row(can_split=False)
                self._dispatch_preg(c, preg, tam_preg, esp_preg,
                                    esp_preg, es_poema,
                                    first_p=c.paragraphs[0])

        # Espacio final
        c = _add_content_row(can_split=True, bot_m=80)
        p = c.paragraphs[0]; _p0(p)
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after  = Pt(0)

    def _add_evaluacion(self, doc, ficha, imagen_bytes=None):
        numero      = ficha.get("numero", "")
        instruccion = ficha.get("instruccion", "")
        pasaje      = ficha.get("pasaje", "")
        preguntas   = ficha.get("preguntas", [])
        w = PAGE_W - 0.4

        # ── TABLA HEADER (fija — solo aparece en pág 1) ───────────────────
        hdr_outer = doc.add_table(rows=1, cols=1)
        hdr_outer.alignment = WD_TABLE_ALIGNMENT.CENTER
        _tbl_w(hdr_outer, PAGE_W)
        _set_tbl_no_spacing(hdr_outer)
        _tbl_borders(hdr_outer,
                     sides=("top", "left", "right"),   # sin bottom
                     color="000000", sz=10)

        hdr_cell = hdr_outer.rows[0].cells[0]
        hdr_cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        _cell_borders(hdr_cell, sides=())
        _cell_margin(hdr_cell, top=60, start=110, bottom=0, end=110)

        # Logo | Título | AMR
        hdr = doc.add_table(rows=1, cols=3)
        _tbl_w(hdr, w)
        _set_tbl_no_spacing(hdr)
        _tbl_borders(hdr, sides=())
        hdr.rows[0].height = Cm(1.2)

        c0, c1, c2 = hdr.rows[0].cells
        _col_w(hdr, 0, 2.8); _col_w(hdr, 2, 2.2)
        for c in (c0, c1, c2):
            _cell_borders(c, sides=())
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        p_logo = c0.paragraphs[0]; _p0(p_logo)
        p_logo.alignment = WD_ALIGN_PARAGRAPH.LEFT
        if LOGO_PATH:
            p_logo.add_run().add_picture(LOGO_PATH, height=Cm(1.1))
        else:
            r = p_logo.add_run("Colegio Cristiano")
            r.font.size = Pt(7); r.font.bold = True

        p_tit = c1.paragraphs[0]; _p0(p_tit)
        p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p_tit.add_run(f"FICHA DE COMPRENSIÓN LECTORA N° {numero}")
        r.font.bold = True; r.font.size = Pt(12)

        p_amr = c2.paragraphs[0]; _p0(p_amr)
        p_amr.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if AMR_PATH:
            p_amr.add_run().add_picture(AMR_PATH, height=Cm(1.1))
        else:
            r = p_amr.add_run("AMR")
            r.font.bold = True; r.font.size = Pt(18)

        _move_into_cell(hdr, hdr_cell)

        # Fila datos: Nombre | _ | Curso | _ | Fecha | _
        datos = doc.add_table(rows=1, cols=6)
        _tbl_w(datos, w)
        _set_tbl_no_spacing(datos)
        _tbl_borders(datos,
                     sides=("top","left","bottom","right","insideH","insideV"),
                     color="000000", sz=8)
        datos.rows[0].height = Cm(0.65)

        specs = [
            ("Nombre", True,  1.9),
            ("",       False, 7.5),
            ("Curso",  True,  1.4),
            ("",       False, 1.8),
            ("Fecha",  True,  1.4),
            ("",       False, 3.0),
        ]
        for idx, (lbl, bold, cw) in enumerate(specs):
            c = datos.rows[0].cells[idx]
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            _col_w(datos, idx, cw)
            _cell_margin(c, top=0, start=80, bottom=0, end=40)
            p = c.paragraphs[0]; _p0(p)
            if lbl:
                r = p.add_run(lbl)
                r.font.bold = bold; r.font.size = Pt(10)

        _move_into_cell(datos, hdr_cell)

        # ── TABLA CONTENIDO (fluye entre páginas sin repetir header) ──────
        content_outer = doc.add_table(rows=1, cols=1)
        content_outer.alignment = WD_TABLE_ALIGNMENT.CENTER
        _tbl_w(content_outer, PAGE_W)
        _set_tbl_no_spacing(content_outer)
        _tbl_borders(content_outer,
                     sides=("left", "bottom", "right"),  # sin top
                     color="000000", sz=10)

        row0 = content_outer.rows[0]
        trPr = row0._tr.get_or_add_trPr()
        cant_split = OxmlElement("w:cantSplit")
        cant_split.set(qn("w:val"), "0")
        trPr.append(cant_split)

        cell = content_outer.rows[0].cells[0]
        cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        _cell_borders(cell, sides=())
        _cell_margin(cell, top=0, start=110, bottom=80, end=110)

        # ── INSTRUCCIÓN ───────────────────────────────────────────────────
        es_poema    = '\n' in pasaje
        tam_pasaje  = Pt(8.5) if es_poema else Pt(10)
        esp_despues = Pt(2)   if es_poema else Pt(4)
        esp_preg    = Pt(2)   if es_poema else Pt(4)
        tam_preg    = Pt(9)   if es_poema else Pt(10)
        esp_alt     = Pt(0)   if es_poema else Pt(1)

        p = cell.add_paragraph()
        p.paragraph_format.space_before  = Pt(4)
        p.paragraph_format.space_after   = Pt(2)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(instruccion)
        r.font.size = Pt(9.5); r.font.italic = True

        # ── PASAJE ────────────────────────────────────────────────────────
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after  = esp_despues
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT if es_poema else WD_ALIGN_PARAGRAPH.JUSTIFY
        for j, linea in enumerate(pasaje.split('\n') if es_poema else [pasaje]):
            if j > 0:
                p.add_run().add_break()
            r = p.add_run(linea)
            r.font.size = tam_pasaje; r.font.bold = True

        # ── IMAGEN ADJUNTA ────────────────────────────────────────────────
        if imagen_bytes:
            pi = cell.add_paragraph(); _p0(pi)
            pi.paragraph_format.space_before = Pt(6)
            pi.paragraph_format.space_after  = Pt(6)
            pi.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pi.add_run().add_picture(BytesIO(imagen_bytes), width=Cm(14))

        # ── PREGUNTAS ─────────────────────────────────────────────────────
        banco = ficha.get("banco_palabras", [])
        if banco and any(pr.get("tipo", "seleccion_multiple") == "completar"
                         for pr in preguntas):
            self._render_banco(cell, banco, tam_preg)

        for preg in preguntas:
            self._dispatch_preg(cell, preg, tam_preg, esp_preg, esp_alt, es_poema)

        # Espacio final
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after  = Pt(0)


    # ── Helpers de renderizado por tipo ──────────────────────────────────────

    def _render_banco(self, cell, banco, tam_preg, first_p=None):
        """Recuadro de palabras para preguntas de completar."""
        p = first_p if first_p is not None else cell.add_paragraph()
        _p0(p)
        p.paragraph_format.space_before  = Pt(4)
        p.paragraph_format.space_after   = Pt(6)
        p.paragraph_format.left_indent   = Cm(0.2)
        p.paragraph_format.right_indent  = Cm(0.2)
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        for lado in ["top", "left", "bottom", "right"]:
            el = OxmlElement(f"w:{lado}")
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), "6")
            el.set(qn("w:space"), "4")
            el.set(qn("w:color"), "000000")
            pBdr.append(el)
        pPr.append(pBdr)
        r = p.add_run("Palabras:  " + "   ·   ".join(banco))
        r.font.size = Pt(9); r.font.italic = True

    def _render_sm_paras(self, cell, preg, tam_preg, esp_preg, esp_alt,
                          es_poema, first_p=None):
        """Selección múltiple A/B/C/D."""
        p = first_p if first_p is not None else cell.add_paragraph()
        _p0(p)
        p.paragraph_format.space_before   = esp_preg
        p.paragraph_format.space_after    = Pt(0) if es_poema else Pt(1)
        p.paragraph_format.keep_with_next = True
        p.paragraph_format.keep_together  = True
        r = p.add_run(f"{preg.get('numero', '?')}. ")
        r.font.bold = True; r.font.size = tam_preg
        r = p.add_run(preg.get("enunciado", ""))
        r.font.bold = True; r.font.size = tam_preg
        alts = preg.get("alternativas", [])
        for idx, alt in enumerate(alts):
            pa = cell.add_paragraph(); _p0(pa)
            pa.paragraph_format.left_indent    = Cm(0.8)
            pa.paragraph_format.space_before   = Pt(0) if es_poema else Pt(1)
            pa.paragraph_format.space_after    = Pt(0) if es_poema else Pt(1)
            pa.paragraph_format.keep_with_next = (idx < len(alts) - 1)
            pa.paragraph_format.keep_together  = True
            r = pa.add_run(alt); r.font.size = tam_preg

    def _render_vf_paras(self, cell, preg, tam_preg, esp_preg,
                          es_poema, first_p=None):
        """Verdadero / Falso — V [   ]   F [   ] al final del enunciado."""
        p = first_p if first_p is not None else cell.add_paragraph()
        _p0(p)
        p.paragraph_format.space_before  = esp_preg
        p.paragraph_format.space_after   = Pt(0) if es_poema else Pt(3)
        p.paragraph_format.keep_together = True
        r = p.add_run(f"{preg.get('numero', '?')}. ")
        r.font.bold = True; r.font.size = tam_preg
        r = p.add_run(preg.get("enunciado", ""))
        r.font.size = tam_preg
        r = p.add_run("          V [   ]     F [   ]")
        r.font.bold = True; r.font.size = tam_preg

    def _render_completar_paras(self, cell, preg, tam_preg, esp_preg,
                                 es_poema, first_p=None):
        """Completar texto — la línea contiene _____ donde va la palabra."""
        p = first_p if first_p is not None else cell.add_paragraph()
        _p0(p)
        p.paragraph_format.space_before  = esp_preg
        p.paragraph_format.space_after   = Pt(0) if es_poema else Pt(2)
        p.paragraph_format.keep_together = True
        r = p.add_run(f"{preg.get('numero', '?')}. ")
        r.font.bold = True; r.font.size = tam_preg
        r = p.add_run(preg.get("enunciado", ""))
        r.font.size = tam_preg

    def _render_desarrollo_paras(self, cell, preg, tam_preg, esp_preg,
                                  first_p=None):
        """Desarrollo — pregunta abierta con líneas en blanco para escribir."""
        p = first_p if first_p is not None else cell.add_paragraph()
        _p0(p)
        p.paragraph_format.space_before   = esp_preg
        p.paragraph_format.space_after    = Pt(2)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(f"{preg.get('numero', '?')}. ")
        r.font.bold = True; r.font.size = tam_preg
        r = p.add_run(preg.get("enunciado", ""))
        r.font.bold = True; r.font.size = tam_preg
        lineas = max(1, preg.get("lineas") or 3)
        for _ in range(lineas):
            lp = cell.add_paragraph(" "); _p0(lp)
            lp.paragraph_format.space_before = Pt(9)
            lp.paragraph_format.space_after  = Pt(0)
            pPr = lp._p.get_or_add_pPr()
            pBdr = OxmlElement("w:pBdr")
            bot = OxmlElement("w:bottom")
            bot.set(qn("w:val"), "single")
            bot.set(qn("w:sz"), "4")
            bot.set(qn("w:space"), "1")
            bot.set(qn("w:color"), "999999")
            pBdr.append(bot)
            pPr.append(pBdr)

    def _dispatch_preg(self, cell, preg, tam_preg, esp_preg, esp_alt,
                        es_poema, first_p=None):
        """Elige el renderer correcto según tipo."""
        tipo = preg.get("tipo", "seleccion_multiple")
        if tipo == "verdadero_falso":
            self._render_vf_paras(cell, preg, tam_preg, esp_preg,
                                   es_poema, first_p)
        elif tipo == "completar":
            self._render_completar_paras(cell, preg, tam_preg, esp_preg,
                                          es_poema, first_p)
        elif tipo == "desarrollo":
            self._render_desarrollo_paras(cell, preg, tam_preg,
                                           esp_preg, first_p)
        else:  # seleccion_multiple (default)
            self._render_sm_paras(cell, preg, tam_preg, esp_preg, esp_alt,
                                   es_poema, first_p)


# ════════════════════════════════════════════════════════════════════════
#  GUI
# ════════════════════════════════════════════════════════════════════════

class App(tk.Tk):
    BG          = "#F5F7FA"
    HEADER_BG   = "#1F3864"
    BTN_COPY    = "#1565C0"
    BTN_DOBLE   = "#2E7D32"
    BTN_NORMAL  = "#1565C0"
    LOG_BG      = "#1E1E1E"
    LOG_INFO    = "#CCCCCC"
    LOG_OK      = "#4CAF50"
    LOG_WARN    = "#FFC107"
    LOG_ERR     = "#F44336"

    def __init__(self):
        super().__init__()
        self.title("Generador de Evaluaciones — AMR")
        self.resizable(True, True)
        self.configure(bg=self.BG)
        self.minsize(720, 680)

        # Icono de la ventana
        if AMR_PATH:
            try:
                icon = tk.PhotoImage(file=AMR_PATH)
                self.iconphoto(True, icon)
                self._icon = icon  # evitar garbage collection
            except Exception:
                pass

        self._build_ui()
        self._log("Listo.", "info")

    def _build_ui(self):
        self.columnconfigure(0, weight=1)

        # ── Header ───────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=self.HEADER_BG, pady=10)
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.columnconfigure(1, weight=1)

        # Logo AMR en header de la app
        if AMR_PATH:
            try:
                from PIL import Image, ImageTk
                img = Image.open(AMR_PATH).convert("RGBA")
                img.thumbnail((48, 48))
                self._hdr_img = ImageTk.PhotoImage(img)
                tk.Label(hdr, image=self._hdr_img,
                         bg="white").grid(row=0, column=0,
                                          rowspan=2, padx=(14, 8))
            except Exception:
                pass

        tk.Label(hdr, text="GENERADOR DE EVALUACIONES",
                 bg=self.HEADER_BG, fg="white",
                 font=("Segoe UI", 14, "bold")).grid(row=0, column=1, sticky="w")
        tk.Label(hdr,
                 text="Pega el JSON de ChatGPT / Gemini y genera tu prueba en Word",
                 bg=self.HEADER_BG, fg="#9DC3E6",
                 font=("Segoe UI", 9, "italic")).grid(row=1, column=1, sticky="w")

        # ── Botones copiar formato ────────────────────────────────────────
        fr_copy = tk.Frame(self, bg=self.BG, pady=6, padx=10)
        fr_copy.grid(row=1, column=0, sticky="ew")
        tk.Label(fr_copy, text="Copiar formato para IA:",
                 bg=self.BG, fg="#333",
                 font=("Segoe UI", 9)).pack(side="left", padx=(0, 8))
        tk.Button(fr_copy, text="📋  CON ENCABEZADO",
                  bg="#1565C0", fg="white",
                  font=("Segoe UI", 9, "bold"),
                  relief="flat", cursor="hand2", padx=10, pady=5,
                  command=lambda: self._copiar(FORMATO_DOBLE)).pack(side="left", padx=(0, 6))
        tk.Button(fr_copy, text="📋  SIN ENCABEZADO",
                  bg="#6A1B9A", fg="white",
                  font=("Segoe UI", 9, "bold"),
                  relief="flat", cursor="hand2", padx=10, pady=5,
                  command=lambda: self._copiar(FORMATO_NORMAL)).pack(side="left")
        tk.Label(fr_copy,
                 text="  ← pégalo en ChatGPT/Gemini",
                 bg=self.BG, fg="#555",
                 font=("Segoe UI", 9, "italic")).pack(side="left", padx=8)

        # ── Área JSON ─────────────────────────────────────────────────────
        fr_json = tk.LabelFrame(self, text="  JSON de la IA  ",
                                bg=self.BG, fg=self.HEADER_BG,
                                font=("Segoe UI", 9, "bold"),
                                pady=4, padx=8)
        fr_json.grid(row=2, column=0, sticky="nsew", padx=10, pady=4)
        fr_json.columnconfigure(0, weight=1)
        fr_json.rowconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        self._txt = tk.Text(fr_json, font=("Consolas", 10),
                            bg="white", fg="#1E1E1E",
                            insertbackground="#1F3864",
                            relief="flat", bd=1, wrap="none")
        self._txt.grid(row=0, column=0, sticky="nsew")
        sb_y = ttk.Scrollbar(fr_json, orient="vertical",   command=self._txt.yview)
        sb_x = ttk.Scrollbar(fr_json, orient="horizontal", command=self._txt.xview)
        self._txt.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)
        sb_y.grid(row=0, column=1, sticky="ns")
        sb_x.grid(row=1, column=0, sticky="ew")

        # ── Botones generar ───────────────────────────────────────────────
        fr_gen = tk.Frame(self, bg=self.BG, pady=10)
        fr_gen.grid(row=3, column=0)

        tk.Label(fr_gen, text="Tipo de impresión:",
                 bg=self.BG, fg="#333",
                 font=("Segoe UI", 10)).pack(side="left", padx=(0, 10))

        tk.Button(fr_gen,
                  text="📄  CON ENCABEZADO  (1 ficha por página)",
                  bg=self.BTN_DOBLE, fg="white",
                  font=("Segoe UI", 11, "bold"),
                  relief="flat", cursor="hand2", padx=16, pady=9,
                  command=lambda: self._generar("doble")).pack(side="left", padx=6)

        tk.Button(fr_gen,
                  text="📋  SIN ENCABEZADO  (fluye continuo)",
                  bg=self.BTN_NORMAL, fg="white",
                  font=("Segoe UI", 11, "bold"),
                  relief="flat", cursor="hand2", padx=16, pady=9,
                  command=lambda: self._generar("normal")).pack(side="left", padx=6)

        # ── Log ───────────────────────────────────────────────────────────
        fr_log = tk.LabelFrame(self, text="  Estado  ",
                               bg=self.BG, fg=self.HEADER_BG,
                               font=("Segoe UI", 9, "bold"),
                               pady=4, padx=8)
        fr_log.grid(row=4, column=0, sticky="ew", padx=10, pady=(0, 10))
        fr_log.columnconfigure(0, weight=1)

        self._log_txt = tk.Text(fr_log, height=5, font=("Consolas", 9),
                                bg=self.LOG_BG, fg=self.LOG_INFO,
                                state="disabled", relief="flat", wrap="word")
        self._log_txt.grid(row=0, column=0, sticky="ew")
        for tag, color in [("info", self.LOG_INFO), ("ok",   self.LOG_OK),
                            ("warn", self.LOG_WARN), ("err",  self.LOG_ERR)]:
            self._log_txt.tag_config(tag, foreground=color)
        sb_log = ttk.Scrollbar(fr_log, orient="vertical",
                               command=self._log_txt.yview)
        self._log_txt.configure(yscrollcommand=sb_log.set)
        sb_log.grid(row=0, column=1, sticky="ns")

        # ── Footer autoría ────────────────────────────────────────────────
        tk.Label(self, text="© 2026 Levil — Todos los derechos reservados",
                 bg=self.BG, fg="#AAAAAA",
                 font=("Segoe UI", 8, "italic")).grid(row=5, column=0,
                                                      pady=(0, 6))

    # ── Acciones ──────────────────────────────────────────────────────────

    def _copiar(self, formato):
        self.clipboard_clear()
        self.clipboard_append(formato)
        self._log("Formato copiado al portapapeles.", "ok")

    def _generar(self, layout: str):
        raw = self._txt.get("1.0", "end").strip()
        if not raw:
            self._log("El área JSON está vacía.", "err")
            messagebox.showerror("Error", "Pega el JSON antes de generar.")
            return
        try:
            datos = json.loads(raw)
        except json.JSONDecodeError as e:
            self._log(f"JSON inválido: {e}", "err")
            messagebox.showerror("JSON inválido", str(e))
            return
        if not isinstance(datos, list):
            self._log("El JSON debe ser una lista [ ... ].", "err")
            messagebox.showerror("Error", "El JSON debe ser una lista.")
            return

        ts  = datetime.now().strftime("%Y%m%d_%H%M%S")
        out = filedialog.asksaveasfilename(
            title="Guardar evaluación como…",
            initialfile=f"evaluacion_{ts}.docx",
            defaultextension=".docx",
            filetypes=[("Word", "*.docx")],
        )
        if not out:
            self._log("Cancelado.", "warn")
            return

        modo = "con encabezado (1 por página)" if layout == "doble" else "sin encabezado (fluye continuo)"
        self._log(f"Generando en modo {modo}…", "info")
        self.update_idletasks()

        try:
            Generador().generar(datos, out, layout)
            self._log(f"Guardado: {out}", "ok")
            messagebox.showinfo("¡Listo!", f"Evaluación guardada en:\n{out}")
        except Exception as e:
            self._log(f"Error: {e}", "err")
            messagebox.showerror("Error", str(e))

    def _log(self, msg, nivel="info"):
        tags = {"info":"[INFO]","ok":"[ OK ]","warn":"[WARN]","err":"[ERR ]"}
        self._log_txt.configure(state="normal")
        self._log_txt.insert("end", f"{tags.get(nivel,'[INFO]')}  {msg}\n", nivel)
        self._log_txt.see("end")
        self._log_txt.configure(state="disabled")


if __name__ == "__main__":
    App().mainloop()
