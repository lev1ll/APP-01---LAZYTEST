"""
crear_plantilla.py
Genera plantilla.docx para el Colegio Cristiano Gabriela Mistral.
Ejecutar una sola vez: python crear_plantilla.py
"""

import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ═══════════════════════════════════════════════════════════════
#  CONFIGURACIÓN  —  edita aquí si cambias algo
# ═══════════════════════════════════════════════════════════════

NOMBRE_COLEGIO  = "Colegio Cristiano Gabriela Mistral"
NOMBRE_PROFESOR = "Prof. Nombre Apellido"          # ← cambia esto

# Ruta del logo (busca primero junto al script, luego en el Desktop)
_script_dir = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = None
for _candidate in [
    os.path.join(_script_dir, "logo_gabriela_mistral.png"),
    os.path.join(_script_dir, "logo.png"),
    r"C:\Users\amont\Desktop\logo_gabriela_mistral.png",
]:
    if os.path.isfile(_candidate):
        LOGO_PATH = _candidate
        break

# Ruta de firma (opcional — deja None si no tienes firma escaneada)
FIRMA_PATH = None
for _candidate in [
    os.path.join(_script_dir, "firma.png"),
    os.path.join(_script_dir, "firma.jpg"),
]:
    if os.path.isfile(_candidate):
        FIRMA_PATH = _candidate
        break

# ═══════════════════════════════════════════════════════════════
#  PALETA DE COLORES
# ═══════════════════════════════════════════════════════════════

AZUL_OSC   = RGBColor(0x1F, 0x38, 0x64)   # header principal
AZUL_MED   = RGBColor(0x2E, 0x54, 0x96)   # banda secundaria
DORADO     = RGBColor(0xC9, 0xA8, 0x4C)   # acentos dorados
ROJO_LOGO  = RGBColor(0xC0, 0x00, 0x00)   # rojo del logo (acento)
GRIS_CLARO = RGBColor(0xF2, 0xF2, 0xF2)
BLANCO     = RGBColor(0xFF, 0xFF, 0xFF)

HEX_AZUL_OSC   = "1F3864"
HEX_AZUL_MED   = "2E5496"
HEX_DORADO     = "C9A84C"
HEX_ROJO       = "C00000"
HEX_GRIS       = "F2F2F2"
HEX_GRIS_MED   = "D9E1F2"   # azul muy claro para filas alternas


# ═══════════════════════════════════════════════════════════════
#  HELPERS XML
# ═══════════════════════════════════════════════════════════════

def set_cell_bg(cell, hex_color: str):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    # Eliminar shd existente para evitar duplicados
    for old in tcPr.findall(qn("w:shd")):
        tcPr.remove(old)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    tcPr.append(shd)


def set_cell_borders(cell, color_hex: str, sz: int = 8,
                     sides=("top", "left", "bottom", "right")):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:tcBorders")):
        tcPr.remove(old)
    tcBorders = OxmlElement("w:tcBorders")
    for side in sides:
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"),   "single")
        b.set(qn("w:sz"),    str(sz))
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), color_hex)
        tcBorders.append(b)
    tcPr.append(tcBorders)


def set_col_width(table, col_idx: int, width_cm: float):
    for row in table.rows:
        cell = row.cells[col_idx]
        tc   = cell._tc
        tcPr = tc.get_or_add_tcPr()
        for old in tcPr.findall(qn("w:tcW")):
            tcPr.remove(old)
        tcW = OxmlElement("w:tcW")
        tcW.set(qn("w:w"),    str(int(width_cm * 567)))
        tcW.set(qn("w:type"), "dxa")
        tcPr.append(tcW)


def set_table_no_spacing(table):
    tbl  = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    for old in tblPr.findall(qn("w:tblCellSpacing")):
        tblPr.remove(old)
    cs = OxmlElement("w:tblCellSpacing")
    cs.set(qn("w:w"),    "0")
    cs.set(qn("w:type"), "dxa")
    tblPr.append(cs)


def add_para_border(p, sides=("bottom",), color="C9A84C", sz=12, space=4):
    pPr  = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    for side in sides:
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"),   "single")
        el.set(qn("w:sz"),    str(sz))
        el.set(qn("w:space"), str(space))
        el.set(qn("w:color"), color)
        pBdr.append(el)
    pPr.append(pBdr)


def add_page_number_run(paragraph):
    """Inserta campo PAGE de Word en el párrafo dado."""
    run = paragraph.add_run()
    for tag, text in [("begin", None), ("separate", None), ("end", None)]:
        fc = OxmlElement("w:fldChar")
        fc.set(qn("w:fldCharType"), tag)
        run._r.append(fc)
        if tag == "begin":
            it = OxmlElement("w:instrText")
            it.text = " PAGE "
            run._r.append(it)
    return run


def add_num_pages_run(paragraph):
    run = paragraph.add_run()
    for tag, text in [("begin", None), ("separate", None), ("end", None)]:
        fc = OxmlElement("w:fldChar")
        fc.set(qn("w:fldCharType"), tag)
        run._r.append(fc)
        if tag == "begin":
            it = OxmlElement("w:instrText")
            it.text = " NUMPAGES "
            run._r.append(it)
    return run


# ═══════════════════════════════════════════════════════════════
#  BANDA DECORATIVA SUPERIOR (rojo fino)
# ═══════════════════════════════════════════════════════════════

def build_top_stripe(doc):
    """Línea roja delgada encima del header principal."""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_no_spacing(tbl)
    row = tbl.rows[0]
    row.height = Cm(0.25)
    cell = row.cells[0]
    set_cell_bg(cell, HEX_ROJO)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    return tbl


# ═══════════════════════════════════════════════════════════════
#  HEADER INSTITUCIONAL
# ═══════════════════════════════════════════════════════════════

def build_header(doc):
    """
    Tabla 1 fila x 3 columnas:
      Col 0 (3 cm)   → Logo
      Col 1 (resto)  → Nombre colegio + subtítulo "Evaluación"
      Col 2 (2.5 cm) → Espacio libre / decorativo dorado
    """
    table = doc.add_table(rows=1, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_no_spacing(table)

    row = table.rows[0]
    row.height = Cm(3.0)

    # Fondo azul oscuro en col 0 y 1, dorado en col 2
    for i, cell in enumerate(row.cells):
        set_cell_bg(cell, HEX_AZUL_OSC if i < 2 else HEX_DORADO)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # ── Col 0: Logo ──────────────────────────────────────────────────────
    cell_logo = row.cells[0]
    set_col_width(table, 0, 3.2)
    p_logo = cell_logo.paragraphs[0]
    p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_logo.paragraph_format.space_before = Pt(0)
    p_logo.paragraph_format.space_after  = Pt(0)

    if LOGO_PATH:
        run_logo = p_logo.add_run()
        run_logo.add_picture(LOGO_PATH, height=Cm(2.6))
    else:
        run_logo = p_logo.add_run("LOGO")
        run_logo.font.color.rgb = BLANCO
        run_logo.font.size      = Pt(11)
        run_logo.font.bold      = True

    # ── Col 1: Textos ────────────────────────────────────────────────────
    cell_txt = row.cells[1]

    # Nombre del colegio
    p_name = cell_txt.paragraphs[0]
    p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_name.paragraph_format.space_before = Pt(0)
    p_name.paragraph_format.space_after  = Pt(2)
    run_name = p_name.add_run(NOMBRE_COLEGIO.upper())
    run_name.font.color.rgb = BLANCO
    run_name.font.size      = Pt(13)
    run_name.font.bold      = True

    # Línea decorativa bajo el nombre (usando borde XML)
    add_para_border(p_name, sides=("bottom",), color=HEX_DORADO, sz=6, space=3)

    # Subtítulo: EVALUACIÓN
    p_eval = cell_txt.add_paragraph()
    p_eval.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_eval.paragraph_format.space_before = Pt(4)
    p_eval.paragraph_format.space_after  = Pt(0)
    run_eval = p_eval.add_run("E V A L U A C I Ó N")
    run_eval.font.color.rgb = RGBColor(0xFF, 0xD9, 0x66)  # dorado claro
    run_eval.font.size      = Pt(11)
    run_eval.font.bold      = True

    # ── Col 2: Acento dorado con estrellitas decorativas ─────────────────
    cell_deco = row.cells[2]
    set_col_width(table, 2, 1.5)
    p_deco = cell_deco.paragraphs[0]
    p_deco.alignment = WD_ALIGN_PARAGRAPH.CENTER
    # Pequeño texto decorativo institucional
    run_deco = p_deco.add_run("★\n★\n★")
    run_deco.font.color.rgb = AZUL_OSC
    run_deco.font.size      = Pt(10)
    run_deco.font.bold      = True

    return table


# ═══════════════════════════════════════════════════════════════
#  BANDA AZUL MEDIA (debajo del header)
# ═══════════════════════════════════════════════════════════════

def build_mid_stripe(doc):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_no_spacing(tbl)
    row = tbl.rows[0]
    row.height = Cm(0.35)
    cell = row.cells[0]
    set_cell_bg(cell, HEX_AZUL_MED)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    run = p.add_run("Formando personas íntegras para el mañana")
    run.font.color.rgb = RGBColor(0xBF, 0xD7, 0xFF)
    run.font.size      = Pt(7)
    run.font.italic    = True
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return tbl


# ═══════════════════════════════════════════════════════════════
#  RECUADRO DE DATOS DEL ALUMNO
# ═══════════════════════════════════════════════════════════════

def build_datos_box(doc):
    """
    Fila 0: Nombre ________________________  | Fecha: ________
    Fila 1: Curso: ___________  | Asignatura: _______________
    Fila 2: Puntaje: ___ / ___  | Nota: ____  | Firma: ____
    """
    # Espacio encima
    p_pre = doc.add_paragraph()
    p_pre.paragraph_format.space_before = Pt(5)
    p_pre.paragraph_format.space_after  = Pt(0)

    table = doc.add_table(rows=3, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_no_spacing(table)

    filas = [
        [("Nombre:", True),  ("_" * 35, False), ("Fecha:", True),      ("_" * 12, False)],
        [("Curso:",  True),  ("_" * 15, False), ("Asignatura:", True), ("_" * 20, False)],
        [("Puntaje:", True), ("___ / ___", False), ("Nota:", True),    ("_____", False)],
    ]
    anchos = [2.4, 7.2, 2.6, 4.3]

    for r_idx, fila in enumerate(filas):
        row = table.rows[r_idx]
        row.height = Cm(0.8)
        # Alternar fondos: fila 0 y 2 blanco, fila 1 azul muy claro
        bg = HEX_GRIS if r_idx % 2 == 0 else HEX_GRIS_MED

        for c_idx, (texto, es_label) in enumerate(fila):
            cell = row.cells[c_idx]
            set_cell_bg(cell, bg)
            set_cell_borders(cell, HEX_DORADO, sz=6)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after  = Pt(0)
            p.paragraph_format.left_indent  = Cm(0.2)

            run = p.add_run(texto)
            run.font.size = Pt(10)
            if es_label:
                run.font.bold      = True
                run.font.color.rgb = AZUL_OSC
            else:
                run.font.color.rgb = RGBColor(0x44, 0x44, 0x44)

    for c_idx, w in enumerate(anchos):
        set_col_width(table, c_idx, w)

    return table


# ═══════════════════════════════════════════════════════════════
#  LÍNEA SEPARADORA DORADA
# ═══════════════════════════════════════════════════════════════

def build_separator(doc, color=HEX_DORADO, sz=24):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after  = Pt(4)
    add_para_border(p, sides=("bottom",), color=color, sz=sz, space=1)
    return p


# ═══════════════════════════════════════════════════════════════
#  SECCIÓN DE INSTRUCCIONES GENERALES
# ═══════════════════════════════════════════════════════════════

def build_instrucciones(doc):
    """Pequeño recuadro de instrucciones genéricas."""
    p_tit = doc.add_paragraph()
    p_tit.paragraph_format.space_before = Pt(4)
    p_tit.paragraph_format.space_after  = Pt(1)
    run_tit = p_tit.add_run("📌  Instrucciones generales:")
    run_tit.font.bold      = True
    run_tit.font.color.rgb = AZUL_OSC
    run_tit.font.size      = Pt(9.5)

    instrucciones = [
        "• Lee atentamente cada pregunta antes de responder.",
        "• Escribe con letra clara y ordenada.",
        "• Está prohibido el uso de corrector o lápiz pasta.",
    ]
    for txt in instrucciones:
        p_i = doc.add_paragraph()
        p_i.paragraph_format.left_indent   = Cm(0.5)
        p_i.paragraph_format.space_before  = Pt(0)
        p_i.paragraph_format.space_after   = Pt(0)
        run_i = p_i.add_run(txt)
        run_i.font.size      = Pt(9)
        run_i.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
        run_i.font.italic    = True


# ═══════════════════════════════════════════════════════════════
#  SECCIÓN DE FIRMA (al final)
# ═══════════════════════════════════════════════════════════════

def build_firma_section(doc):
    """Área de firma del docente al final de la prueba."""
    # Separador antes de la firma
    build_separator(doc, color=HEX_DORADO, sz=12)

    table = doc.add_table(rows=1, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_no_spacing(table)

    anchos = [5.5, 5.5, 5.5]

    titulos = [
        "Revisado por:",
        "Firma del docente:",
        "Visto Bueno Dirección:",
    ]

    for c_idx, titulo in enumerate(titulos):
        cell = table.rows[0].cells[c_idx]
        cell.vertical_alignment = WD_ALIGN_VERTICAL.BOTTOM

        p_tit = cell.paragraphs[0]
        p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_tit.paragraph_format.space_before = Pt(0)
        p_tit.paragraph_format.space_after  = Pt(2)
        run_tit = p_tit.add_run(titulo)
        run_tit.font.size      = Pt(9)
        run_tit.font.bold      = True
        run_tit.font.color.rgb = AZUL_OSC

        # Si es la columna de firma y tenemos imagen, la insertamos
        if c_idx == 1 and FIRMA_PATH:
            p_firma = cell.add_paragraph()
            p_firma.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_firma.paragraph_format.space_before = Pt(2)
            p_firma.paragraph_format.space_after  = Pt(2)
            run_firma = p_firma.add_run()
            run_firma.add_picture(FIRMA_PATH, height=Cm(1.8))

        # Línea para firmar
        p_linea = cell.add_paragraph()
        p_linea.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_linea.paragraph_format.space_before = Pt(30 if (c_idx == 1 and FIRMA_PATH) else 35)
        p_linea.paragraph_format.space_after  = Pt(2)
        add_para_border(p_linea, sides=("bottom",), color=HEX_AZUL_OSC, sz=6, space=2)

        # Nombre del profesor en col de firma
        if c_idx == 1:
            p_nombre = cell.add_paragraph()
            p_nombre.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_nombre.paragraph_format.space_before = Pt(2)
            p_nombre.paragraph_format.space_after  = Pt(0)
            run_nombre = p_nombre.add_run(NOMBRE_PROFESOR)
            run_nombre.font.size      = Pt(8.5)
            run_nombre.font.italic    = True
            run_nombre.font.color.rgb = RGBColor(0x44, 0x44, 0x44)

        set_cell_borders(cell, "DDDDDD", sz=4)

    for c_idx, w in enumerate(anchos):
        set_col_width(table, c_idx, w)

    return table


# ═══════════════════════════════════════════════════════════════
#  PIE DE PÁGINA
# ═══════════════════════════════════════════════════════════════

def build_footer(doc):
    section = doc.sections[0]
    footer  = section.footer
    footer.is_linked_to_previous = False

    # Limpiar párrafos existentes
    for p in footer.paragraphs[1:]:
        p._element.getparent().remove(p._element)

    p = footer.paragraphs[0]
    p.clear()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after  = Pt(0)

    # Línea azul encima del pie
    add_para_border(p, sides=("top",), color=HEX_AZUL_MED, sz=8, space=4)

    # Texto izquierda: nombre colegio
    run_col = p.add_run(f"{NOMBRE_COLEGIO}   |   ")
    run_col.font.size      = Pt(8)
    run_col.font.color.rgb = AZUL_OSC

    # Página X de Y
    run_pag = p.add_run("Página ")
    run_pag.font.size      = Pt(8)
    run_pag.font.color.rgb = AZUL_OSC

    rp = add_page_number_run(p)
    rp.font.size      = Pt(8)
    rp.font.color.rgb = AZUL_OSC

    run_de = p.add_run(" de ")
    run_de.font.size      = Pt(8)
    run_de.font.color.rgb = AZUL_OSC

    rn = add_num_pages_run(p)
    rn.font.size      = Pt(8)
    rn.font.color.rgb = AZUL_OSC

    # Texto derecha: "Prohibida su reproducción"
    run_der = p.add_run(f"   |   {NOMBRE_PROFESOR}")
    run_der.font.size      = Pt(8)
    run_der.font.color.rgb = RGBColor(0x88, 0x88, 0x88)


# ═══════════════════════════════════════════════════════════════
#  FUNCIÓN PRINCIPAL
# ═══════════════════════════════════════════════════════════════

def build_template(output_path: str = "plantilla.docx"):
    doc = Document()

    # Márgenes
    section = doc.sections[0]
    section.top_margin    = Cm(1.2)
    section.bottom_margin = Cm(2.0)
    section.left_margin   = Cm(2.5)
    section.right_margin  = Cm(2.0)

    # Fuente base
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    # ── 1. Banda roja superior ───────────────────────────────────────────
    build_top_stripe(doc)

    # ── 2. Header con logo ───────────────────────────────────────────────
    build_header(doc)

    # ── 3. Banda azul media con lema ────────────────────────────────────
    build_mid_stripe(doc)

    # ── 4. Recuadro de datos del alumno ─────────────────────────────────
    build_datos_box(doc)

    # ── 5. Separador dorado grueso ───────────────────────────────────────
    build_separator(doc, color=HEX_DORADO, sz=24)

    # ── 6. Instrucciones generales ───────────────────────────────────────
    build_instrucciones(doc)

    # ── 7. Segundo separador más fino ────────────────────────────────────
    build_separator(doc, color=HEX_AZUL_MED, sz=10)

    # ── 8. Marcador de contenido ─────────────────────────────────────────
    p_marker = doc.add_paragraph("[CONTENIDO]")
    p_marker.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_marker.paragraph_format.space_before = Pt(6)
    p_marker.paragraph_format.space_after  = Pt(6)
    run_m = p_marker.runs[0]
    run_m.font.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)
    run_m.font.size      = Pt(10)
    run_m.font.italic    = True

    # ── 9. Sección de firma ──────────────────────────────────────────────
    build_firma_section(doc)

    # ── 10. Pie de página ────────────────────────────────────────────────
    build_footer(doc)

    doc.save(output_path)

    print(f"\n[OK] Plantilla creada: {output_path}")
    if LOGO_PATH:
        print(f"[OK] Logo incluido desde: {LOGO_PATH}")
    else:
        print("[AVISO] Logo no encontrado. Pon 'logo_gabriela_mistral.png' junto al script.")
    if FIRMA_PATH:
        print(f"[OK] Firma incluida desde: {FIRMA_PATH}")
    else:
        print("[INFO] Sin firma escaneada. Pon 'firma.png' junto al script para incluirla.")
    print(f"\n  Edita NOMBRE_PROFESOR en la línea 24 de este script")
    print(f"  para que aparezca tu nombre en la firma y pie de página.\n")


if __name__ == "__main__":
    build_template("plantilla.docx")
