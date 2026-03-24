"""
file_parser.py
Extraccion de texto plano desde PDF, Word (.docx) y TXT.
"""

import os


def extraer_texto(path: str) -> str:
    """Devuelve el texto del archivo como string limpio."""
    ext = os.path.splitext(path)[1].lower()

    if ext == ".pdf":
        return _leer_pdf(path)
    elif ext == ".docx":
        return _leer_docx(path)
    elif ext == ".txt":
        return _leer_txt(path)
    else:
        raise ValueError(f"Formato no soportado: {ext}. Usa PDF, Word (.docx) o TXT.")


def _leer_pdf(path: str) -> str:
    try:
        from pypdf import PdfReader
    except ImportError:
        raise ImportError("pypdf no esta instalado. Ejecuta: pip install pypdf")

    reader = PdfReader(path)
    partes = []
    for page in reader.pages:
        texto = page.extract_text()
        if texto:
            partes.append(texto.strip())
    return "\n\n".join(partes)


def _leer_docx(path: str) -> str:
    from docx import Document
    doc = Document(path)
    parrafos = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    return "\n\n".join(parrafos)


def _leer_txt(path: str) -> str:
    for enc in ("utf-8", "latin-1", "cp1252"):
        try:
            with open(path, "r", encoding=enc) as f:
                return f.read().strip()
        except UnicodeDecodeError:
            continue
    raise ValueError("No se pudo leer el archivo TXT. Guardalo en UTF-8.")
