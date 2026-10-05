"""
layout_variabile.py

Supporto ai line sheet con layout "variabile": pagina A4 orizzontale con da 1
a 4 capi affiancati (foto grande, nome, codice stile, righe P:/Sizes:/Colors:,
miniature colore). Vale per qualsiasi brand che usi questo formato.

Non c'e' una griglia fissa: ogni capo e' ancorato al proprio codice stile e il
riquadro si ricava dalla posizione x del codice. Le misure sono tarate su
842x595 pt e scalate in base alla dimensione reale della pagina.
"""

import io
import re

import fitz  # PyMuPDF
import pdfplumber
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

# Codice stile: 4-6 cifre + 0-2 lettere (es. 26035D, 1234, 123456AB).
CODE_RE = re.compile(r"^\d{4,6}[A-Z]{0,2}$")
_YEAR_RE = re.compile(r"^(19|20)\d\d$")

# Geometria misurata su pagina 842x595 pt
REF_W, REF_H = 842.0, 595.0
COL_LEFT = 10.0       # box x0 = x codice - COL_LEFT
COL_RIGHT = 182.0     # box x1 = x codice + COL_RIGHT
PHOTO_LEFT = 9.0
PHOTO_RIGHT = 180.0
BOX_TOP = 78.0
PHOTO_TOP = 82.0
PHOTO_BOTTOM = 390.0
SWATCH_TOP = 495.0
BOX_BOTTOM = 535.0


def _is_code(text):
    return bool(CODE_RE.match(text)) and not _YEAR_RE.match(text)


def find_items_variabile(pdf_bytes):
    """Stessa struttura di core.find_items_in_pdf."""
    items = {}
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for pi, page in enumerate(pdf.pages):
            sx, sy = float(page.width) / REF_W, float(page.height) / REF_H
            words = page.extract_words()
            codes = sorted((w for w in words if _is_code(w["text"])), key=lambda w: w["x0"])
            # un vero codice sta sotto una riga di nome e sopra "P:"/"Sizes:":
            # scarto numeri isolati nella fascia alta/bassa (titoli, pagina)
            codes = [w for w in codes if BOX_BOTTOM * 0.6 * sy < w["top"] < SWATCH_TOP * sy]

            def has_label_below(cw):
                # sotto un vero codice c'e' una riga con etichetta ("P:", "Sizes:")
                return any(
                    w["text"].endswith(":") and abs(w["x0"] - cw["x0"]) < 5 * sx
                    and 0 < w["top"] - cw["top"] < 40 * sy
                    for w in words
                )
            codes = [w for w in codes if has_label_below(w)]

            for idx, code_w in enumerate(codes):
                code = code_w["text"]
                if code in items:
                    continue
                colx = code_w["x0"]
                bx0 = colx - COL_LEFT * sx
                bx1 = colx + COL_RIGHT * sx
                if idx + 1 < len(codes):  # non invadere la colonna successiva
                    bx1 = min(bx1, codes[idx + 1]["x0"] - COL_LEFT * sx)

                text_bottom = code_w["bottom"]
                for w in words:
                    if (bx0 <= w["x0"] < bx1 and w["top"] > code_w["top"]
                            and w["bottom"] < SWATCH_TOP * sy):
                        text_bottom = max(text_bottom, w["bottom"])

                items[code] = {
                    "product_id": code,
                    "page": pi,
                    "price_x": colx,
                    "price_last_bottom": text_bottom,
                    "box": (bx0, BOX_TOP * sy, bx1, BOX_BOTTOM * sy),
                    "photo_box": (colx - PHOTO_LEFT * sx, PHOTO_TOP * sy,
                                  min(colx + PHOTO_RIGHT * sx, bx1), PHOTO_BOTTOM * sy),
                    "layout": "variabile",
                }
    return items


def build_selection_pdf_variabile(pdf_bytes, items, selected_ids, title="Selezione capi"):
    """Selezione in A4 orizzontale: 4 capi per pagina, in fila."""
    page_w, page_h = 842.0, 595.0
    margin_x, gap = 20.0, 14.0
    per_page = 4
    cell_w = (page_w - 2 * margin_x - (per_page - 1) * gap) / per_page
    top_y = page_h - 55
    cell_h = top_y - 20

    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(page_w, page_h))
    zoom = 3.0
    page_cache = {}
    n = 0
    for pid in selected_ids:
        info = items.get(pid)
        if not info:
            continue
        if n % per_page == 0:
            if n:
                c.showPage()
            c.setFont("Helvetica-Bold", 14)
            c.drawString(margin_x, page_h - 35, title)
        pno = info["page"]
        if pno not in page_cache:
            pix = doc[pno].get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            page_cache[pno] = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        x0, y0, x1, y1 = info["box"]
        crop = page_cache[pno].crop((int(x0 * zoom), int(y0 * zoom), int(x1 * zoom), int(y1 * zoom)))
        b = io.BytesIO()
        crop.save(b, format="PNG")
        b.seek(0)

        aspect = crop.height / crop.width
        draw_w, draw_h = cell_w, cell_w * aspect
        if draw_h > cell_h:
            draw_h, draw_w = cell_h, cell_h / aspect
        x = margin_x + (n % per_page) * (cell_w + gap)
        c.drawImage(ImageReader(b), x, top_y - draw_h, width=draw_w, height=draw_h)
        n += 1
    if n:
        c.showPage()
    c.save()
    doc.close()
    return buf.getvalue()
