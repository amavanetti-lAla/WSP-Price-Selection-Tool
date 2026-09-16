"""
core.py

Funzioni condivise dall'app Streamlit:
- lettura dei prezzi dall'Excel Zedonk (Product ID, Description,
  Retail €, Retail US$, Wholesale €, Wholesale US$)
- individuazione dei capi nel PDF (posizione del blocco Sizes e
  riquadro completo immagine+testo per ogni Product ID)
- generazione del PDF finale con i prezzi scelti inseriti sotto "Sizes"
- generazione di un PDF "selezione" con solo alcuni capi
"""

import re
import io

import xlrd
import pdfplumber
import fitz  # PyMuPDF
from PIL import Image
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.colors import black
from reportlab.lib.utils import ImageReader

# ---- Griglia fissa del template line sheet (in punti PDF, pagina A4) ----
PAGE_W, PAGE_H = 595.0, 842.0
COL_SPLIT = 297.0
ROW_SPLIT = 400.0
HEADER_Y = 62.0
FOOTER_Y = 752.0


# ---------------------------------------------------------------------
# Lettura prezzi da Excel
# ---------------------------------------------------------------------
def load_prices_from_xls(file_bytes):
    """
    Ritorna una lista di dict, uno per Product ID:
        {
          "product_id": "DD1864_IV",
          "description": "PEARL MINI DRESS",
          "retail_eur": 448.0 | None,
          "retail_usd": 439.0 | None,
          "wholesale_eur": 300.0 | None,
          "wholesale_usd": 290.0 | None,
        }

    Riconosce automaticamente due formati Excel diversi:

    1) Formato "prefisso testo" (vecchio export Zedonk):
       colonna A contiene "Product ID: <codice> <descrizione colore>",
       con i prezzi in colonne successive sulla stessa riga o su una
       riga immediatamente sotto (es. "WSP (USD EX WORKS HK): US$ ..").

    2) Formato "tabella pulita" (nuovo export): una riga di intestazione
       con celle tipo "Product ID", "Description", "Retail €",
       "Retail US$", "Wholesale €", "Wholesale US$" (l'ordine delle
       colonne puo' variare), seguita da una riga per prodotto.
    """
    wb = xlrd.open_workbook(file_contents=file_bytes)
    sh = wb.sheet_by_index(0)

    def to_num(v):
        if isinstance(v, (int, float)) and v:
            return float(v)
        if isinstance(v, str) and v.strip():
            mm = re.search(r"[\d.,]+", v.replace(",", "."))
            if mm:
                try:
                    return float(mm.group(0))
                except ValueError:
                    return None
        return None

    def norm(s):
        return re.sub(r"[^a-z0-9]", "", str(s).lower())

    # --- Cerca una riga di intestazione "tabella pulita" ---
    header_row_idx = None
    col_map = {}
    for r in range(min(sh.nrows, 20)):
        row_norm = [norm(sh.cell_value(r, c)) for c in range(sh.ncols)]
        if "productid" in row_norm:
            header_row_idx = r
            for c, cell_norm in enumerate(row_norm):
                if cell_norm == "productid":
                    col_map["product_id"] = c
                elif cell_norm == "description":
                    col_map["description"] = c
                elif cell_norm in ("retaileur", "retail"):
                    col_map.setdefault("retail_eur", c)
                elif cell_norm in ("retailusd", "retailus"):
                    col_map["retail_usd"] = c
                elif cell_norm in ("wholesaleeur", "wholesale"):
                    col_map.setdefault("wholesale_eur", c)
                elif cell_norm in ("wholesaleusd", "wholesaleus"):
                    col_map["wholesale_usd"] = c
            break

    records = []

    if header_row_idx is not None and "product_id" in col_map:
        # Formato "tabella pulita": una riga per prodotto
        for r in range(header_row_idx + 1, sh.nrows):
            pid_val = sh.cell_value(r, col_map["product_id"])
            if not (isinstance(pid_val, str) and pid_val.strip()):
                continue
            pid = pid_val.strip()
            if norm(pid) == "productid":
                continue  # riga di intestazione ripetuta

            description = ""
            if "description" in col_map:
                dv = sh.cell_value(r, col_map["description"])
                if isinstance(dv, str):
                    description = dv.strip()

            records.append({
                "product_id": pid,
                "description": description,
                "retail_eur": to_num(sh.cell_value(r, col_map["retail_eur"])) if "retail_eur" in col_map else None,
                "retail_usd": to_num(sh.cell_value(r, col_map["retail_usd"])) if "retail_usd" in col_map else None,
                "wholesale_eur": to_num(sh.cell_value(r, col_map["wholesale_eur"])) if "wholesale_eur" in col_map else None,
                "wholesale_usd": to_num(sh.cell_value(r, col_map["wholesale_usd"])) if "wholesale_usd" in col_map else None,
            })
        return records

    # --- Formato "prefisso testo" (vecchio export) ---
    for r in range(sh.nrows):
        val = sh.cell_value(r, 0)
        if not (isinstance(val, str) and val.strip().startswith("Product ID:")):
            continue
        m = re.match(r"Product ID:\s*(\S+)", val.strip())
        if not m:
            continue
        pid = m.group(1)

        def num(col_idx, _r=r):
            if sh.ncols > col_idx:
                return to_num(sh.cell_value(_r, col_idx))
            return None

        col1_val = sh.cell_value(r, 1) if sh.ncols > 1 else ""
        col1_is_price = isinstance(col1_val, str) and col1_val.strip().startswith("€")

        if col1_is_price:
            # Formato piu' semplice: "Product ID: xxx" | "€188.00"
            description = ""
            retail_eur = None
            wholesale_eur = to_num(col1_val)
            retail_usd = None
            wholesale_usd = None
        else:
            description = col1_val.strip() if isinstance(col1_val, str) else ""
            retail_eur = num(2)
            retail_usd = num(3)
            wholesale_eur = num(4)
            wholesale_usd = num(5)

        records.append({
            "product_id": pid,
            "description": description,
            "retail_eur": retail_eur,
            "retail_usd": retail_usd,
            "wholesale_eur": wholesale_eur,
            "wholesale_usd": wholesale_usd,
        })

    return records


# ---------------------------------------------------------------------
# Individuazione dei capi nel PDF
# ---------------------------------------------------------------------
def collapse(s):
    out = []
    for ch in s:
        if not out or out[-1] != ch:
            out.append(ch)
    return "".join(out)


def find_items_in_pdf(pdf_bytes):
    """
    Analizza il PDF e ritorna una lista di dict, uno per capo trovato:
        {
          "product_id": "DD1864_IV",
          "page": 0,
          "price_x": 212.2,          # posizione dove scrivere il prezzo
          "price_last_bottom": 132.0, # fine del blocco Sizes
          "box": (x0, y0, x1, y1),    # riquadro immagine+testo per i ritagli
        }
    """
    items = {}
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for pi, page in enumerate(pdf.pages):
            words = page.extract_words()
            codes = [
                w for w in words
                if re.match(r"^[A-Z0-9]+[_/][A-Z0-9]+$", w["text"])
                and re.search(r"\d", w["text"])  # un vero codice contiene sempre un numero
            ]

            sizes_labels = []
            for w in words:
                if collapse(w["text"]) == "Sizes:":
                    sizes_labels.append(w)

            lines = {}
            for w in words:
                key = round(w["top"], 1)
                lines.setdefault(key, []).append(w)

            for code_w in codes:
                code = code_w["text"]
                if code in items:
                    continue
                colx = code_w["x0"]
                top = code_w["top"]

                # --- riquadro per il ritaglio immagine+testo (griglia 2x2) ---
                left = colx < COL_SPLIT
                upper = top < ROW_SPLIT
                bx0 = 0.0 if left else COL_SPLIT
                bx1 = COL_SPLIT if left else PAGE_W
                by0 = HEADER_Y if upper else ROW_SPLIT
                by1 = ROW_SPLIT if upper else FOOTER_Y

                # --- posizione dove scrivere il prezzo (sotto Sizes) ---
                # La finestra di ricerca usa il riquadro riga/colonna appena
                # calcolato (by1), non una distanza fissa: cosi' funziona sia
                # con liste colori corte (Catherine Ferraro) sia lunghe fino
                # a 10 colori (Rhea Costa).
                candidates = [
                    s for s in sizes_labels
                    if abs(s["x0"] - colx) < 5
                    and s["top"] > top
                    and s["top"] < by1
                ]
                price_last_bottom = top + 30  # fallback prudente
                if candidates:
                    sizes_w = min(candidates, key=lambda s: s["top"])
                    sizes_top = sizes_w["top"]
                    price_last_bottom = sizes_w["bottom"]
                    for key in sorted(lines.keys()):
                        if key <= sizes_top + 1:
                            continue
                        if key - sizes_top > 12:
                            break
                        line_words = lines[key]
                        near = [w for w in line_words if abs(w["x0"] - colx) < 6]
                        if near and re.match(r"^\d", near[0]["text"]):
                            price_last_bottom = max(
                                w["bottom"] for w in line_words if abs(w["x0"] - colx) < 40
                            )
                            break

                items[code] = {
                    "product_id": code,
                    "page": pi,
                    "price_x": colx,
                    "price_last_bottom": price_last_bottom,
                    "box": (bx0, by0, bx1, by1),
                }
    return items


def crop_item_thumbnail(doc, page_index, box_pts, zoom=1.5):
    """Ritorna i byte PNG di un'anteprima ritagliata (per la UI)."""
    page = doc[page_index]
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat)
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    x0, y0, x1, y1 = box_pts
    crop = img.crop((int(x0 * zoom), int(y0 * zoom), int(x1 * zoom), int(y1 * zoom)))
    buf = io.BytesIO()
    crop.save(buf, format="PNG")
    return buf.getvalue()


# ---------------------------------------------------------------------
# Generazione PDF con prezzi
# ---------------------------------------------------------------------
def build_priced_pdf(pdf_bytes, items, price_rows, price_fields):
    """
    price_fields: lista ordinata di tuple (label, dict_key) da stampare,
        es. [("Retail", "retail_eur"), ("Wholesale", "wholesale_eur")]
    price_rows: { product_id: {"retail_eur": .., "retail_usd": .., ...} }
    """
    reader = PdfReader(io.BytesIO(pdf_bytes))
    page_h = float(reader.pages[0].mediabox.height)
    page_w = float(reader.pages[0].mediabox.width)

    by_page = {}
    for pid, info in items.items():
        by_page.setdefault(info["page"], []).append((pid, info))

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(page_w, page_h))

    for pi in range(len(reader.pages)):
        for pid, info in by_page.get(pi, []):
            x = info["price_x"]
            gap = 9.5
            text_top_from_top = info["price_last_bottom"] + gap
            y = page_h - (text_top_from_top + 7.2)

            row = price_rows.get(pid, {})
            cursor_y = y
            for label, key in price_fields:
                value = row.get(key)
                if value is None:
                    continue
                symbol = "€" if "eur" in key else "US$ "
                text = f"{label}: {symbol}{value:.2f}" if "eur" in key else f"{label}: {symbol}{value:.2f}"
                c.setFont("Helvetica-Bold", 7.2)
                c.setFillColor(black)
                c.drawString(x, cursor_y, text)
                cursor_y -= 8.8
        c.showPage()

    c.save()
    buf.seek(0)

    overlay_reader = PdfReader(buf)
    writer = PdfWriter()
    for i, page in enumerate(reader.pages):
        page.merge_page(overlay_reader.pages[i])
        writer.add_page(page)

    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()


# ---------------------------------------------------------------------
# Generazione PDF selezione (solo alcuni capi)
# ---------------------------------------------------------------------
def build_selection_pdf(pdf_bytes, items, selected_ids, title="Selezione capi"):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(PAGE_W, PAGE_H))

    cell_w = (PAGE_W - 60) / 2
    cell_h = 340
    margin_x = 20
    margin_top = 75
    gap_y = 25
    col_positions = [margin_x, margin_x + cell_w + 20]

    idx_on_page = 0
    for pid in selected_ids:
        info = items.get(pid)
        if not info:
            continue

        if idx_on_page == 0:
            c.setFont("Helvetica-Bold", 14)
            c.drawString(margin_x, PAGE_H - 40, title)

        crop_bytes = crop_item_thumbnail(doc, info["page"], info["box"], zoom=3.0)
        crop = Image.open(io.BytesIO(crop_bytes))

        col = idx_on_page % 2
        row = idx_on_page // 2
        x = col_positions[col]
        y_top = PAGE_H - margin_top - row * (cell_h + gap_y)

        img_w, img_h = crop.size
        aspect = img_h / img_w
        draw_w = cell_w
        draw_h = draw_w * aspect
        if draw_h > cell_h:
            draw_h = cell_h
            draw_w = draw_h / aspect

        c.drawImage(ImageReader(io.BytesIO(crop_bytes)), x, y_top - draw_h, width=draw_w, height=draw_h)

        idx_on_page += 1
        if idx_on_page == 4:
            c.showPage()
            idx_on_page = 0

    if idx_on_page != 0:
        c.showPage()

    c.save()
    buf.seek(0)
    doc.close()
    return buf.getvalue()


# ---------------------------------------------------------------------
# Generazione PDF selezione CON prezzi (solo capi scelti + prezzi scelti)
# ---------------------------------------------------------------------
def build_selection_pdf_with_prices(pdf_bytes, items, selected_ids, price_rows,
                                     price_fields, title="Selezione capi"):
    """
    Combina le due funzionalità: prima inserisce i prezzi scelti nel PDF
    originale (stesso meccanismo di build_priced_pdf), poi ritaglia e
    ricompone SOLO i capi selezionati, foto+testo+prezzo inclusi.
    """
    priced_bytes = build_priced_pdf(pdf_bytes, items, price_rows, price_fields)
    return build_selection_pdf(priced_bytes, items, selected_ids, title=title)
