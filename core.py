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

# ---- Griglia fissa del template line sheet "4 Styles (A)" (pagina A4 verticale) ----
PAGE_W, PAGE_H = 595.0, 842.0
COL_SPLIT = 297.0
ROW_SPLIT = 400.0
HEADER_Y = 62.0
FOOTER_Y = 752.0

# ---- Griglia fissa del template "Landscape (8)" (pagina A4 orizzontale, 4x2) ----
LANDSCAPE8_PAGE_W, LANDSCAPE8_PAGE_H = 842.0, 595.0
LANDSCAPE8_COL_SPLITS = [0.0, 200.5, 396.6, 592.7, 842.0]  # 4 colonne
LANDSCAPE8_ROW_SPLITS = [62.0, 290.0, 550.0]                # 2 righe
LANDSCAPE8_PRICE_ANCHOR = "Colors:"  # il prezzo va sotto questa etichetta

LAYOUTS = {
    "4style": {
        "page_w": PAGE_W, "page_h": PAGE_H,
        "col_splits": [0.0, COL_SPLIT, PAGE_W],
        "row_splits": [HEADER_Y, ROW_SPLIT, FOOTER_Y],
        "price_anchor": "Sizes:",
    },
    "landscape8": {
        "page_w": LANDSCAPE8_PAGE_W, "page_h": LANDSCAPE8_PAGE_H,
        "col_splits": LANDSCAPE8_COL_SPLITS,
        "row_splits": LANDSCAPE8_ROW_SPLITS,
        "price_anchor": LANDSCAPE8_PRICE_ANCHOR,
    },
}


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


def maybe_undouble(word):
    """
    Se 'word' e' scritta in grassetto tramite duplicazione di OGNI carattere
    (es. 'WWhhoolleessaallee:' per 'Wholesale:', stesso trucco usato per 'Sizes:'),
    ritorna la versione "smontata". Altrimenti ritorna la parola invariata
    (per non corrompere parole normali con lettere doppie, es. 'DRESS').
    """
    c = collapse(word)
    if not c:
        return word
    doubled_full = "".join(ch * 2 for ch in c)
    if word == doubled_full:
        return c
    if word == doubled_full[:-1]:
        # l'ultimo carattere (es. ':') a volte non e' raddoppiato
        return c
    return word


def redact_text_in_pdf(pdf_bytes, target_text):
    """
    Cerca 'target_text' (confronto senza distinguere maiuscole/minuscole,
    spazi flessibili) in ogni riga di testo del PDF e la copre con un
    rettangolo bianco, "cancellandola" visivamente. Funziona sia su testo
    normale sia su testo in grassetto (lettere duplicate).

    Ritorna i bytes del nuovo PDF. Se target_text e' vuoto, ritorna
    pdf_bytes invariato.
    """
    target_norm = re.sub(r"\s+", " ", target_text.strip().lower())
    if not target_norm:
        return pdf_bytes

    reader = PdfReader(io.BytesIO(pdf_bytes))
    page_h = float(reader.pages[0].mediabox.height)
    page_w = float(reader.pages[0].mediabox.width)

    redactions_by_page = {}  # {page_index: [(x0, top, x1, bottom), ...]}

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for pi, page in enumerate(pdf.pages):
            words = page.extract_words()
            lines = {}
            for w in words:
                key = round(w["top"], 1)
                lines.setdefault(key, []).append(w)

            for key in sorted(lines.keys()):
                line_words = sorted(lines[key], key=lambda w: w["x0"])
                norm_words = [maybe_undouble(w["text"]) for w in line_words]
                line_text = " ".join(norm_words).lower()
                if target_norm in line_text:
                    x0 = min(w["x0"] for w in line_words)
                    x1 = max(w["x1"] for w in line_words)
                    top = min(w["top"] for w in line_words)
                    bottom = max(w["bottom"] for w in line_words)
                    redactions_by_page.setdefault(pi, []).append((x0, top, x1, bottom))

    if not redactions_by_page:
        return pdf_bytes  # nessuna corrispondenza trovata, nulla da fare

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(page_w, page_h))
    for pi in range(len(reader.pages)):
        for (x0, top, x1, bottom) in redactions_by_page.get(pi, []):
            margin = 1.5
            rect_x = x0 - margin
            rect_y = page_h - (bottom + margin)
            rect_w = (x1 - x0) + 2 * margin
            rect_h = (bottom - top) + 2 * margin
            c.setFillColorRGB(1, 1, 1)
            c.setStrokeColorRGB(1, 1, 1)
            c.rect(rect_x, rect_y, rect_w, rect_h, fill=1, stroke=1)
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


def find_items_in_pdf(pdf_bytes, layout="4style"):
    """
    Analizza il PDF e ritorna una lista di dict, uno per capo trovato:
        {
          "product_id": "DD1864_IV",
          "page": 0,
          "price_x": 212.2,          # posizione dove scrivere il prezzo
          "price_last_bottom": 132.0, # fine del blocco di ancoraggio
          "box": (x0, y0, x1, y1),    # riquadro immagine+testo per i ritagli
          "photo_box": (x0, y0, x1, y1),  # solo la foto grande (per la stella)
        }

    layout: "4style" (griglia 2x2, verticale, prezzo sotto "Sizes:")
            "landscape8" (griglia 4x2, orizzontale, prezzo sotto "Colors:")
    """
    cfg = LAYOUTS[layout]
    col_splits = cfg["col_splits"]
    row_splits = cfg["row_splits"]
    price_anchor = cfg["price_anchor"]

    def find_bounds(value, splits):
        for i in range(len(splits) - 1):
            if splits[i] <= value < splits[i + 1]:
                return splits[i], splits[i + 1]
        # fuori range: aggancia all'estremo piu' vicino
        if value < splits[0]:
            return splits[0], splits[1]
        return splits[-2], splits[-1]

    items = {}
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for pi, page in enumerate(pdf.pages):
            words = page.extract_words()

            def _looks_like_date(text):
                parts = re.split(r"[_/]", text)
                return any(re.match(r"^20\d\d$", p) for p in parts)

            codes = [
                w for w in words
                if re.match(r"^[A-Z0-9]+(?:[_/][A-Z0-9]+)+$", w["text"])
                and re.search(r"\d", w["text"])  # un vero codice contiene sempre un numero
                and not _looks_like_date(w["text"])  # esclude date tipo 15/06/2026
            ]

            anchor_labels = []
            for w in words:
                if collapse(w["text"]) == price_anchor:
                    anchor_labels.append(w)

            lines = {}
            for w in words:
                key = round(w["top"], 1)
                lines.setdefault(key, []).append(w)

            page_images = page.images

            for code_w in codes:
                code = code_w["text"]
                if code in items:
                    continue
                colx = code_w["x0"]
                top = code_w["top"]

                # --- riquadro per il ritaglio immagine+testo (griglia NxM) ---
                bx0, bx1 = find_bounds(colx, col_splits)
                by0, by1 = find_bounds(top, row_splits)

                # --- foto principale (la piu' grande dentro il riquadro,
                #     per distinguerla dalle miniature e dagli swatch colore) ---
                photo_box = (bx0, by0, bx1, by1)  # fallback: l'intero riquadro
                best_area = 0.0
                for im in page_images:
                    ix0, itop, ix1, ibot = im["x0"], im["top"], im["x1"], im["bottom"]
                    if ix0 >= bx0 - 1 and ix1 <= bx1 + 1 and itop >= by0 - 1 and ibot <= by1 + 1:
                        area = (ix1 - ix0) * (ibot - itop)
                        if area > best_area:
                            best_area = area
                            photo_box = (ix0, itop, ix1, ibot)

                # --- posizione dove scrivere il prezzo (sotto l'etichetta di
                # ancoraggio: "Sizes:" per il layout 4style, "Colors:" per
                # landscape8). La finestra di ricerca usa il riquadro
                # riga/colonna appena calcolato (by1), non una distanza
                # fissa: cosi' funziona sia con liste colori corte sia
                # lunghe fino a 10 colori.
                candidates = [
                    s for s in anchor_labels
                    if abs(s["x0"] - colx) < 5
                    and s["top"] > top
                    and s["top"] < by1
                ]
                price_last_bottom = top + 30  # fallback prudente
                if candidates:
                    anchor_w = min(candidates, key=lambda s: s["top"])
                    anchor_top = anchor_w["top"]
                    price_last_bottom = anchor_w["bottom"]
                    for key in sorted(lines.keys()):
                        if key <= anchor_top + 1:
                            continue
                        if key - anchor_top > 12:
                            break
                        line_words = lines[key]
                        near = [w for w in line_words if abs(w["x0"] - colx) < 100]
                        if not near:
                            continue
                        first_word = maybe_undouble(near[0]["text"])
                        if first_word.endswith(":"):
                            break  # e' un nuovo campo (es. "Colors:"), non una continuazione
                        price_last_bottom = max(w["bottom"] for w in near)

                items[code] = {
                    "product_id": code,
                    "page": pi,
                    "price_x": colx,
                    "price_last_bottom": price_last_bottom,
                    "box": (bx0, by0, bx1, by1),
                    "photo_box": photo_box,
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
def _draw_star(c, cx, cy, r_outer, fill_rgb):
    """Disegna una stella a 5 punte piena, centrata in (cx, cy)."""
    import math
    r_inner = r_outer * 0.4
    points = []
    for i in range(10):
        angle = math.pi / 2 + i * math.pi / 5
        r = r_outer if i % 2 == 0 else r_inner
        points.append((cx + r * math.cos(angle), cy + r * math.sin(angle)))

    path = c.beginPath()
    path.moveTo(*points[0])
    for pt in points[1:]:
        path.lineTo(*pt)
    path.close()
    c.setFillColorRGB(*fill_rgb)
    c.setStrokeColorRGB(*fill_rgb)
    c.drawPath(path, fill=1, stroke=1)


def build_priced_pdf(pdf_bytes, items, price_rows, price_fields, highlighted_ids=None, bestseller_ids=None):
    """
    price_fields: lista ordinata di tuple (label, dict_key) da stampare,
        es. [("Retail", "retail_eur"), ("Wholesale", "wholesale_eur")]
    price_rows: { product_id: {"retail_eur": .., "retail_usd": .., ...} }
    highlighted_ids: set/lista di Product ID da evidenziare con un
        rettangolo giallo attorno a foto + testo (opzionale).
    bestseller_ids: set/lista di Product ID da segnare con una stella
        gialla in alto a sinistra (opzionale, indipendente dal rettangolo).
    """
    highlighted_ids = set(highlighted_ids or [])
    bestseller_ids = set(bestseller_ids or [])
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
            if pid in highlighted_ids:
                bx0, by0, bx1, by1 = info["box"]
                margin = 3
                rect_x = bx0 + margin
                rect_y = page_h - (by1 - margin)
                rect_w = (bx1 - bx0) - 2 * margin
                rect_h = (by1 - by0) - 2 * margin
                c.setStrokeColorRGB(0.96, 0.77, 0.09)  # giallo #f5c518
                c.setLineWidth(3)
                c.rect(rect_x, rect_y, rect_w, rect_h, fill=0, stroke=1)

            if pid in bestseller_ids:
                # Stella ancorata all'angolo in alto a destra della FOTO
                # GRANDE (non del riquadro intero, che include anche testo).
                px0, ptop, px1, pbot = info.get("photo_box", info["box"])
                star_r = 11
                star_cx = px1 - star_r - 6
                star_cy = page_h - (ptop + star_r + 6)
                _draw_star(c, star_cx, star_cy, star_r, (0.96, 0.77, 0.09))

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
# Generazione PDF selezione CON prezzi e/o evidenziazioni
# ---------------------------------------------------------------------
def build_selection_pdf_with_prices(pdf_bytes, items, selected_ids, price_rows,
                                     price_fields, highlighted_ids=None,
                                     bestseller_ids=None, title="Selezione capi"):
    """
    Combina le funzionalita': prima inserisce (se richiesti) i prezzi
    scelti, i rettangoli gialli e le stelle best seller nel PDF originale
    (stesso meccanismo di build_priced_pdf), poi ritaglia e ricompone SOLO
    i capi selezionati. Passando price_fields=[] non viene scritto nessun
    prezzo (utile per un PDF gia' completo di prezzi propri, dove serve
    solo selezionare ed eventualmente evidenziare).
    """
    priced_bytes = build_priced_pdf(pdf_bytes, items, price_rows, price_fields, highlighted_ids, bestseller_ids)
    return build_selection_pdf(priced_bytes, items, selected_ids, title=title)
