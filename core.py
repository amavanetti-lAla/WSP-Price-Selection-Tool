"""
core.py

Funzioni condivise dall'app Streamlit:
- lettura dei prezzi dall'Excel Zedonk (Product ID, Description,
  Retail €, Retail US$, Wholesale €, Wholesale US$)
- individuazione dei capi nel PDF (posizione del blocco Sizes e
  riquadro completo immagine+testo per ogni Product ID)
- generazione del PDF finale con i prezzi scelti inseriti sotto "Sizes"
- generazione di un PDF "selezione" con solo alcuni capi
- estrazione capi + immagini da un PDF "linesheet prezzi" (tipo
  JOOR/Zedonk, con Style Name / Style Number / "W: EUR ... | R: EUR ...")
  e generazione del relativo file Excel (Linesheet Name, Style Number,
  Style Name, Price)
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
LANDSCAPE8_COL_SPLITS = [0.0, 224.1, 420.2, 616.3, 842.0]  # 4 colonne (include lo swatch colore)
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
    ritorna la versione "smontata". Altrimenti ritorna la parola invariata.

    Usa una decodifica a coppie (word[0::2]) invece di un semplice collasso
    dei duplicati consecutivi: cosi' funziona correttamente anche su parole
    che hanno gia' lettere doppie "naturali" una volta raddoppiate (es.
    'DRESS' -> 'DDRREESSSS', 'CAMILLE' -> 'CCAAMMIILLLLEE'), che con un
    semplice collasso perderebbero una lettera.
    """
    n = len(word)
    if n < 2:
        return word

    # Tutta la parola raddoppiata (lunghezza pari): ogni coppia di
    # caratteri adiacenti e' identica.
    if n % 2 == 0 and all(word[i] == word[i + 1] for i in range(0, n, 2)):
        return word[0::2]

    # Raddoppiata tranne l'ultimo carattere (es. ':' spesso non e'
    # raddoppiato a fine parola).
    if n % 2 == 1 and all(word[i] == word[i + 1] for i in range(0, n - 1, 2)):
        return word[0:n - 1:2] + word[-1]

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


def find_item_rows(pdf_bytes, items):
    """
    Per ogni capo gia' individuato da find_items_in_pdf, scompone il suo
    riquadro (box) nelle singole righe di testo che contiene (Nome, Codice,
    prezzo, Sizes, Colors, eventuali righe di continuazione), cosi' da
    poterle mostrare, correggere o eliminare una per una.

    items: il dict ritornato da find_items_in_pdf.

    Ritorna { product_id: [ {"text": "...", "box": (x0, top, x1, bottom)}, ... ] }
    con le righe in ordine dall'alto in basso, testo gia' "smontato" dal
    trucco del grassetto (maybe_undouble).
    """
    rows_by_item = {}
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for pid, info in items.items():
            page = pdf.pages[info["page"]]
            bx0, by0, bx1, by1 = info["box"]

            words = [
                w for w in page.extract_words()
                if w["x0"] >= bx0 - 1 and w["x1"] <= bx1 + 1
                and w["top"] >= by0 - 1 and w["bottom"] <= by1 + 1
            ]

            lines = {}
            for w in words:
                key = round(w["top"], 1)
                lines.setdefault(key, []).append(w)

            rows = []
            for top_key in sorted(lines.keys()):
                line_words = sorted(lines[top_key], key=lambda w: w["x0"])
                text = " ".join(maybe_undouble(w["text"]) for w in line_words)
                if not text.strip():
                    continue
                rows.append({
                    "text": text,
                    "box": (
                        min(w["x0"] for w in line_words),
                        min(w["top"] for w in line_words),
                        max(w["x1"] for w in line_words),
                        max(w["bottom"] for w in line_words),
                    ),
                })
            rows_by_item[pid] = rows
    return rows_by_item


def build_edited_pdf(pdf_bytes, items, row_edits):
    """
    Applica correzioni e cancellazioni a righe di testo gia' individuate da
    find_item_rows, e ritorna i bytes del PDF risultante.

    row_edits: { product_id: { row_index: {"text": "nuovo testo", "deleted": bool} } }
        - "deleted": True -> la riga viene coperta con un rettangolo bianco
          (rimossa), il resto del capo resta intatto.
        - "text": se diverso dal testo originale della riga, la riga viene
          coperta e il nuovo testo viene scritto nella stessa posizione.
        Le righe non presenti in row_edits, o identiche all'originale,
        restano invariate.

    Ricalcola find_item_rows sullo stesso pdf_bytes/items per conoscere box
    e testo originale di ogni riga (deve essere lo stesso PDF su cui sono
    stati generati gli indici di riga mostrati all'utente).
    """
    rows_by_item = find_item_rows(pdf_bytes, items)

    reader = PdfReader(io.BytesIO(pdf_bytes))
    page_h = float(reader.pages[0].mediabox.height)
    page_w = float(reader.pages[0].mediabox.width)

    redactions_by_page = {}  # {page_index: [(x0, top, x1, bottom), ...]}
    texts_by_page = {}       # {page_index: [(x0, bottom, text), ...]}

    for pid, edits in row_edits.items():
        rows = rows_by_item.get(pid)
        info = items.get(pid)
        if not rows or not info:
            continue
        page_idx = info["page"]

        for row_idx, edit in edits.items():
            if row_idx < 0 or row_idx >= len(rows):
                continue
            row = rows[row_idx]
            deleted = bool(edit.get("deleted"))
            new_text = edit.get("text")
            original_text = row["text"]

            if not deleted and (new_text is None or new_text == original_text):
                continue  # nessuna modifica reale su questa riga

            x0, top, x1, bottom = row["box"]
            redactions_by_page.setdefault(page_idx, []).append((x0, top, x1, bottom))
            if not deleted and new_text:
                texts_by_page.setdefault(page_idx, []).append((x0, bottom, new_text))

    if not redactions_by_page:
        return pdf_bytes  # nessuna modifica richiesta

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

        for (x0, bottom, text) in texts_by_page.get(pi, []):
            y = page_h - bottom + 1.5
            c.setFont("Helvetica", 7.2)
            c.setFillColor(black)
            c.drawString(x0, y, text)

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
# PDF "linesheet prezzi" (JOOR/Zedonk) -> Excel
# ---------------------------------------------------------------------
# Questi PDF hanno spesso più capi affiancati sulla stessa riga di testo
# (es. 2, 3 o 4 colonne per pagina) e usano lo stesso trucco del
# "grassetto = ogni lettera raddoppiata" già gestito da maybe_undouble()
# per Style Name e per le etichette "W:"/"R:". Per questo l'estrazione
# lavora sulle posizioni (x, top) delle singole parole, non sul testo
# piatto: individua i codici stile, ricostruisce colonne dinamiche in
# base alla loro posizione orizzontale, poi recupera nome (riga sopra)
# e prezzo (riga sotto, dopo l'etichetta "W:") nella stessa colonna.
_CODE_RE = re.compile(r"^[0-9]{4,6}[A-Z]{0,2}$")
_PRICE_W_RE = re.compile(r"W:\s*EUR\s*([\d.,]+)")


def extract_pricing_linesheet_items(pdf_bytes):
    """
    Estrae da un PDF "linesheet prezzi" (Style Name / Style Number / riga
    "W: EUR ... | R: EUR ...") l'elenco dei capi con nome, codice, prezzo
    wholesale (W) e immagine principale del capo.

    Ritorna una lista di dict, nell'ordine in cui i capi compaiono nel PDF:
        {
          "style_number": "26035D",
          "style_name": "LORENE LONG DRESS",
          "price": 1473.00,
          "image_bytes": b"...png...",
        }

    Se in una pagina il conteggio testo/immagini non coincide, i capi in
    eccesso vengono scartati per non rischiare abbinamenti sbagliati tra
    testo e foto.
    """
    results = []
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for pi, page in enumerate(pdf.pages):
                words = page.extract_words()
                if not words:
                    continue

                # Righe raggruppate per "top" arrotondato (come altrove in
                # questo modulo).
                lines = {}
                for w in words:
                    key = round(w["top"], 1)
                    lines.setdefault(key, []).append(w)
                sorted_tops = sorted(lines.keys())

                code_words = [w for w in words if _CODE_RE.match(w["text"])]
                if not code_words:
                    continue

                # Colonne dinamiche: un confine a meta' strada tra ogni
                # coppia di codici adiacenti (per pagine con 1-4+ capi
                # affiancati).
                anchor_xs = sorted(set(round(w["x0"], 1) for w in code_words))
                bounds = [0.0]
                for i in range(len(anchor_xs) - 1):
                    bounds.append((anchor_xs[i] + anchor_xs[i + 1]) / 2)
                bounds.append(float(page.width))

                def bucket(x, bounds=bounds):
                    for i in range(len(bounds) - 1):
                        if bounds[i] <= x < bounds[i + 1]:
                            return i
                    return len(bounds) - 2

                page_text_items = []  # (top, x0, code, style_name, price)
                for code_w in code_words:
                    code = code_w["text"]
                    code_top = code_w["top"]
                    code_x0 = code_w["x0"]
                    col = bucket(code_x0)

                    # Nome: parole nella/e riga/e immediatamente sopra il
                    # codice (entro ~20pt), stessa colonna.
                    name_parts = []
                    for t in sorted_tops:
                        if t >= code_top or code_top - t > 20:
                            continue
                        for w in lines[t]:
                            if bucket(w["x0"]) == col:
                                name_parts.append((w["x0"], maybe_undouble(w["text"])))
                    name_parts.sort(key=lambda p: p[0])
                    style_name = " ".join(p[1] for p in name_parts).strip()

                    # Prezzo: nella/e riga/e sotto il codice (entro ~40pt),
                    # stessa colonna, riga che contiene "W: EUR ...".
                    price_val = None
                    for t in sorted_tops:
                        if t <= code_top:
                            continue
                        if t - code_top > 40:
                            break
                        col_words = sorted(
                            (w for w in lines[t] if bucket(w["x0"]) == col),
                            key=lambda w: w["x0"],
                        )
                        if not col_words:
                            continue
                        line_text = " ".join(maybe_undouble(w["text"]) for w in col_words)
                        pm = _PRICE_W_RE.search(line_text)
                        if pm:
                            try:
                                price_val = float(pm.group(1).replace(",", ""))
                            except ValueError:
                                price_val = None
                            break

                    if not style_name or price_val is None:
                        continue

                    page_text_items.append({
                        "top": code_top,
                        "x0": code_x0,
                        "style_number": code,
                        "style_name": style_name,
                        "price": price_val,
                    })

                if not page_text_items:
                    continue

                # Ordine di lettura: dall'alto in basso, poi da sinistra a
                # destra.
                page_text_items.sort(key=lambda it: (round(it["top"]), it["x0"]))

                fitz_page = doc[pi]

                # Tutte le immagini della pagina, con posizione (rect) e xref.
                img_infos = []
                for img in fitz_page.get_images(full=True):
                    xref = img[0]
                    for rect in fitz_page.get_image_rects(xref):
                        img_infos.append((rect, xref))

                if not img_infos:
                    continue

                # Le "foto principali" sono nettamente piu' grandi degli
                # swatch colore e del logo/firma in alto pagina: si tengono
                # solo le immagini con area >= 50% dell'immagine piu' grande
                # della pagina.
                max_area = max(r.width * r.height for r, _ in img_infos)
                main_imgs = [
                    (r, xref) for r, xref in img_infos
                    if r.width * r.height >= max_area * 0.5
                ]
                main_imgs.sort(key=lambda t: (round(t[0].y0), t[0].x0))

                if len(main_imgs) != len(page_text_items):
                    # Conteggio diverso: per sicurezza abbina solo i primi N
                    # in comune, invece di rischiare un abbinamento sbagliato.
                    n = min(len(main_imgs), len(page_text_items))
                    main_imgs = main_imgs[:n]
                    page_text_items = page_text_items[:n]

                for text_item, (rect, xref) in zip(page_text_items, main_imgs):
                    pix = fitz.Pixmap(doc, xref)
                    if pix.n - pix.alpha >= 4:
                        pix = fitz.Pixmap(fitz.csRGB, pix)
                    image_bytes = pix.tobytes("png")

                    results.append({
                        "style_number": text_item["style_number"],
                        "style_name": text_item["style_name"],
                        "price": text_item["price"],
                        "image_bytes": image_bytes,
                    })
    finally:
        doc.close()

    return results


def build_linesheet_xlsx(items):
    """
    Costruisce un file Excel (bytes) con la stessa struttura del template
    linesheet di riferimento:

        Colonna A: Linesheet Name (immagine del capo)
        Colonna B: Style Number
        Colonna C: Style Name
        Colonna D: Silhouette (vuota)
        Colonna E: Materials (vuota)
        Colonna F: Price (prezzo W estratto dal PDF)

    Le colonne G:O (e oltre, fino a XFD) restano vuote e vengono nascoste,
    come nel file di riferimento.

    items: lista di dict come ritornati da extract_pricing_linesheet_items,
        cioe' con le chiavi "style_number", "style_name", "price",
        "image_bytes".
    """
    import openpyxl
    from openpyxl.styles import Font
    from openpyxl.drawing.image import Image as XLImage
    from PIL import Image as PILImage

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    header_font = Font(name="Calibri", size=11, bold=True)
    headers = ["Linesheet Name", "Style Number", "Style Name", "Silhouette", "Materials", "Price "]
    for col_idx, htext in enumerate(headers, start=1):
        ws.cell(row=3, column=col_idx, value=htext).font = header_font

    col_widths = {"A": 25.5, "B": 14.33, "C": 12.16, "D": 11.0, "E": 18.66, "F": 18.66}
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    ws.row_dimensions[2].height = 17
    ws.row_dimensions[3].height = 14.25

    price_format = '_([$€-2]\\ * #,##0.00_);_([$€-2]\\ * \\(#,##0.00\\);_([$€-2]\\ * "-"??_);_(@_)'
    normal_font = Font(name="Calibri", size=11)

    start_row = 4
    for i, item in enumerate(items):
        row = start_row + i
        ws.row_dimensions[row].height = 145

        ws.cell(row=row, column=2, value=item["style_number"]).font = normal_font
        ws.cell(row=row, column=3, value=item["style_name"]).font = normal_font
        price_cell = ws.cell(row=row, column=6, value=item["price"])
        price_cell.font = normal_font
        price_cell.number_format = price_format

        pil_img = PILImage.open(io.BytesIO(item["image_bytes"]))
        target_height_px = 190  # ~ altezza riga 145pt
        ratio = target_height_px / pil_img.height
        target_width_px = int(pil_img.width * ratio)

        xl_img = XLImage(pil_img)
        xl_img.height = target_height_px
        xl_img.width = target_width_px
        ws.add_image(xl_img, f"A{row}")

    # Nasconde le colonne dalla Q in poi, come richiesto per il template.
    ws.column_dimensions.group("Q", "XFD", hidden=True)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()
