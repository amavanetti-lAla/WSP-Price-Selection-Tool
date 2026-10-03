"""
linesheet_export.py

Scheda "PDF -> Excel Linesheet": legge uno o piu' PDF (qualunque brand,
stagione o impaginazione, tipicamente scaricati da JOOR) e produce SEMPRE
lo stesso Excel, identico al file di riferimento
"Rhea_Costa_SS27_Linesheet.xlsx":

    A  Linesheet Name (foto)     G  25% OFF  (=ROUNDUP(F*75%,0))
    B  Style Number              H  Finance  (=ROUND(G*H$3,0))
    C  Style Name                I  Duty     (=ROUND(G*I$3,0))
    D  Silhouette (vuota)        J  Ware/ship(=ROUND(G*J$3,0))
    E  Materials  (vuota)        K  Cost     (=ROUNDUP(G+H+I+J,0))
    F  Price                     L  Mark up  (=M-K)
                                 M  Price EURO (da compilare)
                                 N  Price DOLLAR (=M*$M$1)
                                 O  Margin   (=L/M)

Il file e' costruito PARTENDO da linesheet_template.xlsx (stesso formato,
colori, larghezze, intestazioni e parametri di riga 2-3).

ESTRAZIONE (indipendente dal layout): si parte dalle FOTO principali del
capo, non da una griglia fissa. Per ogni foto si cerca il testo
associato, sotto la foto oppure accanto (a destra):
    nome  -> righe sopra il codice
    codice-> riga con un solo "token" tipo 26035D, DD1933_PI_PS27, 26039D_FW25
    prezzo-> "W: EUR 185.00", "P: EUR 1,473.00", "R: EUR ..." (se presente)
Funziona quindi con 1, 2, 3, 4, 8 capi per pagina, con testo sotto o a
fianco della foto, con o senza prezzi.

Richiede linesheet_template.xlsx nella stessa cartella di questo file.
"""

import base64
import io
import os
import re
from copy import copy

import fitz  # PyMuPDF
from PIL import Image

TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "linesheet_template.xlsx")

# Dimensione foto nel file di riferimento (px visualizzati) e risoluzione di salvataggio
IMG_DISPLAY_W, IMG_DISPLAY_H = 139, 190
IMG_STORE_MAX = (330, 450)
ROW_HEIGHT = 144.95


# ---------------------------------------------------------------------
# Utilita' di testo
# ---------------------------------------------------------------------
def _to_float(s):
    """'1,473.00' / '1.473,00' / '1473' / '448,50' -> float."""
    s = s.strip().rstrip(".,")
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".") if re.search(r",\d{1,2}$", s) else s.replace(",", "")
    elif s.count(".") > 1 or re.search(r"\.\d{3}$", s):
        s = s.replace(".", "")
    return float(s)


def _undouble(word):
    """Alcuni export scrivono il grassetto raddoppiando ogni lettera (DDRREESSSS)."""
    n = len(word)
    if n >= 4 and n % 2 == 0 and all(word[i] == word[i + 1] for i in range(0, n, 2)):
        return word[0::2]
    return word


_PRICE_RE = re.compile(
    r"\b(W|P|R|WHOLESALE|RETAIL|PRICE|WSP|RRP)\b\s*:?\s*(?:EUR|€)\s*([\d][\d.,]*)",
    re.IGNORECASE,
)
_PRICE_KIND = {"W": "W", "WHOLESALE": "W", "WSP": "W",
               "R": "R", "RETAIL": "R", "RRP": "R",
               "P": "P", "PRICE": "P"}
_DATE_RE = re.compile(r"^\d{1,2}[/.\-]\d{1,2}[/.\-]\d{2,4}$")
_CODE_RE = re.compile(r"^[A-Z0-9][A-Z0-9_/\-.]{3,}$")


def _is_code(token):
    """Codice stile: un solo token maiuscolo/cifre con almeno una cifra (no date, no numeri puri)."""
    t = token.strip()
    if not _CODE_RE.match(t) or not re.search(r"\d", t):
        return False
    if _DATE_RE.match(t):
        return False
    if re.fullmatch(r"[\d.\-/]+", t) and not re.fullmatch(r"\d{4,}", t):
        return False
    return True


def _parse_prices(text):
    """'W: EUR 185.00 | R: EUR 300' -> {'W': 185.0, 'R': 300.0}"""
    found = {}
    for m in _PRICE_RE.finditer(text):
        kind = _PRICE_KIND[m.group(1).upper()]
        try:
            found.setdefault(kind, _to_float(m.group(2)))
        except ValueError:
            pass
    return found


def _lines(words):
    """Raggruppa le parole (tuple PyMuPDF) in righe, dall'alto in basso."""
    ws = sorted(words, key=lambda w: ((w[1] + w[3]) / 2, w[0]))
    lines = []
    for w in ws:
        yc = (w[1] + w[3]) / 2
        if lines and abs(lines[-1]["y"] - yc) < 3:
            lines[-1]["w"].append(w)
        else:
            lines.append({"y": yc, "w": [w]})
    for ln in lines:
        ln["w"].sort(key=lambda w: w[0])
        ln["tokens"] = [_undouble(w[4]) for w in ln["w"]]
        ln["text"] = " ".join(ln["tokens"])
    return lines


def _prep_image(img_bytes):
    """Foto in JPEG leggero, max 330x450 px (come nel file di riferimento)."""
    img = Image.open(io.BytesIO(img_bytes))
    if img.mode in ("RGBA", "LA", "P"):
        rgba = img.convert("RGBA")
        bg = Image.new("RGB", rgba.size, (255, 255, 255))
        bg.paste(rgba, mask=rgba.split()[-1])
        img = bg
    elif img.mode != "RGB":
        img = img.convert("RGB")
    img.thumbnail(IMG_STORE_MAX, Image.LANCZOS)
    out = io.BytesIO()
    img.save(out, "JPEG", quality=85, optimize=True)
    return out.getvalue()


# ---------------------------------------------------------------------
# Estrazione: parte dalle foto principali di ogni pagina
# ---------------------------------------------------------------------
def _main_photos(page):
    """Rettangoli delle foto principali, in ordine di lettura (righe, poi sinistra->destra)."""
    page_area = page.rect.width * page.rect.height
    rects = []
    for info in page.get_image_info():
        r = fitz.Rect(info["bbox"])
        if r.width < 20 or r.height < 20:
            continue
        if r.width * r.height > 0.6 * page_area:  # sfondi a tutta pagina
            continue
        if any(abs(r.x0 - q.x0) < 1 and abs(r.y0 - q.y0) < 1 for q in rects):
            continue
        rects.append(r)
    if not rects:
        return []
    max_area = max(r.width * r.height for r in rects)
    main = [r for r in rects if r.width * r.height >= 0.5 * max_area and r.height / r.width >= 0.6]
    main.sort(key=lambda r: (r.y0, r.x0))

    rows, cur = [], []
    for r in main:  # raggruppa per riga (le foto della stessa riga hanno y0 simile)
        if cur and abs(r.y0 - cur[0].y0) > 40:
            rows.append(cur)
            cur = []
        cur.append(r)
    if cur:
        rows.append(cur)
    return [r for row in rows for r in sorted(row, key=lambda r: r.x0)]


def _overlap_v(a, b):
    ov = min(a.y1, b.y1) - max(a.y0, b.y0)
    return ov > 0.5 * min(a.height, b.height)


def _regions(ph, photos, page):
    """Due zone di ricerca del testo: SOTTO la foto e A DESTRA della foto."""
    same_row = [q for q in photos if q is not ph and _overlap_v(ph, q)]
    left = [q for q in same_row if q.x0 < ph.x0]
    right = [q for q in same_row if q.x0 > ph.x0]
    cx = (ph.x0 + ph.x1) / 2
    xl = (cx + max(((q.x0 + q.x1) / 2 for q in left), default=-cx)) / 2 if left else 0.0
    xr = (cx + min(((q.x0 + q.x1) / 2 for q in right), default=0)) / 2 if right else page.rect.width

    below_neighbors = [q for q in photos if q.y0 > ph.y1 - 1 and q.x0 < ph.x1 and q.x1 > ph.x0]
    ylim = min([q.y0 for q in below_neighbors] + [ph.y1 + 130, page.rect.height])
    below = fitz.Rect(xl, ph.y1 - 1, xr, ylim)

    nxt = min((q.x0 for q in right), default=page.rect.width)
    side = fitz.Rect(ph.x1 - 2, ph.y0 - 3, nxt, ph.y1 + 3)
    return below, side


def _words_in(words, rect):
    inside = [w for w in words
              if rect.x0 <= (w[0] + w[2]) / 2 < rect.x1 and rect.y0 <= (w[1] + w[3]) / 2 < rect.y1]
    # Alcuni export simulano il grassetto stampando la stessa parola due volte,
    # quasi sovrapposta: tengo una sola copia.
    out = []
    for w in inside:
        if any(w[4] == q[4] and abs(w[0] - q[0]) < 3 and abs(w[1] - q[1]) < 3 for q in out):
            continue
        out.append(w)
    return out


def _read_block(words, rect):
    """Cerca nel rettangolo: nome (righe sopra il codice), codice, prezzi. None se manca il codice."""
    lines = _lines(_words_in(words, rect))
    code_idx = next((i for i, ln in enumerate(lines) if len(ln["tokens"]) == 1 and _is_code(ln["tokens"][0])), None)
    if code_idx is None:
        return None
    name = " ".join(ln["text"] for ln in lines[:code_idx]).strip()
    prices = {}
    for ln in lines[code_idx + 1:]:
        for k, v in _parse_prices(ln["text"]).items():
            prices.setdefault(k, v)
    return {"code": lines[code_idx]["tokens"][0], "name": name, "prices": prices}


def extract_items(pdf_bytes):
    """
    Ritorna (items, notes). Ogni item:
        {"style_number", "style_name", "prices": {"W":.., "P":.., "R":..},
         "image_bytes", "page"}
    notes: avvisi leggibili (foto senza codice, ecc.).
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    items, notes = [], []
    try:
        for pno, page in enumerate(doc, start=1):
            photos = _main_photos(page)
            if not photos:
                continue
            words = page.get_text("words")
            page_items, orphans = [], 0
            for ph in photos:
                below, side = _regions(ph, photos, page)
                blk = _read_block(words, below) or _read_block(words, side)
                if blk is None:
                    orphans += 1
                    continue
                clip = fitz.Rect(ph.x0 + 1.5, ph.y0 + 1.5, ph.x1 - 1.5, ph.y1 - 1.5)  # evita bordi delle foto vicine
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=clip, alpha=False)
                page_items.append({
                    "style_number": blk["code"],
                    "style_name": blk["name"],
                    "prices": blk["prices"],
                    "image_bytes": pix.tobytes("png"),
                    "page": pno,
                })
            if page_items and orphans:
                notes.append(f"pagina {pno}: {orphans} foto senza codice stile (ignorate)")
            items.extend(page_items)
    finally:
        doc.close()
    return items, notes


def pick_price(prices, pref="W"):
    """Sceglie il prezzo da mettere in colonna F. Ritorna (valore, tipo) o (None, '')."""
    order = {"W": ["W", "P", "R"], "R": ["R", "P", "W"]}[pref]
    for k in order:
        if prices.get(k) is not None:
            return prices[k], k
    return None, ""


# ---------------------------------------------------------------------
# Costruzione dell'Excel (identico al file di riferimento)
# ---------------------------------------------------------------------
def build_linesheet_xlsx(items, template_path=TEMPLATE_PATH):
    """
    items: lista di dict con "style_number", "style_name", "price" (numero o
    None) e "image_bytes" (o None). Ritorna i bytes del file .xlsx.
    """
    import openpyxl
    from openpyxl.drawing.image import Image as XLImage

    wb = openpyxl.load_workbook(template_path)
    ws = wb.active
    ws._images = []

    # stile di riferimento = riga 4 del template (A..O)
    proto = {c: copy(ws.cell(row=4, column=c)._style) for c in range(1, 16)}

    for i, it in enumerate(items):
        r = 4 + i
        for c in range(1, 16):
            ws.cell(row=r, column=c)._style = copy(proto[c])
        ws.row_dimensions[r].height = ROW_HEIGHT

        ws.cell(row=r, column=2, value=it.get("style_number"))
        ws.cell(row=r, column=3, value=it.get("style_name"))
        ws.cell(row=r, column=6, value=it.get("price"))
        ws.cell(row=r, column=7, value=f"=ROUNDUP(F{r}*75%,0)")
        ws.cell(row=r, column=8, value=f"=ROUND(G{r}*H$3,0)")
        ws.cell(row=r, column=9, value=f"=ROUND(G{r}*I$3,0)")
        ws.cell(row=r, column=10, value=f"=ROUND(G{r}*J$3,0)")
        ws.cell(row=r, column=11, value=f"=ROUNDUP(G{r}+H{r}+I{r}+J{r},0)")
        ws.cell(row=r, column=12, value=f"=M{r}-K{r}")
        ws.cell(row=r, column=14, value=f"=M{r}*$M$1")
        ws.cell(row=r, column=15, value=f'=IFERROR(L{r}/M{r},"")')

        if it.get("image_bytes"):
            xl = XLImage(io.BytesIO(_prep_image(it["image_bytes"])))
            xl.width, xl.height = IMG_DISPLAY_W, IMG_DISPLAY_H
            ws.add_image(xl, f"A{r}")

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# ---------------------------------------------------------------------
# Scheda Streamlit
# ---------------------------------------------------------------------
def render_tab():
    import pandas as pd
    import streamlit as st

    st.subheader("Genera il linesheet Excel da uno o più PDF")
    st.caption(
        "Carica uno o più PDF scaricati da JOOR (anche di brand, stagioni e impaginazioni "
        "diverse): il layout viene riconosciuto da solo. Ottieni sempre lo stesso Excel, con "
        "foto, Style Number, Style Name, Price e le colonne di calcolo (25% OFF, Finance, Duty, "
        "Cost, Mark up, Price EURO/DOLLAR, Margin) già con le formule."
    )

    pref_label = st.radio(
        "Se un capo ha più prezzi, metti in colonna F il prezzo",
        ["Wholesale (W)", "Retail (R)"],
        horizontal=True,
        key="lsx_price_pref",
        help="Se nel PDF c'è un solo prezzo (W, R oppure P) viene usato quello.",
    )
    pref = "W" if pref_label.startswith("Wholesale") else "R"

    files = st.file_uploader("PDF (uno o più)", type=["pdf"], accept_multiple_files=True, key="lsx_files")
    if not files:
        st.info("Carica almeno un PDF per continuare.")
        return
    if not os.path.exists(TEMPLATE_PATH):
        st.error("Manca linesheet_template.xlsx: mettilo nella stessa cartella di app.py.")
        return

    @st.cache_data(show_spinner="Estrazione capi e foto dai PDF...")
    def _extract_all(file_tuples):
        all_items, report, seen = [], [], set()
        for name, data in file_tuples:
            items, notes = extract_items(data)
            kept = 0
            for it in items:
                if it["style_number"] in seen:
                    notes.append(f"{it['style_number']}: duplicato (già presente), ignorato")
                    continue
                seen.add(it["style_number"])
                it["source"] = name
                all_items.append(it)
                kept += 1
            no_price = sum(1 for it in items if not it["prices"])
            if no_price:
                notes.append(f"{no_price} capi senza prezzo nel PDF: compila la colonna Price a mano")
            report.append({"file": name, "capi": kept, "notes": notes})
        return all_items, report

    all_items, report = _extract_all(tuple((f.name, f.getvalue()) for f in files))

    for rep in report:
        st.write(f"**{rep['file']}** → {rep['capi']} capi")
        for n in rep["notes"]:
            st.warning(n)

    if not all_items:
        st.error("Non ho trovato capi in questi PDF (servono foto con codice stile e nome).")
        return

    rows = []
    for idx, it in enumerate(all_items):
        thumb = Image.open(io.BytesIO(it["image_bytes"]))
        thumb.thumbnail((120, 160))
        buf = io.BytesIO()
        thumb.convert("RGB").save(buf, "JPEG", quality=70)
        price, kind = pick_price(it["prices"], pref)
        rows.append({
            "id": idx,
            "Foto": "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode(),
            "Style Number": it["style_number"],
            "Style Name": it["style_name"],
            "Price": price,
            "Tipo": kind,
            "File": it["source"],
        })

    st.caption("Controlla e correggi i dati prima di generare. Per eliminare un capo selezionane la riga.")
    edited = st.data_editor(
        pd.DataFrame(rows),
        column_config={
            "Foto": st.column_config.ImageColumn("Foto"),
            "Price": st.column_config.NumberColumn("Price", format="%.2f"),
            "Tipo": st.column_config.TextColumn("Tipo prezzo", help="W = wholesale, R = retail, P = prezzo unico"),
        },
        column_order=["Foto", "Style Number", "Style Name", "Price", "Tipo", "File"],
        disabled=["Foto", "Tipo", "File"],
        num_rows="dynamic",
        hide_index=True,
        use_container_width=True,
        key=f"lsx_editor_{pref}_{len(all_items)}",
    )

    if st.button("Genera Excel", type="primary", key="lsx_go"):
        final = []
        for rec in edited.to_dict("records"):
            idx = rec.get("id")
            img = all_items[int(idx)]["image_bytes"] if pd.notna(idx) else None
            price = rec.get("Price")
            final.append({
                "style_number": rec.get("Style Number"),
                "style_name": rec.get("Style Name"),
                "price": float(price) if pd.notna(price) else None,
                "image_bytes": img,
            })
        st.download_button(
            "Scarica Excel linesheet",
            data=build_linesheet_xlsx(final),
            file_name="linesheet.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="lsx_dl",
        )
