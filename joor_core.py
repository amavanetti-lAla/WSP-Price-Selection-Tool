"""
joor_core.py - logica per i line sheet in formato JOOR (4 capi per pagina,
codici tipo 26035D, righe "Wholesale: EUR ..." / "Retail: EUR ...", elenco
varianti colore con campioncino, riga "Sizes:" e rettangolo CCxx).

Nessuna dipendenza da Streamlit: si puo' testare da solo.
"""

import math
import re
import datetime

import fitz  # PyMuPDF

DEFAULT_CODE_PATTERN = r"\d{5}[A-Z]"

# Griglia 2x2 espressa in frazioni della pagina (misurata sul PDF JOOR)
SLOT_X = {"L": (0.050, 0.500), "R": (0.505, 0.955)}
SLOT_Y = {"T": (0.070, 0.440), "B": (0.490, 0.860)}
ROW_SPLIT = 0.47
HEADER_H = 0.074
# Rettangolo "CCxx" standard (frazioni di pagina)
CC_X = {"L": 0.272, "R": 0.728}
CC_Y = {"T": 0.370, "B": 0.783}
CC_W, CC_H = 0.067, 0.0277

YELLOW = (0.96, 0.77, 0.09)  # #f5c518, come in core.py


# ------------------------------------------------------------------
# Utilita'
# ------------------------------------------------------------------
def slot_rect(pw, ph, col, row):
    return (SLOT_X[col][0] * pw, SLOT_Y[row][0] * ph, SLOT_X[col][1] * pw, SLOT_Y[row][1] * ph)


def cc_std_rect(pw, ph, col, row):
    x0 = CC_X[col] * pw
    y0 = CC_Y[row] * ph
    return (x0, y0, x0 + CC_W * pw, y0 + CC_H * ph)


def _infl(r, d):
    r = fitz.Rect(r)
    return fitz.Rect(r.x0 - d, r.y0 - d, r.x1 + d, r.y1 + d)


def fmt_price(v):
    return f"{v:,.2f}"


def _tlen(text, fs, bold=False):
    return fitz.get_text_length(text, fontname="hebo" if bold else "helv", fontsize=fs)


def _wrap(text, max_w, fs):
    lines, cur = [], ""
    for w in text.split():
        t = (cur + " " + w).strip()
        if cur and _tlen(t, fs) > max_w:
            lines.append(cur)
            cur = w
        else:
            cur = t
    if cur:
        lines.append(cur)
    return lines or [""]


def _read_lines(page, rect):
    d = page.get_text("dict", clip=rect)
    out = []
    for b in d["blocks"]:
        if b.get("type") != 0:
            continue
        for l in b["lines"]:
            spans = [s for s in l["spans"] if s["text"].strip()]
            if not spans:
                continue
            text = "".join(s["text"] for s in l["spans"]).strip()
            x0, y0, x1, y1 = l["bbox"]
            cy = (y0 + y1) / 2
            if not (rect.y0 <= cy <= rect.y1 and x0 >= rect.x0 - 1):
                continue
            out.append({
                "text": text,
                "bbox": (x0, y0, x1, y1),
                "ox": spans[0]["origin"][0],
                "oy": spans[0]["origin"][1],
                "size": spans[0]["size"],
            })
    out.sort(key=lambda l: (round(l["bbox"][1], 1), l["bbox"][0]))
    return out


def _price_of(lines):
    full = " ".join(l["text"] for l in lines)
    cur = re.search(r"(EUR|USD|US\$|GBP|€|\$|£)", full)
    val = re.search(r"(\d[\d,]*\.\d{2})", full)
    return (cur.group(1) if cur else "EUR", float(val.group(1).replace(",", "")) if val else None)


# ------------------------------------------------------------------
# Analisi PDF
# ------------------------------------------------------------------
def parse_items(pdf_bytes, pattern=DEFAULT_CODE_PATTERN):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    rx = re.compile(r"^(?:%s)$" % pattern)
    items, seen = [], {}
    for pno, page in enumerate(doc):
        pw, ph = page.rect.width, page.rect.height
        imgs = [fitz.Rect(i["bbox"]) for i in page.get_image_info()]
        ccs = []
        for a in page.annots() or []:
            if a.type[0] == 2:  # FreeText
                ccs.append({"code": (a.info.get("content") or "").strip(), "rect": tuple(a.rect), "xref": a.xref})
        found = []
        for w in page.get_text("words"):
            if rx.match(w[4]):
                found.append(w)
        for w in found:
            x0, y0, _, _, pid = w[:5]
            col = "L" if x0 < pw * 0.5 else "R"
            row = "T" if y0 < ph * ROW_SPLIT else "B"
            slot = slot_rect(pw, ph, col, row)
            trect = fitz.Rect(x0 - 3, slot[1], slot[2], slot[3])
            lines = _read_lines(page, trect)
            idx = next((i for i, l in enumerate(lines) if l["text"] == pid), None)
            if idx is None:
                continue
            name = " ".join(l["text"] for l in lines[:idx])
            w_lines, r_lines, s_lines, entries = [], [], [], []
            mode = None
            for l in lines[idx + 1:]:
                t = l["text"]
                if t.startswith("Wholesale"):
                    mode, w_lines = "w", [l]
                elif t.startswith("Retail"):
                    mode, r_lines = "r", [l]
                elif t.startswith("Sizes"):
                    mode, s_lines = "s", [l]
                elif mode in ("w", "r", "c", None) and re.match(r"^\d{2}(\s|$)", t) and not (
                        mode in ("w", "r") and re.match(r"^[\d,]*\.\d{2}$", t)):
                    mode = "c"
                    entries.append({"lines": [l]})
                elif mode == "w":
                    w_lines.append(l)
                elif mode == "r":
                    r_lines.append(l)
                elif mode == "c" and entries:
                    entries[-1]["lines"].append(l)
                elif mode == "s":
                    s_lines.append(l)

            # campioncini colore (immagini piccole nella colonna testo)
            sw = [r for r in imgs
                  if 8 <= r.width <= 30 and 8 <= r.height <= 30
                  and x0 - 6 <= r.x0 <= x0 + 8 and slot[1] <= r.y0 <= slot[3]]
            used = set()
            for en in entries:
                en["text"] = " ".join(l["text"] for l in en["lines"])
                cy = (en["lines"][0]["bbox"][1] + en["lines"][-1]["bbox"][3]) / 2
                best = None
                for j, r in enumerate(sw):
                    if j in used:
                        continue
                    d = abs((r.y0 + r.y1) / 2 - cy)
                    if d < 12 and (best is None or d < best[0]):
                        best = (d, j)
                if best:
                    used.add(best[1])
                    en["swatch"] = tuple(sw[best[1]])
                else:
                    en["swatch"] = (x0 + 1, cy - 8, x0 + 17, cy + 8)

            fs = entries[0]["lines"][0]["size"] if entries else 7.5
            wc, wv = _price_of(w_lines) if w_lines else ("EUR", None)
            rc, rv = _price_of(r_lines) if r_lines else ("EUR", None)
            sizes_text = " ".join(l["text"] for l in s_lines)
            sizes_text = re.sub(r"^Sizes:?\s*", "", sizes_text)

            cands = [r for r in imgs if r.x0 >= slot[0] - 1 and r.x1 <= slot[2] + 1 and r.y0 >= slot[1] - 1 and r.y1 <= slot[3] + 1]
            photo = tuple(max(cands, key=lambda r: r.width * r.height)) if cands else slot
            cc = None
            for c in ccs:
                cx, cy = (c["rect"][0] + c["rect"][2]) / 2, (c["rect"][1] + c["rect"][3]) / 2
                if slot[0] <= cx <= slot[2] and slot[1] <= cy <= slot[3]:
                    cc = c
                    break
            key = pid if pid not in seen else f"{pid}_{pno}_{col}{row}"
            seen[pid] = True
            first_sw = entries[0]["swatch"] if entries else (x0, 0, x0 + 16, 16)
            items.append({
                "key": key, "pid": pid, "page": pno, "col": col, "row": row,
                "slot": slot, "trect": tuple(trect), "code_x": x0, "name": name,
                "wholesale": wv, "retail": rv, "currency": wc,
                "w_lines": w_lines, "r_lines": r_lines,
                "entries": entries, "size_lines": s_lines, "sizes_text": sizes_text,
                "fs": fs,
                "x_text": min([e["lines"][0]["bbox"][0] for e in entries], default=x0 + 19),
                "max_w": max([l["bbox"][2] - l["bbox"][0] for e in entries for l in e["lines"]], default=60) + 0.5,
                "sw_w": first_sw[2] - first_sw[0], "sw_h": first_sw[3] - first_sw[1],
                "sw_x0": first_sw[0], "cc": cc, "photo": photo,
            })
    order = {"T": 0, "B": 1, "L": 0, "R": 1}
    items.sort(key=lambda i: (i["page"], order[i["row"]], order[i["col"]]))
    return items


# ------------------------------------------------------------------
# Costruzione del PDF modificato
# ------------------------------------------------------------------
def _apply_redactions(page):
    try:
        page.apply_redactions(images=0, graphics=0)
    except TypeError:
        page.apply_redactions(images=0)


def redact_text(pdf_bytes, text):
    """Rimuove davvero (non solo copre) ogni occorrenza di `text` nel PDF."""
    text = (text or "").strip()
    if not text:
        return pdf_bytes
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    found = False
    for page in doc:
        hits = page.search_for(text)
        if hits:
            found = True
            for r in hits:
                page.add_redact_annot(r, fill=(1, 1, 1))
            _apply_redactions(page)
    return doc.tobytes(garbage=3, deflate=True) if found else pdf_bytes


def _write_lv(page, x, y, label, rest, fs):
    """Etichetta in grassetto + resto in regolare sulla stessa riga."""
    page.insert_text((x, y), label, fontname="hebo", fontsize=fs)
    page.insert_text((x + _tlen(label, fs, True), y), rest, fontname="helv", fontsize=fs)


def _rewrite_prices(page, it, e, currency, hide_retail):
    fs = it["w_lines"][0]["size"] if it["w_lines"] else it["fs"]
    x = it["w_lines"][0]["ox"]
    cur = currency
    if it["w_lines"] and e.get("wholesale") is not None:
        v = fmt_price(e["wholesale"])
        wl = it["w_lines"]
        if len(wl) >= 2:
            _write_lv(page, x, wl[0]["oy"], "Wholesale", f": {cur}", fs)
            page.insert_text((x, wl[1]["oy"]), v, fontname="helv", fontsize=fs)
        else:
            _write_lv(page, x, wl[0]["oy"], "Wholesale", f": {cur} {v}", fs)
    if it["r_lines"] and not hide_retail and e.get("retail") is not None:
        v = fmt_price(e["retail"])
        rl = it["r_lines"]
        if len(rl) >= 2:
            _write_lv(page, x, rl[0]["oy"], "Retail", f": {cur}", fs)
            page.insert_text((x, rl[1]["oy"]), v, fontname="helv", fontsize=fs)
        else:
            _write_lv(page, x, rl[0]["oy"], "Retail", f": {cur} {v}", fs)


def _price_region(it):
    ls = it["w_lines"] + it["r_lines"]
    y0 = min(l["bbox"][1] for l in ls) + 1.5
    y1 = max(l["bbox"][3] for l in ls) - 1.5
    return fitz.Rect(it["code_x"] - 2, y0, it["trect"][2], y1)


def _variant_region(it):
    ys0, ys1 = [], []
    for en in it["entries"]:
        ys0 += [l["bbox"][1] for l in en["lines"]] + [en["swatch"][1]]
        ys1 += [l["bbox"][3] for l in en["lines"]] + [en["swatch"][3]]
    for l in it["size_lines"]:
        ys0.append(l["bbox"][1])
        ys1.append(l["bbox"][3])
    return fitz.Rect(it["code_x"] - 3, min(ys0), it["trect"][2], max(ys1))


def _rebuild_variants(page, spage, it, e):
    fs = it["fs"]
    lh = fs * 1.15
    reg = _variant_region(it)
    page.draw_rect(reg, color=None, fill=(1, 1, 1))
    ents = it["entries"]
    y = reg.y0
    for v in e["variants"]:
        o = v.get("o")
        if o is not None and v["text"] == ents[o]["text"]:
            tl = [l["text"] for l in ents[o]["lines"]]
        else:
            tl = _wrap(v["text"], it["max_w"], fs)
        n = len(tl)
        h = max(it["sw_h"], n * lh)
        cy = y + h / 2
        sr = fitz.Rect(it["sw_x0"], cy - it["sw_h"] / 2, it["sw_x0"] + it["sw_w"], cy + it["sw_h"] / 2)
        if o is not None:
            sw = ents[o]["swatch"]
            png = spage.get_pixmap(clip=fitz.Rect(sw), dpi=300).tobytes("png")
            page.insert_image(sr, stream=png)
        else:
            hx = v.get("hex", "#cccccc").lstrip("#")
            rgb = tuple(int(hx[i:i + 2], 16) / 255 for i in (0, 2, 4))
            page.draw_rect(sr, color=(0.75, 0.75, 0.75), fill=rgb, width=0.8)
        top = cy - n * lh / 2
        for k, t in enumerate(tl):
            page.insert_text((it["x_text"], top + k * lh + fs * 0.88), t, fontname="helv", fontsize=fs)
        y += h + 6
    # Sizes
    sizes = e["sizes"].strip()
    if sizes:
        if sizes == it["sizes_text"] and it["size_lines"]:
            tl = [l["text"] for l in it["size_lines"]]
        else:
            ow = max([l["bbox"][2] - l["bbox"][0] for l in it["size_lines"]], default=80)
            tl = _wrap("Sizes: " + sizes, ow + 0.5, fs)
        top = y + 5
        for k, t in enumerate(tl):
            by = top + k * lh + fs * 0.88
            if k == 0 and t.startswith("Sizes"):
                _write_lv(page, it["code_x"], by, "Sizes", t[5:], fs)
            else:
                page.insert_text((it["code_x"], by), t, fontname="helv", fontsize=fs)


def _star(page, cx, cy, R):
    pts = []
    for k in range(10):
        ang = -math.pi / 2 + k * math.pi / 5
        rad = R if k % 2 == 0 else R * 0.4
        pts.append(fitz.Point(cx + rad * math.cos(ang), cy + rad * math.sin(ang)))
    sh = page.new_shape()
    sh.draw_polyline(pts + [pts[0]])
    sh.finish(fill=YELLOW, color=YELLOW, width=1, closePath=True)
    sh.commit()


def _variants_changed(it, e):
    orig = [(i, en["text"]) for i, en in enumerate(it["entries"])]
    new = [(v.get("o"), v["text"]) for v in e["variants"]]
    return orig != new


def build_pdf(pdf_bytes, items, edits, currency="EUR", hide_retail=False):
    src = fitz.open(stream=pdf_bytes, filetype="pdf")
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    by_page = {}
    for it in items:
        by_page.setdefault(it["page"], []).append(it)

    for pno, its in by_page.items():
        page, spage = doc[pno], src[pno]
        pw, ph = page.rect.width, page.rect.height
        todo = []
        for it in its:
            e = edits.get(it["key"])
            if not e:
                continue
            price_ch = bool(it["w_lines"]) and (
                e.get("wholesale") != it["wholesale"] or e.get("retail") != it["retail"]
                or currency != it["currency"] or (hide_retail and it["r_lines"]))
            var_ch = bool(it["entries"]) and (_variants_changed(it, e) or e["sizes"].strip() != it["sizes_text"])
            if price_ch:
                page.add_redact_annot(_price_region(it), fill=(1, 1, 1))
            if var_ch:
                r = _variant_region(it)
                page.add_redact_annot(fitz.Rect(r.x0, r.y0 + 1.2, r.x1, r.y1 - 1.2), fill=(1, 1, 1))
            todo.append((it, e, price_ch, var_ch))
        if any(t[2] or t[3] for t in todo):
            _apply_redactions(page)

        for it, e, price_ch, var_ch in todo:
            if price_ch:
                _rewrite_prices(page, it, e, currency, hide_retail)
            if var_ch:
                _rebuild_variants(page, spage, it, e)
            # rettangolo CC (annotazione FreeText nel PDF originale)
            code = (e.get("cc_code") or "").strip()
            dx, dy = e.get("cc_dx", 0), e.get("cc_dy", 0)
            orig = it["cc"]
            if orig is not None and (code != orig["code"] or dx or dy):
                for a in list(page.annots() or []):
                    if a.xref == orig["xref"]:
                        page.delete_annot(a)
                base = orig["rect"]
            elif orig is None and code:
                base = cc_std_rect(pw, ph, it["col"], it["row"])
            else:
                base = None
            if base is not None and code and (orig is None or code != orig["code"] or dx or dy):
                fs = 13
                tw = _tlen(code, fs)
                w = max(base[2] - base[0], tw + 10)
                h = base[3] - base[1]
                rx0, ry0 = base[0] + dx, base[1] + dy
                page.draw_rect(fitz.Rect(rx0, ry0, rx0 + w, ry0 + h), color=(0, 0, 0), fill=(1, 1, 1), width=1)
                page.insert_text((rx0 + (w - tw) / 2, ry0 + h / 2 + fs * 0.35), code, fontname="helv", fontsize=fs)
            sx0, sy0, sx1, sy1 = it["slot"]
            if e.get("highlight"):
                page.draw_rect(fitz.Rect(sx0 + 3, sy0 + 3, sx1 - 3, sy1 - 3), color=YELLOW, width=3)
            if e.get("bestseller"):
                px0, py0, px1, py1 = it["photo"]
                _star(page, px1 - 11 - 6, py0 + 11 + 6, 11)
    return doc.tobytes(garbage=3, deflate=True)


def build_selection_pdf(edited_bytes, items, selected_keys):
    src = fitz.open(stream=edited_bytes, filetype="pdf")
    try:
        src.bake(annots=True, widgets=False)  # le annotazioni (CCxx) diventano contenuto
    except Exception:
        pass
    out = fitz.open()
    sel = [it for it in items if it["key"] in selected_keys]
    if not sel:
        return None
    pw, ph = src[0].rect.width, src[0].rect.height
    m = re.search(r"Date Printed:\s*([\d/]+)", src[0].get_text())
    date = m.group(1) if m else datetime.date.today().strftime("%d/%m/%Y")
    slots = [("L", "T"), ("R", "T"), ("L", "B"), ("R", "B")]
    n_pages = math.ceil(len(sel) / 4)
    hdr = fitz.Rect(0, 0, pw, ph * HEADER_H)
    for pi in range(n_pages):
        page = out.new_page(width=pw, height=ph)
        page.show_pdf_page(hdr, src, 0, clip=hdr)
        for k, it in enumerate(sel[pi * 4:(pi + 1) * 4]):
            c, r = slots[k]
            page.show_pdf_page(fitz.Rect(slot_rect(pw, ph, c, r)), src, it["page"], clip=fitz.Rect(it["slot"]))
        txt = f"Date Printed: {date}    Page {pi + 1} of {n_pages}"
        page.insert_text((pw * 0.954 - _tlen(txt, 8), ph * 0.911), txt, fontname="helv", fontsize=8)
    return out.tobytes(garbage=3, deflate=True)


def render_slot(doc, it, zoom=1.3):
    page = doc[it["page"]]
    return page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=fitz.Rect(it["slot"])).tobytes("png")
