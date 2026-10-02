"""
image_compressor.py

Scheda "Riduci immagini" per l'app Streamlit:
- carica immagini (anche molte insieme) oppure uno ZIP che contiene
  cartelle di immagini (la struttura delle cartelle viene mantenuta)
- riduce il peso (da MB a KB) in WebP, JPEG o PNG
- scarica tutto in un unico ZIP

Dipendenze: solo Pillow (gia' presente) e la libreria standard.
"""

import io
import os
import zipfile

from PIL import Image, ImageOps

IMG_EXT = (".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff", ".gif")
FORMATS = {
    "WebP (il più leggero)": ("WEBP", "webp"),
    "JPEG (compatibile ovunque)": ("JPEG", "jpg"),
    "PNG (senza perdita, file più pesanti)": ("PNG", "png"),
}


def _encode(img, fmt, quality):
    """Codifica un'immagine PIL nel formato richiesto e ritorna i bytes."""
    out = io.BytesIO()
    if fmt == "JPEG":
        if img.mode in ("RGBA", "LA", "P"):
            rgba = img.convert("RGBA")
            bg = Image.new("RGB", rgba.size, (255, 255, 255))
            bg.paste(rgba, mask=rgba.split()[-1])
            img = bg
        elif img.mode != "RGB":
            img = img.convert("RGB")
        img.save(out, "JPEG", quality=quality, optimize=True, progressive=True)
    elif fmt == "WEBP":
        if img.mode not in ("RGB", "RGBA"):
            img = img.convert("RGBA" if "A" in img.mode or img.mode == "P" else "RGB")
        img.save(out, "WEBP", quality=quality, method=6)
    else:  # PNG
        if img.mode not in ("RGB", "RGBA", "L", "LA"):
            img = img.convert("RGBA" if img.mode == "P" else "RGB")
        img.save(out, "PNG", optimize=True)
    return out.getvalue()


def compress_image(data, fmt, quality=80, max_side=2000, target_kb=0):
    """
    Riduce un'immagine. Ritorna (bytes, (larghezza, altezza)).

    - max_side: lato massimo in px (0 = lascia la dimensione originale)
    - target_kb: peso massimo desiderato (0 = nessun limite).
        JPEG/WebP: abbassa la qualita' a passi, fino a un minimo di 45.
        PNG: non ha "qualita'", quindi rimpicciolisce a passi del 10%
        (minimo 30% della dimensione di partenza).
    """
    img = Image.open(io.BytesIO(data))
    img = ImageOps.exif_transpose(img)  # rispetta l'orientamento delle foto da telefono
    img.load()

    if max_side and max(img.size) > max_side:
        img.thumbnail((max_side, max_side), Image.LANCZOS)

    target = int(target_kb * 1024) if target_kb else 0
    base = img
    size = base.size
    out = _encode(base, fmt, quality)

    if fmt == "PNG":
        scale = 1.0
        while target and len(out) > target and scale > 0.3:
            scale -= 0.1
            size = (max(1, round(base.width * scale)), max(1, round(base.height * scale)))
            out = _encode(base.resize(size, Image.LANCZOS), fmt, quality)
    else:
        q = quality
        while target and len(out) > target and q > 45:
            q = max(45, q - 7)
            out = _encode(base, fmt, q)

    # Se il risultato pesa piu' dell'originale (ed e' lo stesso formato), tieni l'originale.
    try:
        src_fmt = Image.open(io.BytesIO(data)).format
    except Exception:
        src_fmt = None
    if len(out) >= len(data) and src_fmt == fmt:
        return data, base.size
    return out, size


def _collect_inputs(uploaded_files):
    """Ritorna una lista di (percorso, bytes) da immagini singole e ZIP."""
    items = []
    for uf in uploaded_files:
        name = uf.name
        raw = uf.getvalue()
        if name.lower().endswith(".zip"):
            with zipfile.ZipFile(io.BytesIO(raw)) as z:
                for info in z.infolist():
                    p = info.filename
                    if info.is_dir() or p.startswith("__MACOSX") or os.path.basename(p).startswith("."):
                        continue
                    if p.lower().endswith(IMG_EXT):
                        items.append((p, z.read(info)))
        elif name.lower().endswith(IMG_EXT):
            items.append((name, raw))
    return items


def _fmt_size(n):
    return f"{n / 1048576:.1f} MB" if n >= 1048576 else f"{round(n / 1024)} KB"


def render_tab():
    import streamlit as st

    st.subheader("Riduci il peso di immagini e cartelle")
    st.caption(
        "Carica immagini (anche molte insieme) oppure uno ZIP con dentro le cartelle: "
        "ogni foto viene ridotta da MB a KB senza rovinarla. Le sottocartelle "
        "dello ZIP vengono mantenute."
    )

    files = st.file_uploader(
        "Immagini o ZIP di cartelle",
        type=["jpg", "jpeg", "png", "webp", "bmp", "tif", "tiff", "gif", "zip"],
        accept_multiple_files=True,
        key="imgc_uploader",
    )

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        fmt_label = st.selectbox("Formato", list(FORMATS.keys()), key="imgc_fmt")
    fmt, ext = FORMATS[fmt_label]
    with c2:
        quality = st.slider("Qualità %", 40, 95, 80, key="imgc_q", disabled=(fmt == "PNG"))
    with c3:
        max_side = st.number_input("Lato massimo px (0 = originale)", 0, 10000, 2000, 100, key="imgc_max")
    with c4:
        target_kb = st.number_input("Peso massimo KB (0 = nessuno)", 0, 20000, 300, 50, key="imgc_kb")

    if not files:
        st.info("Carica almeno un'immagine o uno ZIP per continuare.")
        return

    if st.button("Riduci", type="primary", key="imgc_go"):
        inputs = _collect_inputs(files)
        if not inputs:
            st.warning("Non ho trovato immagini nei file caricati.")
            return

        results, errors = [], []
        bar = st.progress(0.0)
        for i, (path, data) in enumerate(inputs):
            try:
                out, (w, h) = compress_image(data, fmt, quality, max_side, target_kb)
                out_ext = ext if out is not data else os.path.splitext(path)[1].lstrip(".").lower()
                results.append({
                    "path": os.path.splitext(path)[0] + "." + out_ext,
                    "before": len(data),
                    "after": len(out),
                    "size": f"{w}×{h}",
                    "data": out,
                })
            except Exception as e:  # file non leggibile
                errors.append(f"{path}: {e}")
            bar.progress((i + 1) / len(inputs))
        bar.empty()
        st.session_state["imgc_results"] = results
        st.session_state["imgc_errors"] = errors

    results = st.session_state.get("imgc_results")
    if not results:
        return

    tot_b = sum(r["before"] for r in results)
    tot_a = sum(r["after"] for r in results)
    st.success(
        f"{len(results)} immagini: {_fmt_size(tot_b)} → {_fmt_size(tot_a)} "
        f"(−{round((1 - tot_a / tot_b) * 100)}%)"
    )
    for err in st.session_state.get("imgc_errors", []):
        st.warning("Non letto: " + err)

    st.dataframe(
        [
            {
                "File": r["path"],
                "Prima": _fmt_size(r["before"]),
                "Dopo": _fmt_size(r["after"]),
                "Risparmio": f"−{round((1 - r['after'] / r['before']) * 100)}%",
                "Dimensioni": r["size"],
            }
            for r in results
        ],
        use_container_width=True,
        hide_index=True,
    )

    if len(results) == 1:
        r = results[0]
        st.download_button(
            "Scarica immagine",
            data=r["data"],
            file_name=os.path.basename(r["path"]),
            key="imgc_dl_one",
        )
    else:
        buf = io.BytesIO()
        used = set()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_STORED) as z:
            for r in results:
                name, k = r["path"], 1
                while name in used:
                    base, e = os.path.splitext(r["path"])
                    name = f"{base}_{k}{e}"
                    k += 1
                used.add(name)
                z.writestr(name, r["data"])
        st.download_button(
            "Scarica tutto (ZIP)",
            data=buf.getvalue(),
            file_name="immagini-ridotte.zip",
            mime="application/zip",
            key="imgc_dl_zip",
        )
