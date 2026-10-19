"""
joor_tool.py - scheda Streamlit per i line sheet JOOR (es. RHEA COSTA SS27).

Permette di: modificare prezzi (Wholesale/Retail/valuta), modificare / eliminare /
aggiungere varianti colore e taglie, creare/modificare/spostare/eliminare il
rettangolo CCxx, evidenziare (giallo / best seller), scegliere i capi e scaricare
il PDF completo oppure solo la selezione.

Uso in app.py:   import joor_tool   ...   with tab_joor: joor_tool.render_tab()
"""

import hashlib
import json
import re

import fitz
import streamlit as st

import joor_core as jc


@st.cache_data(show_spinner="Analisi del PDF in corso...")
def _parse(pdf_bytes, pattern):
    return jc.parse_items(pdf_bytes, pattern)


@st.cache_data(show_spinner="Aggiorno il PDF...")
def _build(pdf_bytes, pattern, edits_json, currency, hide_retail):
    items = _parse(pdf_bytes, pattern)
    return jc.build_pdf(pdf_bytes, items, json.loads(edits_json), currency, hide_retail)


def _excel_prices(xls_file):
    """Prezzi opzionali da Excel (Zedonk) tramite core.load_prices_from_xls."""
    if not xls_file:
        return {}
    try:
        import core
        recs = core.load_prices_from_xls(xls_file.read())
    except Exception as exc:  # noqa: BLE001
        st.warning(f"Impossibile leggere l'Excel: {exc}")
        return {}
    return {r["product_id"]: r for r in recs}


def _find_excel(prices, pid):
    if pid in prices:
        return prices[pid]
    for k, v in prices.items():
        if k.startswith(pid) or pid in k:
            return v
    return None


def _inc(key):
    st.session_state[key] = st.session_state.get(key, 0) + 1


def render_tab():
    st.subheader("Line sheet JOOR (4 capi per pagina)")
    st.caption(
        "Carica il PDF esportato da JOOR (codici tipo 26035D). Puoi modificare prezzi, varianti, "
        "taglie e il rettangolo CCxx, evidenziare i capi e scaricare il PDF completo o solo la selezione."
    )

    pdf_file = st.file_uploader("PDF JOOR", type=["pdf"], key="joor_pdf")
    with st.expander("Opzioni avanzate"):
        pattern = st.text_input("Formato del codice capo (espressione regolare)", jc.DEFAULT_CODE_PATTERN, key="joor_pattern")
        xls_file = st.file_uploader(
            "Prezzi da Excel (facoltativo, .xls Zedonk): precompila Wholesale/Retail € se il Product ID corrisponde al codice",
            type=["xls"], key="joor_xls",
        )
    if not pdf_file:
        st.info("Carica il PDF per continuare.")
        return

    pdf_bytes = pdf_file.read()
    items = _parse(pdf_bytes, pattern)
    if not items:
        st.warning("Nessun capo trovato: controlla il formato del codice nelle opzioni avanzate.")
        return

    prices = _excel_prices(xls_file)
    P = "jr" + hashlib.md5(pdf_bytes).hexdigest()[:8] + ("x" if prices else "") + "_"
    st.success(f"{len(items)} capi trovati.")

    # ---------------- stato iniziale dei widget ----------------
    for it in items:
        k = it["key"]
        rec = _find_excel(prices, it["pid"]) if prices else None
        w0, r0 = it["wholesale"], it["retail"]
        if rec:
            w0 = rec.get("wholesale_eur") or w0
            r0 = rec.get("retail_eur") or r0
        st.session_state.setdefault(f"{P}sel_{k}", True)
        st.session_state.setdefault(f"{P}hl_{k}", False)
        st.session_state.setdefault(f"{P}bs_{k}", False)
        st.session_state.setdefault(f"{P}w_{k}", float(w0 or 0.0))
        st.session_state.setdefault(f"{P}r_{k}", float(r0 or 0.0))
        st.session_state.setdefault(f"{P}sz_{k}", it["sizes_text"])
        st.session_state.setdefault(f"{P}cc_{k}", (it["cc"] or {}).get("code", ""))
        st.session_state.setdefault(f"{P}dx_{k}", 0.0)
        st.session_state.setdefault(f"{P}dy_{k}", 0.0)
        st.session_state.setdefault(f"{P}nn_{k}", 0)
        for i, en in enumerate(it["entries"]):
            st.session_state.setdefault(f"{P}vt_{k}_{i}", en["text"])
            st.session_state.setdefault(f"{P}vd_{k}_{i}", False)

    # ---------------- opzioni globali ----------------
    g1, g2 = st.columns(2)
    orig_cur = next((i["currency"] for i in items if i["w_lines"]), "EUR")
    cur_opts = [orig_cur] + [c for c in ["EUR", "USD", "US$", "GBP"] if c != orig_cur]
    currency = g1.selectbox("Valuta dei prezzi", cur_opts, key=f"{P}cur",
                            help="Cambiare la valuta riscrive l'etichetta (es. EUR → USD) su tutti i capi; i valori vanno inseriti a mano.")
    hide_retail = g2.checkbox("Nascondi la riga Retail in tutti i capi", key=f"{P}hide")

    # ---------------- trova e sostituisci nelle varianti ----------------
    with st.expander("🔁 Trova e sostituisci nelle varianti di tutti i capi (es. LILIAC → LILAC)"):
        f1, f2, f3 = st.columns([2, 2, 1])
        find = f1.text_input("Trova", key=f"{P}find")
        repl = f2.text_input("Sostituisci con", key=f"{P}repl")
        if f3.button("Applica", key=f"{P}apply") and find:
            n = 0
            for it in items:
                for i in range(len(it["entries"])):
                    kk = f"{P}vt_{it['key']}_{i}"
                    new = re.sub(re.escape(find), repl, st.session_state[kk], flags=re.I)
                    if new != st.session_state[kk]:
                        st.session_state[kk] = new
                        n += 1
            st.toast(f"{n} varianti modificate")

    search = st.text_input("Filtra per codice o descrizione", key=f"{P}search")
    b1, b2 = st.columns(2)
    if b1.button("Seleziona tutti", key=f"{P}all"):
        for it in items:
            st.session_state[f"{P}sel_{it['key']}"] = True
    if b2.button("Deseleziona tutti", key=f"{P}none"):
        for it in items:
            st.session_state[f"{P}sel_{it['key']}"] = False

    # ---------------- elenco capi ----------------
    placeholders = {}
    for it in items:
        k, pid = it["key"], it["pid"]
        if search and search.lower() not in pid.lower() and search.lower() not in it["name"].lower():
            continue
        with st.container(border=True):
            img_col, info_col = st.columns([1, 2])
            placeholders[k] = img_col.empty()
            with info_col:
                t = st.columns([1, 1.3, 1.3])
                t[0].checkbox("Includi", key=f"{P}sel_{k}")
                t[1].checkbox("🟨 Rettangolo giallo", key=f"{P}hl_{k}")
                t[2].checkbox("⭐ Best seller", key=f"{P}bs_{k}")
                st.markdown(f"**{it['name']}**  \n`{pid}` · pag. {it['page'] + 1}")

                if it["w_lines"]:
                    p1, p2 = st.columns(2)
                    p1.number_input("Wholesale", key=f"{P}w_{k}", step=1.0, format="%.2f", min_value=0.0)
                    p2.number_input("Retail", key=f"{P}r_{k}", step=1.0, format="%.2f", min_value=0.0)

                with st.expander(f"Varianti ({len(it['entries'])}) e taglie"):
                    for i in range(len(it["entries"])):
                        c1, c2 = st.columns([4, 1])
                        c1.text_input(f"Variante {i + 1}", key=f"{P}vt_{k}_{i}", label_visibility="collapsed")
                        c2.checkbox("Elimina", key=f"{P}vd_{k}_{i}")
                    for j in range(st.session_state[f"{P}nn_{k}"]):
                        st.session_state.setdefault(f"{P}nt_{k}_{j}", "")
                        st.session_state.setdefault(f"{P}nc_{k}_{j}", "#cccccc")
                        st.session_state.setdefault(f"{P}nd_{k}_{j}", False)
                        c1, c2, c3 = st.columns([4, 1, 1])
                        c1.text_input("Nuova variante (es. 11 CORAL)", key=f"{P}nt_{k}_{j}")
                        c2.color_picker("Colore", key=f"{P}nc_{k}_{j}")
                        c3.checkbox("Elimina", key=f"{P}nd_{k}_{j}")
                    st.button("➕ Aggiungi variante", key=f"{P}add_{k}", on_click=_inc, args=(f"{P}nn_{k}",))
                    st.text_input("Taglie", key=f"{P}sz_{k}")

                with st.expander("Rettangolo codice (CCxx)"):
                    c1, c2, c3 = st.columns(3)
                    c1.text_input("Codice (vuoto = nessun rettangolo)", key=f"{P}cc_{k}")
                    c2.number_input("Sposta a destra (pt)", key=f"{P}dx_{k}", step=5.0)
                    c3.number_input("Sposta in basso (pt)", key=f"{P}dy_{k}", step=5.0)
                    st.caption("Se il capo non ha il rettangolo, scrivi un codice per crearlo accanto alle miniature; usa gli spostamenti per metterlo in un punto vuoto vicino.")

    # ---------------- raccolta modifiche ----------------
    edits = {}
    for it in items:
        k = it["key"]
        variants = []
        for i, en in enumerate(it["entries"]):
            txt = st.session_state[f"{P}vt_{k}_{i}"].strip()
            if st.session_state[f"{P}vd_{k}_{i}"] or not txt:
                continue
            variants.append({"o": i, "text": txt})
        for j in range(st.session_state[f"{P}nn_{k}"]):
            txt = st.session_state[f"{P}nt_{k}_{j}"].strip()
            if txt and not st.session_state[f"{P}nd_{k}_{j}"]:
                variants.append({"o": None, "text": txt, "hex": st.session_state[f"{P}nc_{k}_{j}"]})
        edits[k] = {
            "wholesale": st.session_state[f"{P}w_{k}"] if it["w_lines"] else None,
            "retail": st.session_state[f"{P}r_{k}"] if it["r_lines"] else None,
            "variants": variants,
            "sizes": st.session_state[f"{P}sz_{k}"],
            "cc_code": st.session_state[f"{P}cc_{k}"],
            "cc_dx": st.session_state[f"{P}dx_{k}"],
            "cc_dy": st.session_state[f"{P}dy_{k}"],
            "highlight": st.session_state[f"{P}hl_{k}"],
            "bestseller": st.session_state[f"{P}bs_{k}"],
        }
    selected = {it["key"] for it in items if st.session_state[f"{P}sel_{it['key']}"]}

    edited = _build(pdf_bytes, pattern, json.dumps(edits, sort_keys=True), currency, hide_retail)

    # anteprime dal PDF gia' modificato
    pdoc = fitz.open(stream=edited, filetype="pdf")
    by_key = {it["key"]: it for it in items}
    for k, ph in placeholders.items():
        ph.image(jc.render_slot(pdoc, by_key[k]), use_container_width=True)

    # ---------------- download ----------------
    st.divider()
    st.caption(f"{len(selected)} capi selezionati su {len(items)}")
    d1, d2 = st.columns(2)
    d1.download_button("⬇️ Scarica PDF completo", data=edited, file_name="line_sheet_joor.pdf",
                       mime="application/pdf", type="primary", key=f"{P}dl_all")
    if selected:
        sel_pdf = jc.build_selection_pdf(edited, items, selected)
        d2.download_button("⬇️ Scarica PDF selezione", data=sel_pdf, file_name="selezione_joor.pdf",
                           mime="application/pdf", key=f"{P}dl_sel")
    else:
        d2.info("Seleziona almeno un capo per il PDF selezione.")
