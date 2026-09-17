"""
app.py - WSP Price & Selection Tool

Interfaccia web (Streamlit) per:
  1. caricare il line sheet PDF e il report prezzi Excel (Zedonk)
  2. vedere l'anteprima di ogni capo con i prezzi precompilati
     (Retail €/US$, Wholesale €/US$), modificabili a mano
  3. scegliere quali prezzi inserire nel PDF (con etichetta "Retail:"
     o "Wholesale:" davanti al valore)
  4. selezionare solo alcuni capi (checkbox) per creare un PDF ridotto
  5. scaricare il PDF finale

AVVIO IN LOCALE:
    pip install streamlit pdfplumber pymupdf reportlab pillow pypdf xlrd
    streamlit run app.py
"""

import base64
import io

import streamlit as st
import streamlit.components.v1 as components

import core

st.set_page_config(
    page_title="WSP Price & Selection Tool",
    page_icon="static/icon-192.png",
    layout="wide",
)

# ------------------------------------------------------------------
# Collega manifest.json e icona per "Aggiungi a Home" da smartphone.
# Richiede [server] enableStaticServing = true in .streamlit/config.toml
# e la cartella static/ con manifest.json + icone accanto a questo file.
# ------------------------------------------------------------------
components.html(
    """
    <script>
      const head = window.parent.document.querySelector('head');
      const cdnBase = 'https://cdn.jsdelivr.net/gh/amavanetti-lAla/WSP-Price-Selection-Tool@main/static/';

      function addLink(rel, href, attrs) {
        if (head.querySelector('link[rel="' + rel + '"]')) return;
        const link = document.createElement('link');
        link.rel = rel;
        link.href = href;
        if (attrs) {
          Object.entries(attrs).forEach(([k, v]) => link.setAttribute(k, v));
        }
        head.appendChild(link);
      }

      addLink('manifest', cdnBase + 'manifest.json');
      addLink('apple-touch-icon', cdnBase + 'apple-touch-icon.png', {sizes: '180x180'});

      if (!head.querySelector('meta[name="theme-color"]')) {
        const meta = document.createElement('meta');
        meta.name = 'theme-color';
        meta.content = '#1f2937';
        head.appendChild(meta);
      }
    </script>
    """,
    height=0,
    width=0,
)

st.title("WSP Price & Selection Tool")
st.caption("Carica il line sheet PDF (con o senza prezzi) e genera il PDF finale.")

# ------------------------------------------------------------------
# 1. Modalita' di lavoro
# ------------------------------------------------------------------
st.subheader("1. Il PDF che carichi ha già i prezzi?")
mode = st.radio(
    "Modalità",
    [
        "No: aggiungi i prezzi da un file Excel",
        "Sì: il PDF ha già i prezzi (seleziona ed evidenzia soltanto)",
    ],
    index=0,
    label_visibility="collapsed",
)
already_priced = mode.startswith("Sì")

# ------------------------------------------------------------------
# 2. Caricamento file
# ------------------------------------------------------------------
if already_priced:
    pdf_file = st.file_uploader("Line sheet PDF (già con i prezzi)", type=["pdf"])
    xls_file = None
else:
    col_up1, col_up2 = st.columns(2)
    with col_up1:
        pdf_file = st.file_uploader("Line sheet PDF", type=["pdf"])
    with col_up2:
        xls_file = st.file_uploader("Report prezzi Excel (.xls)", type=["xls"])

if not pdf_file or (not already_priced and not xls_file):
    st.info("Carica " + ("il PDF" if already_priced else "entrambi i file") + " per continuare.")
    st.stop()

pdf_bytes = pdf_file.read()
xls_bytes = xls_file.read() if xls_file else None

# ------------------------------------------------------------------
# 2b. Rimozione di un testo specifico dal PDF (facoltativo)
# ------------------------------------------------------------------
redact_text = st.text_input(
    "Testo esatto da rimuovere dal PDF (facoltativo, es. \"Wholesale: EUR 0.01\")",
    "",
    help="Cerca questa riga in ogni capo del PDF e la copre con un rettangolo "
         "bianco prima di procedere. Lascia vuoto per non rimuovere nulla.",
)
if redact_text.strip():
    @st.cache_data(show_spinner="Rimozione testo dal PDF...")
    def _redact(pdf_bytes, text):
        return core.redact_text_in_pdf(pdf_bytes, text)

    pdf_bytes = _redact(pdf_bytes, redact_text)

# ------------------------------------------------------------------
# 3. Analisi file (con cache per evitare di rifare il lavoro ad ogni click)
# ------------------------------------------------------------------
@st.cache_data(show_spinner="Analisi del PDF in corso...")
def _find_items(pdf_bytes):
    return core.find_items_in_pdf(pdf_bytes)


@st.cache_data(show_spinner="Lettura prezzi da Excel...")
def _load_prices(xls_bytes):
    return core.load_prices_from_xls(xls_bytes)


@st.cache_resource
def _open_doc(pdf_bytes):
    import fitz
    return fitz.open(stream=pdf_bytes, filetype="pdf")


items = _find_items(pdf_bytes)
price_records = _load_prices(xls_bytes) if xls_bytes else []
doc = _open_doc(pdf_bytes)

price_by_id = {r["product_id"]: r for r in price_records}

if already_priced:
    st.success(f"{len(items)} capi trovati nel PDF")
else:
    missing_in_excel = [pid for pid in items if pid not in price_by_id]
    missing_in_pdf = [r["product_id"] for r in price_records if r["product_id"] not in items]

    st.success(f"{len(items)} capi trovati nel PDF - {len(price_records)} Product ID trovati nell'Excel")
    if missing_in_excel:
        with st.expander(f"ATTENZIONE: {len(missing_in_excel)} capi nel PDF senza prezzo nell'Excel"):
            st.write(", ".join(missing_in_excel))
    if missing_in_pdf:
        with st.expander(f"ATTENZIONE: {len(missing_in_pdf)} Product ID nell'Excel non trovati nel PDF"):
            st.write(", ".join(missing_in_pdf))

# ------------------------------------------------------------------
# 4. Stato modificabile (inizializzato una sola volta per ogni PDF/Excel caricato)
# ------------------------------------------------------------------
state_key = f"price_state_{'priced' if already_priced else 'excel'}"
if state_key not in st.session_state:
    state = {}
    for pid in items:
        rec = price_by_id.get(pid, {})
        state[pid] = {
            "retail_eur": rec.get("retail_eur"),
            "retail_usd": rec.get("retail_usd"),
            "wholesale_eur": rec.get("wholesale_eur"),
            "wholesale_usd": rec.get("wholesale_usd"),
            "description": rec.get("description", ""),
            "selected": True,
            "highlighted": False,
            "bestseller": False,
        }
    st.session_state[state_key] = state

state = st.session_state[state_key]

# ------------------------------------------------------------------
# 5. Opzioni di stampa (quali prezzi inserire nel PDF) - solo se NON already_priced
# ------------------------------------------------------------------
price_fields = []
if not already_priced:
    st.subheader("2. Quali prezzi vuoi inserire nel PDF?")
    field_options = {
        "Retail €": ("Retail", "retail_eur"),
        "Retail US$": ("Retail", "retail_usd"),
        "Wholesale €": ("Wholesale", "wholesale_eur"),
        "Wholesale US$": ("Wholesale", "wholesale_usd"),
    }
    chosen_labels = st.multiselect(
        "Seleziona uno o più prezzi (verranno scritti uno sotto l'altro, con l'etichetta davanti)",
        options=list(field_options.keys()),
        default=["Wholesale €"],
    )
    price_fields = [field_options[lbl] for lbl in chosen_labels]

# ------------------------------------------------------------------
# 6. Elenco capi con anteprima, selezione ed evidenziazione
# ------------------------------------------------------------------
step_num = "2" if already_priced else "3"
step_label = "Scegli ed evidenzia i capi" if already_priced else "Controlla / modifica i prezzi e scegli i capi"
st.subheader(f"{step_num}. {step_label}")

search = st.text_input("Filtra per Product ID o descrizione", "")

select_all_col, deselect_all_col = st.columns(2)
with select_all_col:
    if st.button("Seleziona tutti"):
        for pid in state:
            state[pid]["selected"] = True
            st.session_state[f"sel_{pid}"] = True
        st.rerun()
with deselect_all_col:
    if st.button("Deseleziona tutti"):
        for pid in state:
            state[pid]["selected"] = False
            st.session_state[f"sel_{pid}"] = False
        st.rerun()

sorted_ids = sorted(items.keys())

for pid in sorted_ids:
    rec = state[pid]
    desc = rec["description"] or ""
    if search and search.lower() not in pid.lower() and search.lower() not in desc.lower():
        continue

    with st.container(border=True):
        img_col, info_col = st.columns([1, 3])
        with img_col:
            thumb = core.crop_item_thumbnail(doc, items[pid]["page"], items[pid]["box"])
            thumb_b64 = base64.b64encode(thumb).decode("ascii")
            border_style = "4px solid #f5c518" if rec["highlighted"] else "4px solid transparent"
            st.markdown(
                f"""
                <div style="border:{border_style}; border-radius:8px; padding:2px;">
                    <img src="data:image/png;base64,{thumb_b64}" style="width:100%; display:block; border-radius:4px;" />
                </div>
                """,
                unsafe_allow_html=True,
            )
        with info_col:
            top_row = st.columns([1, 1, 1, 2])
            with top_row[0]:
                rec["selected"] = st.checkbox("Includi", value=rec["selected"], key=f"sel_{pid}")
            with top_row[1]:
                rec["highlighted"] = st.checkbox(
                    "🟨 Rettangolo giallo", value=rec["highlighted"], key=f"hl_{pid}"
                )
            with top_row[2]:
                rec["bestseller"] = st.checkbox(
                    "⭐ Best seller", value=rec["bestseller"], key=f"bs_{pid}"
                )
            with top_row[3]:
                st.markdown(f"**{desc or '(senza descrizione)'}**  \n`{pid}`")

            if not already_priced:
                p1, p2, p3, p4 = st.columns(4)
                rec["retail_eur"] = p1.number_input(
                    "Retail €", value=rec["retail_eur"] or 0.0, step=1.0, key=f"re_{pid}"
                )
                rec["retail_usd"] = p2.number_input(
                    "Retail US$", value=rec["retail_usd"] or 0.0, step=1.0, key=f"ru_{pid}"
                )
                rec["wholesale_eur"] = p3.number_input(
                    "Wholesale €", value=rec["wholesale_eur"] or 0.0, step=1.0, key=f"we_{pid}"
                )
                rec["wholesale_usd"] = p4.number_input(
                    "Wholesale US$", value=rec["wholesale_usd"] or 0.0, step=1.0, key=f"wu_{pid}"
                )

selected_ids = [pid for pid in sorted_ids if state[pid]["selected"]]
highlighted_ids = [pid for pid in sorted_ids if state[pid]["highlighted"]]
bestseller_ids = [pid for pid in sorted_ids if state[pid]["bestseller"]]
st.caption(
    f"{len(selected_ids)} capi selezionati su {len(sorted_ids)} - "
    f"{len(highlighted_ids)} evidenziati in giallo - "
    f"{len(bestseller_ids)} best seller"
)

# ------------------------------------------------------------------
# 7. Generazione PDF
# ------------------------------------------------------------------
next_step = "3" if already_priced else "4"
st.subheader(f"{next_step}. Genera il PDF")

price_rows = {pid: state[pid] for pid in state}

if already_priced:
    gen_col1, gen_col2 = st.columns(2)

    with gen_col1:
        st.markdown("**PDF completo** (tutti i capi, layout originale, con le evidenziazioni scelte sopra)")
        if st.button("Genera PDF completo", type="primary"):
            result = core.build_priced_pdf(pdf_bytes, items, {}, [], highlighted_ids, bestseller_ids)
            st.download_button(
                "Scarica PDF completo",
                data=result,
                file_name="line_sheet_evidenziato.pdf",
                mime="application/pdf",
            )

    with gen_col2:
        st.markdown("**PDF selezione** (solo i capi spuntati sopra, con le evidenziazioni)")
        if st.button("Genera PDF selezione"):
            if not selected_ids:
                st.warning("Seleziona almeno un capo (punto 2).")
            else:
                result = core.build_selection_pdf_with_prices(
                    pdf_bytes, items, selected_ids, {}, [], highlighted_ids, bestseller_ids
                )
                st.download_button(
                    "Scarica PDF selezione",
                    data=result,
                    file_name="selezione_capi.pdf",
                    mime="application/pdf",
                )

else:
    gen_col1, gen_col2, gen_col3 = st.columns(3)

    with gen_col1:
        st.markdown("**PDF completo** (tutti i capi, layout originale, con i prezzi scelti sopra)")
        if st.button("Genera PDF completo con prezzi", type="primary"):
            if not price_fields:
                st.warning("Seleziona almeno un prezzo da inserire (punto 2).")
            else:
                result = core.build_priced_pdf(pdf_bytes, items, price_rows, price_fields, highlighted_ids, bestseller_ids)
                st.download_button(
                    "Scarica PDF completo",
                    data=result,
                    file_name="line_sheet_con_prezzi.pdf",
                    mime="application/pdf",
                )

    with gen_col2:
        st.markdown("**PDF selezione** (solo i capi spuntati sopra, con le evidenziazioni, senza prezzi aggiuntivi)")
        if st.button("Genera PDF selezione"):
            if not selected_ids:
                st.warning("Seleziona almeno un capo (punto 3).")
            else:
                result = core.build_selection_pdf_with_prices(
                    pdf_bytes, items, selected_ids, {}, [], highlighted_ids, bestseller_ids
                )
                st.download_button(
                    "Scarica PDF selezione",
                    data=result,
                    file_name="selezione_capi.pdf",
                    mime="application/pdf",
                )

    with gen_col3:
        st.markdown("**PDF selezione con prezzi** (solo i capi spuntati sopra, CON i prezzi scelti al punto 2)")
        if st.button("Genera PDF selezione con prezzi"):
            if not selected_ids:
                st.warning("Seleziona almeno un capo (punto 3).")
            elif not price_fields:
                st.warning("Seleziona almeno un prezzo da inserire (punto 2).")
            else:
                result = core.build_selection_pdf_with_prices(
                    pdf_bytes, items, selected_ids, price_rows, price_fields, highlighted_ids, bestseller_ids
                )
                st.download_button(
                    "Scarica PDF selezione con prezzi",
                    data=result,
                    file_name="selezione_capi_con_prezzi.pdf",
                    mime="application/pdf",
                )
