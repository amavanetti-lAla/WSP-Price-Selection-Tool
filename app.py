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

import io

import streamlit as st

import core

st.set_page_config(page_title="WSP Price & Selection Tool", layout="wide")

st.title("WSP Price & Selection Tool")
st.caption("Carica il line sheet PDF e il report prezzi Excel per generare il PDF finale.")

# ------------------------------------------------------------------
# 1. Caricamento file
# ------------------------------------------------------------------
col_up1, col_up2 = st.columns(2)
with col_up1:
    pdf_file = st.file_uploader("Line sheet PDF", type=["pdf"])
with col_up2:
    xls_file = st.file_uploader("Report prezzi Excel (.xls)", type=["xls"])

if not pdf_file or not xls_file:
    st.info("Carica entrambi i file per continuare.")
    st.stop()

pdf_bytes = pdf_file.read()
xls_bytes = xls_file.read()

# ------------------------------------------------------------------
# 2. Analisi file (con cache per evitare di rifare il lavoro ad ogni click)
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
price_records = _load_prices(xls_bytes)
doc = _open_doc(pdf_bytes)

price_by_id = {r["product_id"]: r for r in price_records}

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
# 3. Stato modificabile dei prezzi (inizializzato una sola volta)
# ------------------------------------------------------------------
if "price_state" not in st.session_state:
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
        }
    st.session_state.price_state = state

state = st.session_state.price_state

# ------------------------------------------------------------------
# 4. Opzioni di stampa (quali prezzi inserire nel PDF)
# ------------------------------------------------------------------
st.subheader("1. Quali prezzi vuoi inserire nel PDF?")
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
# 5. Elenco capi con anteprima, prezzi modificabili, selezione
# ------------------------------------------------------------------
st.subheader("2. Controlla / modifica i prezzi e scegli i capi")

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
            st.image(thumb, use_container_width=True)
        with info_col:
            top_row = st.columns([1, 3])
            with top_row[0]:
                rec["selected"] = st.checkbox("Includi", value=rec["selected"], key=f"sel_{pid}")
            with top_row[1]:
                st.markdown(f"**{desc or '(senza descrizione)'}**  \n`{pid}`")

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
st.caption(f"{len(selected_ids)} capi selezionati su {len(sorted_ids)}")

# ------------------------------------------------------------------
# 6. Generazione PDF
# ------------------------------------------------------------------
st.subheader("3. Genera il PDF")

price_rows = {pid: state[pid] for pid in state}

gen_col1, gen_col2, gen_col3 = st.columns(3)

with gen_col1:
    st.markdown("**PDF completo** (tutti i capi, layout originale, con i prezzi scelti sopra)")
    if st.button("Genera PDF completo con prezzi", type="primary"):
        if not price_fields:
            st.warning("Seleziona almeno un prezzo da inserire (punto 1).")
        else:
            result = core.build_priced_pdf(pdf_bytes, items, price_rows, price_fields)
            st.download_button(
                "Scarica PDF completo",
                data=result,
                file_name="line_sheet_con_prezzi.pdf",
                mime="application/pdf",
            )

with gen_col2:
    st.markdown("**PDF selezione** (solo i capi spuntati sopra, foto + testo, senza prezzi aggiuntivi)")
    if st.button("Genera PDF selezione"):
        if not selected_ids:
            st.warning("Seleziona almeno un capo (punto 2).")
        else:
            result = core.build_selection_pdf(pdf_bytes, items, selected_ids)
            st.download_button(
                "Scarica PDF selezione",
                data=result,
                file_name="selezione_capi.pdf",
                mime="application/pdf",
            )

with gen_col3:
    st.markdown("**PDF selezione con prezzi** (solo i capi spuntati sopra, CON i prezzi scelti al punto 1)")
    if st.button("Genera PDF selezione con prezzi"):
        if not selected_ids:
            st.warning("Seleziona almeno un capo (punto 2).")
        elif not price_fields:
            st.warning("Seleziona almeno un prezzo da inserire (punto 1).")
        else:
            result = core.build_selection_pdf_with_prices(
                pdf_bytes, items, selected_ids, price_rows, price_fields
            )
            st.download_button(
                "Scarica PDF selezione con prezzi",
                data=result,
                file_name="selezione_capi_con_prezzi.pdf",
                mime="application/pdf",
            )
