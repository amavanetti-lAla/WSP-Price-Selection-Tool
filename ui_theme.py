"""
ui_theme.py - aspetto condiviso con le altre app (Topsheet, Todo clienti):
sfondo crema, barra laterale colorata e pulsanti di navigazione colorati
in alto a destra.

Uso in app.py:
    import ui_theme
    ui_theme.apply_theme()
    ... ui_theme.nav_buttons(links) ...
    ui_theme.sidebar("WSP Price & Selection")
"""

import html

import streamlit as st

# Colori campionati dalle altre app (modificali qui per cambiare l'aspetto)
BG = "#f8f5ee"            # sfondo crema (uguale alle altre app)
SIDEBAR = "#dde8d8"       # barra laterale: verde WSP
ACCENT = "#6f8f6a"        # verde piu' scuro: linea schede, pulsanti principali
ACCENT_TEXT = "#4f6b4a"
UPLOADER = "#e7eee3"      # riquadro caricamento file

# Colori dei pulsanti di navigazione (gli stessi usati nelle altre app)
BTN_COLORS = {
    "wsp": "#dde8d8",       # verde
    "libretto": "#dceaf4",  # azzurro
    "topsheet": "#f4e6d3",  # pesca
    "todo": "#e6dff0",      # lavanda
}

_CSS = f"""
<style>
.stApp, [data-testid="stAppViewContainer"] {{ background: {BG}; }}
[data-testid="stHeader"] {{ background: transparent; }}
[data-testid="stSidebar"], [data-testid="stSidebar"] > div:first-child {{ background: {SIDEBAR}; }}
div[data-baseweb="tab-highlight"] {{ background-color: {ACCENT} !important; }}
button[data-baseweb="tab"][aria-selected="true"] {{ color: {ACCENT_TEXT} !important; }}
[data-testid="stFileUploaderDropzone"] {{ background: {UPLOADER}; }}
button[kind="primary"], [data-testid="stBaseButton-primary"],
[data-testid="stDownloadButton"] button[kind="primary"] {{
    background-color: {ACCENT} !important; border-color: {ACCENT} !important; color: #fff !important;
}}
.wsp-nav a {{
    display: block; text-align: center; padding: .5rem .8rem; margin-bottom: .55rem;
    border-radius: .5rem; border: 1px solid rgba(0,0,0,.06); color: #31333f !important;
    text-decoration: none !important; font-size: .85rem; transition: filter .15s;
}}
.wsp-nav a:hover {{ filter: brightness(.95); }}
</style>
"""


def apply_theme():
    st.markdown(_CSS, unsafe_allow_html=True)


def nav_buttons(links):
    """links: lista di (etichetta, url, chiave_colore). Le voci senza url vengono saltate."""
    parts = []
    for label, url, color in links:
        if not url:
            continue
        bg = BTN_COLORS.get(color, "#ffffff")
        parts.append(
            f'<a href="{html.escape(url)}" target="_blank" rel="noopener" '
            f'style="background:{bg}">{html.escape(label)} ↗</a>'
        )
    if parts:
        st.markdown('<div class="wsp-nav">' + "".join(parts) + "</div>", unsafe_allow_html=True)


def sidebar(title):
    with st.sidebar:
        st.markdown(f"### {title}")
        if st.button("Ricarica dati", key="wsp_reload", help="Svuota la cache e rilegge i file"):
            st.cache_data.clear()
            st.cache_resource.clear()
            st.rerun()
