from pathlib import Path

import streamlit as st

# ==========================================================
# DINOPETANIMAL PRO
# ENTRADA PRINCIPAL E NAVEGAÇÃO
# ==========================================================

st.set_page_config(
    page_title="DinoPetAnimal PRO",
    page_icon="🐾",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE_DIR = Path(__file__).resolve().parent
PASTA_PETS = BASE_DIR / "pets"
PASTA_RELATORIOS = BASE_DIR / "relatorios"
ARQUIVO_CSS = BASE_DIR / "styles.css"
ARQUIVO_LOGO = BASE_DIR / "animal.png"

PASTA_PETS.mkdir(parents=True, exist_ok=True)
PASTA_RELATORIOS.mkdir(parents=True, exist_ok=True)

# CSS global: carregado uma única vez para todas as páginas.
if ARQUIVO_CSS.exists() and ARQUIVO_CSS.stat().st_size > 0:
    st.markdown(
        f"<style>{ARQUIVO_CSS.read_text(encoding='utf-8')}</style>",
        unsafe_allow_html=True,
    )

# Identidade visual no topo da barra lateral.
with st.sidebar:
    if ARQUIVO_LOGO.exists():
        st.image(str(ARQUIVO_LOGO), use_container_width=True)
    else:
        st.markdown(
            """
            <div class="sidebar-brand">
                <div class="sidebar-brand-icon">🐾</div>
                <div>
                    <div class="sidebar-brand-title">DinoPetAnimal</div>
                    <div class="sidebar-brand-pro">PRO</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

# Ordem profissional seguindo o fluxo de uso do sistema.
paginas = [
    st.Page("pages/Home.py", title="Início", icon="🏠", default=True),
    st.Page("pages/Cadastro.py", title="Cadastrar Pet", icon="➕"),
    st.Page("pages/Pets.py", title="Meus Pets", icon="🐾"),
    st.Page("pages/Avaliacao.py", title="Avaliações", icon="🩺"),
    st.Page("pages/Relatorios.py", title="Relatórios", icon="📄"),
    st.Page("pages/Dashboard.py", title="Dashboard", icon="📊"),
]

navegacao = st.navigation(paginas, position="sidebar")

with st.sidebar:
    st.markdown(
        """
        <div class="sidebar-footer">
            <strong>DinoPetAnimal PRO</strong><br>
            Versão 2.0<br>
            Professor Dinaldo Guedes
        </div>
        """,
        unsafe_allow_html=True,
    )

navegacao.run()
