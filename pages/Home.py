from pathlib import Path

import streamlit as st

# ==========================================================
# DINOPETANIMAL PRO
# PÁGINA: HOME
# ==========================================================


# ==========================================================
# CAMINHOS
# ==========================================================

BASE_DIR = Path(__file__).resolve().parent.parent
PASTA_PETS = BASE_DIR / "pets"
PASTA_RELATORIOS = BASE_DIR / "relatorios"
ARQUIVO_CSS = BASE_DIR / "styles.css"
ARQUIVO_BANNER = BASE_DIR / "animal.png"
MODELO_PET = BASE_DIR / "modelo_pet.xlsx"

PASTA_PETS.mkdir(parents=True, exist_ok=True)
PASTA_RELATORIOS.mkdir(parents=True, exist_ok=True)

# ==========================================================
# CSS GLOBAL
# ==========================================================

if ARQUIVO_CSS.exists() and ARQUIVO_CSS.stat().st_size > 0:
    st.markdown(
        f"<style>{ARQUIVO_CSS.read_text(encoding='utf-8')}</style>",
        unsafe_allow_html=True,
    )

st.markdown(
    """
    <style>
    .home-topo {
        background: linear-gradient(135deg, #fff7ed, #ffedd5);
        border: 1px solid #fdba74;
        border-radius: 20px;
        padding: 24px 28px;
        margin-bottom: 22px;
    }

    .home-topo h1 {
        margin: 0;
        color: #1f2937;
        font-size: 2.35rem;
    }

    .home-topo p {
        margin: 8px 0 0 0;
        color: #475569;
        font-size: 1.08rem;
    }

    .menu-aviso {
        background: linear-gradient(90deg, #7c2d12, #c2410c);
        border: 1px solid #fb923c;
        border-radius: 16px;
        padding: 18px 22px;
        margin-bottom: 22px;
        box-shadow: 0 4px 16px rgba(124, 45, 18, .18);
    }

    .menu-aviso h3 {
        color: white;
        margin: 0 0 6px 0;
    }

    .menu-aviso p {
        color: white;
        font-size: 1.02rem;
        margin: 0;
    }

    .atalho-card {
        background: white;
        border: 1px solid #fed7aa;
        border-radius: 16px;
        padding: 18px 20px;
        min-height: 145px;
        box-shadow: 0 2px 10px rgba(0, 0, 0, .04);
        margin-bottom: 10px;
    }

    .atalho-card h3 {
        color: #7c2d12;
        margin: 0 0 8px 0;
    }

    .atalho-card p {
        color: #475569;
        margin: 0;
    }

    .status-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 14px;
        padding: 16px 18px;
        height: 100%;
    }

    .stButton > button {
        border-radius: 10px;
        min-height: 44px;
        font-weight: 700;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ==========================================================
# DADOS DO SISTEMA
# ==========================================================

try:
    from excel import resumo_sistema, sistema_saudavel, sobre

    resumo = resumo_sistema()
    sistema_ok = sistema_saudavel()
    informacoes = sobre()

except Exception as erro:
    resumo = {
        "pets": 0,
        "ultimo_pet": "-",
        "especies": 0,
        "tutores": 0,
        "relatorios": 0,
        "peso_medio": 0,
    }
    sistema_ok = False
    informacoes = {
        "sistema": "DinoPetAnimal PRO",
        "versao": "1.0",
        "autor": "Professor Dinaldo Guedes",
    }
    erro_sistema = erro
else:
    erro_sistema = None

# ==========================================================
# BANNER
# ==========================================================

if ARQUIVO_BANNER.exists():
    st.image(str(ARQUIVO_BANNER), use_container_width=True)

st.markdown(
    """
    <div class="home-topo">
        <h1>🏠 Bem-vindo ao DinoPetAnimal PRO</h1>
        <p>Sistema inteligente para cadastro, avaliação, acompanhamento e relatórios veterinários.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="menu-aviso">
        <h3>📱 MENU DO SISTEMA</h3>
        <p>Use o menu lateral esquerdo para acessar Cadastro, Pets, Avaliação, Dashboard e Relatórios.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ==========================================================
# STATUS
# ==========================================================

st.subheader("🏥 Status do sistema")

if sistema_ok:
    st.success("✅ Sistema carregado e estrutura principal verificada.")
else:
    st.warning("⚠️ O sistema foi aberto, mas existem itens que precisam ser verificados.")

if erro_sistema is not None:
    st.error(f"Não foi possível carregar todas as informações: {erro_sistema}")

# ==========================================================
# INDICADORES
# ==========================================================

st.subheader("📊 Resumo geral")

col1, col2, col3, col4, col5 = st.columns(5)

with col1:
    st.metric("🐾 Pets", resumo.get("pets", 0))

with col2:
    st.metric("🦴 Espécies", resumo.get("especies", 0))

with col3:
    st.metric("👤 Tutores", resumo.get("tutores", 0))

with col4:
    st.metric("📄 Relatórios", resumo.get("relatorios", 0))

with col5:
    st.metric(
        "⚖️ Peso médio",
        f"{resumo.get('peso_medio', 0)} kg",
    )

col_info1, col_info2 = st.columns(2)

with col_info1:
    st.info(
        f"🐾 Último pet cadastrado: "
        f"**{resumo.get('ultimo_pet', '-')}**"
    )

with col_info2:
    quantidade_planilhas = len(
        [
            arquivo
            for arquivo in PASTA_PETS.glob("*.xlsx")
            if not arquivo.name.startswith("~$")
        ]
    )
    st.info(
        f"📂 Planilhas individuais encontradas: "
        f"**{quantidade_planilhas}**"
    )

st.divider()

# ==========================================================
# ACESSOS RÁPIDOS
# ==========================================================

st.subheader("🚀 Acesso rápido")

col_a1, col_a2, col_a3 = st.columns(3)

with col_a1:
    st.markdown(
        """
        <div class="atalho-card">
            <h3>➕ Cadastro</h3>
            <p>Cadastre um novo pet, os dados do tutor e o profissional responsável.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if st.button(
        "Abrir Cadastro",
        key="home_cadastro",
        use_container_width=True,
    ):
        st.switch_page("pages/Cadastro.py")

with col_a2:
    st.markdown(
        """
        <div class="atalho-card">
            <h3>🐾 Pets</h3>
            <p>Consulte, edite, baixe ou exclua os cadastros existentes.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if st.button(
        "Abrir Pets",
        key="home_pets",
        use_container_width=True,
    ):
        st.switch_page("pages/Pets.py")

with col_a3:
    st.markdown(
        """
        <div class="atalho-card">
            <h3>🩺 Avaliação</h3>
            <p>Registre as notas por domínio e acompanhe o resultado clínico.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if st.button(
        "Abrir Avaliação",
        key="home_avaliacao",
        use_container_width=True,
    ):
        st.switch_page("pages/Avaliacao.py")

col_a4, col_a5 = st.columns(2)

with col_a4:
    st.markdown(
        """
        <div class="atalho-card">
            <h3>📊 Dashboard</h3>
            <p>Visualize indicadores, classificações, gráficos e desempenho dos pets.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if st.button(
        "Abrir Dashboard",
        key="home_dashboard",
        use_container_width=True,
    ):
        st.switch_page("pages/Dashboard.py")

with col_a5:
    st.markdown(
        """
        <div class="atalho-card">
            <h3>📄 Relatórios</h3>
            <p>Gere e baixe relatórios veterinários individuais em PDF.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if st.button(
        "Abrir Relatórios",
        key="home_relatorios",
        use_container_width=True,
    ):
        st.switch_page("pages/Relatorios.py")

st.divider()

# ==========================================================
# VERIFICAÇÃO DA ESTRUTURA
# ==========================================================

st.subheader("🔎 Verificação da estrutura")

verificacoes = {
    "Arquivo modelo_pet.xlsx": MODELO_PET.exists(),
    "Imagem animal.png": ARQUIVO_BANNER.exists(),
    "Pasta pets": PASTA_PETS.exists(),
    "Pasta relatorios": PASTA_RELATORIOS.exists(),
    "Arquivo styles.css": ARQUIVO_CSS.exists(),
}

col_v1, col_v2 = st.columns(2)

itens = list(verificacoes.items())

with col_v1:
    for nome, existe in itens[:3]:
        st.markdown(
            f'<div class="status-card">{"✅" if existe else "❌"} '
            f"<b>{nome}</b></div>",
            unsafe_allow_html=True,
        )
        st.write("")

with col_v2:
    for nome, existe in itens[3:]:
        st.markdown(
            f'<div class="status-card">{"✅" if existe else "❌"} '
            f"<b>{nome}</b></div>",
            unsafe_allow_html=True,
        )
        st.write("")

if not ARQUIVO_BANNER.exists():
    st.caption(
        "Coloque a imagem enviada com o nome animal.png na mesma pasta do app.py."
    )

# ==========================================================
# RODAPÉ
# ==========================================================

st.divider()

st.caption(
    f"{informacoes.get('sistema', 'DinoPetAnimal PRO')} • "
    f"Versão {informacoes.get('versao', '1.0')} • "
    f"{informacoes.get('autor', 'Professor Dinaldo Guedes')}"
)
