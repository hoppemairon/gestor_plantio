import streamlit as st
import uuid

# --- Funções Auxiliares ---
def inicializar_plantios():
    """Inicializa o dicionário de plantios no session_state."""
    if 'plantios' not in st.session_state:
        st.session_state['plantios'] = {}

def adicionar_plantio(ano, cultura, hectares, sacas_por_hectare, preco_saca):
    """Adiciona um novo plantio ao session_state."""
    plantio_id = str(uuid.uuid4())[:8]
    st.session_state['plantios'][plantio_id] = {
        'ano': ano,
        'cultura': cultura,
        'hectares': hectares,
        'sacas_por_hectare': sacas_por_hectare,
        'preco_saca': preco_saca
    }
    return plantio_id

def atualizar_plantio(pid, hectares, cultura, sacas_por_hectare, preco_saca):
    """Atualiza os dados de um plantio existente."""
    st.session_state['plantios'][pid].update({
        'hectares': hectares,
        'cultura': cultura,
        'sacas_por_hectare': sacas_por_hectare,
        'preco_saca': preco_saca
    })

def excluir_plantio(pid):
    """Exclui um plantio pelo ID."""
    del st.session_state['plantios'][pid]

# --- Inicialização ---
st.title("Cadastro de Plantio 🌱")
inicializar_plantios()

# --- Formulário de Cadastro ---
st.markdown("### Adicionar Novo Plantio")
with st.form("form_plantio"):
    ano = st.number_input("Ano do plantio", min_value=2000, max_value=2100, step=1, value=2025)
    cultura = st.selectbox("Tipo de cultura", ["Soja", "Arroz", "Trigo", "Outros"])
    hectares = st.number_input("Área plantada (hectares)", min_value=0.1, step=0.1)
    sacas_por_hectare = st.number_input("Produtividade (sacas/ha)", min_value=1.0, step=1.0)
    preco_saca = st.number_input("Valor da saca (R$)", min_value=0.5, step=0.5)
    submitted = st.form_submit_button("Cadastrar Plantio")

    if submitted:
        plantio_id = adicionar_plantio(ano, cultura, hectares, sacas_por_hectare, preco_saca)
        st.success(f"Plantio cadastrado com sucesso! (ID: {plantio_id})")

# --- Visualização e Edição de Plantios ---
if st.session_state['plantios']:
    st.markdown("### Plantios Cadastrados")
    for pid, dados in st.session_state['plantios'].items():
        with st.expander(f"{dados['ano']} - {dados['cultura']}"):
            col1, col2 = st.columns(2)

            with col1:
                novo_hectares = st.number_input(
                    f"Área (ha)", value=dados['hectares'], key=f"ha_{pid}"
                )
                nova_cultura = st.selectbox(
                    f"Cultura",
                    ["Soja", "Arroz", "Trigo", "Outros"],
                    index=["Soja", "Arroz", "Trigo", "Outros"].index(dados['cultura']),
                    key=f"cult_{pid}"
                )

            with col2:
                nova_sacas = st.number_input(
                    f"Sacas/ha", value=dados['sacas_por_hectare'], key=f"sph_{pid}"
                )
                novo_preco = st.number_input(
                    f"Preço saca (R$)", value=dados['preco_saca'], key=f"ps_{pid}"
                )

            # Botões de ação
            col1, col2 = st.columns([1, 1])
            with col1:
                if st.button("💾 Salvar alterações", key=f"save_{pid}"):
                    atualizar_plantio(pid, novo_hectares, nova_cultura, nova_sacas, novo_preco)
                    st.success(f"Plantio {pid} atualizado com sucesso!")

            with col2:
                if st.button("🗑️ Excluir", key=f"delete_{pid}"):
                    excluir_plantio(pid)
                    st.success(f"Plantio {pid} excluído.")
                    st.experimental_rerun()

# --- Botão "Limpar Tudo" ---
st.markdown("---")
st.markdown("### ⚙️ Ações Gerais")
if st.button("Limpar Todos os Plantios"):
    st.session_state['plantios'] = {}
    st.success("Todos os plantios foram removidos!")
    st.experimental_rerun()