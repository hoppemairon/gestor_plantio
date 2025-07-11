import streamlit as st
import uuid

# --- Funções Auxiliares ---
def inicializar_dados():
    """Inicializa os dicionários de plantios e receitas adicionais no session_state."""
    if 'plantios' not in st.session_state:
        st.session_state['plantios'] = {}
    if 'receitas_adicionais' not in st.session_state:
        st.session_state['receitas_adicionais'] = {}

def adicionar_plantio(ano, cultura, hectares, sacas_por_hectare, preco_saca):
    """Adiciona um novo plantio ao session_state."""
    plantio_id = str(uuid.uuid4())[:8]
    st.session_state['plantios'][plantio_id] = {
        'ano': ano,
        'cultura': cultura,
        'hectares': hectares,
        'sacas_por_hectare': sacas_por_hectare,
        'preco_saca': preco_saca,
        'tipo': 'Plantio'
    }
    return plantio_id

def adicionar_receita_adicional(nome, valor, categoria, anos_aplicacao):
    """Adiciona uma nova receita adicional ao session_state."""
    receita_id = str(uuid.uuid4())[:8]
    st.session_state['receitas_adicionais'][receita_id] = {
        'nome': nome,
        'valor': valor,
        'categoria': categoria,
        'anos_aplicacao': anos_aplicacao
    }
    return receita_id

def atualizar_plantio(pid, hectares, cultura, sacas_por_hectare, preco_saca):
    """Atualiza os dados de um plantio existente."""
    st.session_state['plantios'][pid].update({
        'hectares': hectares,
        'cultura': cultura,
        'sacas_por_hectare': sacas_por_hectare,
        'preco_saca': preco_saca
    })

def atualizar_receita_adicional(rid, nome, valor, categoria, anos_aplicacao):
    """Atualiza os dados de uma receita adicional existente."""
    st.session_state['receitas_adicionais'][rid].update({
        'nome': nome,
        'valor': valor,
        'categoria': categoria,
        'anos_aplicacao': anos_aplicacao
    })

def excluir_plantio(pid):
    """Exclui um plantio pelo ID."""
    del st.session_state['plantios'][pid]

def excluir_receita_adicional(rid):
    """Exclui uma receita adicional pelo ID."""
    del st.session_state['receitas_adicionais'][rid,]

# --- Inicialização ---
st.title("Cadastro de Plantio e Receitas 🌱")
inicializar_dados()

# --- Formulário de Cadastro de Plantio ---
st.markdown("### Adicionar Novo Plantio")
with st.form("form_plantio"):
    ano = st.number_input("Ano do plantio", min_value=2000, max_value=2100, step=1, value=2025)
    cultura = st.selectbox("Tipo de cultura", ["Soja", "Arroz", "Trigo", "Outros"])
    hectares = st.number_input("Área plantada (hectares)", min_value=0.1, step=0.1, value=1200.0)
    sacas_por_hectare = st.number_input("Produtividade (sacas/ha)", min_value=1.0, step=1.0, value=40.0)
    preco_saca = st.number_input("Valor da saca (R$)", min_value=0.5, step=0.5, value=120.0)
    submitted = st.form_submit_button("Cadastrar Plantio")

    if submitted:
        plantio_id = adicionar_plantio(ano, cultura, hectares, sacas_por_hectare, preco_saca)
        st.success(f"Plantio cadastrado com sucesso! (ID: {plantio_id})")

# --- Formulário de Cadastro de Receitas Adicionais ---
st.markdown("### Adicionar Nova Receita Adicional")
with st.form("form_receita_adicional"):
    nome_receita = st.text_input("Nome da Receita (ex: Venda de Gado, Empréstimo)", value="")
    valor_receita = st.number_input("Valor Anual (R$)", min_value=0.0, step=100.0, value=0.0)
    categoria_receita = st.selectbox("Categoria", ["Operacional", "Extra Operacional"])
    anos_disponiveis = [f"Ano {i+1}" for i in range(5)]
    anos_aplicacao = st.multiselect("Anos de Aplicação", anos_disponiveis, default=anos_disponiveis)
    submitted_receita = st.form_submit_button("Cadastrar Receita")

    if submitted_receita:
        if not nome_receita or valor_receita <= 0 or not anos_aplicacao:
            st.warning("Preencha todos os campos corretamente.")
        else:
            receita_id = adicionar_receita_adicional(nome_receita, valor_receita, categoria_receita, anos_aplicacao)
            st.success(f"Receita adicional cadastrada com sucesso! (ID: {receita_id})")

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

            col1, col2 = st.columns([1, 1])
            with col1:
                if st.button("💾 Salvar alterações", key=f"save_{pid}"):
                    atualizar_plantio(pid, novo_hectares, nova_cultura, nova_sacas, novo_preco)
                    st.success(f"Plantio {pid} atualizado com sucesso!")

            with col2:
                if st.button("🗑️ Excluir", key=f"delete_{pid}"):
                    excluir_plantio(pid)
                    st.success(f"Plantio {pid} excluído.")
                    st.rerun()

# --- Visualização e Edição de Receitas Adicionais ---
if st.session_state['receitas_adicionais']:
    st.markdown("### Receitas Adicionais Cadastradas")
    for rid, dados in st.session_state['receitas_adicionais'].items():
        with st.expander(f"{dados['nome']} ({dados['categoria']})"):
            col1, col2 = st.columns(2)

            with col1:
                novo_nome = st.text_input("Nome da Receita", value=dados['nome'], key=f"nome_{rid}")
                novo_valor = st.number_input(
                    "Valor Anual (R$)", min_value=0.0, step=100.0, value=dados['valor'], key=f"valor_{rid}"
                )

            with col2:
                nova_categoria = st.selectbox(
                    "Categoria",
                    ["Operacional", "Extra Operacional"],
                    index=["Operacional", "Extra Operacional"].index(dados['categoria']),
                    key=f"cat_{rid}"
                )
                novos_anos = st.multiselect(
                    "Anos de Aplicação",
                    anos_disponiveis,
                    default=dados['anos_aplicacao'],
                    key=f"anos_{rid}"
                )

            col1, col2 = st.columns([1, 1])
            with col1:
                if st.button("💾 Salvar alterações", key=f"save_rec_{rid}"):
                    atualizar_receita_adicional(rid, novo_nome, novo_valor, nova_categoria, novos_anos)
                    st.success(f"Receita {rid} atualizada com sucesso!")

            with col2:
                if st.button("🗑️ Excluir", key=f"delete_rec_{rid}"):
                    excluir_receita_adicional(rid)
                    st.success(f"Receita {rid} excluída.")
                    st.rerun()

# --- Botão "Limpar Tudo" ---
st.markdown("---")
st.markdown("### ⚙️ Ações Gerais")
if st.button("Limpar Todos os Plantios e Receitas"):
    st.session_state['plantios'] = {}
    st.session_state['receitas_adicionais'] = {}
    st.success("Todos os plantios e receitas foram removidos!")
    st.rerun()