import streamlit as st
import uuid


st.title("Cadastro de Plantio")

# Inicializa estrutura se não existir
if 'plantios' not in st.session_state:
    st.session_state['plantios'] = {}

with st.form("form_plantio"):
    ano = st.number_input("Ano do plantio", min_value=2000, max_value=2100, step=1, value=2025)
    cultura = st.selectbox("Tipo de cultura", ["Soja", "Arroz", "Trigo", "Outros"])
    hectares = st.number_input("Área plantada (hectares)", min_value=0.0, step=0.1)
    sacas_por_hectare = st.number_input("Produtividade (sacas/ha)", min_value=0.0, step=1.0)
    preco_saca = st.number_input("Valor da saca (R$)", min_value=0.0, step=0.5)
    
    submitted = st.form_submit_button("Cadastrar Plantio")

    if submitted:
        plantio_id = str(uuid.uuid4())[:8]
        st.session_state['plantios'][plantio_id] = {
            'ano': ano,
            'cultura': cultura,
            'hectares': hectares,
            'sacas_por_hectare': sacas_por_hectare,
            'preco_saca': preco_saca
        }
        st.success(f"Plantio cadastrado com sucesso (ID: {plantio_id})")

# Visualização dos cadastros
# Visualização dos cadastros com opção de editar/excluir
if st.session_state['plantios']:
    st.markdown("### Plantios cadastrados")

    for pid in list(st.session_state['plantios'].keys()):
        dados = st.session_state['plantios'][pid]

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

            if st.button("💾 Salvar alterações", key=f"save_{pid}"):
                st.session_state['plantios'][pid] = {
                    'ano': dados['ano'],
                    'cultura': nova_cultura,
                    'hectares': novo_hectares,
                    'sacas_por_hectare': nova_sacas,
                    'preco_saca': novo_preco
                }
                st.success(f"Plantio {pid} atualizado!")

            if st.button("🗑️ Excluir", key=f"delete_{pid}"):
                del st.session_state['plantios'][pid]
                st.success(f"Plantio {pid} excluído.")
                st.experimental_rerun()

# --- Botão "Limpar Tudo" ---
st.markdown("---")
st.markdown("### ⚙️ Ações Gerais")
if st.button("Limpar Todos os Plantios", key="btn_clear_all_plantios"):
    st.session_state['plantios'] = {} # Limpa o dicionário de plantios
    st.success("Todos os plantios foram removidos!")
    st.rerun() # Força o re-render para atualizar a interface