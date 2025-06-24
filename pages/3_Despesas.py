import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(layout="wide", page_title="Planejamento de Despesas")

st.title("Planejamento de Despesas")

anos = [f"Ano {i+1}" for i in range(5)]
inflacao_padrao = 0.04  # 4%

# --- INFLAÇÃO ---
st.markdown("### 📈 Taxas de Inflação Anual")
col1, col2 = st.columns(2)
inflacoes = []
for i, ano in enumerate(anos):
    with (col1 if i % 2 == 0 else col2):
        # Usar session_state.get para persistir o valor da inflação entre reruns
        valor = st.number_input(
            f"Inflação estimada para {ano} (%)",
            min_value=0.0,
            value=st.session_state.get(f"inf_{i}", inflacao_padrao * 100),
            key=f"inf_{i}"
        )
        inflacoes.append(valor)

# --- GERENCIAMENTO DE DESPESAS ---
st.markdown("### 💰 Gerenciamento de Despesas")

# Inicializa o estado das despesas e do modo de edição
if 'despesas' not in st.session_state:
    st.session_state['despesas'] = []
if 'editing_expense_index' not in st.session_state:
    st.session_state['editing_expense_index'] = None

# Determina se estamos no modo de edição
is_editing = st.session_state['editing_expense_index'] is not None
expense_to_edit = None
if is_editing:
    expense_to_edit = st.session_state['despesas'][st.session_state['editing_expense_index']]

# Formulário para adicionar/editar despesas
# clear_on_submit=True apenas se não estiver editando
with st.form("form_despesa", clear_on_submit=not is_editing):
    st.subheader(f"{'Editar' if is_editing else 'Adicionar'} Despesa")

    # Preenche os campos com os valores da despesa sendo editada, ou vazios para nova despesa
    default_nome = expense_to_edit['Despesa'] if is_editing else ""
    default_valor = expense_to_edit['Valor'] if is_editing else 0.0
    # Encontra o índice da categoria padrão para o selectbox
    categorias_disponiveis = ["Operacional", "RH", "Administrativa", "Extra Operacional", "Dividendos", "Impostos"]
    default_categoria_index = categorias_disponiveis.index(default_categoria) if (default_categoria := expense_to_edit['Categoria'] if is_editing else "Operacional") in categorias_disponiveis else 0

    nome_despesa = st.text_input("Nome da Despesa", value=default_nome, key="nome_despesa_input")
    valor_despesa = st.number_input("Valor Anual (R$)", min_value=0.0, step=100.0, value=default_valor, key="valor_despesa_input")
    categoria = st.selectbox(
        "Categoria",
        categorias_disponiveis,
        index=default_categoria_index,
        key="categoria_select"
    )

    col_buttons = st.columns([1, 1, 4]) # Colunas para os botões do formulário

    with col_buttons[0]:
        submit_button = st.form_submit_button(
            f"{'Atualizar' if is_editing else 'Adicionar'} Despesa"
        )
    with col_buttons[1]:
        if is_editing:
            cancel_button = st.form_submit_button("Cancelar Edição")

    # Lógica de submissão do formulário
    if submit_button:
        if not nome_despesa or valor_despesa <= 0:
            st.warning("Preencha todos os campos corretamente (Nome e Valor > 0).")
        else:
            new_expense_data = {
                "Despesa": nome_despesa.strip(),
                "Valor": valor_despesa,
                "Categoria": categoria
            }
            if is_editing:
                st.session_state['despesas'][st.session_state['editing_expense_index']] = new_expense_data
                st.session_state['editing_expense_index'] = None # Sai do modo de edição
                st.success("Despesa atualizada com sucesso!")
            else:
                st.session_state['despesas'].append(new_expense_data)
                st.success("Despesa adicionada com sucesso!")
            st.rerun() # Força o re-render para atualizar a lista e limpar o formulário

    # Lógica do botão Cancelar Edição
    if is_editing and cancel_button:
        st.session_state['editing_expense_index'] = None
        st.rerun() # Força o re-render para sair do modo de edição

# Exibição das despesas cadastradas com botões de ação
st.markdown("### Despesas Cadastradas")
with st.expander("Despesas Cadastradas"):
    if not st.session_state['despesas']:
        st.info("Nenhuma despesa cadastrada ainda.")
    else:
        # Cabeçalho da tabela de despesas
        header_cols = st.columns([3, 2, 2, 1, 1])
        header_cols[0].write("**Despesa**")
        header_cols[1].write("**Valor Anual**")
        header_cols[2].write("**Categoria**")
        header_cols[3].write("") # Coluna para botões
        header_cols[4].write("") # Coluna para botões

        # Itera sobre as despesas para exibi-las com botões
        for i, expense in enumerate(st.session_state['despesas']):
            cols = st.columns([3, 2, 2, 1, 1]) # Colunas para cada linha de despesa
            cols[0].write(expense["Despesa"])
            cols[1].write(f"R$ {expense['Valor']:,.2f}".replace(",", "v").replace(".", ",").replace("v", "."))
            cols[2].write(expense["Categoria"])

            with cols[3]:
                # Botão Editar - define o índice da despesa a ser editada e re-renderiza
                if st.button("Editar", key=f"edit_{i}"):
                    st.session_state['editing_expense_index'] = i
                    st.rerun()
            with cols[4]:
                # Botão Excluir - remove a despesa da lista e re-renderiza
                if st.button("Excluir", key=f"delete_{i}"):
                    st.session_state['despesas'].pop(i)
                    st.rerun()

# --- CÁLCULO E EXIBIÇÃO DO FLUXO DE DESPESAS ---
st.markdown("---")
st.markdown("### 📊 Projeção de Despesas com Inflação")

if not st.session_state['despesas']:
    st.info("Adicione despesas para ver a projeção.")
else:
    # Cria um DataFrame temporário a partir das despesas do session_state
    df_despesas_calc = pd.DataFrame(st.session_state['despesas'])

    # Normaliza o nome da despesa e agrupa para somar valores duplicados
    df_despesas_calc['Despesa_Normalized'] = df_despesas_calc['Despesa'].astype(str).str.strip()
    # Se quiser agrupar ignorando maiúsculas/minúsculas, descomente a linha abaixo:
    # df_despesas_calc['Despesa_Normalized'] = df_despesas_calc['Despesa_Normalized'].str.lower()

    df_despesas_grouped = df_despesas_calc.groupby('Despesa_Normalized')['Valor'].sum().reset_index()
    df_despesas_grouped.set_index('Despesa_Normalized', inplace=True)

    fluxo = {}
    for i, ano in enumerate(anos):
        fator = np.prod([1 + inflacoes[j] / 100 for j in range(i + 1)])
        fluxo[ano] = df_despesas_grouped["Valor"] * fator

    df_fluxo = pd.DataFrame(fluxo)
    df_fluxo.index.name = "Despesa" # Renomeia o índice para clareza na exibição

    # Formatação brasileira
    def format_brl(valor):
        return f"R$ {valor:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")

    st.dataframe(df_fluxo.style.format(format_brl))

    # Guardar no session_state para outras páginas
    st.session_state['fluxo_caixa'] = df_fluxo

# --- BOTÃO "LIMPAR TUDO" ---
st.markdown("### ⚙️ Ações Gerais")
if st.button("Limpar Tudo", key="btn_clear_all"):
    st.session_state['despesas'] = []  # Limpa todas as despesas
    st.session_state['fluxo_caixa'] = None  # Limpa o fluxo de caixa
    st.session_state['editing_expense_index'] = None  # Sai do modo de edição
    st.success("Todos os dados foram limpos!")
    st.rerun()  # Força o re-render para atualizar a interface