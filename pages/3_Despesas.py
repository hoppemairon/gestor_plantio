import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from dateutil.relativedelta import relativedelta

st.set_page_config(layout="wide", page_title="Planejamento de Despesas")
st.title("Planejamento de Despesas")

anos = [f"Ano {i+1}" for i in range(5)]
inflacao_padrao = 0.04

# --- INFLAÇÃO ---
st.markdown("### 📈 Taxas de Inflação Anual")
col1, col2 = st.columns(2)
inflacoes = []
for i, ano in enumerate(anos):
    with (col1 if i % 2 == 0 else col2):
        valor = st.number_input(
            f"Inflação estimada para {ano} (%)",
            min_value=0.0,
            value=st.session_state.get(f"inf_{i}", inflacao_padrao * 100),
            key=f"inf_{i}"
        )
        inflacoes.append(valor)

# --- DESPESAS ---
st.markdown("### 💰 Gerenciamento de Despesas")

if 'despesas' not in st.session_state:
    st.session_state['despesas'] = []
if 'editing_expense_index' not in st.session_state:
    st.session_state['editing_expense_index'] = None

is_editing = st.session_state['editing_expense_index'] is not None
expense_to_edit = st.session_state['despesas'][st.session_state['editing_expense_index']] if is_editing else None

with st.form("form_despesa", clear_on_submit=not is_editing):
    st.subheader(f"{'Editar' if is_editing else 'Adicionar'} Despesa")

    default_nome = expense_to_edit['Despesa'] if is_editing else ""
    default_valor = expense_to_edit['Valor'] if is_editing else 0.0
    categorias = ["Operacional", "RH", "Administrativa", "Extra Operacional", "Dividendos", "Impostos"]
    default_categoria_index = categorias.index(
        expense_to_edit['Categoria']) if is_editing and expense_to_edit['Categoria'] in categorias else 0

    nome = st.text_input("Nome da Despesa", value=default_nome, key="nome_despesa_input")
    valor = st.number_input("Valor Anual (R$)", min_value=0.0, step=100.0, value=default_valor, key="valor_despesa_input")
    categoria = st.selectbox("Categoria", categorias, index=default_categoria_index, key="categoria_select")

    col_buttons = st.columns([1, 1, 4])
    with col_buttons[0]:
        submit = st.form_submit_button("Atualizar" if is_editing else "Adicionar")
    with col_buttons[1]:
        if is_editing:
            cancel = st.form_submit_button("Cancelar Edição")

    if submit:
        if not nome or valor <= 0:
            st.warning("Preencha todos os campos corretamente.")
        else:
            nova_despesa = {"Despesa": nome.strip(), "Valor": valor, "Categoria": categoria}
            if is_editing:
                st.session_state['despesas'][st.session_state['editing_expense_index']] = nova_despesa
                st.session_state['editing_expense_index'] = None
            else:
                st.session_state['despesas'].append(nova_despesa)
            st.rerun()

    if is_editing and cancel:
        st.session_state['editing_expense_index'] = None
        st.rerun()

# --- EMPRÉSTIMOS ---
st.markdown("### 🏦 Cadastro de Empréstimos e Financiamentos")

def format_brl(valor):
    return f"R$ {valor:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")

if "emprestimos" not in st.session_state:
    st.session_state["emprestimos"] = []
if "editing_loan_index" not in st.session_state:
    st.session_state["editing_loan_index"] = None

editing_loan = st.session_state["editing_loan_index"] is not None
emprestimo_to_edit = st.session_state["emprestimos"][st.session_state["editing_loan_index"]] if editing_loan else {}

with st.form("form_emprestimo", clear_on_submit=not editing_loan):
    st.subheader(f"{'Editar' if editing_loan else 'Cadastrar'} Empréstimo")

    col1, col2 = st.columns(2)
    with col1:
        banco = st.text_input("Banco", value=emprestimo_to_edit.get("banco", ""))
        titular = st.text_input("Titular", value=emprestimo_to_edit.get("titular", ""))
        contrato = st.text_input("Contrato Número", value=emprestimo_to_edit.get("contrato", ""))
        data = st.date_input("Data do Contrato", format="DD/MM/YYYY",
                             value=pd.to_datetime(emprestimo_to_edit.get("data", datetime.today())))
        valor_total = st.number_input("Valor Total do Empréstimo", min_value=0.0, step=1000.0, format="%.2f",
                                      value=emprestimo_to_edit.get("valor_total", 0.0))
        objeto = st.text_input("Objeto", value=emprestimo_to_edit.get("objeto", ""))

    with col2:
        recursos = st.text_input("Recursos", value=emprestimo_to_edit.get("recursos", ""))
        encargos = st.number_input("Encargos (% ao ano)", min_value=0.0, step=0.1,
                                   value=emprestimo_to_edit.get("encargos", 0.0))
        parcelas = st.number_input("Quantidade de Parcelas", min_value=1, step=1,
                                   value=emprestimo_to_edit.get("parcelas", 1))
        valor_parcela = st.number_input("Valor Parcela (R$)", min_value=0.0, step=100.0,
                                        value=emprestimo_to_edit.get("valor_parcela", 0.0))
        periodicidade = st.selectbox("Período de Pagamento", ["ANUAL", "SEMESTRAL", "MENSAL"],
                                     index=["ANUAL", "SEMESTRAL", "MENSAL"].index(
                                         emprestimo_to_edit.get("periodo", "ANUAL")))

        meses_intervalo = {"ANUAL": 12, "SEMESTRAL": 6, "MENSAL": 1}[periodicidade]
        primeira_parcela = datetime(data.year + 1, 5, 15)
        ultima_parcela = primeira_parcela + relativedelta(months=meses_intervalo * (parcelas - 1))
        data_final = st.date_input("Data da Última Parcela", format="DD/MM/YYYY",
                                   value=pd.to_datetime(datetime.today() if not editing_loan else ultima_parcela))

    enviar = st.form_submit_button("Atualizar" if editing_loan else "Cadastrar")

    if enviar:
        novo = {
            "banco": banco,
            "titular": titular,
            "contrato": contrato,
            "data": str(data),
            "valor_total": valor_total,
            "objeto": objeto,
            "recursos": recursos,
            "encargos": encargos,
            "parcelas": parcelas,
            "valor_parcela": valor_parcela,
            "periodo": periodicidade,
            "data_ultima_parcela": data_final.strftime("%d/%m/%Y")
        }

        if editing_loan:
            st.session_state["emprestimos"][st.session_state["editing_loan_index"]] = novo
            st.session_state["editing_loan_index"] = None
        else:
            st.session_state["emprestimos"].append(novo)
        st.success("Empréstimo salvo com sucesso!")
        st.rerun()

# --- EXIBIÇÃO DE DESPESAS ---
st.markdown("### Despesas Cadastradas")
with st.expander("Despesas Cadastradas"):
    if not st.session_state['despesas']:
        st.info("Nenhuma despesa cadastrada.")
    else:
        for i, d in enumerate(st.session_state['despesas']):
            cols = st.columns([3, 2, 2, 1, 1])
            cols[0].write(d["Despesa"])
            cols[1].write(format_brl(d["Valor"]))
            cols[2].write(d["Categoria"])
            if cols[3].button("Editar", key=f"edit_{i}"):
                st.session_state['editing_expense_index'] = i
                st.rerun()
            if cols[4].button("Excluir", key=f"del_{i}"):
                st.session_state['despesas'].pop(i)
                st.rerun()

# --- EXIBIÇÃO DOS EMPRÉSTIMOS ---
st.markdown("### Empréstimos Cadastrados")
with st.expander("Empréstimos"):
    if not st.session_state["emprestimos"]:
        st.info("Nenhum empréstimo cadastrado.")
    else:
        for i, e in enumerate(st.session_state["emprestimos"]):
            cols = st.columns([2, 2, 2, 2, 2, 1, 1])
            cols[0].write(e["banco"])
            cols[1].write(e["titular"])
            cols[2].write(e["objeto"])
            cols[3].write(format_brl(e["valor_total"]))
            cols[4].write(e.get("data_ultima_parcela", "Não informado"))
            if cols[5].button("Editar", key=f"edit_loan_{i}"):
                st.session_state["editing_loan_index"] = i
                st.rerun()
            if cols[6].button("Excluir", key=f"del_loan_{i}"):
                st.session_state["emprestimos"].pop(i)
                st.rerun()

# --- PROJEÇÃO ---
st.markdown("---")
st.markdown("### 📊 Projeção de Despesas com Inflação")

if not st.session_state['despesas'] and not st.session_state['emprestimos']:
    st.info("Adicione despesas ou empréstimos para ver a projeção.")
else:
    df_desp = pd.DataFrame(st.session_state['despesas'])
    df_desp['Despesa_Normalized'] = df_desp['Despesa'].str.strip()
    group = df_desp.groupby('Despesa_Normalized')['Valor'].sum()

    fluxo = {}
    for i, ano in enumerate(anos):
        fator = np.prod([1 + inflacoes[j] / 100 for j in range(i + 1)])
        fluxo[ano] = group * fator if not group.empty else pd.Series(dtype=float)

    df_fluxo = pd.DataFrame(fluxo)

    for emp in st.session_state["emprestimos"]:
        linha = f"Empréstimo: {emp['objeto']}"
        for i in range(min(emp["parcelas"], 5)):
            ano = f"Ano {i+1}"
            if linha not in df_fluxo.index:
                df_fluxo.loc[linha] = 0
            df_fluxo.loc[linha, ano] += emp["valor_parcela"]

    df_fluxo.index.name = "Despesa"
    st.dataframe(df_fluxo.style.format(format_brl))
    st.session_state['fluxo_caixa'] = df_fluxo

# --- LIMPAR TUDO ---
st.markdown("### ⚙️ Ações Gerais")
if st.button("Limpar Tudo", key="btn_clear_all"):
    st.session_state.clear()
    st.success("Todos os dados foram limpos!")
    st.rerun()
