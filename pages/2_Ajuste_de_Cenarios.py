import streamlit as st
import json
import os

st.set_page_config(layout="wide", page_title="Configurações de Cenário")
st.title("⚙️ Configurações de Cenário e Inflação")

CONFIG_PATH = "config.json"

# Valores padrão
defaults = {
    "pess_receita": 15,
    "pess_despesas": 10,
    "otm_receita": 10,
    "otm_despesas": 10,
    **{f"inf_{i}": 4.0 for i in range(5)}
}

# Carrega configurações do arquivo se existir
if os.path.exists(CONFIG_PATH):
    with open(CONFIG_PATH, "r") as f:
        config_data = json.load(f)
else:
    config_data = defaults.copy()

# Carrega para o session_state
for key, value in config_data.items():
    if key not in st.session_state:
        st.session_state[key] = value

# Interface de ajustes
with st.form("form_configuracoes"):
    st.subheader("📉 Ajustes de Cenário")
    col1, col2 = st.columns(2)
    with col1:
        pess_receita = st.number_input("Pessimista: Receita - redução (%)", 0, 50, value=st.session_state.get("pess_receita", 15), key="input_pess_receita")
        pess_despesas = st.number_input("Pessimista: Despesas - aumento (%)", 0, 50, value=st.session_state.get("pess_despesas", 10), key="input_pess_despesas")
    with col2:
        otm_receita = st.number_input("Otimista: Receita - aumento (%)", 0, 50, value=st.session_state.get("otm_receita", 10), key="input_otm_receita")
        otm_despesas = st.number_input("Otimista: Despesas - redução (%)", 0, 50, value=st.session_state.get("otm_despesas", 10), key="input_otm_despesas")

    st.subheader("📈 Inflação Projetada por Ano")
    cols = st.columns(5)
    inflacoes = []
    for i, col in enumerate(cols):
        with col:
            inflacao_ano = st.number_input(f"Ano {i+1}", min_value=0.0, max_value=100.0, value=st.session_state.get(f"inf_{i}", 4.0), key=f"input_inf_{i}")
            inflacoes.append(inflacao_ano)

    if st.form_submit_button("Salvar Configurações"):
        st.session_state["pess_receita"] = pess_receita
        st.session_state["pess_despesas"] = pess_despesas
        st.session_state["otm_receita"] = otm_receita
        st.session_state["otm_despesas"] = otm_despesas

        for i in range(5):
            st.session_state[f"inf_{i}"] = inflacoes[i]

        with open(CONFIG_PATH, "w") as f:
            json.dump({k: st.session_state[k] for k in defaults.keys()}, f)

        st.success("Configurações salvas com sucesso. As projeções já podem ser utilizadas nas demais páginas.")