import streamlit as st
import pandas as pd

st.title("Histórico de Plantios")

plantios = st.session_state.get('plantios', {})

if not plantios:
    st.warning("Nenhum plantio cadastrado ainda.")
    st.stop()

# Transformar em DataFrame para visualização
dados = []
for pid, dados_plantio in plantios.items():
    receita_bruta = (
        dados_plantio['hectares']
        * dados_plantio['sacas_por_hectare']
        * dados_plantio['preco_saca']
    )
    dados.append({
        "ID": pid,
        "Ano": dados_plantio['ano'],
        "Cultura": dados_plantio['cultura'],
        "Hectares": dados_plantio['hectares'],
        "Sacas/ha": dados_plantio['sacas_por_hectare'],
        "Preço/saca (R$)": dados_plantio['preco_saca'],
        "Receita bruta (R$)": receita_bruta
    })

df = pd.DataFrame(dados)
df = df.sort_values(by="Ano", ascending=False)

# Função de formatação brasileira
def format_brl(valor):
    return f"R$ {valor:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")

def format_num(valor):
    return f"{valor:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")


st.dataframe(df.style.format({
    "Hectares": format_num,
    "Sacas/ha": format_num,
    "Preço/saca (R$)": format_brl,
    "Receita bruta (R$)": format_brl
}))