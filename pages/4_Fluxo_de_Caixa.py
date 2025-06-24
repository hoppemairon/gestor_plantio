import streamlit as st
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px

st.set_page_config(layout="wide", page_title="Fluxo de Caixa Projeção") # Configura a página para layout amplo

st.title("🎯 Fluxo de Caixa Projeção 5 Anos")

# ✅ Checagem dos dados necessários
if "plantios" not in st.session_state or not st.session_state["plantios"]:
    st.warning("Cadastre ao menos um plantio para gerar o fluxo de caixa.")
    st.stop()

if "fluxo_caixa" not in st.session_state:
    st.warning("Você precisa preencher as despesas antes de acessar esta página.")
    st.stop()

# 🔢 Anos e inflação
anos = [f"Ano {i+1}" for i in range(5)]
inflacoes = [
    st.session_state.get(f"inf_{i}", 4.0)  # fallback: 4%
    for i in range(5)
]

# 📌 Dados base
# df_despesas já vem sem "Receita Estimada" e "Lucro Líquido"
df_despesas = st.session_state["fluxo_caixa"].drop(["Receita Estimada", "Lucro Líquido"], errors="ignore")
plantios = st.session_state["plantios"] # Assumindo que plantios é um dict como {"Milho": {...}, "Soja": {...}}

# 📈 Cálculo da média da receita por hectare (para o cálculo da receita total)
total_sacas = 0
preco_total = 0
hectares_total = 0

for p_data in plantios.values(): # Itera sobre os valores do dicionário de plantios
    hectares_total += p_data.get("hectares", 0)
    total_sacas += p_data.get("sacas_por_hectare", 0) * p_data.get("hectares", 0)
    preco_total += p_data.get("preco_saca", 0) * p_data.get("sacas_por_hectare", 0) * p_data.get("hectares", 0)

if hectares_total == 0 or total_sacas == 0:
    st.warning("Dados de plantio incompletos para estimar receita.")
    st.stop()

media_sacas = total_sacas / hectares_total
media_preco = preco_total / total_sacas
media_receita_hectare = media_sacas * media_preco

# 📊 Projeção da Receita Total
receita_proj = []
for i in range(5):
    fator = np.prod([1 + inflacoes[j] / 100 for j in range(i + 1)])
    receita_proj.append(hectares_total * media_receita_hectare * fator)

# --- INÍCIO DAS NOVAS SEÇÕES ---

st.markdown("---")
st.markdown("### 📊 Detalhamento e Análise da Projeção de Receitas")

# 1. Explicação do Cálculo das Receitas Futuras
with st.expander("Entenda como a receita futura é calculada por cultura"):
    st.markdown("""
    Para projetar as receitas futuras, consideramos a contribuição individual de cada cultura plantada e aplicamos a taxa de inflação anual.

    **Passos do Cálculo:**

    1.  **Receita Base por Cultura:** Para cada tipo de plantio que você cadastrou (ex: Milho, Soja), calculamos a receita esperada no ano base (sem inflação) multiplicando:
        *   `Hectares Plantados` (para aquela cultura)
        *   `Sacas por Hectare` (produtividade esperada para aquela cultura)
        *   `Preço por Saca` (preço de venda esperado para aquela cultura)

        Isso nos dá a **receita bruta inicial** que cada cultura contribui.

    2.  **Projeção com Inflação:** A receita base de cada cultura é então projetada para os próximos 5 anos. Para cada ano, aplicamos a taxa de inflação acumulada. Isso significa que a receita do Ano 2 considera a inflação do Ano 1 e do Ano 2, e assim por diante.

        *   **Exemplo (Ano 1):** `Receita Base da Cultura X * (1 + Inflação Ano 1)`
        *   **Exemplo (Ano 2):** `Receita Base da Cultura X * (1 + Inflação Ano 1) * (1 + Inflação Ano 2)`

    3.  **Receita Total Estimada:** A "Receita Estimada" que você vê no Fluxo de Caixa Consolidado é a **soma das receitas projetadas de todas as suas culturas** para cada ano. Isso garante uma visão abrangente do seu faturamento futuro.
    """)

# 2. Cálculo e DataFrame para Receitas por Cultura (para os gráficos)
plantio_receitas_data = {} # {nome_cultura: [proj_ano1, proj_ano2, ...]}

if plantios: # Garante que há plantios para iterar
    for nome_plantio, p_data in plantios.items():
        plantio_hectares = p_data.get("hectares", 0)
        plantio_sacas_por_hectare = p_data.get("sacas_por_hectare", 0)
        plantio_preco_saca = p_data.get("preco_saca", 0)

        # Garante que os dados para o cálculo de cada plantio são válidos
        if plantio_hectares == 0 or plantio_sacas_por_hectare == 0 or plantio_preco_saca == 0:
            plantio_receitas_data[nome_plantio] = [0] * 5
            continue

        base_receita_plantio = plantio_hectares * plantio_sacas_por_hectare * plantio_preco_saca

        proj_plantio_anual = []
        for i in range(5):
            fator_inflacao_acumulado = np.prod([1 + inflacoes[j] / 100 for j in range(i + 1)])
            proj_plantio_anual.append(base_receita_plantio * fator_inflacao_acumulado)
        plantio_receitas_data[nome_plantio] = proj_plantio_anual

# Converte para DataFrame para facilitar a plotagem (anos como índice, culturas como colunas)
df_temp = pd.DataFrame.from_dict(plantio_receitas_data, orient='columns')
df_temp.columns = [
    plantios[k].get("cultura", k) for k in plantio_receitas_data.keys()
]
df_temp.index = anos

# Agrupa colunas com mesmo nome (mesma cultura) somando
df_receitas_por_cultura = df_temp.groupby(level=0, axis=1).sum()
# --- FIM DAS NOVAS SEÇÕES ---

# �� Novo DataFrame consolidado
df_fluxo = df_despesas.copy()

# --- INÍCIO DA CORREÇÃO PARA DUPLICAÇÃO DE VALORES NO FLUXO DE CAIXA ---
# 1. Normalizar os nomes dos índices para garantir que duplicatas sejam identificadas corretamente
df_fluxo.index = df_fluxo.index.astype(str).str.strip() # Remove espaços em branco
# Opcional: Se houver problemas com maiúsculas/minúsculas, descomente a linha abaixo:
# df_fluxo.index = df_fluxo.index.str.lower()

# 2. Consolidar quaisquer linhas com índices duplicados somando seus valores
if not df_fluxo.index.is_unique:
    df_fluxo = df_fluxo.groupby(df_fluxo.index).sum()
# --- FIM DA CORREÇÃO PARA DUPLICAÇÃO DE VALORES NO FLUXO DE CAIXA ---

# Adicionar/Atualizar as linhas calculadas
df_fluxo.loc["Receita Estimada"] = receita_proj

df_fluxo.loc["Impostos Sobre Venda"] = [r * 0.0485 for r in receita_proj]

# Calcular lucro_operacional para Impostos Sobre Resultado
lucro_operacional = df_fluxo.loc["Receita Estimada"] - df_fluxo.drop(["Receita Estimada", "Lucro Líquido"], errors="ignore").sum()

df_fluxo.loc["Impostos Sobre Resultado"] = [
    lo * 0.15 if lo > 0 else 0 for lo in lucro_operacional
]

# Calcular o Lucro Líquido final (esta é a única e definitiva linha de cálculo)
df_fluxo.loc["Lucro Líquido"] = df_fluxo.loc["Receita Estimada"] - df_fluxo.drop("Receita Estimada").sum()

# 💰 Formatação brasileira (mantido do seu código original)
def format_brl(valor):
    return f"R$ {valor:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")

# A garantia de índice único já foi feita acima, esta parte é redundante agora.
# if not df_fluxo.index.is_unique:
#     df_fluxo = df_fluxo.groupby(df_fluxo.index).sum()

# Estilização e ordenação
ordem = ["Receita Estimada"] + [idx for idx in df_fluxo.index if idx not in ["Receita Estimada", "Lucro Líquido"]] + ["Lucro Líquido"]
df_fluxo = df_fluxo.loc[ordem]

def aplicar_estilo(linha):
    if linha.name == "Receita Estimada":
        return ["background-color: #003366; color: white;" for _ in linha]
    elif linha.name == "Lucro Líquido":
        return ["background-color: #006400; color: white;" for _ in linha]
    else:
        return ["" for _ in linha]

style_idx = [
    {
        "selector": f"tbody tr:nth-child({df_fluxo.index.get_loc('Receita Estimada') + 1}) th",
        "props": [("background-color", "#003366"), ("color", "white")]
    },
    {
        "selector": f"tbody tr:nth-child({df_fluxo.index.get_loc('Lucro Líquido') + 1}) th",
        "props": [("background-color", "#006400"), ("color", "white")]
    }
]

st.markdown("### Resultado Consolidado")
st.dataframe(
    df_fluxo.style
        .format(format_brl)
        .apply(aplicar_estilo, axis=1)
        .set_table_styles(style_idx, overwrite=False)
)

# =====================
# DEMONSTRAÇÃO DE RESULTADO DO EXERCÍCIO (DRE) (mantido do seu código original)
# =====================
st.markdown("### Demonstração do Resultado do Exercício (DRE)")

# Recupera as despesas detalhadas
df_despesas_info = pd.DataFrame(st.session_state['despesas'])
# 'anos' já está definido no início do script

# Mapeia categorias para o DRE
categorias_dre = {
    "Impostos Sobre Venda": "Impostos",
    "Despesas Operacionais": "Operacional",
    "Despesas Administrativas": "Administrativa",
    "Despesas RH": "RH",
    "Despesas Extra Operacional": "Extra Operacional",
    "Dividendos": "Dividendos"
}

# Inicializa o dicionário do DRE
dre = {}

# Processa cada categoria de despesa
for nome_dre, cat in categorias_dre.items():
    # Normaliza a coluna 'Categoria' para garantir correspondência
    if not df_despesas_info.empty:
        df_despesas_info['Categoria_Normalized'] = df_despesas_info['Categoria'].astype(str).str.strip()
        # Opcional: df_despesas_info['Categoria_Normalized'] = df_despesas_info['Categoria_Normalized'].str.lower()
        valores = df_despesas_info[df_despesas_info["Categoria_Normalized"] == cat]["Valor"].tolist()
    else:
        valores = []

    linha = []
    for i in range(5):
        fator = np.prod([1 + inflacoes[j] / 100 for j in range(i + 1)])
        # Garante que a soma seja feita apenas se houver valores
        valor_base_categoria = sum(valores) if valores else 0
        linha.append(valor_base_categoria * fator)
    dre[nome_dre] = linha

# RECEITA direto do fluxo
dre["Receita"] = df_fluxo.loc["Receita Estimada"].tolist()

# ==== 📘 Explicação dos Impostos ====
with st.expander("🧾 Entenda os impostos aplicados no DRE e Fluxo de Caixa"):
    st.markdown("""
    O sistema aplica os principais impostos com base na **receita estimada** e no **lucro operacional**, simulando uma empresa agropecuária no regime presumido.

    **1. Impostos sobre Venda (4,85%)**
    - **FUNRURAL (1,2%)**: contribuição previdenciária sobre a receita bruta da comercialização.
    - **PIS/COFINS (3,65%)**: contribuições sociais sobre faturamento, comuns no agronegócio.

    Esses impostos incidem diretamente sobre a **Receita Estimada** e são somados como "Impostos Sobre Venda".

    **2. Impostos sobre Resultado (15%)**
    - Refere-se a uma estimativa de **IRPJ + CSLL**, aplicada **somente sobre o Lucro Operacional**, se positivo.

    > 💡 Estes valores são simulações aproximadas, ideais para análise financeira e não substituem a apuração contábil real.
    """)

# ==== 🔢 Recalcular DRE com impostos ====
dre["Impostos Sobre Venda"] = [dre["Receita"][i] * 0.0485 for i in range(5)]

dre["Margem de Contribuição"] = [
    dre["Receita"][i] - dre["Impostos Sobre Venda"][i] - dre["Despesas Operacionais"][i]
    for i in range(5)
]

dre["Resultado Operacional"] = [
    dre["Margem de Contribuição"][i] - dre["Despesas Administrativas"][i] - dre["Despesas RH"][i]
    for i in range(5)
]

dre["Lucro Operacional"] = [
    dre["Resultado Operacional"][i] - dre["Despesas Extra Operacional"][i]
    for i in range(5)
]

dre["Impostos Sobre Resultado"] = [
    dre["Lucro Operacional"][i] * 0.15 if dre["Lucro Operacional"][i] > 0 else 0
    for i in range(5)
]

dre["Lucro Líquido"] = [
    dre["Lucro Operacional"][i] - dre["Impostos Sobre Resultado"][i] - dre["Dividendos"][i]
    for i in range(5)
]

# Ordem correta das linhas
ordem_dre = [ # Renomeado para evitar conflito com 'ordem' anterior
    "Receita",
    "Impostos Sobre Venda",
    "Despesas Operacionais",
    "Margem de Contribuição",
    "Despesas Administrativas",
    "Despesas RH",
    "Resultado Operacional",
    "Despesas Extra Operacional",
    "Lucro Operacional",
    "Impostos Sobre Resultado",
    "Dividendos",
    "Lucro Líquido"
]

# Construir DataFrame final do DRE
df_dre = pd.DataFrame(dre, index=anos).T.loc[ordem_dre]

def aplicar_estilo_dre(linha):
    """Aplica estilos de fundo e fonte às linhas específicas da DRE."""
    if linha.name == "Receita":
        return ["background-color: #003366; color: white;" for _ in linha] # Azul Escuro
    elif linha.name in ["Margem de Contribuição", "Resultado Operacional", "Lucro Operacional", "Lucro Líquido"]:
        return ["background-color: #006400; color: white;" for _ in linha] # Verde Escuro
    else:
        return ["" for _ in linha]

# Lista de estilos para os cabeçalhos das linhas (índices) da DRE
style_idx_dre = []

# Estilo para a linha "Receita"
if "Receita" in df_dre.index:
    style_idx_dre.append({
        "selector": f"tbody tr:nth-child({df_dre.index.get_loc('Receita') + 1}) th",
        "props": [("background-color", "#003366"), ("color", "white")]
    })

# Estilos para as linhas verdes
green_rows_dre = ["Margem de Contribuição", "Resultado Operacional", "Lucro Operacional", "Lucro Líquido"]
for row_name in green_rows_dre:
    if row_name in df_dre.index: # Garante que a linha existe antes de tentar estilizar
        style_idx_dre.append({
            "selector": f"tbody tr:nth-child({df_dre.index.get_loc(row_name) + 1}) th",
            "props": [("background-color", "#006400"), ("color", "white")]
        })

st.dataframe(
    df_dre.style
        .format(format_brl)
        .apply(aplicar_estilo_dre, axis=1) # Aplica estilo às células da linha
        .set_table_styles(style_idx_dre, overwrite=False), # Aplica estilo aos cabeçalhos das linhas
    height=458
)

# 3. Geração dos Gráficos
if not df_receitas_por_cultura.empty and df_receitas_por_cultura.sum().sum() > 0: # Verifica se há dados válidos para plotar
    st.markdown("### Visualização das Projeções de Receita")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("#### Projeção da Receita Total Estimada")
        fig_total_receita = go.Figure()
        fig_total_receita.add_trace(go.Scatter(x=anos, y=receita_proj, mode='lines+markers', name='Receita Total Estimada',
                                               line=dict(color='#003366', width=3), marker=dict(size=8)))
        # --- MODIFICAÇÃO AQUI: hovertemplate para Receita Total ---
        fig_total_receita.update_layout(
            title_text='Evolução da Receita Total ao Longo dos Anos',
            xaxis_title='Ano',
            yaxis_title='Receita (R$)',
            hovermode="x unified",
            template="plotly_white", # Tema claro
            height=400
        )
        fig_total_receita.update_traces(hovertemplate="<b>Ano: %{x}</b><br>Receita Total: R$ %{y:,.2f}<extra></extra>")
        st.plotly_chart(fig_total_receita, use_container_width=True)

    with col2:
        st.markdown("#### Contribuição de Receita por Cultura (Empilhado)")
        fig_cultura_stacked = px.bar(
            df_receitas_por_cultura,
            x=df_receitas_por_cultura.index, # Anos
            y=df_receitas_por_cultura.columns, # Nomes das culturas
            title='Projeção de Receita por Cultura (5 Anos)',
            labels={'value': 'Receita (R$)', 'variable': 'Cultura', 'index': 'Ano'},
            height=400,
            template="plotly_white",
            color_discrete_sequence=px.colors.qualitative.Pastel # Paleta de cores suave
        )
        fig_cultura_stacked.update_layout(
            yaxis_title='Receita (R$)',
            barmode='stack',
            hovermode="x unified"
        )
        # --- MODIFICAÇÃO AQUI: hovertemplate para Barras Empilhadas ---
        fig_cultura_stacked.update_traces(hovertemplate="<b>Ano: %{x}</b><br>Cultura: %{name}<br>Receita: R$ %{y:,.2f}<extra></extra>")
        st.plotly_chart(fig_cultura_stacked, use_container_width=True)

    st.markdown("#### Tendência de Receita por Cultura Individual")
    fig_cultura_lines = px.line(
        df_receitas_por_cultura,
        x=df_receitas_por_cultura.index, # Anos
        y=df_receitas_por_cultura.columns, # Nomes das culturas
        title='Evolução da Receita de Cada Cultura ao Longo dos Anos',
        labels={'value': 'Receita (R$)', 'variable': 'Cultura', 'index': 'Ano'},
        height=500,
        template="plotly_white",
        color_discrete_sequence=px.colors.qualitative.Dark24 # Paleta de cores mais distinta
    )
    fig_cultura_lines.update_layout(
        xaxis_title='Ano',
        yaxis_title='Receita (R$)',
        hovermode="x unified"
    )
    # --- MODIFICAÇÃO AQUI: hovertemplate para Linhas Individuais ---
    fig_cultura_lines.update_traces(hovertemplate="<b>Ano: %{x}</b><br>Cultura: %{name}<br>Receita: R$ %{y:,.2f}<extra></extra>")
    st.plotly_chart(fig_cultura_lines, use_container_width=True)

else:
    st.info("Não há dados de plantio suficientes ou válidos para gerar os gráficos de receita por cultura.")