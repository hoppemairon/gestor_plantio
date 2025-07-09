import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from io import BytesIO

from utils.session import carregar_configuracoes
carregar_configuracoes()

st.set_page_config(layout="wide", page_title="Indicadores Financeiros")
st.title("📈 Indicadores Financeiros e Análise - Agronegócio")

# Explicações dos Indicadores
with st.expander("🧾 Entenda os Indicadores Financeiros"):
    st.markdown("""
    Abaixo, apresentamos indicadores financeiros avançados, adaptados ao contexto do agronegócio, com explicações sobre sua importância e interpretação:

    **1. Margem Líquida (%)**
    - **O que é?** Percentual do lucro líquido em relação à receita total após todas as despesas e impostos.
    - **Por que é importante?** Mede a rentabilidade líquida do negócio. No agronegócio, margens são impactadas por custos sazonais (insumos, colheita) e preços de mercado.
    - **Exemplo:** Receita de R\\$ 1.000.000 e lucro líquido de R\\$ 150.000 resultam em 15% de margem.
    - **Interpretação:** Margens abaixo de 10% sugerem necessidade de otimizar custos ou preços.

    **2. Retorno por Real Gasto**
    - **O que é?** Quanto cada real gasto (despesas + impostos) gera de lucro líquido.
    - **Por que é importante?** Avalia a eficiência dos investimentos, crucial em um setor com altos custos fixos e variáveis.
    - **Exemplo:** Gastos de R\\$ 500.000 geram R\\$ 100.000 de lucro (retorno de 0,2).
    - **Interpretação:** Valores abaixo de 0,2 indicam ineficiência; analise despesas operacionais.

    **3. Liquidez Operacional**
    - **O que é?** Quantas vezes a receita cobre as despesas operacionais.
    - **Por que é importante?** Mostra a capacidade de sustentar operações sem depender de financiamentos, essencial em safras incertas.
    - **Exemplo:** Receita de R\\$ 1.000.000 e despesas operacionais de R\\$ 400.000 resultam em liquidez de 2,5.
    - **Interpretação:** Valores abaixo de 1,5 indicam risco de fluxo de caixa.

    **4. Endividamento (%)**
    - **O que é?** Proporção das parcelas de empréstimos em relação à receita total.
    - **Por que é importante?** Avalia o peso das dívidas, comum no agronegócio para custeio e investimentos.
    - **Exemplo:** Parcelas de R\\$ 200.000 com receita de R\\$ 1.000.000 resultam em 20% de endividamento.
    - **Interpretação:** Acima de 30% é arriscado; priorize redução de dívidas.

    **5. Produtividade por Hectare (R$/ha)**
    - **O que é?** Receita gerada por hectare plantado.
    - **Por que é importante?** Mede a eficiência do uso da terra, um recurso crítico no agronegócio.
    - **Exemplo:** Receita de R\\$ 1.000.000 em 500 ha resulta em R\\$ 2.000/ha.
    - **Interpretação:** Valores baixos podem indicar baixa produtividade ou preços de mercado desfavoráveis.

    **6. Custo por Receita (%)**
    - **O que é?** Proporção dos custos operacionais em relação à receita.
    - **Por que é importante?** Indica a eficiência na gestão de custos, essencial em um setor com margens apertadas.
    - **Exemplo:** Custos operacionais de R\\$ 600.000 e receita de R\\$ 1.000.000 resultam em 60%.
    - **Interpretação:** Acima de 70% sugere necessidade de redução de custos.

    **7. Debt Service Coverage Ratio (DSCR)**
    - **O que é?** Razão entre o lucro operacional e as parcelas anuais de dívidas.
    - **Por que é importante?** Avalia a capacidade de pagar dívidas com o lucro gerado, crucial para financiamentos agrícolas.
    - **Exemplo:** Lucro operacional de R\\$ 300.000 e parcelas de R\\$ 150.000 resultam em DSCR de 2,0.
    - **Interpretação:** Valores abaixo de 1,25 indicam risco de inadimplência.

    **8. Break-Even Yield (sacas/ha)**
    - **O que é?** Produtividade mínima por hectare necessária para cobrir todos os custos.
    - **Por que é importante?** Ajuda a avaliar o risco de safras em cenários de baixa produtividade.
    - **Exemplo:** Custos totais de R\\$ 1.000.000, 500 ha e preço da saca de R\\$ 100 resultam em 20 sacas/ha.
    - **Interpretação:** Valores altos indicam maior vulnerabilidade a falhas na safra.

    **9. Return on Assets (ROA) (%)**
    - **O que é?** Percentual do lucro líquido em relação ao total de ativos (estimado).
    - **Por que é importante?** Mede a eficiência do uso de ativos (terra, máquinas) para gerar lucro.
    - **Exemplo:** Lucro líquido de R\\$ 150.000 e ativos de R\\$ 5.000.000 resultam em 3% de ROA.
    - **Interpretação:** Valores abaixo de 5% sugerem baixa eficiência no uso de ativos.

    **10. CAGR Receita (%)**
    - **O que é?** Taxa de crescimento anual composta da receita ao longo de 5 anos.
    - **Por que é importante?** Indica a tendência de crescimento do faturamento, útil para planejamento.
    - **Exemplo:** Receita inicial de R\\$ 1.000.000 e final de R\\$ 1.300.000 resultam em ~5,4%.
    - **Interpretação:** Valores negativos indicam retração; revise preços ou produtividade.

    **11. CAGR Lucro Líquido (%)**
    - **O que é?** Taxa de crescimento anual composta do lucro líquido ao longo de 5 anos.
    - **Por que é importante?** Reflete a sustentabilidade do lucro em um setor volátil.
    - **Exemplo:** Lucro inicial de R\\$ 100.000 e final de R\\$ 150.000 resultam em ~8,4%.
    - **Interpretação:** Valores negativos requerem revisão de custos ou estratégias.
    """)

# Verificação de dados essenciais
if "fluxo_caixa" not in st.session_state:
    st.warning("Dados do fluxo de caixa não foram encontrados. Preencha as informações na página de Fluxo de Caixa.")
    st.stop()
if "plantios" not in st.session_state or not st.session_state["plantios"]:
    st.warning("Nenhum plantio cadastrado. Cadastre ao menos um plantio para gerar os indicadores.")
    st.stop()

inflacao_padrao = 0.04
anos = [f"Ano {i+1}" for i in range(5)]
inflacoes = [st.session_state.get(f"inf_{i}", inflacao_padrao * 100) / 100 for i in range(5)]
nomes_cenarios = ["Projetado", "Pessimista", "Otimista"]

with st.expander("🔧 Cenário e Inflação"):
    st.markdown("### 📈 Inflação Estimada por Ano")
    cols = st.columns(5)
    for i, col in enumerate(cols):
        with col:
            st.metric(f"Ano {i+1}", f"{inflacoes[i]*100:.2f}%")

    pess_receita = st.session_state.get("pess_receita", 15)
    pess_despesas = st.session_state.get("pess_despesas", 10)
    otm_receita = st.session_state.get("otm_receita", 10)
    otm_despesas = st.session_state.get("otm_despesas", 10)

    st.markdown("### 🔧 Parâmetros de Cenário Atuais")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("💸 Receita Pessimista", f"-{pess_receita}%")
    with col2:    
        st.metric("💰 Despesa Pessimista", f"+{pess_despesas}%")
    with col3:
        st.metric("💸 Receita Otimista", f"+{otm_receita}%")
    with col4:
        st.metric("💰 Despesa Otimista", f"-{otm_despesas}%")

# === RECEITA POR CULTURA ===
st.markdown("### 🌱 Receita por Cultura Agrícola (Ano Base)")
culturas = {}
for p in st.session_state["plantios"].values():
    cultura = p["cultura"]
    receita = p["hectares"] * p["sacas_por_hectare"] * p["preco_saca"]
    if cultura not in culturas:
        culturas[cultura] = {"receita": 0, "hectares": 0}
    culturas[cultura]["receita"] += receita
    culturas[cultura]["hectares"] += p["hectares"]

# Calcular receitas base e despesas para custo por hectare
total_sacas = preco_total = hectares_total = 0
for p_data in st.session_state["plantios"].values():
    hectares = p_data.get("hectares", 0)
    sacas = p_data.get("sacas_por_hectare", 0)
    preco = p_data.get("preco_saca", 0)
    hectares_total += hectares
    total_sacas += sacas * hectares
    preco_total += preco * sacas * hectares

if hectares_total == 0 or total_sacas == 0:
    st.error("Dados de plantio incompletos para estimar receita.")
    st.stop()

media_receita_hectare = (preco_total / total_sacas) * (total_sacas / hectares_total) if total_sacas > 0 else 0
receita_base = [hectares_total * media_receita_hectare * np.prod([1 + inflacoes[j] for j in range(i + 1)]) for i in range(5)]
receitas = {
    "Projetado": receita_base,
    "Pessimista": [r * (1 - pess_receita / 100) for r in receita_base],
    "Otimista": [r * (1 + otm_receita / 100) for r in receita_base]
}

# Calcular custo por hectare usando o cenário Projetado
def ajustar_despesas(df_base, ajuste_percentual):
    df = df_base.copy()
    for row in df.index:
        if row not in ["Receita Estimada", "Lucro Líquido", "Impostos Sobre Resultado"]:
            df.loc[row] = df.loc[row] * (1 + ajuste_percentual / 100)
    return df

fluxos = {
    "Projetado": st.session_state["fluxo_caixa"].copy(),
    "Pessimista": ajustar_despesas(st.session_state["fluxo_caixa"], pess_despesas),
    "Otimista": ajustar_despesas(st.session_state["fluxo_caixa"], -otm_despesas)
}

# Estimativa de ativos totais
total_ativos = hectares_total * 20000 + 1000000  # R$20.000/ha + ativos fixos (máquinas, etc.)

# Calcular DRE para custo por hectare
dre_por_cenario = {}
for cenario in nomes_cenarios:
    df_fluxo = fluxos[cenario].copy()
    df_fluxo.loc["Receita Estimada"] = receitas[cenario]
    df_fluxo.loc["Impostos Sobre Venda"] = df_fluxo.loc["Receita Estimada"] * 0.0485

    # Recalcular DRE
    df_despesas_info = pd.DataFrame(st.session_state.get("despesas", []))
    if not df_despesas_info.empty and "Categoria" in df_despesas_info.columns:
        df_despesas_info["Categoria"] = df_despesas_info["Categoria"].astype(str).str.strip()
    else:
        df_despesas_info = pd.DataFrame(columns=["Categoria", "Valor"])

    dre_calc = {}
    dre_calc["Receita"] = receitas[cenario]
    dre_calc["Impostos Sobre Venda"] = [r * 0.0485 for r in receitas[cenario]]

    def linha_despesa(cat):
        total = sum(df_despesas_info[df_despesas_info["Categoria"] == cat]["Valor"]) if not df_despesas_info.empty else 0
        ajuste = pess_despesas if cenario == "Pessimista" else (-otm_despesas if cenario == "Otimista" else 0)
        return [total * np.prod([1 + inflacoes[j] for j in range(i + 1)]) * (1 + ajuste / 100) for i in range(5)]

    dre_calc["Despesas Operacionais"] = linha_despesa("Operacional")
    dre_calc["Despesas Administrativas"] = linha_despesa("Administrativa")
    dre_calc["Despesas RH"] = linha_despesa("RH")
    extra_operacional = linha_despesa("Extra Operacional")

    if "emprestimos" in st.session_state:
        for emp in st.session_state.get("emprestimos", []):
            try:
                start_year_index = anos.index(emp["ano_inicial"])
                end_year_index = anos.index(emp["ano_final"])
                num_years = end_year_index - start_year_index + 1
                ajuste = pess_despesas if cenario == "Pessimista" else (-otm_despesas if cenario == "Otimista" else 0)
                for i in range(start_year_index, min(start_year_index + min(emp["parcelas"], num_years), len(anos))):
                    extra_operacional[i] += emp["valor_parcela"] * (1 + ajuste / 100)
            except ValueError:
                continue

    dre_calc["Despesas Extra Operacional"] = extra_operacional
    dre_calc["Dividendos"] = linha_despesa("Dividendos")
    dre_calc["Margem de Contribuição"] = [
        dre_calc["Receita"][i] - dre_calc["Impostos Sobre Venda"][i] - dre_calc["Despesas Operacionais"][i]
        for i in range(5)
    ]
    dre_calc["Resultado Operacional"] = [
        dre_calc["Margem de Contribuição"][i] - dre_calc["Despesas Administrativas"][i] - dre_calc["Despesas RH"][i]
        for i in range(5)
    ]
    dre_calc["Lucro Operacional"] = [
        dre_calc["Resultado Operacional"][i] - dre_calc["Despesas Extra Operacional"][i]
        for i in range(5)
    ]
    dre_calc["Impostos Sobre Resultado"] = [
        dre_calc["Lucro Operacional"][i] * 0.15 if dre_calc["Lucro Operacional"][i] > 0 else 0
        for i in range(5)
    ]
    dre_calc["Lucro Líquido"] = [
        dre_calc["Lucro Operacional"][i] - dre_calc["Impostos Sobre Resultado"][i] - dre_calc["Dividendos"][i]
        for i in range(5)
    ]

    df_fluxo.loc["Lucro Líquido"] = dre_calc["Lucro Líquido"]
    dre_por_cenario[cenario] = dre_calc


# Adicionar custo por hectare à tabela de culturas
df_culturas = pd.DataFrame([
    {
        "Cultura": cultura,
        "Receita Total": dados["receita"],
        "Área (ha)": dados["hectares"],
        "Receita por ha": dados["receita"] / dados["hectares"] if dados["hectares"] > 0 else 0
    }
    for cultura, dados in culturas.items()
])

# Função para formatar valores em BRL
def format_brl(x):
    try:
        return f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return x

# Exibir tabela de culturas
st.dataframe(
    df_culturas.style.format({
        "Receita Total": format_brl,
        "Área (ha)": "{:.2f}",
        "Receita por ha": format_brl,
        "Custo por ha": format_brl
    }),
    use_container_width=True
)

# Função para calcular CAGR
def calcular_cagr(valor_inicial, valor_final, periodos):
    if valor_inicial <= 0 or valor_final <= 0 or periodos <= 0:
        return 0
    return ((valor_final / valor_inicial) ** (1 / periodos) - 1) * 100

# Calcular indicadores
indicadores = {cenario: {} for cenario in nomes_cenarios}
for cenario in nomes_cenarios:
    # Indicadores Avançados
    despesas_totais = [
        dre_por_cenario[cenario]["Impostos Sobre Venda"][i] + dre_por_cenario[cenario]["Despesas Operacionais"][i] +
        dre_por_cenario[cenario]["Despesas Administrativas"][i] + dre_por_cenario[cenario]["Despesas RH"][i] +
        dre_por_cenario[cenario]["Despesas Extra Operacional"][i] + dre_por_cenario[cenario]["Dividendos"][i] +
        dre_por_cenario[cenario]["Impostos Sobre Resultado"][i]
        for i in range(5)
    ]  
      
    indicadores[cenario]["Produtividade por Hectare (R$/ha)"] = [
        r / hectares_total if hectares_total > 0 else 0 for r in dre_por_cenario[cenario]["Receita"]
    ]

    indicadores[cenario]["Custo por Hectare (R$/ha)"] = [
        (
            dre_por_cenario[cenario]["Impostos Sobre Venda"][i] +
            dre_por_cenario[cenario]["Despesas Operacionais"][i] +
            dre_por_cenario[cenario]["Despesas Administrativas"][i] +
            dre_por_cenario[cenario]["Despesas RH"][i] +
            dre_por_cenario[cenario]["Despesas Extra Operacional"][i] +
            dre_por_cenario[cenario]["Dividendos"][i] +
            dre_por_cenario[cenario]["Impostos Sobre Resultado"][i]
        ) / hectares_total if hectares_total > 0 else 0
        for i in range(5)
    ]
    
    indicadores[cenario]["Margem Líquida (%)"] = [
        l / r * 100 if r > 0 else 0 for l, r in zip(dre_por_cenario[cenario]["Lucro Líquido"], dre_por_cenario[cenario]["Receita"])
    ]

    indicadores[cenario]["Retorno por Real Gasto"] = [
        l / d if d > 0 else 0 for l, d in zip(dre_por_cenario[cenario]["Lucro Líquido"], despesas_totais)
    ]

    endividamento = [
        sum(emp["valor_parcela"] for emp in st.session_state.get("emprestimos", [])
            if i >= anos.index(emp["ano_inicial"]) and i <= anos.index(emp["ano_final"])) / r if r > 0 else 0
        for i, r in enumerate(dre_por_cenario[cenario]["Receita"])
    ]
    
    indicadores[cenario]["Liquidez Operacional"] = [
        r / d if d > 0 else 0 for r, d in zip(dre_por_cenario[cenario]["Receita"], dre_por_cenario[cenario]["Despesas Operacionais"])
    ]

    indicadores[cenario]["Endividamento (%)"] = [e * 100 for e in endividamento]

    indicadores[cenario]["Custo por Receita (%)"] = [
        d / r * 100 if r > 0 else 0 for d, r in zip(dre_por_cenario[cenario]["Despesas Operacionais"], dre_por_cenario[cenario]["Receita"])
    ]

    indicadores[cenario]["DSCR"] = [
        lo / sum(emp["valor_parcela"] for emp in st.session_state.get("emprestimos", [])
                 if i >= anos.index(emp["ano_inicial"]) and i <= anos.index(emp["ano_final"]))
        if sum(emp["valor_parcela"] for emp in st.session_state.get("emprestimos", [])
               if i >= anos.index(emp["ano_inicial"]) and i <= anos.index(emp["ano_final"])) > 0 else float("inf")
        for i, lo in enumerate(dre_por_cenario[cenario]["Lucro Operacional"])
    ]

    indicadores[cenario]["Break-Even Yield (sacas/ha)"] = [
        (d + dre_por_cenario[cenario]["Impostos Sobre Venda"][i] + dre_por_cenario[cenario]["Despesas Administrativas"][i] +
         dre_por_cenario[cenario]["Despesas RH"][i] + dre_por_cenario[cenario]["Despesas Extra Operacional"][i] +
         dre_por_cenario[cenario]["Dividendos"][i] + dre_por_cenario[cenario]["Impostos Sobre Resultado"][i]) /
        (hectares_total * (preco_total / total_sacas)) if hectares_total > 0 and total_sacas > 0 else 0
        for i, d in enumerate(dre_por_cenario[cenario]["Despesas Operacionais"])
    ]
    
    indicadores[cenario]["ROA (%)"] = [
        l / total_ativos * 100 if total_ativos > 0 else 0 for l in dre_por_cenario[cenario]["Lucro Líquido"]
    ]
    
    indicadores[cenario]["CAGR Receita (%)"] = calcular_cagr(dre_por_cenario[cenario]["Receita"][0], dre_por_cenario[cenario]["Receita"][-1], len(anos) - 1)
    
    indicadores[cenario]["CAGR Lucro Líquido (%)"] = calcular_cagr(dre_por_cenario[cenario]["Lucro Líquido"][0], dre_por_cenario[cenario]["Lucro Líquido"][-1], len(anos) - 1)

# Exibição de indicadores
st.markdown("### 📊 Indicadores Financeiros")
parecer_por_cenario = {}
for cenario in nomes_cenarios:
    st.subheader(f"Cenário {cenario}")
    df_indicadores = pd.DataFrame({
        k: v for k, v in indicadores[cenario].items()
        if k not in ["CAGR Receita (%)", "CAGR Lucro Líquido (%)"]
    }, index=anos)

    styled_df = df_indicadores.style.format({
        "Margem Líquida (%)": "{:.2f}%",
        "Retorno por Real Gasto": "{:.2f}",
        "Endividamento (%)": "{:.2f}%",
        "Liquidez Operacional": "{:.2f}",
        "Produtividade por Hectare (R$/ha)": format_brl,
        "Custo por Receita (%)": "{:.2f}%",
        "DSCR": "{:.2f}",
        "Break-Even Yield (sacas/ha)": "{:.2f}",
        "ROA (%)": "{:.2f}%",
        "Custo por Hectare (R$/ha)": format_brl
    })

    st.dataframe(styled_df, use_container_width=True)

    # Exibir CAGRs separadamente
    col_a, col_b = st.columns(2)
    with col_a:
        st.metric("📈 CAGR Receita (5 anos)", f"{indicadores[cenario]['CAGR Receita (%)']:.2f}%")
    with col_b:
        st.metric("📈 CAGR Lucro Líquido (5 anos)", f"{indicadores[cenario]['CAGR Lucro Líquido (%)']:.2f}%")

    # Tabela Resumo por Ano
    st.markdown("### 📘 Resumo Financeiro por Ano")
    dados_resumo = {
        "Ano": anos,
        "📈 Receita": [dre_por_cenario[cenario]["Receita"][i] for i in range(5)],
        "💸 Despesas": [
            dre_por_cenario[cenario]["Impostos Sobre Venda"][i] +
            dre_por_cenario[cenario]["Despesas Operacionais"][i] +
            dre_por_cenario[cenario]["Despesas Administrativas"][i] +
            dre_por_cenario[cenario]["Despesas RH"][i] +
            dre_por_cenario[cenario]["Despesas Extra Operacional"][i] +
            dre_por_cenario[cenario]["Dividendos"][i] +
            dre_por_cenario[cenario]["Impostos Sobre Resultado"][i]
            for i in range(5)
        ],
        "💰 Resultado": [dre_por_cenario[cenario]["Lucro Líquido"][i] for i in range(5)],
    }
    df_resumo = pd.DataFrame(dados_resumo)
    df_resumo_formatado = df_resumo.style.format({
        "📈 Receita": format_brl,
        "💸 Despesas": format_brl,
        "💰 Resultado": format_brl,
    }).highlight_between(
        subset=["💰 Resultado"],
        left=0, right=float("inf"),
        color="#1f5142"
    ).highlight_between(
        subset=["💰 Resultado"],
        left=float("-inf"), right=0,
        color="#800000"
    )
    st.dataframe(df_resumo_formatado, use_container_width=True, hide_index=True)

# Gráficos Avançados
st.markdown("### 📈 Visualizações")


st.subheader("Receita vs. Lucro Líquido")
fig1 = go.Figure()
for cenario in nomes_cenarios:
    fig1.add_trace(go.Bar(
        x=anos, y=receitas[cenario], name=f"Receita ({cenario})",
        marker_color="#1f77b4" if cenario == "Projetado" else "#ff7f0e" if cenario == "Pessimista" else "#2ca02c"
    ))
    fig1.add_trace(go.Bar(
        x=anos, y=dre_por_cenario[cenario]["Lucro Líquido"], name=f"Lucro Líquido ({cenario})",
        marker_color="#aec7e8" if cenario == "Projetado" else "#ffbb78" if cenario == "Pessimista" else "#98df8a"
    ))
    
fig1.update_layout(barmode="group", title="Comparação de Receita e Lucro Líquido", yaxis_title="R$", template="plotly_white")
st.plotly_chart(fig1, use_container_width=True)

col1, col2 = st.columns(2)

with col1:
    st.subheader("Distribuição de Despesas (Cenário Projetado)")
    despesas_categorias = pd.DataFrame(st.session_state.get("despesas", [])).groupby("Categoria")["Valor"].sum()
    if not despesas_categorias.empty:
        fig4 = px.pie(values=despesas_categorias.values, names=despesas_categorias.index, title="Distribuição de Despesas por Categoria")
        fig4.update_layout(template="plotly_white")
        st.plotly_chart(fig4, use_container_width=True)
    else:
        st.warning("Nenhuma despesa cadastrada para exibir a distribuição.")

with col2:
    st.subheader("Margem Líquida vs. Custo por Receita")
    fig2 = go.Figure()
    for cenario in nomes_cenarios:
        fig2.add_trace(go.Scatter(
            x=anos, y=indicadores[cenario]["Margem Líquida (%)"], name=f"Margem Líquida ({cenario})",
            line=dict(color="#1f77b4" if cenario == "Projetado" else "#ff7f0e" if cenario == "Pessimista" else "#2ca02c")
        ))
        fig2.add_trace(go.Scatter(
            x=anos, y=indicadores[cenario]["Custo por Receita (%)"], name=f"Custo por Receita ({cenario})",
            line=dict(color="#aec7e8" if cenario == "Projetado" else "#ffbb78" if cenario == "Pessimista" else "#98df8a", dash="dash")
        ))
    fig2.update_layout(title="Margem Líquida vs. Custo por Receita", yaxis_title="Percentual (%)", template="plotly_white")
    st.plotly_chart(fig2, use_container_width=True)

st.subheader("Produtividade por Hectare e Break-Even Yield")
fig3 = go.Figure()
for cenario in nomes_cenarios:
    fig3.add_trace(go.Bar(
        x=anos, y=indicadores[cenario]["Produtividade por Hectare (R$/ha)"], name=f"Produtividade por Hectare ({cenario})",
        marker_color="#1f77b4" if cenario == "Projetado" else "#ff7f0e" if cenario == "Pessimista" else "#2ca02c"
    ))
    fig3.add_trace(go.Scatter(
        x=anos, y=indicadores[cenario]["Break-Even Yield (sacas/ha)"], name=f"Break-Even Yield ({cenario})",
        line=dict(color="#aec7e8" if cenario == "Projetado" else "#ffbb78" if cenario == "Pessimista" else "#98df8a")
    ))
fig3.update_layout(barmode="group", title="Produtividade por Hectare vs. Break-Even Yield", yaxis_title="R$/ha e Sacas/ha", template="plotly_white")
st.plotly_chart(fig3, use_container_width=True)



# Parecer Financeiro
st.markdown("### 📝 Parecer Financeiro")
for cenario in nomes_cenarios:
    st.subheader(f"Parecer - Cenário {cenario}")
    margem_media = np.mean(indicadores[cenario]["Margem Líquida (%)"])
    retorno_medio = np.mean(indicadores[cenario]["Retorno por Real Gasto"])
    liquidez_media = np.mean(indicadores[cenario]["Liquidez Operacional"])
    endividamento_medio = np.mean(indicadores[cenario]["Endividamento (%)"])
    produtividade_media = np.mean(indicadores[cenario]["Produtividade por Hectare (R$/ha)"])
    custo_receita_media = np.mean(indicadores[cenario]["Custo por Receita (%)"])
    dscr_medio = np.mean([x for x in indicadores[cenario]["DSCR"] if x != float("inf")])
    break_even_media = np.mean(indicadores[cenario]["Break-Even Yield (sacas/ha)"])
    roa_medio = np.mean(indicadores[cenario]["ROA (%)"])

    parecer = []
    if margem_media < 10:
        parecer.append(f"Margem Líquida Baixa ({margem_media:.2f}%): Rentabilidade abaixo do ideal. Considere renegociar preços com fornecedores ou investir em culturas de maior valor agregado.")
    else:
        parecer.append(f"Margem Líquida Saudável ({margem_media:.2f}%): Boa rentabilidade. Monitore custos para manter a consistência.")

    if retorno_medio < 0.2:
        parecer.append(f"Retorno por Real Gasto Baixo ({retorno_medio:.2f}): Gastos com baixo retorno. Avalie a redução de despesas operacionais ou otimize processos agrícolas.")
    else:
        parecer.append(f"Retorno por Real Gasto Adequado ({retorno_medio:.2f}): Investimentos geram retorno satisfatório. Considere reinvestir em tecnologia para aumentar a produtividade.")

    if liquidez_media < 1.5:
        parecer.append(f"Liquidez Operacional Baixa ({liquidez_media:.2f}): Risco de dificuldades para cobrir custos operacionais. Negocie prazos de pagamento ou busque linhas de crédito de curto prazo.")
    else:
        parecer.append(f"Liquidez Operacional Confortável ({liquidez_media:.2f}): Boa capacidade de sustentar operações. Mantenha reservas para safras incertas.")

    if endividamento_medio > 30:
        parecer.append(f"Alto Endividamento ({endividamento_medio:.2f}%): Dívidas elevadas. Priorize a quitação de empréstimos de alto custo ou renegocie taxas de juros.")
    else:
        parecer.append(f"Endividamento Controlado ({endividamento_medio:.2f}%): Dívidas em nível gerenciável. Considere investimentos estratégicos, como expansão de área plantada.")

    if produtividade_media < media_receita_hectare:
        parecer.append(f"Produtividade por Hectare Baixa ({format_brl(produtividade_media)}): A receita por hectare está abaixo da média esperada. Avalie técnicas de cultivo ou rotação de culturas.")
    else:
        parecer.append(f"Produtividade por Hectare Boa ({format_brl(produtividade_media)}): Boa eficiência no uso da terra. Considere investir em tecnologia para manter ou aumentar a produtividade.")

    if custo_receita_media > 70:
        parecer.append(f"Custo por Receita Alto ({custo_receita_media:.2f}%): Custos operacionais consomem grande parte da receita. Analise insumos e processos para reduzir despesas.")
    else:
        parecer.append(f"Custo por Receita Controlado ({custo_receita_media:.2f}%): Boa gestão de custos. Continue monitorando preços de insumos.")

    if dscr_medio < 1.25 and dscr_medio != float("inf"):
        parecer.append(f"DSCR Baixo ({dscr_medio:.2f}): Risco de dificuldades no pagamento de dívidas. Considere reestruturar financiamentos ou aumentar a receita.")
    else:
        parecer.append(f"DSCR Adequado ({dscr_medio:.2f}): Boa capacidade de cobrir dívidas. Mantenha o lucro operacional estável.")

    if break_even_media > (total_sacas / hectares_total) * 0.8:
        parecer.append(f"Break-Even Yield Alto ({break_even_media:.2f} sacas/ha): Alta dependência de produtividade para cobrir custos. Considere culturas mais resilientes ou seguros agrícolas.")
    else:
        parecer.append(f"Break-Even Yield Seguro ({break_even_media:.2f} sacas/ha): Margem de segurança confortável contra falhas na safra.")

    if roa_medio < 5:
        parecer.append(f"ROA Baixo ({roa_medio:.2f}%): Baixa eficiência no uso de ativos. Avalie a venda de ativos ociosos ou investimentos em equipamentos mais produtivos.")
    else:
        parecer.append(f"ROA Adequado ({roa_medio:.2f}%): Boa utilização dos ativos. Considere expansão controlada ou modernização.")

    if indicadores[cenario]["CAGR Lucro Líquido (%)"] < 0:
        parecer.append(f"Crescimento Negativo do Lucro ({indicadores[cenario]['CAGR Lucro Líquido (%)']:.2f}%): Lucro em queda. Revisar estratégias de custo, preço e produtividade.")
    else:
        parecer.append(f"Crescimento do Lucro ({indicadores[cenario]['CAGR Lucro Líquido (%)']:.2f}%): Lucro em trajetória positiva. Considere reinvestir em áreas estratégicas.")

    st.markdown("\n".join([f"- {item}" for item in parecer]))

# Exportação para Excel
st.markdown("### ⬇️ Exportar Indicadores")

output_excel = BytesIO()
with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
    for cenario in nomes_cenarios:
        df_indicadores = pd.DataFrame({
            k: v for k, v in indicadores[cenario].items()
            if k not in ["CAGR Receita (%)", "CAGR Lucro Líquido (%)"]
        }, index=anos)
        df_indicadores.loc["CAGR"] = [indicadores[cenario]["CAGR Receita (%)"], indicadores[cenario]["CAGR Lucro Líquido (%)"]] + [np.nan] * (len(df_indicadores.columns) - 2)
        df_indicadores.to_excel(writer, sheet_name=f"Indicadores_{cenario}")
        df_dre = pd.DataFrame(dre_por_cenario[cenario], index=anos).T
        df_dre.to_excel(writer, sheet_name=f"DRE_{cenario}")
    df_culturas.to_excel(writer, sheet_name="Receita_por_Cultura")
output_excel.seek(0)
st.download_button(
    label="Baixar Indicadores em Excel",
    data=output_excel,
    file_name="indicadores_financeiros.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)