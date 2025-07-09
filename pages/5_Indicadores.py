import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from io import BytesIO

st.set_page_config(layout="wide", page_title="Indicadores Financeiros")
st.title("📈 Indicadores Financeiros e Análise")

# Explicações dos Indicadores
with st.expander("🧾 Entenda os Indicadores Financeiros"):
    st.markdown("""
    Abaixo, explicamos cada indicador financeiro apresentado, sua importância e como interpretá-lo no contexto do agronegócio:

    **1. Margem Líquida (%)**
    - **O que é?** Representa a porcentagem do lucro líquido (após todas as despesas e impostos) em relação à receita total.
    - **Por que é importante?** Indica a rentabilidade do negócio. Uma margem líquida alta sugere que a operação é eficiente em converter receitas em lucro.
    - **Exemplo:** Se a receita é R\$ 100.000 e o lucro líquido é R$ 20.000, a margem líquida é 20%. No agronegócio, margens podem variar devido a custos sazonais (ex.: insumos, colheita).
    - **Interpretação:** Margens abaixo de 10% podem indicar necessidade de reduzir custos ou aumentar preços de venda.

    **2. Retorno por Real Gasto**
    - **O que é?** Mede quanto cada real gasto (em despesas e impostos) gera de lucro líquido.
    - **Por que é importante?** Avalia a eficiência dos investimentos. Um valor alto indica que os gastos estão gerando bons retornos.
    - **Exemplo:** Se você gasta R\$ 50.000 e obtém R$ 10.000 de lucro líquido, o retorno é 0,2 (ou 20 centavos por real gasto). No agronegócio, isso reflete a eficiência de gastos com insumos, mão de obra, etc.
    - **Interpretação:** Valores abaixo de 0,2 sugerem ineficiência nos gastos; avalie despesas não essenciais.

    **3. Liquidez Operacional**
    - **O que é?** Indica quantas vezes a receita cobre as despesas operacionais (ex.: insumos, manutenção).
    - **Por que é importante?** Mostra a capacidade do negócio de sustentar suas operações sem depender de financiamentos. No agronegócio, é crucial para lidar com safras incertas.
    - **Exemplo:** Se a receita é R\$ 100.000 e as despesas operacionais são R$ 40.000, a liquidez é 2,5 (a receita cobre 2,5 vezes as despesas).
    - **Interpretação:** Valores abaixo de 1,5 indicam risco de dificuldades para cobrir custos operacionais.

    **4. Endividamento (%)**
    - **O que é?** Representa a proporção das parcelas de empréstimos em relação à receita total.
    - **Por que é importante?** Avalia o peso das dívidas no faturamento. No agronegócio, empréstimos são comuns para custeio, mas níveis altos podem comprometer a saúde financeira.
    - **Exemplo:** Se as parcelas anuais de empréstimos somam R\$ 30.000 e a receita é R$ 100.000, o endividamento é 30%.
    - **Interpretação:** Valores acima de 30% sugerem alto risco financeiro; priorize redução de dívidas.

    **5. CAGR Receita (%)**
    - **O que é?** Taxa de Crescimento Anual Composta da receita ao longo de 5 anos, considerando a inflação.
    - **Por que é importante?** Mostra a tendência de crescimento do faturamento, útil para planejar expansões ou investimentos no agronegócio.
    - **Exemplo:** Se a receita inicial é R\$ 100.000 e, após 5 anos, é R$ 130.000, o CAGR é cerca de 5,4% ao ano.
    - **Interpretação:** Valores positivos indicam crescimento; valores negativos sugerem retração no faturamento.

    **6. CAGR Lucro Líquido (%)**
    - **O que é?** Taxa de Crescimento Anual Composta do lucro líquido ao longo de 5 anos.
    - **Por que é importante?** Reflete a sustentabilidade do lucro. No agronegócio, é afetado por preços de mercado, custos e condições climáticas.
    - **Exemplo:** Se o lucro líquido inicial é R\$ 20.000 e, após 5 anos, é R$ 30.000, o CAGR é cerca de 8,4% ao ano.
    - **Interpretação:** Valores negativos indicam queda no lucro, exigindo revisão de custos ou estratégias de venda.
    """)

# Verificação de dados essenciais
if "fluxo_caixa" not in st.session_state:
    st.warning("Dados do fluxo de caixa não foram encontrados. Preencha as informações na página de Fluxo de Caixa.")
    st.stop()
if "plantios" not in st.session_state or not st.session_state["plantios"]:
    st.warning("Nenhum plantio cadastrado. Cadastre ao menos um plantio para gerar os indicadores.")
    st.stop()

anos = [f"Ano {i+1}" for i in range(5)]
inflacoes = [st.session_state.get(f"inf_{i}", 4.0) for i in range(5)]
nomes_cenarios = ["Projetado", "Pessimista", "Otimista"]

# Ajustes de cenários (valores padrão se não estiverem no session_state)
pess_receita = st.session_state.get("pess_receita", 15)
pess_despesas = st.session_state.get("pess_despesas", 10)
otm_receita = st.session_state.get("otm_receita", 10)
otm_despesas = st.session_state.get("otm_despesas", 10)

# Função para formatar valores em BRL
def format_brl(x):
    try:
        return f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return x

# Função para calcular CAGR
def calcular_cagr(valor_inicial, valor_final, periodos):
    if valor_inicial == 0 or valor_final <= 0 or periodos <= 0:
        return 0
    return ((valor_final / valor_inicial) ** (1 / periodos) - 1) * 100

# Calcular receitas base
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

media_receita_hectare = (preco_total / total_sacas) * (total_sacas / hectares_total)
receita_base = [hectares_total * media_receita_hectare * np.prod([1 + inflacoes[j] / 100 for j in range(i + 1)]) for i in range(5)]
receitas = {
    "Projetado": receita_base,
    "Pessimista": [r * (1 - pess_receita / 100) for r in receita_base],
    "Otimista": [r * (1 + otm_receita / 100) for r in receita_base]
}

# Ajustar despesas para cenários
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

# Calcular indicadores e DRE
indicadores = {cenario: {} for cenario in nomes_cenarios}
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
        return [total * np.prod([1 + inflacoes[j] / 100 for j in range(i + 1)]) * (1 + ajuste / 100) for i in range(5)]

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

    # Indicadores
    indicadores[cenario]["Margem Líquida (%)"] = [
        l / r * 100 if r > 0 else 0 for l, r in zip(dre_calc["Lucro Líquido"], dre_calc["Receita"])
    ]
    despesas_totais = [
        dre_calc["Impostos Sobre Venda"][i] + dre_calc["Despesas Operacionais"][i] +
        dre_calc["Despesas Administrativas"][i] + dre_calc["Despesas RH"][i] +
        dre_calc["Despesas Extra Operacional"][i] + dre_calc["Dividendos"][i] +
        dre_calc["Impostos Sobre Resultado"][i]
        for i in range(5)
    ]
    indicadores[cenario]["Retorno por Real Gasto"] = [
        l / d if d > 0 else 0 for l, d in zip(dre_calc["Lucro Líquido"], despesas_totais)
    ]
    indicadores[cenario]["Liquidez Operacional"] = [
        r / d if d > 0 else 0 for r, d in zip(dre_calc["Receita"], dre_calc["Despesas Operacionais"])
    ]
    endividamento = [
        sum(emp["valor_parcela"] for emp in st.session_state.get("emprestimos", [])
            if i >= anos.index(emp["ano_inicial"]) and i <= anos.index(emp["ano_final"])) / r if r > 0 else 0
        for i, r in enumerate(dre_calc["Receita"])
    ]
    indicadores[cenario]["Endividamento (%)"] = [e * 100 for e in endividamento]
    indicadores[cenario]["CAGR Receita (%)"] = calcular_cagr(dre_calc["Receita"][0], dre_calc["Receita"][-1], len(anos) - 1)
    indicadores[cenario]["CAGR Lucro Líquido (%)"] = calcular_cagr(dre_calc["Lucro Líquido"][0], dre_calc["Lucro Líquido"][-1], len(anos) - 1)

# Exibição de indicadores
st.markdown("### 📊 Indicadores Financeiros")
for cenario in nomes_cenarios:
    st.subheader(f"Cenário {cenario}")
    # Criar DataFrame de indicadores, excluindo CAGRs para exibição anual
    df_indicadores = pd.DataFrame({
        k: v for k, v in indicadores[cenario].items()
        if k not in ["CAGR Receita (%)", "CAGR Lucro Líquido (%)"]
    }, index=anos)
    # Adicionar CAGRs como uma linha separada
    df_cagr = pd.DataFrame({
        "CAGR Receita (%)": [indicadores[cenario]["CAGR Receita (%)"]],
        "CAGR Lucro Líquido (%)": [indicadores[cenario]["CAGR Lucro Líquido (%)"]]
    }, index=["CAGR"])
    # Concatenar DataFrames
    df_display = pd.concat([df_indicadores, df_cagr])
    # Estilizar e formatar
    styled_df = df_display.style.format({
        "Margem Líquida (%)": "{:.2f}%",
        "Retorno por Real Gasto": "{:.2f}",
        "Liquidez Operacional": "{:.2f}",
        "Endividamento (%)": "{:.2f}%",
        "CAGR Receita (%)": "{:.2f}%",
        "CAGR Lucro Líquido (%)": "{:.2f}%"
    })
    # Usar CSS para esconder o índice
    st.markdown(
        """
        <style>
        .dataframe th:first-child {
            display: none;
        }
        .dataframe td:first-child {
            display: none;
        }
        </style>
        """,
        unsafe_allow_html=True
    )
    st.dataframe(styled_df, use_container_width=True)

# Gráficos
st.markdown("### 📈 Visualizações")
col1, col2 = st.columns(2)

with col1:
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

with col2:
    st.subheader("Margem Líquida ao Longo do Tempo")
    fig2 = px.line(
        pd.DataFrame({cenario: indicadores[cenario]["Margem Líquida (%)"] for cenario in nomes_cenarios}, index=anos),
        title="Evolução da Margem Líquida",
        labels={"value": "Margem Líquida (%)", "index": "Ano"}
    )
    fig2.update_layout(template="plotly_white")
    st.plotly_chart(fig2, use_container_width=True)

st.subheader("Distribuição de Despesas (Cenário Projetado)")
despesas_categorias = pd.DataFrame(st.session_state.get("despesas", [])).groupby("Categoria")["Valor"].sum()
if not despesas_categorias.empty:
    fig3 = px.pie(values=despesas_categorias.values, names=despesas_categorias.index, title="Distribuição de Despesas por Categoria")
    fig3.update_layout(template="plotly_white")
    st.plotly_chart(fig3, use_container_width=True)
else:
    st.warning("Nenhuma despesa cadastrada para exibir a distribuição.")

st.subheader("Endividamento por Ano")
fig4 = go.Figure()
for cenario in nomes_cenarios:
    fig4.add_trace(go.Bar(
        x=anos, y=indicadores[cenario]["Endividamento (%)"], name=f"Endividamento ({cenario})",
        marker_color="#1f77b4" if cenario == "Projetado" else "#ff7f0e" if cenario == "Pessimista" else "#2ca02c"
    ))
fig4.update_layout(barmode="stack", title="Endividamento por Ano", yaxis_title="Endividamento (%)", template="plotly_white")
st.plotly_chart(fig4, use_container_width=True)

# Pareceres
st.markdown("### 📝 Parecer Financeiro")
for cenario in nomes_cenarios:
    st.subheader(f"Parecer - Cenário {cenario}")
    margem_media = np.mean(indicadores[cenario]["Margem Líquida (%)"])
    retorno_medio = np.mean(indicadores[cenario]["Retorno por Real Gasto"])
    liquidez_media = np.mean(indicadores[cenario]["Liquidez Operacional"])
    endividamento_medio = np.mean(indicadores[cenario]["Endividamento (%)"])
    
    parecer = []
    if margem_media < 10:
        parecer.append(f"**Margem Líquida Baixa ({margem_media:.2f}%):** A rentabilidade está abaixo do ideal. Considere reduzir custos operacionais ou aumentar preços de venda.")
    else:
        parecer.append(f"**Margem Líquida Saudável ({margem_media:.2f}%):** Boa rentabilidade. Continue monitorando para manter consistência.")
    
    if retorno_medio < 0.2:
        parecer.append(f"**Retorno por Real Gasto Baixo ({retorno_medio:.2f}):** Cada real investido gera pouco retorno. Avalie despesas não essenciais ou renegocie contratos.")
    else:
        parecer.append(f"**Retorno por Real Gasto Adequado ({retorno_medio:.2f}):** Investimentos estão gerando retorno satisfatório.")
    
    if liquidez_media < 1.5:
        parecer.append(f"**Liquidez Operacional Baixa ({liquidez_media:.2f}):** Receitas podem não cobrir despesas operacionais. Considere renegociar prazos de pagamento ou aumentar vendas.")
    else:
        parecer.append(f"**Liquidez Operacional Confortável ({liquidez_media:.2f}):** Boa capacidade de cobrir despesas operacionais.")
    
    if endividamento_medio > 30:
        parecer.append(f"**Alto Endividamento ({endividamento_medio:.2f}%):** Nível de dívida elevado em relação à receita. Priorize redução de empréstimos ou aumento da receita.")
    else:
        parecer.append(f"**Endividamento Controlado ({endividamento_medio:.2f}%):** Dívidas estão em nível gerenciável.")
    
    if indicadores[cenario]["CAGR Lucro Líquido (%)"] < 0:
        parecer.append(f"**Crescimento Negativo do Lucro ({indicadores[cenario]['CAGR Lucro Líquido (%)']:.2f}%):** Atenção! Lucro está diminuindo. Revisar estratégia de custos e vendas.")
    else:
        parecer.append(f"**Crescimento do Lucro ({indicadores[cenario]['CAGR Lucro Líquido (%)']:.2f}%):** Lucro em trajetória positiva.")
    
    st.markdown("\n".join(parecer))

# Exportação para Excel
st.markdown("### ⬇️ Exportar Indicadores")
output = BytesIO()
with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
    for cenario in nomes_cenarios:
        df_indicadores = pd.DataFrame({
            k: v for k, v in indicadores[cenario].items()
            if k not in ["CAGR Receita (%)", "CAGR Lucro Líquido (%)"]
        }, index=anos)
        df_indicadores.loc["CAGR"] = [indicadores[cenario]["CAGR Receita (%)"], indicadores[cenario]["CAGR Lucro Líquido (%)"]] + [np.nan] * (len(df_indicadores.columns) - 2)
        df_indicadores.to_excel(writer, sheet_name=f"Indicadores_{cenario}")
        df_dre = pd.DataFrame(dre_por_cenario[cenario], index=anos).T
        df_dre.to_excel(writer, sheet_name=f"DRE_{cenario}")
output.seek(0)
st.download_button(
    label="Baixar Indicadores em Excel",
    data=output,
    file_name="indicadores_financeiros.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)