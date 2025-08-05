import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from io import BytesIO
import base64
from datetime import datetime

# Importa as configurações de sessão e a função de cálculo do DRE existente
from utils.session import carregar_configuracoes
from utils.dre import calcular_dre # <--- ESSENCIAL: Reutilizar a função existente!

carregar_configuracoes()

st.set_page_config(layout="wide", page_title="Indicadores Financeiros")
st.title("📈 Indicadores Financeiros e Análise - Agronegócio")

# --- Funções Auxiliares ---

def format_brl(x):
    """Formata um número para o formato de moeda brasileira (R\$)."""
    try:
        # Verifica se x é um número antes de formatar
        if pd.isna(x) or not isinstance(x, (int, float)):
            return x
        return f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return x # Retorna o valor original em caso de erro

def calcular_cagr(valor_inicial, valor_final, periodos):
    """Calcula a Taxa de Crescimento Anual Composta (CAGR)."""
    if valor_inicial <= 0 or periodos <= 0:
        return 0.0 # Evita divisão por zero ou log de números não positivos
    if valor_final <= 0: # Se o valor final for negativo, é um declínio
        # Calcula a taxa de declínio como um CAGR negativo
        # Usamos o valor absoluto para o cálculo da base, mas o resultado é negativo
        return ((abs(valor_final) / valor_inicial) ** (1 / periodos) - 1) * -100
    return ((valor_final / valor_inicial) ** (1 / periodos) - 1) * 100

def display_indicator_explanation():
    """Exibe a seção de explicação dos indicadores financeiros."""
    with st.expander("🧾 Entenda os Indicadores Financeiros"):
        st.markdown("""
        Abaixo, apresentamos indicadores financeiros avançados, adaptados ao contexto do agronegócio, com explicações sobre sua importância e interpretação:

        **1. Margem Líquida (%)**
        - **O que é?** Percentual do lucro líquido em relação à receita total após todas as despesas e impostos.
        - **Por que é importante?** Mede a rentabilidade líquida do negócio. No agronegócio, margens são impactadas por custos sazonais (insumos, colheita) e preços de mercado.
        - **Exemplo:** Receita de R\$ 1.000.000 e lucro líquido de R\$ 150.000 resultam em 15% de margem.
        - **Interpretação:** Margens abaixo de 10% sugerem necessidade de otimizar custos ou preços.

        **2. Retorno por Real Gasto**
        - **O que é?** Quanto cada real gasto (despesas + impostos) gera de lucro líquido.
        - **Por que é importante?** Avalia a eficiência dos investimentos, crucial em um setor com altos custos fixos e variáveis.
        - **Exemplo:** Gastos de R\$ 500.000 geram R\$ 100.000 de lucro (retorno de 0,2).
        - **Interpretação:** Valores abaixo de 0,2 indicam ineficiência; analise despesas operacionais.

        **3. Liquidez Operacional**
        - **O que é?** Quantas vezes a receita cobre as despesas operacionais.
        - **Por que é importante?** Mostra a capacidade de sustentar operações sem depender de financiamentos, essencial em safras incertas.
        - **Exemplo:** Receita de R\$ 1.000.000 e despesas operacionais de R\$ 400.000 resultam em liquidez de 2,5.
        - **Interpretação:** Valores abaixo de 1,5 indicam risco de fluxo de caixa.

        **4. Endividamento (%)**
        - **O que é?** Proporção das parcelas de empréstimos em relação à receita total.
        - **Por que é importante?** Avalia o peso das dívidas, comum no agronegócio para custeio e investimentos.
        - **Exemplo:** Parcelas de R\$ 200.000 com receita de R\$ 1.000.000 resultam em 20% de endividamento.
        - **Interpretação:** Acima de 30% é arriscado; priorize redução de dívidas.

        **5. Produtividade por Hectare (R\$/ha)**
        - **O que é?** Receita gerada por hectare plantado.
        - **Por que é importante?** Mede a eficiência do uso da terra, um recurso crítico no agronegócio.
        - **Exemplo:** Receita de R\$ 1.000.000 em 500 ha resulta em R\$ 2.000/ha.
        - **Interpretação:** Valores baixos podem indicar baixa produtividade ou preços de mercado desfavoráveis.

        **6. Custo por Receita (%)**
        - **O que é?** Proporção dos custos operacionais em relação à receita.
        - **Por que é importante?** Indica a eficiência na gestão de custos, essencial em um setor com margens apertadas.
        - **Exemplo:** Custos operacionais de R\$ 600.000 e receita de R\$ 1.000.000 resultam em 60%.
        - **Interpretação:** Acima de 70% sugere necessidade de redução de custos.

        **7. Debt Service Coverage Ratio (DSCR)**
        - **O que é?** Razão entre o lucro operacional e as parcelas anuais de dívidas.
        - **Por que é importante?** Avalia a capacidade de pagar dívidas com o lucro gerado, crucial para financiamentos agrícolas.
        - **Exemplo:** Lucro operacional de R\$ 300.000 e parcelas de R\$ 150.000 resultam em DSCR de 2,0.
        - **Interpretação:** Valores abaixo de 1,25 indicam risco de inadimplência.

        **8. Break-Even Yield (sacas/ha)**
        - **O que é?** Produtividade mínima por hectare necessária para cobrir todos os custos.
        - **Por que é importante?** Ajuda a avaliar o risco de safras em cenários de baixa produtividade.
        - **Exemplo:** Custos totais de R\$ 1.000.000, 500 ha e preço da saca de R\$ 100 resultam em 20 sacas/ha.
        - **Interpretação:** Valores altos indicam maior vulnerabilidade a falhas na safra.

        **9. Return on Assets (ROA) (%)**
        - **O que é?** Percentual do lucro líquido em relação ao total de ativos (estimado).
        - **Por que é importante?** Mede a eficiência do uso de ativos (terra, máquinas) para gerar lucro.
        - **Exemplo:** Lucro líquido de R\$ 150.000 e ativos de R\$ 5.000.000 resultam em 3% de ROA.
        - **Interpretação:** Valores abaixo de 5% sugerem baixa eficiência no uso de ativos.

        **10. CAGR Receita (%)**
        - **O que é?** Taxa de crescimento anual composta da receita ao longo de 5 anos.
        - **Por que é importante?** Indica a tendência de crescimento do faturamento, útil para planejamento.
        - **Exemplo:** Receita inicial de R\$ 1.000.000 e final de R\$ 1.300.000 resultam em ~5,4%.
        - **Interpretação:** Valores negativos indicam retração; revise preços ou produtividade.

        **11. CAGR Lucro Líquido (%)**
        - **O que é?** Taxa de crescimento anual composta do lucro líquido ao longo de 5 anos.
        - **Por que é importante?** Reflete a sustentabilidade do lucro em um setor volátil.
        - **Exemplo:** Lucro inicial de R\$ 100.000 e final de R\$ 150.000 resultam em ~8,4%.
        - **Interpretação:** Valores negativos requerem revisão de custos ou estratégias.
        """)

def get_base_financial_data():
    """
    Consolida a extração de dados financeiros base e o cálculo de receitas.
    Retorna um dicionário com hectares_total, total_sacas, preco_total_base,
    receitas_cenarios, receitas_extras_projetadas e total_ativos.
    """
    required_keys = ["plantios", "dre_cenarios", "receitas_cenarios", "inflacoes", "anos"]
    for key in required_keys:
        if key not in st.session_state:
            st.warning(f"Dados essenciais não encontrados: '{key}'. Por favor, verifique as páginas anteriores (especialmente a de 'Fluxo de Caixa').")
            st.stop()
    if not st.session_state["plantios"]:
        st.warning("Nenhum plantio cadastrado. Cadastre ao menos um plantio para gerar os indicadores.")
        st.stop()

    plantios = st.session_state["plantios"]
    inflacoes = st.session_state["inflacoes"]
    anos = st.session_state["anos"]

    hectares_total = 0
    total_sacas = 0
    preco_total_base = 0 # Valor total da produção no ano base (sem inflação)

    # Calcula totais de plantio
    for p_data in plantios.values():
        hectares = p_data.get("hectares", 0)
        sacas = p_data.get("sacas_por_hectare", 0)
        preco = p_data.get("preco_saca", 0)
        hectares_total += hectares
        total_sacas += sacas * hectares
        preco_total_base += preco * sacas * hectares # Soma de (preço * sacas * hectares)

    if hectares_total == 0 or total_sacas == 0:
        st.error("Dados de plantio incompletos para estimar receita e indicadores. Verifique o cadastro de plantios.")
        st.stop()

    # Estimativa de ativos totais (pode ser refinada ou configurável pelo usuário)
    # Estimativa aproximada: R\$20.000/ha para terra + R\$1.000.000 para ativos fixos (máquinas, etc.)
    total_ativos = hectares_total * 20000 + 1000000

    # Receitas já calculadas e armazenadas em 4_Fluxo_de_Caixa.py
    receitas_cenarios = st.session_state["receitas_cenarios"]

    # Receitas extras projetadas (também já calculadas em 4_Fluxo_de_Caixa.py)
    # Precisamos recriar a estrutura aqui, pois o dre_calc usa um formato diferente
    receitas_extras_projetadas = {"Operacional": [0] * len(anos), "Extra Operacional": [0] * len(anos)}
    if "receitas_adicionais" in st.session_state:
        for receita_add in st.session_state["receitas_adicionais"].values():
            valor = receita_add["valor"]
            categoria = receita_add["categoria"]
            for ano_aplicacao in receita_add["anos_aplicacao"]:
                try:
                    idx = anos.index(ano_aplicacao)
                    if categoria == "Operacional":
                        # Aplica inflação para receitas operacionais
                        fator = np.prod([1 + inflacoes[j] for j in range(idx + 1)])
                        receitas_extras_projetadas["Operacional"][idx] += valor * fator
                    else:
                        # Receitas extra operacionais não sofrem ajuste de inflação
                        receitas_extras_projetadas["Extra Operacional"][idx] += valor
                except ValueError:
                    # Ano não encontrado, ignora
                    continue

    return {
        "plantios": plantios,
        "dre_cenarios": st.session_state["dre_cenarios"],
        "receitas_cenarios": receitas_cenarios,
        "inflacoes": inflacoes,
        "anos": anos,
        "pess_receita": st.session_state.get("pess_receita", 15),
        "pess_despesas": st.session_state.get("pess_despesas", 10),
        "otm_receita": st.session_state.get("otm_receita", 10),
        "otm_despesas": st.session_state.get("otm_despesas", 10),
        "despesas_info": pd.DataFrame(st.session_state.get("despesas", [])),
        "emprestimos": st.session_state.get("emprestimos", []),
        "hectares_total": hectares_total,
        "total_sacas": total_sacas,
        "preco_total_base": preco_total_base, # Valor total da produção no ano base
        "total_ativos": total_ativos,
        "receitas_extras_projetadas": receitas_extras_projetadas
    }

def calculate_indicators_for_scenario(scenario_name, dre_data, session_data):
    """Calcula todos os indicadores financeiros para um dado cenário."""
    anos = session_data["anos"]
    hectares_total = session_data["hectares_total"]
    total_sacas = session_data["total_sacas"]
    preco_total_base = session_data["preco_total_base"]
    # Preço médio por saca no ano base (para Break-Even Yield)
    preco_medio_saca_base = preco_total_base / total_sacas if total_sacas > 0 else 0
    total_ativos = session_data["total_ativos"]

    # Extrair dados relevantes do DRE
    receita = dre_data["Receita"]
    impostos_sobre_venda = dre_data["Impostos Sobre Venda"]
    despesas_operacionais = dre_data["Despesas Operacionais"]
    despesas_administrativas = dre_data["Despesas Administrativas"]
    despesas_rh = dre_data["Despesas RH"]
    despesas_extra_operacional = dre_data["Despesas Extra Operacional"]
    dividendos = dre_data["Dividendos"]
    impostos_sobre_resultado = dre_data["Impostos Sobre Resultado"]
    lucro_liquido = dre_data["Lucro Líquido"]
    lucro_operacional = dre_data["Lucro Operacional"] # Já vem calculado do dre.py

    # Calcular despesas totais para cada ano
    despesas_totais = [
        impostos_sobre_venda[i] +
        despesas_operacionais[i] +
        despesas_administrativas[i] +
        despesas_rh[i] +
        despesas_extra_operacional[i] +
        dividendos[i] +
        impostos_sobre_resultado[i]
        for i in range(len(anos))
    ]

    indicators = {}

    # 1. Margem Líquida (%)
    indicators["Margem Líquida (%)"] = [
        (l / r * 100) if r != 0 else 0 for l, r in zip(lucro_liquido, receita)
    ]

    # 2. Retorno por Real Gasto
    indicators["Retorno por Real Gasto"] = [
        (l / d) if d != 0 else 0 for l, d in zip(lucro_liquido, despesas_totais)
    ]

    # 3. Liquidez Operacional (Receita / Despesas Operacionais)
    indicators["Liquidez Operacional"] = [
        (r / d) if d != 0 else 0 for r, d in zip(receita, despesas_operacionais)
    ]

    # 4. Endividamento (%) (Total Parcelas Empréstimos / Receita Total)
    # Usando Despesas Extra Operacional como proxy para o serviço da dívida,
    # pois o dre.py já incorpora os ajustes de cenário para empréstimos nessa linha.
    endividamento_anual = []
    for i in range(len(anos)):
        # Assumindo que Despesas Extra Operacional inclui principalmente pagamentos de empréstimos
        # Se houver outras despesas extra operacionais significativas, isso pode ser menos preciso.
        total_servico_divida_ano = despesas_extra_operacional[i]
        endividamento_anual.append((total_servico_divida_ano / receita[i] * 100) if receita[i] != 0 else 0)
    indicators["Endividamento (%)"] = endividamento_anual

    # 5. Produtividade por Hectare (R\$/ha)
    indicators["Produtividade por Hectare (R$/ha)"] = [
        (r / hectares_total) if hectares_total != 0 else 0 for r in receita
    ]

    # 6. Custo por Receita (%)
    indicators["Custo por Receita (%)"] = [
        (d / r * 100) if r != 0 else 0 for d, r in zip(despesas_operacionais, receita)
    ]

    # 7. Debt Service Coverage Ratio (DSCR)
    # DSCR = Lucro Operacional / Total Debt Service (pagamentos de empréstimos)
    dscr_anual = []
    for i in range(len(anos)):
        total_servico_divida_ano = despesas_extra_operacional[i] # Usando proxy novamente
        if total_servico_divida_ano != 0:
            dscr_anual.append(lucro_operacional[i] / total_servico_divida_ano)
        else:
            dscr_anual.append(float("inf")) # Sem serviço da dívida, então a cobertura é infinita
    indicators["DSCR"] = dscr_anual

    # 8. Break-Even Yield (sacas/ha)
    # Break-Even Yield = (Custos Totais / (Hectares * Preço por Saca))
    # Usando o preço médio por saca do ano base para a projeção
    indicators["Break-Even Yield (sacas/ha)"] = [
        (despesas_totais[i] / (hectares_total * preco_medio_saca_base)) if (hectares_total != 0 and preco_medio_saca_base != 0) else 0
        for i in range(len(anos))
    ]

    # 9. Return on Assets (ROA) (%)
    indicators["ROA (%)"] = [
        (l / total_ativos * 100) if total_ativos != 0 else 0 for l in lucro_liquido
    ]

    # 10. CAGR Receita (%)
    indicators["CAGR Receita (%)"] = calcular_cagr(receita[0], receita[-1], len(anos) - 1)

    # 11. CAGR Lucro Líquido (%)
    indicators["CAGR Lucro Líquido (%)"] = calcular_cagr(lucro_liquido[0], lucro_liquido[-1], len(anos) - 1)

    # 12. Custo por Hectare (R\$/ha)
    indicators["Custo por Hectare (R$/ha)"] = [
        (d / hectares_total) if hectares_total != 0 else 0 for d in despesas_totais
    ]

    return indicators

def display_scenario_parameters(session_data):
    with st.expander("###  ⚖️ Inflação e Parâmetros de Cenário Atuais"):
        """Exibe os parâmetros atuais do cenário (inflação, ajustes pessimistas/otimistas)."""
        st.markdown("### 📈 Inflação Estimada por Ano")
        cols = st.columns(len(session_data["anos"]))
        for i, col in enumerate(cols):
            with col:
                st.metric(f"Ano {i+1}", f"{session_data['inflacoes'][i]:.2f}%")

        st.markdown("### 🔧 Parâmetros de Cenário Atuais")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("💸 Receita Pessimista", f"-{session_data['pess_receita']}%")
        with col2:
            st.metric("💰 Despesa Pessimista", f"+{session_data['pess_despesas']}%")
        with col3:
            st.metric("💸 Receita Otimista", f"+{session_data['otm_receita']}%")
        with col4:
            st.metric("💰 Despesa Otimista", f"-{session_data['otm_despesas']}%")

def display_revenue_by_crop(session_data):
    """Exibe o detalhamento da receita por cultura e receitas adicionais."""
    st.markdown("### 💰 Receita por Cultura Agrícola e Receitas Extras (Ano Base)")
    culturas = {}
    for p in session_data["plantios"].values():
        cultura = p["cultura"]
        receita = p["hectares"] * p["sacas_por_hectare"] * p["preco_saca"]
        if cultura not in culturas:
            culturas[cultura] = {"receita": 0, "hectares": 0}
        culturas[cultura]["receita"] += receita
        culturas[cultura]["hectares"] += p["hectares"]

    df_culturas_data = []
    for cultura, dados in culturas.items():
        df_culturas_data.append({
            "Cultura": cultura,
            "Receita Total": dados["receita"],
            "Área (ha)": dados["hectares"],
            "Receita por ha": dados["receita"] / dados["hectares"] if dados["hectares"] != 0 else 0
        })

    # Adicionar receitas adicionais se existirem no session_state
    if "receitas_adicionais" in st.session_state:
        receitas_extras_base = {"Operacional": 0, "Extra Operacional": 0}
        for receita_add in st.session_state["receitas_adicionais"].values():
            valor = receita_add["valor"]
            categoria = receita_add["categoria"]
            # Soma o valor base, não o projetado
            receitas_extras_base[categoria] += valor

        for cat in ["Operacional", "Extra Operacional"]:
            if receitas_extras_base[cat] > 0:
                df_culturas_data.append({
                    "Cultura": f"Receita Extra ({cat})",
                    "Receita Total": receitas_extras_base[cat],
                    "Área (ha)": 0,
                    "Receita por ha": 0
                })

    df_culturas = pd.DataFrame(df_culturas_data)

    st.dataframe(
        df_culturas.style.format({
            "Receita Total": format_brl,
            "Área (ha)": "{:.2f}",
            "Receita por ha": format_brl,
        }),
        use_container_width=True,
        hide_index=True
    )
    return df_culturas # Retorna para uso na exportação

def display_indicators_table(all_indicators, anos):
    """Exibe a tabela formatada de indicadores financeiros para cada cenário."""
    st.markdown("### 📊 Indicadores Financeiros")
    nomes_cenarios = ["Projetado", "Pessimista", "Otimista"]
    for cenario_name in nomes_cenarios:
        indicators = all_indicators[cenario_name]
        st.subheader(f"Cenário {cenario_name}")
        df_indicadores = pd.DataFrame({
            k: v for k, v in indicators.items()
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

        #col_a, col_b = st.columns(2)
        #with col_a:
        #    st.metric("📈 CAGR Receita (5 anos)", f"{indicators['CAGR Receita (%)']:.2f}%")
        #with col_b:
        #    st.metric("📈 CAGR Lucro Líquido (5 anos)", f"{indicators['CAGR Lucro Líquido (%)']:.2f}%")

def display_financial_summary(all_dre_data, anos):
    """Exibe o resumo financeiro anual (Receita, Despesas Totais, Lucro Líquido)."""
    st.markdown("### 📘 Resumo Financeiro por Ano")
    nomes_cenarios = ["Projetado", "Pessimista", "Otimista"]
    for cenario_name in nomes_cenarios:
        dre_data = all_dre_data[cenario_name]
        st.subheader(f"Resumo - Cenário {cenario_name}")

        # Recalcular despesas totais para consistência do resumo
        despesas_totais = [
            dre_data["Impostos Sobre Venda"][i] +
            dre_data["Despesas Operacionais"][i] +
            dre_data["Despesas Administrativas"][i] +
            dre_data["Despesas RH"][i] +
            dre_data["Despesas Extra Operacional"][i] +
            dre_data["Dividendos"][i] +
            dre_data["Impostos Sobre Resultado"][i]
            for i in range(len(anos))
        ]

        resumo = pd.DataFrame({
            "Receita": dre_data["Receita"],
            "Despesas Totais": despesas_totais,
            "Lucro Líquido": dre_data["Lucro Líquido"]
        }, index=anos)

        st.dataframe(
            resumo.style.format({
                "Receita": format_brl,
                "Despesas Totais": format_brl,
                "Lucro Líquido": format_brl,
            }),
            use_container_width=True
        )

def generate_visualizations(all_dre_data, all_indicators, anos, nomes_cenarios, session_data):
    """Gera e exibe os gráficos Plotly."""
    st.markdown("### 📈 Visualizações")

    st.subheader("Receita vs. Lucro Líquido")
    fig1 = go.Figure()
    for cenario in nomes_cenarios:
        fig1.add_trace(go.Bar(
            x=anos, y=all_dre_data[cenario]["Receita"], name=f"Receita ({cenario})",
            marker_color="#1f77b4" if cenario == "Projetado" else "#ff7f0e" if cenario == "Pessimista" else "#2ca02c"
        ))
        fig1.add_trace(go.Bar(
            x=anos, y=all_dre_data[cenario]["Lucro Líquido"], name=f"Lucro Líquido ({cenario})",
            marker_color="#aec7e8" if cenario == "Projetado" else "#ffbb78" if cenario == "Pessimista" else "#98df8a"
        ))
    fig1.update_layout(barmode="group", title="Comparação de Receita e Lucro Líquido", yaxis_title="R$", template="plotly_white")
    st.plotly_chart(fig1, use_container_width=True)

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Distribuição de Despesas (Cenário Projetado)")
        despesas_info = session_data.get("despesas_info") # Usar o df_despesas_info do session_data
        if not despesas_info.empty:
            df_despesas_categorias = despesas_info.groupby("Categoria")["Valor"].sum()
            if not df_despesas_categorias.empty:
                fig4 = px.pie(values=df_despesas_categorias.values, names=df_despesas_categorias.index, title="Distribuição de Despesas por Categoria")
                fig4.update_layout(template="plotly_white")
                st.plotly_chart(fig4, use_container_width=True)
            else:
                st.warning("Nenhuma despesa cadastrada para exibir a distribuição.")
        else:
            st.warning("Nenhuma despesa cadastrada para exibir a distribuição.")


    with col2:
        st.subheader("Margem Líquida vs. Custo por Receita")
        fig2 = go.Figure()
        for cenario in nomes_cenarios:
            fig2.add_trace(go.Scatter(
                x=anos, y=all_indicators[cenario]["Margem Líquida (%)"], name=f"Margem Líquida ({cenario})",
                line=dict(color="#1f77b4" if cenario == "Projetado" else "#ff7f0e" if cenario == "Pessimista" else "#2ca02c")
            ))
            fig2.add_trace(go.Scatter(
                x=anos, y=all_indicators[cenario]["Custo por Receita (%)"], name=f"Custo por Receita ({cenario})",
                line=dict(color="#aec7e8" if cenario == "Projetado" else "#ffbb78" if cenario == "Pessimista" else "#98df8a", dash="dash")
            ))
        fig2.update_layout(title="Margem Líquida vs. Custo por Receita", yaxis_title="Percentual (%)", template="plotly_white")
        st.plotly_chart(fig2, use_container_width=True)

    st.subheader("Produtividade por Hectare e Break-Even Yield")
    fig3 = go.Figure()
    for cenario in nomes_cenarios:
        fig3.add_trace(go.Bar(
            x=anos, y=all_indicators[cenario]["Produtividade por Hectare (R$/ha)"], name=f"Produtividade por Hectare ({cenario})",
            marker_color="#1f77b4" if cenario == "Projetado" else "#ff7f0e" if cenario == "Pessimista" else "#2ca02c"
        ))
        fig3.add_trace(go.Scatter(
            x=anos, y=all_indicators[cenario]["Break-Even Yield (sacas/ha)"], name=f"Break-Even Yield ({cenario})",
            line=dict(color="#aec7e8" if cenario == "Projetado" else "#ffbb78" if cenario == "Pessimista" else "#98df8a")
        ))
    fig3.update_layout(barmode="group", title="Produtividade por Hectare vs. Break-Even Yield", yaxis_title="R$/ha e Sacas/ha", template="plotly_white")
    st.plotly_chart(fig3, use_container_width=True)

def generate_financial_opinion(all_indicators, session_data):
    """Gera um parecer financeiro textual para cada cenário."""
    st.markdown("### 📝 Parecer Financeiro")
    nomes_cenarios = ["Projetado", "Pessimista", "Otimista"]
    hectares_total = session_data["hectares_total"]
    total_sacas = session_data["total_sacas"]
    preco_total_base = session_data["preco_total_base"] # Valor total da produção no ano base
    
    # Calcular a receita média por hectare no ano base para comparação
    media_receita_hectare_base = (preco_total_base / total_sacas) * (total_sacas / hectares_total) if total_sacas != 0 and hectares_total != 0 else 0
    current_avg_yield = total_sacas / hectares_total if hectares_total != 0 else 0

    for cenario in nomes_cenarios:
        st.subheader(f"Parecer - Cenário {cenario}")
        indicators = all_indicators[cenario]

        margem_media = np.mean(indicators["Margem Líquida (%)"])
        retorno_medio = np.mean(indicators["Retorno por Real Gasto"])
        liquidez_media = np.mean(indicators["Liquidez Operacional"])
        endividamento_medio = np.mean(indicators["Endividamento (%)"])
        produtividade_media = np.mean(indicators["Produtividade por Hectare (R$/ha)"])
        custo_receita_media = np.mean(indicators["Custo por Receita (%)"])
        # Filtrar valores inf para média do DSCR
        dscr_values = [x for x in indicators["DSCR"] if x != float("inf")]
        dscr_medio = np.mean(dscr_values) if dscr_values else float("inf")
        break_even_media = np.mean(indicators["Break-Even Yield (sacas/ha)"])
        roa_medio = np.mean(indicators["ROA (%)"])

        parecer = []
        # Limiares para a opinião (podem ser configuráveis)
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

        if produtividade_media < media_receita_hectare_base * 0.8: # Comparar com uma linha de base ou média
            parecer.append(f"Produtividade por Hectare Baixa ({format_brl(produtividade_media)}): A receita por hectare está abaixo da média esperada. Avalie técnicas de cultivo ou rotação de culturas.")
        else:
            parecer.append(f"Produtividade por Hectare Boa ({format_brl(produtividade_media)}): Boa eficiência no uso da terra. Considere investir em tecnologia para manter ou aumentar a produtividade.")

        if custo_receita_media > 70:
            parecer.append(f"Custo por Receita Alto ({custo_receita_media:.2f}%): Custos operacionais consomem grande parte da receita. Analise insumos e processos para reduzir despesas.")
        else:
            parecer.append(f"Custo por Receita Controlado ({custo_receita_media:.2f}%): Boa gestão de custos. Continue monitorando preços de insumos.")

        if dscr_medio != float("inf") and dscr_medio < 1.25:
            parecer.append(f"DSCR Baixo ({dscr_medio:.2f}): Risco de dificuldades no pagamento de dívidas. Considere reestruturar financiamentos ou aumentar a receita.")
        else:
            parecer.append(f"DSCR Adequado ({dscr_medio:.2f}): Boa capacidade de cobrir dívidas. Mantenha o lucro operacional estável.")

        # Comparação do Break-Even Yield: definir o que significa "alto".
        # Vamos comparar com uma porcentagem da produtividade média atual.
        if break_even_media > current_avg_yield * 0.8 and current_avg_yield != 0: # Se o break-even estiver próximo de 80% da produtividade atual
            parecer.append(f"Break-Even Yield Alto ({break_even_media:.2f} sacas/ha): Alta dependência de produtividade para cobrir custos. Considere culturas mais resilientes ou seguros agrícolas.")
        else:
            parecer.append(f"Break-Even Yield Seguro ({break_even_media:.2f} sacas/ha): Margem de segurança confortável contra falhas na safra.")

        if roa_medio < 5:
            parecer.append(f"ROA Baixo ({roa_medio:.2f}%): Baixa eficiência no uso de ativos. Avalie a venda de ativos ociosos ou investimentos em equipamentos mais produtivos.")
        else:
            parecer.append(f"ROA Adequado ({roa_medio:.2f}%): Boa utilização dos ativos. Considere expansão controlada ou modernização.")

        if indicators["CAGR Lucro Líquido (%)"] < 0:
            parecer.append(f"Crescimento Negativo do Lucro ({indicators['CAGR Lucro Líquido (%)']:.2f}%): Lucro em queda. Revisar estratégias de custo, preço e produtividade.")
        else:
            parecer.append(f"Crescimento do Lucro ({indicators['CAGR Lucro Líquido (%)']:.2f}%): Lucro em trajetória positiva. Considere reinvestir em áreas estratégicas.")

        st.markdown("\n".join([f"- {item}" for item in parecer]))

def generate_excel_export(all_indicators, all_dre_data, df_culturas_for_excel, nomes_cenarios, anos):
    """Gera e fornece um botão de download do Excel para todos os dados."""
    st.markdown("### ⬇️ Exportar Relatório Completo")
    
    def criar_relatorio_completo():
        """Cria um arquivo Excel completo com Fluxo de Caixa e Indicadores"""
        output_excel = BytesIO()
        
        with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
            # === DADOS DE INDICADORES ===
            for cenario in nomes_cenarios:
                # Indicadores
                indicators_df_for_excel = pd.DataFrame({
                    k: v for k, v in all_indicators[cenario].items()
                    if k not in ["CAGR Receita (%)", "CAGR Lucro Líquido (%)"]
                }, index=anos)
                # Adicionar CAGR como uma linha separada para clareza no Excel
                cagr_row_data = [all_indicators[cenario]["CAGR Receita (%)"], all_indicators[cenario]["CAGR Lucro Líquido (%)"]] + \
                                [np.nan] * (len(indicators_df_for_excel.columns) - 2)
                cagr_row = pd.Series(cagr_row_data, index=indicators_df_for_excel.columns, name="CAGR")
                indicators_df_for_excel = pd.concat([indicators_df_for_excel, pd.DataFrame(cagr_row).T])
                indicators_df_for_excel.to_excel(writer, sheet_name=f"Indicadores_{cenario}")

                # DRE
                df_dre_for_excel = pd.DataFrame(all_dre_data[cenario], index=anos).T
                df_dre_for_excel.to_excel(writer, sheet_name=f"DRE_{cenario}")

                # Resumo
                despesas_totais_summary = [
                    all_dre_data[cenario]["Impostos Sobre Venda"][i] +
                    all_dre_data[cenario]["Despesas Operacionais"][i] +
                    all_dre_data[cenario]["Despesas Administrativas"][i] +
                    all_dre_data[cenario]["Despesas RH"][i] +
                    all_dre_data[cenario]["Despesas Extra Operacional"][i] +
                    all_dre_data[cenario]["Dividendos"][i] +
                    all_dre_data[cenario]["Impostos Sobre Resultado"][i]
                    for i in range(len(anos))
                ]
                summary_df = pd.DataFrame({
                    "Receita": all_dre_data[cenario]["Receita"],
                    "Despesas Totais": despesas_totais_summary,
                    "Lucro Líquido": all_dre_data[cenario]["Lucro Líquido"]
                }, index=anos)
                summary_df.to_excel(writer, sheet_name=f"Resumo_{cenario}")

            # === DADOS DE FLUXO DE CAIXA ===
            # Incluir dados das despesas se existirem
            if st.session_state.get('fluxo_caixa') is not None and not st.session_state['fluxo_caixa'].empty:
                df_fluxo_despesas = st.session_state['fluxo_caixa'].copy()
                df_fluxo_despesas.to_excel(writer, sheet_name='Fluxo_Despesas', index=True)
                
                # Totais de despesas por ano
                totais_despesas = df_fluxo_despesas.sum(axis=0)
                df_totais_despesas = pd.DataFrame({
                    'Ano': totais_despesas.index,
                    'Total Despesas (R$)': totais_despesas.values
                })
                df_totais_despesas.to_excel(writer, sheet_name='Totais_Despesas', index=False)
            
            # Incluir despesas cadastradas
            if st.session_state.get('despesas'):
                df_despesas_cadastradas = pd.DataFrame(st.session_state['despesas'])
                df_despesas_cadastradas.to_excel(writer, sheet_name='Despesas_Cadastradas', index=False)
            
            # Incluir empréstimos cadastrados
            if st.session_state.get('emprestimos'):
                df_emprestimos_cadastrados = pd.DataFrame(st.session_state['emprestimos'])
                df_emprestimos_cadastrados.to_excel(writer, sheet_name='Emprestimos_Cadastrados', index=False)
        
            # Incluir plantios cadastrados
            if st.session_state.get('plantios'):
                plantios_list = []
                for nome, dados in st.session_state['plantios'].items():
                    plantio_row = {'Nome': nome}
                    plantio_row.update(dados)
                    plantios_list.append(plantio_row)
                df_plantios = pd.DataFrame(plantios_list)
                df_plantios.to_excel(writer, sheet_name='Plantios_Cadastrados', index=False)
            
            # Incluir receitas adicionais se existirem
            if st.session_state.get('receitas_adicionais'):
                receitas_list = []
                for nome, dados in st.session_state['receitas_adicionais'].items():
                    receita_row = {'Nome': nome}
                    receita_row.update(dados)
                    # Converter lista de anos para string
                    if 'anos_aplicacao' in receita_row:
                        receita_row['anos_aplicacao'] = ', '.join(receita_row['anos_aplicacao'])
                    receitas_list.append(receita_row)
                df_receitas_adicionais = pd.DataFrame(receitas_list)
                df_receitas_adicionais.to_excel(writer, sheet_name='Receitas_Adicionais', index=False)

            # Receita por Cultura
            df_culturas_for_excel.to_excel(writer, sheet_name="Receita_por_Cultura", index=False)
            
            # Configurações (Inflação e parâmetros)
            df_config = pd.DataFrame({
                'Ano': anos,
                'Inflacao (%)': st.session_state.get('inflacoes', [4.0] * len(anos))
            })
            df_config.to_excel(writer, sheet_name='Configuracoes_Inflacao', index=False)
            
            df_config_params = pd.DataFrame({
                'Parametro': ['Receita Pessimista (%)', 'Despesa Pessimista (%)', 'Receita Otimista (%)', 'Despesa Otimista (%)'],
                'Valor': [
                    st.session_state.get('pess_receita', 15),
                    st.session_state.get('pess_despesas', 10),
                    st.session_state.get('otm_receita', 10),
                    st.session_state.get('otm_despesas', 10)
                ]
            })
            df_config_params.to_excel(writer, sheet_name='Parametros_Cenario', index=False)

        output_excel.seek(0)
        return output_excel
    
    def criar_relatorio_pdf(all_indicators, all_dre_data, df_culturas_for_excel, nomes_cenarios, anos):
        """Cria um relatório PDF completo com todos os dados, gráficos e tabelas"""
        try:
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import letter, A4
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.units import inch
            import matplotlib.pyplot as plt
            import matplotlib
            matplotlib.use('Agg')  # Use non-interactive backend
            
            buffer = BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=50, leftMargin=50, topMargin=50, bottomMargin=50)
            
            # Estilos
            styles = getSampleStyleSheet()
            title_style = ParagraphStyle('CustomTitle', parent=styles['Heading1'], fontSize=16, spaceAfter=20, alignment=1)
            heading_style = ParagraphStyle('CustomHeading', parent=styles['Heading2'], fontSize=12, spaceAfter=10)
            subheading_style = ParagraphStyle('CustomSubHeading', parent=styles['Heading3'], fontSize=10, spaceAfter=8)
            normal_style = ParagraphStyle('CustomNormal', parent=styles['Normal'], fontSize=8)
            
            story = []
            
            # === PÁGINA DE TÍTULO ===
            timestamp = datetime.now().strftime("%d/%m/%Y às %H:%M")
            story.append(Paragraph("📈 RELATÓRIO COMPLETO", title_style))
            story.append(Paragraph("GESTOR DE PLANTIO - INDICADORES FINANCEIROS", title_style))
            story.append(Spacer(1, 30))
            story.append(Paragraph(f"Gerado em: {timestamp}", normal_style))
            story.append(PageBreak())
            
            # === GRÁFICO: RECEITA vs LUCRO LÍQUIDO ===
            story.append(Paragraph("RECEITA vs LUCRO LÍQUIDO POR CENÁRIO", heading_style))
            
            # Criar gráfico matplotlib
            fig, ax = plt.subplots(figsize=(10, 6))
            x = range(len(anos))
            width = 0.25
            
            colors_map = {'Projetado': '#1f77b4', 'Pessimista': '#ff7f0e', 'Otimista': '#2ca02c'}
            
            for i, cenario in enumerate(nomes_cenarios):
                receitas = all_dre_data[cenario]["Receita"]
                lucros = all_dre_data[cenario]["Lucro Líquido"]
                
                ax.bar([p + width*i for p in x], receitas, width, 
                       label=f'Receita ({cenario})', color=colors_map[cenario], alpha=0.7)
                ax.bar([p + width*i for p in x], lucros, width,
                       label=f'Lucro ({cenario})', color=colors_map[cenario], alpha=0.4)
            
            ax.set_xlabel('Anos')
            ax.set_ylabel('Valores (R$)')
            ax.set_title('Receita vs Lucro Líquido por Cenário')
            ax.set_xticks([p + width for p in x])
            ax.set_xticklabels(anos)
            ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
            ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'R$ {x/1e6:.1f}M'))
            
            plt.tight_layout()
            
            # Salvar gráfico temporariamente
            img_buffer = BytesIO()
            plt.savefig(img_buffer, format='png', dpi=150, bbox_inches='tight')
            img_buffer.seek(0)
            plt.close()
            
            # Adicionar ao PDF
            story.append(Image(img_buffer, width=6*inch, height=3.6*inch))
            story.append(Spacer(1, 20))
            story.append(PageBreak())
            
            # === SLIDE: FLUXO DE CAIXA DETALHADO ===
            if st.session_state.get('fluxo_caixa') is not None and not st.session_state['fluxo_caixa'].empty:
                story.append(Paragraph("FLUXO DE CAIXA - DESPESAS POR CATEGORIA", heading_style))
                
                # Tabela de fluxo de caixa (primeiras 10 linhas)
                df_fluxo = st.session_state['fluxo_caixa']
                rows = min(len(df_fluxo) + 1, 12)  # Máximo 11 categorias + cabeçalho
                cols = len(anos) + 1  # Anos + coluna categoria
                
                # Criar tabela
                table_data = [(['Categoria'] + anos)]
                
                # Dados do fluxo de caixa
                for categoria, valores in df_fluxo.head(11).iterrows():
                    row_data = [categoria]
                    row_data += [f"R$ {valores[ano]:,.0f}".replace(",", ".") for ano in anos]
                    table_data.append(row_data)
                
                # Adicionar linha de totais
                total_row = ["Total"] + [f"R$ {df_fluxo[ano].sum():,.0f}".replace(",", ".") for ano in anos]
                table_data.append(total_row)
                
                # Criar tabela
                table = Table(table_data, colWidths=[2.5*inch] + [0.8*inch]*len(anos))
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, -1), 7),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ]))
                story.append(table)
                story.append(PageBreak())
            
            # === INDICADORES FINANCEIROS DETALHADOS ===
            for cenario in nomes_cenarios:
                story.append(Paragraph(f"INDICADORES FINANCEIROS - CENÁRIO {cenario.upper()}", heading_style))
                
                indicators = all_indicators[cenario]
                
                # Criar tabela de indicadores
                indicators_data = [['Indicador', 'Ano 1', 'Ano 2', 'Ano 3', 'Ano 4', 'Ano 5']]
                
                for indicator_name, values in indicators.items():
                    if indicator_name not in ["CAGR Receita (%)", "CAGR Lucro Líquido (%)"]:
                        if 'R$' in indicator_name:
                            formatted_values = [f"R$ {v:,.0f}".replace(",", ".") for v in values]
                        elif '%' in indicator_name or indicator_name == 'DSCR' or 'Real Gasto' in indicator_name or 'Operacional' in indicator_name:
                            formatted_values = [f"{v:.2f}" + ("%" if "%" in indicator_name else "") for v in values]
                        else:
                            formatted_values = [f"{v:.2f}" for v in values]
                        
                        indicators_data.append([indicator_name] + formatted_values)
                
                # Adicionar CAGR
                indicators_data.append(['CAGR Receita (%)', f"{indicators['CAGR Receita (%)']:.2f}%", '-', '-', '-', '-'])
                indicators_data.append(['CAGR Lucro Líquido (%)', f"{indicators['CAGR Lucro Líquido (%)']:.2f}%", '-', '-', '-', '-'])
                
                # Criar tabela
                indicators_table = Table(indicators_data, colWidths=[2.5*inch, 0.8*inch, 0.8*inch, 0.8*inch, 0.8*inch, 0.8*inch])
                indicators_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, -1), 7),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ]))
                story.append(indicators_table)
                story.append(Spacer(1, 15))
                
                # === DRE COMPLETO ===
                story.append(Paragraph(f"DRE DETALHADO - CENÁRIO {cenario.upper()}", subheading_style))
                
                dre_data = all_dre_data[cenario]
                dre_table_data = [['Item', 'Ano 1', 'Ano 2', 'Ano 3', 'Ano 4', 'Ano 5']]
                
                for item, values in dre_data.items():
                    formatted_values = [f"R$ {v:,.0f}".replace(",", ".") for v in values]
                    dre_table_data.append([item] + formatted_values)
                
                dre_table = Table(dre_table_data, colWidths=[2.5*inch, 0.8*inch, 0.8*inch, 0.8*inch, 0.8*inch, 0.8*inch])
                dre_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, -1), 7),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.lightblue),
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ]))
                story.append(dre_table)
                story.append(PageBreak())
            
            # === RECEITA POR CULTURA ===
            story.append(Paragraph("RECEITA POR CULTURA (ANO BASE)", heading_style))
            
            cultura_data = [['Cultura', 'Receita Total (R$)', 'Área (ha)', 'Receita/ha (R$)']]
            for _, row in df_culturas_for_excel.iterrows():
                cultura_data.append([
                    str(row['Cultura']),
                    f"R$ {row['Receita Total']:,.0f}".replace(",", "."),
                    f"{row['Área (ha)']:.1f}",
                    f"R$ {row['Receita por ha']:,.0f}".replace(",", ".")
                ])
            
            cultura_table = Table(cultura_data, colWidths=[2*inch, 1.5*inch, 1*inch, 1.5*inch])
            cultura_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.green),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 9),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
                ('BACKGROUND', (0, 1), (-1, -1), colors.lightgreen),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ]))
            story.append(cultura_table)
            story.append(Spacer(1, 20))
            
            # === DADOS BASE ===
            story.append(PageBreak())
            story.append(Paragraph("DADOS BASE DO SISTEMA", heading_style))
            
            # Plantios
            if st.session_state.get('plantios'):
                story.append(Paragraph("Plantios Cadastrados:", subheading_style))
                plantios_data = [['Nome', 'Cultura', 'Hectares', 'Sacas/ha', 'Preço/Saca']]
                for nome, dados in st.session_state['plantios'].items():
                    plantios_data.append([
                        nome,
                        dados.get('cultura', ''),
                        f"{dados.get('hectares', 0):.1f}",
                        f"{dados.get('sacas_por_hectare', 0):.1f}",
                        f"R$ {dados.get('preco_saca', 0):.2f}".replace(".", ",")
                    ])
                
                plantios_table = Table(plantios_data)
                plantios_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, -1), 8),
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.black)
                ]))
                story.append(plantios_table)
                story.append(Spacer(1, 15))
            
            # Despesas resumidas
            if st.session_state.get('despesas'):
                story.append(Paragraph("Resumo de Despesas por Categoria:", subheading_style))
                df_despesas = pd.DataFrame(st.session_state['despesas'])
                despesas_resumo = df_despesas.groupby('Categoria')['Valor'].sum()
                
                despesas_data = [['Categoria', 'Valor Total (R$)']]
                for categoria, valor in despesas_resumo.items():
                    despesas_data.append([categoria, f"R$ {valor:,.0f}".replace(",", ".")])
                
                despesas_table = Table(despesas_data)
                despesas_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.orange),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, -1), 8),
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.black)
                ]))
                story.append(despesas_table)
                story.append(Spacer(1, 15))
            
            # Configurações
            story.append(Paragraph("Configurações do Sistema:", subheading_style))
            config_data = [['Parâmetro', 'Valor']]
            config_data.append(['Receita Pessimista (%)', f"-{st.session_state.get('pess_receita', 15)}%"])
            config_data.append(['Despesa Pessimista (%)', f"+{st.session_state.get('pess_despesas', 10)}%"])
            config_data.append(['Receita Otimista (%)', f"+{st.session_state.get('otm_receita', 10)}%"])
            config_data.append(['Despesa Otimista (%)', f"-{st.session_state.get('otm_despesas', 10)}%"])
            
            for i, ano in enumerate(anos):
                inflacao = st.session_state.get('inflacoes', [4.0] * len(anos))[i]
                config_data.append([f'Inflação {ano}', f"{inflacao:.2f}%"])
            
            config_table = Table(config_data)
            config_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.purple),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.black)
            ]))
            story.append(config_table)
            
            # Construir PDF
            doc.build(story)
            buffer.seek(0)
            return buffer
            
        except ImportError as e:
            st.error(f"Bibliotecas necessárias não encontradas. Instale com: pip install reportlab matplotlib")
            return None
        except Exception as e:
            st.error(f"Erro ao gerar PDF: {str(e)}")
            return None

    # Criar as colunas para os botões
    col_export1, col_export2, col_export3, col_export4 = st.columns([1, 1, 1, 1])
    
    with col_export1:
        if st.button("📊 Gerar Excel", type="primary", key="relatorio_excel"):
            try:
                excel_buffer = criar_relatorio_completo()
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"relatorio_completo_gestor_plantio_{timestamp}.xlsx"
                
                st.download_button(
                    label="⬇️ Baixar Excel",
                    data=excel_buffer,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_excel"
                )
                st.success("Excel gerado!")
            except Exception as e:
                st.error(f"Erro ao gerar Excel: {e}")
    
    with col_export2:
        if st.button("📄 Gerar PDF", type="secondary", key="relatorio_pdf"):
            try:
                pdf_buffer = criar_relatorio_pdf(all_indicators, all_dre_data, df_culturas_for_excel, nomes_cenarios, anos)
                if pdf_buffer:
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = f"relatorio_indicadores_{timestamp}.pdf"
                    
                    st.download_button(
                        label="⬇️ Baixar PDF", 
                        data=pdf_buffer,
                        file_name=filename,
                        mime="application/pdf",
                        key="download_pdf"
                    )
                    st.success("PDF gerado!")
            except Exception as e:
                st.error(f"Erro ao gerar PDF: {e}")
    
    with col_export3:
        if st.button("🎯 Gerar PPT", key="relatorio_ppt"):
            try:
                # Importar o gerador de PPT
                from utils.ppt_generator import criar_relatorio_ppt
                
                ppt_buffer = criar_relatorio_ppt(all_indicators, all_dre_data, df_culturas_for_excel, nomes_cenarios, anos)
                if ppt_buffer:
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = f"apresentacao_indicadores_{timestamp}.pptx"
                    
                    st.download_button(
                        label="⬇️ Baixar PPT",
                        data=ppt_buffer,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        key="download_ppt"
                    )
                    st.success("PowerPoint gerado!")
                else:
                    st.error("Não foi possível gerar o PowerPoint. Verifique se as bibliotecas estão instaladas.")
            except ImportError as e:
                st.error(f"""
                ❌ **Erro de importação:** {str(e)}
                
                **Para usar a exportação PPT, instale:**
                ```bash
                pip install python-pptx
                ```
                """)
            except Exception as e:
                st.error(f"Erro ao gerar PowerPoint: {e}")
    
    with col_export4:
        st.info("""
        📋 **Formatos:
        
        **Excel:** Dados completos
        **PDF:** Relatório visual  
        **PPT:** Apresentação editável
        
        Todos incluem gráficos, tabelas e indicadores.
        """)

def main():
    display_indicator_explanation()

    # Carrega todos os dados necessários do session_state e calcula os totais base
    session_data = get_base_financial_data()

    # Exibe os parâmetros de cenário
    display_scenario_parameters(session_data)

    # Exibe a receita por cultura
    df_culturas_for_excel = display_revenue_by_crop(session_data)

    # Nomes dos cenários
    nomes_cenarios = ["Projetado", "Pessimista", "Otimista"]
    anos = session_data["anos"]

    # Calcula os indicadores para cada cenário
    all_indicators = {}
    for cenario_name in nomes_cenarios:
        # Acessa os dados do DRE que foram salvos pelo 4_Fluxo_de_Caixa.py
        dre_data = session_data["dre_cenarios"][cenario_name]
        all_indicators[cenario_name] = calculate_indicators_for_scenario(cenario_name, dre_data, session_data)

    # Exibe as tabelas de indicadores
    display_indicators_table(all_indicators, anos)

    # Exibe o resumo financeiro
    display_financial_summary(session_data["dre_cenarios"], anos)

    # Gera e exibe as visualizações 
    generate_visualizations(session_data["dre_cenarios"], all_indicators, anos, nomes_cenarios, session_data)

    # Gera o parecer financeiro
    generate_financial_opinion(all_indicators, session_data)
    
    # Botão de exportar para Excel
    generate_excel_export(all_indicators, session_data["dre_cenarios"], df_culturas_for_excel, nomes_cenarios, anos)

if __name__ == "__main__":
    main()