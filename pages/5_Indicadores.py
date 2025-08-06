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
    Inclui dados por cultura quando disponíveis.
    """
    required_keys = ["plantios", "dre_cenarios", "receitas_cenarios", "inflacoes", "anos"]
    for key in required_keys:
        if key not in st.session_state:
            st.warning(f"Dados essenciais não encontrados: '{key}'. Por favor, verifique as páginas anteriores.")
            st.stop()
    if not st.session_state["plantios"]:
        st.warning("Nenhum plantio cadastrado. Cadastre ao menos um plantio para gerar os indicadores.")
        st.stop()

    plantios = st.session_state["plantios"]
    inflacoes = st.session_state["inflacoes"]
    anos = st.session_state["anos"]

    hectares_total = 0
    total_sacas = 0
    preco_total_base = 0

    # Calcular totais de plantio
    for p_data in plantios.values():
        hectares = p_data.get("hectares", 0)
        sacas = p_data.get("sacas_por_hectare", 0)
        preco = p_data.get("preco_saca", 0)
        hectares_total += hectares
        total_sacas += sacas * hectares
        preco_total_base += preco * sacas * hectares

    if hectares_total == 0 or total_sacas == 0:
        st.error("Dados de plantio incompletos para estimar receita e indicadores. Verifique o cadastro de plantios.")
        st.stop()

    # Estimativa de ativos totais
    total_ativos = hectares_total * 20000 + 1000000

    # Receitas já calculadas
    receitas_cenarios = st.session_state["receitas_cenarios"]

    # Receitas extras projetadas
    receitas_extras_projetadas = {"Operacional": [0] * len(anos), "Extra Operacional": [0] * len(anos)}
    if "receitas_adicionais" in st.session_state:
        for receita_add in st.session_state["receitas_adicionais"].values():
            valor = receita_add["valor"]
            categoria = receita_add["categoria"]
            for ano_aplicacao in receita_add["anos_aplicacao"]:
                try:
                    idx = anos.index(ano_aplicacao)
                    if categoria == "Operacional":
                        fator = np.prod([1 + inflacoes[j] for j in range(idx + 1)])
                        receitas_extras_projetadas["Operacional"][idx] += valor * fator
                    else:
                        receitas_extras_projetadas["Extra Operacional"][idx] += valor
                except ValueError:
                    continue

    # Dados por cultura (se disponíveis)
    custos_por_cultura = st.session_state.get('custos_por_cultura', {})
    rateio_administrativo = st.session_state.get('rateio_administrativo', {})
    
    # Calcular receitas por cultura para TODOS OS CENÁRIOS
    receitas_por_cultura_cenarios = {}
    
    if plantios and custos_por_cultura:
        for cenario_name in ["Projetado", "Pessimista", "Otimista"]:
            receitas_por_cultura_cenarios[cenario_name] = {}
            
            # Obter fatores de ajuste do cenário
            pess_receita = st.session_state.get("pess_receita", 15)
            otm_receita = st.session_state.get("otm_receita", 10)
            
            fator_receita = 1.0
            if cenario_name == "Pessimista":
                fator_receita = 1 - (pess_receita / 100)
            elif cenario_name == "Otimista":
                fator_receita = 1 + (otm_receita / 100)
            
            for plantio_nome, plantio_data in plantios.items():
                cultura = plantio_data.get('cultura', '')
                if cultura and cultura in custos_por_cultura:
                    if cultura not in receitas_por_cultura_cenarios[cenario_name]:
                        receitas_por_cultura_cenarios[cenario_name][cultura] = {}
                        for ano in anos:
                            receitas_por_cultura_cenarios[cenario_name][cultura][ano] = 0
                    
                    # Calcular receita da cultura para cada ano com ajuste de cenário
                    hectares = plantio_data.get('hectares', 0)
                    sacas_por_ha = plantio_data.get('sacas_por_hectare', 0)
                    preco_saca = plantio_data.get('preco_saca', 0)
                    
                    for i, ano in enumerate(anos):
                        fator_inflacao = np.prod([1 + inflacoes[j] / 100 for j in range(i + 1)])
                        receita_base = hectares * sacas_por_ha * preco_saca * fator_inflacao
                        receita_cenario = receita_base * fator_receita
                        receitas_por_cultura_cenarios[cenario_name][cultura][ano] += receita_cenario

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
        "preco_total_base": preco_total_base,
        "total_ativos": total_ativos,
        "receitas_extras_projetadas": receitas_extras_projetadas,
        "custos_por_cultura": custos_por_cultura,
        "rateio_administrativo": rateio_administrativo,
        "receitas_por_cultura_cenarios": receitas_por_cultura_cenarios
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

def calculate_custos_cultura_por_cenario(custos_por_cultura_base, cenario_name, session_data):
    """Calcula custos por cultura ajustados pelo cenário."""
    if not custos_por_cultura_base:
        return {}
    
    # Obter fatores de ajuste do cenário
    pess_despesas = session_data.get("pess_despesas", 10)
    otm_despesas = session_data.get("otm_despesas", 10)
    
    fator_custo = 1.0
    if cenario_name == "Pessimista":
        fator_custo = 1 + (pess_despesas / 100)
    elif cenario_name == "Otimista":
        fator_custo = 1 - (otm_despesas / 100)
    
    # Aplicar fator aos custos por cultura
    custos_ajustados = {}
    for cultura, df_custos in custos_por_cultura_base.items():
        if not df_custos.empty:
            custos_ajustados[cultura] = df_custos * fator_custo
        else:
            custos_ajustados[cultura] = df_custos
    
    return custos_ajustados

def display_indicators_by_cultura(session_data):
    """Exibe indicadores detalhados por cultura para todos os cenários."""
    custos_por_cultura = session_data.get("custos_por_cultura", {})
    receitas_por_cultura_cenarios = session_data.get("receitas_por_cultura_cenarios", {})
    
    if not custos_por_cultura or not receitas_por_cultura_cenarios:
        st.info("📌 Para visualizar indicadores por cultura, cadastre despesas/empréstimos com centros de custo na página de Despesas.")
        return {}
    
    st.markdown("### 🌱 Análise Financeira por Cultura e Cenário")
    
    total_ativos = session_data["total_ativos"]
    anos = session_data["anos"]
    hectares_total = session_data["hectares_total"]
    
    # Criar abas para cada cenário
    tab1, tab2, tab3 = st.tabs(["📊 Projetado", "📉 Pessimista", "📈 Otimista"])
    tabs = [tab1, tab2, tab3]
    cenarios = ["Projetado", "Pessimista", "Otimista"]
    
    all_indicators_cultura_cenarios = {}
    
    for tab, cenario_name in zip(tabs, cenarios):
        with tab:
            st.markdown(f"#### Indicadores por Cultura - Cenário {cenario_name}")
            
            # Calcular custos ajustados pelo cenário
            custos_ajustados = calculate_custos_cultura_por_cenario(custos_por_cultura, cenario_name, session_data)
            
            # Obter receitas do cenário
            receitas_cenario = receitas_por_cultura_cenarios.get(cenario_name, {})
            
            all_indicators_cultura_cenarios[cenario_name] = {}
            
            for cultura in custos_ajustados.keys():
                if cultura in receitas_cenario:
                    # Calcular hectares da cultura
                    hectares_cultura = sum(
                        plantio.get('hectares', 0) 
                        for plantio in session_data["plantios"].values() 
                        if plantio.get('cultura') == cultura
                    )
                    
                    # Calcular ativos proporcionais
                    total_ativos_cultura = (hectares_cultura / hectares_total * total_ativos) if hectares_total > 0 else 0
                    
                    # Calcular indicadores para a cultura no cenário
                    indicators_cultura = calculate_indicators_for_cultura(
                        cultura, 
                        receitas_cenario[cultura], 
                        custos_ajustados[cultura], 
                        anos, 
                        total_ativos_cultura
                    )
                    
                    all_indicators_cultura_cenarios[cenario_name][cultura] = indicators_cultura
                    
                    # Exibir tabela de indicadores da cultura
                    st.subheader(f"🌿 {cultura}")
                    
                    df_indicadores_cultura = pd.DataFrame({
                        k: v for k, v in indicators_cultura.items()
                        if k not in ["CAGR Receita (%)", "CAGR Lucro Líquido (%)"]
                    }, index=anos)
                    
                    styled_df_cultura = df_indicadores_cultura.style.format({
                        "Margem Líquida (%)": "{:.2f}%",
                        "Retorno por Real Gasto": "{:.2f}",
                        "Liquidez Operacional": "{:.2f}",
                        "Custo por Receita (%)": "{:.2f}%",
                        "ROA (%)": "{:.2f}%"
                    })
                    
                    st.dataframe(styled_df_cultura, use_container_width=True)
                    
                    # Métricas CAGR
                    col_cagr1, col_cagr2 = st.columns(2)
                    with col_cagr1:
                        st.metric("📈 CAGR Receita (5 anos)", f"{indicators_cultura['CAGR Receita (%)']:.2f}%")
                    with col_cagr2:
                        st.metric("📈 CAGR Lucro Líquido (5 anos)", f"{indicators_cultura['CAGR Lucro Líquido (%)']:.2f}%")
                    
                    # Parecer da cultura
                    generate_financial_opinion_cultura(indicators_cultura, cultura, hectares_cultura)
                    
                    st.markdown("---")
    
    return all_indicators_cultura_cenarios

def generate_fluxo_caixa_consolidado_e_culturas(session_data, all_indicators_cultura_cenarios):
    """Gera fluxo de caixa consolidado e por cultura."""
    st.markdown("### 💰 Fluxo de Caixa Projetado")
    
    anos = session_data["anos"]
    dre_cenarios = session_data["dre_cenarios"]
    
    # Criar abas para fluxo de caixa
    tab_geral, tab_culturas = st.tabs(["💼 Consolidado", "🌱 Por Cultura"])
    
    with tab_geral:
        st.markdown("#### 💼 Fluxo de Caixa Consolidado")
        
        # Criar DataFrame do fluxo de caixa consolidado
        fluxo_consolidado = {}
        
        for cenario in ["Projetado", "Pessimista", "Otimista"]:
            dre_data = dre_cenarios[cenario]
            
            fluxo_consolidado[cenario] = {
                "Receita Operacional": dre_data["Receita"],
                "(-) Impostos sobre Venda": [-x for x in dre_data["Impostos Sobre Venda"]],
                "(-) Despesas Operacionais": [-x for x in dre_data["Despesas Operacionais"]],
                "(-) Despesas Administrativas": [-x for x in dre_data["Despesas Administrativas"]],
                "(-) Despesas RH": [-x for x in dre_data["Despesas RH"]],
                "(=) EBITDA": [
                    dre_data["Receita"][i] 
                    - dre_data["Impostos Sobre Venda"][i]
                    - dre_data["Despesas Operacionais"][i] 
                    - dre_data["Despesas Administrativas"][i]
                    - dre_data["Despesas RH"][i]
                    for i in range(len(anos))
                ],
                "(-) Despesas Extra Operacionais": [-x for x in dre_data["Despesas Extra Operacional"]],
                "(-) Dividendos": [-x for x in dre_data["Dividendos"]],
                "(-) Impostos sobre Resultado": [-x for x in dre_data["Impostos Sobre Resultado"]],
                "(=) FLUXO DE CAIXA LÍQUIDO": dre_data["Lucro Líquido"]
            }
        
        # Exibir fluxo de caixa por cenário
        for cenario in ["Projetado", "Pessimista", "Otimista"]:
            emoji = "📊" if cenario == "Projetado" else "📉" if cenario == "Pessimista" else "📈"
            
            with st.expander(f"{emoji} Fluxo de Caixa - {cenario}"):
                df_fluxo = pd.DataFrame(fluxo_consolidado[cenario], index=anos).T
                
                # Aplicar formatação
                styled_fluxo = df_fluxo.style.format(lambda x: f"R$ {x:,.0f}".replace(",", "v").replace(".", ",").replace("v", "."))
                
                # Destacar linhas importantes
                def highlight_important_rows(s):
                    styles = []
                    for idx in s.index:
                        if "(=)" in str(idx):
                            styles.append('background-color: #e6f3ff; font-weight: bold')
                        elif "(-)" in str(idx):
                            styles.append('color: #d32f2f')
                        else:
                            styles.append('')
                    return styles
                
                styled_fluxo = styled_fluxo.apply(highlight_important_rows, axis=1)
                st.dataframe(styled_fluxo, use_container_width=True)
                
                # Resumo do cenário
                total_5_anos = sum(fluxo_consolidado[cenario]["(=) FLUXO DE CAIXA LÍQUIDO"])
                media_anual = total_5_anos / len(anos)
                
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Total 5 Anos", f"R$ {total_5_anos:,.0f}".replace(",", "v").replace(".", ",").replace("v", "."))
                with col2:
                    st.metric("Média Anual", f"R$ {media_anual:,.0f}".replace(",", "v").replace(".", ",").replace("v", "."))
    
    with tab_culturas:
        st.markdown("#### 🌱 Fluxo de Caixa por Cultura")
        
        if not all_indicators_cultura_cenarios:
            st.info("📌 Dados por cultura não disponíveis. Configure centros de custo na página de Despesas.")
            return fluxo_consolidado, {}
        
        # Selecionar cenário para análise por cultura
        cenario_selecionado = st.selectbox(
            "Selecione o cenário para análise:",
            ["Projetado", "Pessimista", "Otimista"],
            key="fluxo_cultura_cenario"
        )
        
        if cenario_selecionado not in all_indicators_cultura_cenarios:
            st.warning("Dados do cenário selecionado não disponíveis.")
            return fluxo_consolidado, {}
        
        receitas_por_cultura = session_data.get("receitas_por_cultura_cenarios", {}).get(cenario_selecionado, {})
        custos_por_cultura = session_data.get("custos_por_cultura", {})
        
        # Calcular custos ajustados pelo cenário
        custos_ajustados = calculate_custos_cultura_por_cenario(custos_por_cultura, cenario_selecionado, session_data)
        
        fluxos_por_cultura = {}
        
        for cultura in receitas_por_cultura.keys():
            if cultura in custos_ajustados:
                
                # Obter receitas da cultura
                receitas_cultura = [receitas_por_cultura[cultura].get(ano, 0) for ano in anos]
                
                # Obter custos da cultura
                if isinstance(custos_ajustados[cultura], pd.DataFrame) and not custos_ajustados[cultura].empty:
                    custos_totais = custos_ajustados[cultura].sum(axis=0).tolist()
                else:
                    custos_totais = [0] * len(anos)
                
                # Estimar impostos (aproximação baseada na proporção geral)
                impostos_receita = session_data["dre_cenarios"][cenario_selecionado]["Impostos Sobre Venda"]
                receita_geral = session_data["dre_cenarios"][cenario_selecionado]["Receita"]
                
                impostos_cultura = []
                for i, ano in enumerate(anos):
                    if receita_geral[i] > 0:
                        proporcao_receita = receitas_cultura[i] / receita_geral[i]
                        imposto_estimado = impostos_receita[i] * proporcao_receita
                        impostos_cultura.append(imposto_estimado)
                    else:
                        impostos_cultura.append(0)
                
                # Estimar impostos sobre resultado
                lucro_bruto = [receitas_cultura[i] - impostos_cultura[i] - custos_totais[i] for i in range(len(anos))]
                impostos_resultado = [max(0, lucro * 0.25) for lucro in lucro_bruto]  # Aproximação de 25%
                
                # Fluxo líquido
                fluxo_liquido = [
                    receitas_cultura[i] - impostos_cultura[i] - custos_totais[i] - impostos_resultado[i]
                    for i in range(len(anos))
                ]
                
                fluxos_por_cultura[cultura] = {
                    "(+) Receita Operacional": receitas_cultura,
                    "(-) Impostos sobre Venda": [-x for x in impostos_cultura],
                    "(-) Custos Diretos": [-x for x in custos_totais],
                    "(=) Lucro Bruto": lucro_bruto,
                    "(-) Impostos sobre Resultado": [-x for x in impostos_resultado],
                    "(=) FLUXO DE CAIXA LÍQUIDO": fluxo_liquido
                }
        
        # Exibir fluxo de caixa por cultura
        for cultura, fluxo_data in fluxos_por_cultura.items():
            with st.expander(f"🌿 {cultura} - Fluxo de Caixa"):
                df_fluxo_cultura = pd.DataFrame(fluxo_data, index=anos).T
                
                # Aplicar formatação
                styled_fluxo_cultura = df_fluxo_cultura.style.format(
                    lambda x: f"R$ {x:,.0f}".replace(",", "v").replace(".", ",").replace("v", ".")
                )
                
                # Destacar linhas importantes
                def highlight_cultura_rows(s):
                    styles = []
                    for idx in s.index:
                        if "(=)" in str(idx):
                            styles.append('background-color: #e8f5e8; font-weight: bold')
                        elif "(-)" in str(idx):
                            styles.append('color: #d32f2f')
                        else:
                            styles.append('color: #2e7d32')
                    return styles
                
                styled_fluxo_cultura = styled_fluxo_cultura.apply(highlight_cultura_rows, axis=1)
                st.dataframe(styled_fluxo_cultura, use_container_width=True)
                
                # Métricas da cultura
                total_cultura = sum(fluxo_data["(=) FLUXO DE CAIXA LÍQUIDO"])
                media_cultura = total_cultura / len(anos)
                
                # Calcular hectares da cultura
                hectares_cultura = sum(
                    plantio.get('hectares', 0) 
                    for plantio in session_data["plantios"].values() 
                    if plantio.get('cultura') == cultura
                )
                
                fluxo_por_hectare = media_cultura / hectares_cultura if hectares_cultura > 0 else 0
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Total 5 Anos", f"R$ {total_cultura:,.0f}".replace(",", "v").replace(".", ",").replace("v", "."))
                with col2:
                    st.metric("Média Anual", f"R$ {media_cultura:,.0f}".replace(",", "v").replace(".", ",").replace("v", "."))
                with col3:
                    st.metric("Fluxo/Hectare/Ano", f"R$ {fluxo_por_hectare:,.0f}".replace(",", "v").replace(".", ",").replace("v", "."))
        
        # Comparativo entre culturas
        if len(fluxos_por_cultura) > 1:
            st.markdown("#### 📊 Comparativo de Fluxo de Caixa entre Culturas")
            
            comparativo_data = []
            for cultura, fluxo_data in fluxos_por_cultura.items():
                total_fluxo = sum(fluxo_data["(=) FLUXO DE CAIXA LÍQUIDO"])
                media_anual = total_fluxo / len(anos)
                
                # Calcular hectares
                hectares = sum(
                    plantio.get('hectares', 0) 
                    for plantio in session_data["plantios"].values() 
                    if plantio.get('cultura') == cultura
                )
                
                fluxo_por_ha = media_anual / hectares if hectares > 0 else 0
                
                comparativo_data.append({
                    "Cultura": cultura,
                    "Área (ha)": hectares,
                    "Fluxo Total (5 anos)": total_fluxo,
                    "Fluxo Médio Anual": media_anual,
                    "Fluxo por Hectare/Ano": fluxo_por_ha
                })
            
            df_comparativo = pd.DataFrame(comparativo_data)
            df_comparativo = df_comparativo.sort_values("Fluxo por Hectare/Ano", ascending=False)
            
            # Aplicar formatação
            styled_comparativo = df_comparativo.style.format({
                "Área (ha)": "{:.1f}",
                "Fluxo Total (5 anos)": lambda x: f"R$ {x:,.0f}".replace(",", "v").replace(".", ",").replace("v", "."),
                "Fluxo Médio Anual": lambda x: f"R$ {x:,.0f}".replace(",", "v").replace(".", ",").replace("v", "."),
                "Fluxo por Hectare/Ano": lambda x: f"R$ {x:,.0f}".replace(",", "v").replace(".", ",").replace("v", ".")
            })
            
            st.dataframe(styled_comparativo, use_container_width=True)
            
            # Análise do comparativo
            if not df_comparativo.empty:
                melhor_cultura = df_comparativo.iloc[0]
                pior_cultura = df_comparativo.iloc[-1]
                
                st.markdown("##### 📈 Análise Comparativa:")
                
                analise_comparativa = f"""
                **🥇 Melhor Performance:** {melhor_cultura['Cultura']}
                - Fluxo por hectare/ano: R$ {melhor_cultura['Fluxo por Hectare/Ano']:,.0f}
                - Área: {melhor_cultura['Área (ha)']:.1f} ha
                
                **🔻 Menor Performance:** {pior_cultura['Cultura']}
                - Fluxo por hectare/ano: R$ {pior_cultura['Fluxo por Hectare/Ano']:,.0f}
                - Área: {pior_cultura['Área (ha)']:.1f} ha
                
                **💡 Oportunidade:** Diferença de R$ {melhor_cultura['Fluxo por Hectare/Ano'] - pior_cultura['Fluxo por Hectare/Ano']:,.0f} por hectare/ano
                """
                
                st.markdown(analise_comparativa.replace(",", "v").replace(".", ",").replace("v", "."))
    
    return fluxo_consolidado, fluxos_por_cultura

def generate_excel_export_with_cultura(all_indicators, all_dre_data, df_culturas_for_excel, nomes_cenarios, anos, all_indicators_cultura_cenarios=None, fluxo_consolidado=None, fluxos_por_cultura=None):
    """Gera exportação Excel incluindo dados por cultura e fluxos de caixa."""
    st.markdown("### ⬇️ Exportar Relatório Completo")
    
    def criar_relatorio_completo():
        output_excel = BytesIO()
        
        with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
            # === DADOS GERAIS POR CENÁRIO ===
            for cenario in nomes_cenarios:
                # Indicadores gerais
                indicators_df_for_excel = pd.DataFrame({
                    k: v for k, v in all_indicators[cenario].items()
                    if k not in ["CAGR Receita (%)", "CAGR Lucro Líquido (%)"]
                }, index=anos)
                cagr_row_data = [all_indicators[cenario]["CAGR Receita (%)"], all_indicators[cenario]["CAGR Lucro Líquido (%)"]] + \
                                [np.nan] * (len(indicators_df_for_excel.columns) - 2)
                cagr_row = pd.Series(cagr_row_data, index=indicators_df_for_excel.columns, name="CAGR")
                indicators_df_for_excel = pd.concat([indicators_df_for_excel, pd.DataFrame(cagr_row).T])
                indicators_df_for_excel.to_excel(writer, sheet_name=f"Indicadores_Geral_{cenario}")

                # DRE geral
                df_dre_for_excel = pd.DataFrame(all_dre_data[cenario], index=anos).T
                df_dre_for_excel.to_excel(writer, sheet_name=f"DRE_Geral_{cenario}")

            # === DADOS POR CULTURA E CENÁRIO ===
            if all_indicators_cultura_cenarios:
                for cenario_name in nomes_cenarios:
                    if cenario_name in all_indicators_cultura_cenarios:
                        for cultura, indicators_cultura in all_indicators_cultura_cenarios[cenario_name].items():
                            # Indicadores por cultura e cenário
                            df_cultura_indicators = pd.DataFrame({
                                k: v for k, v in indicators_cultura.items()
                                if k not in ["CAGR Receita (%)", "CAGR Lucro Líquido (%)"]
                            }, index=anos)
                            
                            # Adicionar CAGR
                            cagr_cultura_data = [indicators_cultura["CAGR Receita (%)"], indicators_cultura["CAGR Lucro Líquido (%)"]] + \
                                               [np.nan] * (len(df_cultura_indicators.columns) - 2)
                            cagr_cultura_row = pd.Series(cagr_cultura_data, index=df_cultura_indicators.columns, name="CAGR")
                            df_cultura_indicators = pd.concat([df_cultura_indicators, pd.DataFrame(cagr_cultura_row).T])
                            df_cultura_indicators.to_excel(writer, sheet_name=f"Indicadores_{cultura}_{cenario_name}")
            
            # === NOVOS DADOS: FLUXOS DE CAIXA ===
            if fluxo_consolidado:
                for cenario, fluxo_data in fluxo_consolidado.items():
                    df_fluxo = pd.DataFrame(fluxo_data, index=anos).T
                    df_fluxo.to_excel(writer, sheet_name=f'FluxoCaixa_Geral_{cenario}', index=True)
            
            if fluxos_por_cultura:
                for cultura, fluxo_data in fluxos_por_cultura.items():
                    df_fluxo_cultura = pd.DataFrame(fluxo_data, index=anos).T
                    df_fluxo_cultura.to_excel(writer, sheet_name=f'FluxoCaixa_{cultura}', index=True)
                
                # Criar comparativo de fluxos por cultura
                comparativo_fluxos = []
                for cultura, fluxo_data in fluxos_por_cultura.items():
                    total_fluxo = sum(fluxo_data["(=) FLUXO DE CAIXA LÍQUIDO"])
                    media_anual = total_fluxo / len(anos)
                    
                    hectares = sum(
                        plantio.get('hectares', 0) 
                        for plantio_nome, plantio in st.session_state.get('plantios', {}).items() 
                        if plantio.get('cultura') == cultura
                    )
                    
                    comparativo_fluxos.append({
                        'Cultura': cultura,
                        'Area_ha': hectares,
                        'Fluxo_Total_5anos': total_fluxo,
                        'Fluxo_Medio_Anual': media_anual,
                        'Fluxo_por_Hectare_Ano': media_anual / hectares if hectares > 0 else 0
                    })
                
                if comparativo_fluxos:
                    df_comparativo_fluxos = pd.DataFrame(comparativo_fluxos)
                    df_comparativo_fluxos.to_excel(writer, sheet_name='Comparativo_FluxoCaixa_Culturas', index=False)
            
            # === DADOS EXISTENTES ===
            if st.session_state.get('custos_por_cultura'):
                for cultura, df_cultura in st.session_state['custos_por_cultura'].items():
                    if not df_cultura.empty:
                        df_cultura.to_excel(writer, sheet_name=f'Custos_{cultura}', index=True)
            
            if st.session_state.get('fluxo_caixa') is not None and not st.session_state['fluxo_caixa'].empty:
                df_fluxo_despesas = st.session_state['fluxo_caixa'].copy()
                df_fluxo_despesas.to_excel(writer, sheet_name='Fluxo_Despesas', index=True)
            
            if st.session_state.get('despesas'):
                df_despesas_cadastradas = pd.DataFrame(st.session_state['despesas'])
                df_despesas_cadastradas.to_excel(writer, sheet_name='Despesas_Cadastradas', index=False)

            if st.session_state.get('emprestimos'):
                df_emprestimos_cadastrados = pd.DataFrame(st.session_state['emprestimos'])
                df_emprestimos_cadastrados.to_excel(writer, sheet_name='Emprestimos_Cadastrados', index=False)
            
            if st.session_state.get('plantios'):
                plantios_list = []
                for nome, dados in st.session_state['plantios'].items():
                    plantio_row = {'Nome': nome}
                    plantio_row.update(dados)
                    plantios_list.append(plantio_row)
                df_plantios = pd.DataFrame(plantios_list)
                df_plantios.to_excel(writer, sheet_name='Plantios_Cadastrados', index=False)
            
            # Receita por Cultura
            df_culturas_for_excel.to_excel(writer, sheet_name="Receita_por_Cultura", index=False)
            
            # Configurações
            df_config = pd.DataFrame({
                'Ano': anos,
                'Inflacao (%)': st.session_state.get('inflacoes', [4.0] * len(anos))
            })
            df_config.to_excel(writer, sheet_name='Configuracoes_Inflacao', index=False)

        output_excel.seek(0)
        return output_excel
    
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
            st.info("PDF com dados por cultura em desenvolvimento...")
    
    with col_export3:
        if st.button("🎯 Gerar PPT", key="relatorio_ppt"):
            try:
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
                st.error(f"Erro de importação: {str(e)}")
            except Exception as e:
                st.error(f"Erro ao gerar PowerPoint: {e}")
    
    with col_export4:
        st.info("""
        📋 **Formatos:**
        
        **Excel:** Dados completos + por cultura + fluxos de caixa
        **PDF:** Relatório visual (em breve)
        **PPT:** Apresentação editável
        
        Inclui análises consolidadas, por cultura e fluxos de caixa.
        """)

def display_scenario_parameters(session_data):
    """Exibe os parâmetros de cenário configurados."""
    st.markdown("### ⚙️ Parâmetros dos Cenários")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("📉 Receita Pessimista", f"-{session_data['pess_receita']}%")
    with col2:
        st.metric("📈 Despesa Pessimista", f"+{session_data['pess_despesas']}%")
    with col3:
        st.metric("📈 Receita Otimista", f"+{session_data['otm_receita']}%")
    with col4:
        st.metric("📉 Despesa Otimista", f"-{session_data['otm_despesas']}%")

def display_revenue_by_crop(session_data):
    """Exibe a receita por cultura e retorna DataFrame para exportação."""
    st.markdown("### 🌾 Receita por Cultura (Ano Base)")
    
    plantios = session_data["plantios"]
    
    # Calcular receita por cultura
    culturas_data = []
    for plantio_nome, plantio_data in plantios.items():
        cultura = plantio_data.get("cultura", "")
        hectares = plantio_data.get("hectares", 0)
        sacas_por_ha = plantio_data.get("sacas_por_hectare", 0)
        preco_saca = plantio_data.get("preco_saca", 0)
        
        receita_total = hectares * sacas_por_ha * preco_saca
        receita_por_ha = receita_total / hectares if hectares > 0 else 0
        
        culturas_data.append({
            "Cultura": cultura,
            "Receita Total": receita_total,
            "Área (ha)": hectares,
            "Receita por ha": receita_por_ha
        })
    
    df_culturas = pd.DataFrame(culturas_data)
    
    # Agrupar por cultura (caso haja múltiplos plantios da mesma cultura)
    df_culturas_grouped = df_culturas.groupby("Cultura").agg({
        "Receita Total": "sum",
        "Área (ha)": "sum"
    }).reset_index()
    
    # Recalcular receita por hectare
    df_culturas_grouped["Receita por ha"] = df_culturas_grouped["Receita Total"] / df_culturas_grouped["Área (ha)"]
    
    # Exibir tabela formatada
    styled_df = df_culturas_grouped.style.format({
        "Receita Total": format_brl,
        "Área (ha)": "{:.1f}",
        "Receita por ha": format_brl
    })
    
    st.dataframe(styled_df, use_container_width=True)
    
    return df_culturas_grouped

def display_indicators_table(all_indicators, anos):
    """Exibe as tabelas de indicadores para todos os cenários."""
    st.markdown("### 📊 Indicadores Financeiros por Cenário")
    
    tab1, tab2, tab3 = st.tabs(["📊 Projetado", "📉 Pessimista", "📈 Otimista"])
    
    tabs = [tab1, tab2, tab3]
    cenarios = ["Projetado", "Pessimista", "Otimista"]
    
    for tab, cenario in zip(tabs, cenarios):
        with tab:
            indicators = all_indicators[cenario]
            
            # Criar DataFrame com indicadores (exceto CAGR)
            df_indicators = pd.DataFrame({
                k: v for k, v in indicators.items()
                if k not in ["CAGR Receita (%)", "CAGR Lucro Líquido (%)"]
            }, index=anos)
            
            # Aplicar formatação
            styled_df = df_indicators.style.format({
                "Margem Líquida (%)": "{:.2f}%",
                "Retorno por Real Gasto": "{:.2f}",
                "Liquidez Operacional": "{:.2f}",
                "Endividamento (%)": "{:.2f}%",
                "Produtividade por Hectare (R$/ha)": format_brl,
                "Custo por Receita (%)": "{:.2f}%",
                "DSCR": "{:.2f}",
                "Break-Even Yield (sacas/ha)": "{:.1f}",
                "ROA (%)": "{:.2f}%",
                "Custo por Hectare (R$/ha)": format_brl
            })
            
            st.dataframe(styled_df, use_container_width=True)
            
            # Métricas CAGR
            col_cagr1, col_cagr2 = st.columns(2)
            with col_cagr1:
                st.metric("📈 CAGR Receita (5 anos)", f"{indicators['CAGR Receita (%)']:.2f}%")
            with col_cagr2:
                st.metric("📈 CAGR Lucro Líquido (5 anos)", f"{indicators['CAGR Lucro Líquido (%)']:.2f}%")

def display_financial_summary(dre_cenarios, anos):
    """Exibe resumo financeiro consolidado."""
    st.markdown("### 💰 Resumo Financeiro Consolidado")
    
    # Criar resumo para cada cenário
    resumo_data = []
    for cenario, dre_data in dre_cenarios.items():
        receita_total = sum(dre_data["Receita"])
        lucro_total = sum(dre_data["Lucro Líquido"])
        margem_media = (lucro_total / receita_total * 100) if receita_total > 0 else 0
        
        resumo_data.append({
            "Cenário": cenario,
            "Receita Total (5 anos)": receita_total,
            "Lucro Total (5 anos)": lucro_total,
            "Margem Média (%)": margem_media
        })
    
    df_resumo = pd.DataFrame(resumo_data)
    
    styled_resumo = df_resumo.style.format({
        "Receita Total (5 anos)": format_brl,
        "Lucro Total (5 anos)": format_brl,
        "Margem Média (%)": "{:.2f}%"
    })
    
    st.dataframe(styled_resumo, use_container_width=True)

def generate_visualizations(dre_cenarios, all_indicators, anos, nomes_cenarios, session_data):
    """Gera visualizações gráficas dos dados."""
    st.markdown("### 📈 Visualizações")
    
    col_viz1, col_viz2 = st.columns(2)
    
    with col_viz1:
        st.markdown("#### Receita vs Lucro por Cenário")
        
        fig_receita_lucro = go.Figure()
        
        for cenario in nomes_cenarios:
            dre_data = dre_cenarios[cenario]
            fig_receita_lucro.add_trace(go.Scatter(
                x=anos,
                y=dre_data["Receita"],
                mode='lines+markers',
                name=f'Receita {cenario}',
                line=dict(width=3)
            ))
            
            fig_receita_lucro.add_trace(go.Scatter(
                x=anos,
                y=dre_data["Lucro Líquido"],
                mode='lines+markers',
                name=f'Lucro {cenario}',
                line=dict(dash='dot')
            ))
        
        fig_receita_lucro.update_layout(
            title="Receita e Lucro Líquido por Cenário",
            xaxis_title="Anos",
            yaxis_title="Valores (R$)",
            hovermode='x unified'
        )
        
        st.plotly_chart(fig_receita_lucro, use_container_width=True)
    
    with col_viz2:
        st.markdown("#### Margem Líquida por Cenário")
        
        fig_margem = go.Figure()
        
        for cenario in nomes_cenarios:
            indicators = all_indicators[cenario]
            fig_margem.add_trace(go.Scatter(
                x=anos,
                y=indicators["Margem Líquida (%)"],
                mode='lines+markers',
                name=f'Margem {cenario}',
                line=dict(width=3)
            ))
        
        fig_margem.update_layout(
            title="Margem Líquida (%) por Cenário",
            xaxis_title="Anos",
            yaxis_title="Margem (%)",
            hovermode='x unified'
        )
        
        st.plotly_chart(fig_margem, use_container_width=True)

def generate_financial_opinion(all_indicators, session_data):
    """Gera parecer financeiro consolidado."""
    st.markdown("#### 📝 Análise Consolidada dos Cenários")
    
    # Análise do cenário projetado
    indicators_proj = all_indicators["Projetado"]
    
    margem_media = np.mean(indicators_proj["Margem Líquida (%)"])
    retorno_medio = np.mean(indicators_proj["Retorno por Real Gasto"])
    liquidez_media = np.mean(indicators_proj["Liquidez Operacional"])
    endividamento_medio = np.mean(indicators_proj["Endividamento (%)"])
    produtividade_media = np.mean(indicators_proj["Produtividade por Hectare (R$/ha)"])
    custo_receita_media = np.mean(indicators_proj["Custo por Receita (%)"])
    dscr_values = [x for x in indicators_proj["DSCR"] if x != float("inf")]
    dscr_medio = np.mean(dscr_values) if dscr_values else float("inf")
    break_even_media = np.mean(indicators_proj["Break-Even Yield (sacas/ha)"])
    roa_medio = np.mean(indicators_proj["ROA (%)"])
    
    parecer = []
    
    # Margem Líquida
    if margem_media < 10:
        parecer.append(f"• **Margem Líquida Baixa ({margem_media:.2f}%)**: Rentabilidade abaixo do ideal. Considere renegociar preços com fornecedores ou investir em culturas de maior valor agregado.")
    else:
        parecer.append(f"• **Margem Líquida Saudável ({margem_media:.2f}%)**: Boa rentabilidade. Monitore custos para manter a consistência.")

    # Retorno por Real Gasto
    if retorno_medio < 0.2:
        parecer.append(f"• **Retorno por Real Gasto Baixo ({retorno_medio:.2f})**: Gastos com baixo retorno. Avalie a redução de despesas operacionais ou otimize processos agrícolas.")
    else:
        parecer.append(f"• **Retorno por Real Gasto Adequado ({retorno_medio:.2f})**: Investimentos geram retorno satisfatório. Considere reinvestir em tecnologia para aumentar a produtividade.")

    # Liquidez Operacional
    if liquidez_media < 1.5:
        parecer.append(f"• **Liquidez Operacional Baixa ({liquidez_media:.2f})**: Risco de dificuldades para cobrir custos operacionais. Negocie prazos de pagamento ou busque linhas de crédito de curto prazo.")
    else:
        parecer.append(f"• **Liquidez Operacional Confortável ({liquidez_media:.2f})**: Boa capacidade de sustentar operações. Mantenha reservas para safras incertas.")

    # Endividamento
    if endividamento_medio > 30:
        parecer.append(f"• **Alto Endividamento ({endividamento_medio:.2f}%)**: Dívidas elevadas. Priorize a quitação de empréstimos de alto custo ou renegocie taxas de juros.")
    else:
        parecer.append(f"• **Endividamento Controlado ({endividamento_medio:.2f}%)**: Dívidas em nível gerenciável. Considere investimentos estratégicos, como expansão de área plantada.")

    # Produtividade
    parecer.append(f"• **Produtividade por Hectare ({produtividade_media:,.0f} R$/ha)**: {'Boa produtividade' if produtividade_media > 5000 else 'Produtividade pode ser melhorada'}. Compare com benchmarks da região.")

    # Custo por Receita
    if custo_receita_media > 70:
        parecer.append(f"• **Custo por Receita Alto ({custo_receita_media:.2f}%)**: Custos operacionais consomem grande parte da receita. Analise insumos e processos para reduzir despesas.")
    else:
        parecer.append(f"• **Custo por Receita Controlado ({custo_receita_media:.2f}%)**: Boa gestão de custos. Continue monitorando preços de insumos.")

    # DSCR
    if dscr_medio != float("inf") and dscr_medio < 1.25:
        parecer.append(f"• **DSCR Baixo ({dscr_medio:.2f})**: Risco de dificuldades no pagamento de dívidas. Considere reestruturar financiamentos ou aumentar a receita.")
    else:
        dscr_text = f"{dscr_medio:.2f}" if dscr_medio != float("inf") else "∞"
        parecer.append(f"• **DSCR Adequado ({dscr_text})**: Boa capacidade de cobrir dívidas. Mantenha o lucro operacional estável.")

    # Break-Even Yield
    parecer.append(f"• **Break-Even Yield ({break_even_media:.1f} sacas/ha)**: {'Alto risco em safras ruins' if break_even_media > 50 else 'Risco moderado em cenários adversos'}.")

    # ROA
    if roa_medio < 5:
        parecer.append(f"• **ROA Baixo ({roa_medio:.2f}%)**: Baixa eficiência no uso de ativos. Avalie a venda de ativos ociosos ou investimentos em equipamentos mais produtivos.")
    else:
        parecer.append(f"• **ROA Adequado ({roa_medio:.2f}%)**: Boa utilização dos ativos. Considere expansão controlada ou modernização.")

    # CAGR
    if indicators_proj["CAGR Lucro Líquido (%)"] < 0:
        parecer.append(f"• **Crescimento Negativo do Lucro ({indicators_proj['CAGR Lucro Líquido (%)']:.2f}%)**: Lucro em queda. Revisar estratégias de custo, preço e produtividade.")
    else:
        parecer.append(f"• **Crescimento do Lucro ({indicators_proj['CAGR Lucro Líquido (%)']:.2f}%)**: Lucro em trajetória positiva. Considere reinvestir em áreas estratégicas.")

    # Comparação entre cenários
    lucro_proj = sum(session_data["dre_cenarios"]["Projetado"]["Lucro Líquido"])
    lucro_pess = sum(session_data["dre_cenarios"]["Pessimista"]["Lucro Líquido"])
    lucro_otm = sum(session_data["dre_cenarios"]["Otimista"]["Lucro Líquido"])
    
    diferenca_pess = ((lucro_pess - lucro_proj) / lucro_proj * 100) if lucro_proj != 0 else 0
    diferenca_otm = ((lucro_otm - lucro_proj) / lucro_proj * 100) if lucro_proj != 0 else 0
    
    parecer.append(f"• **Análise de Cenários**: No cenário pessimista, o lucro seria {abs(diferenca_pess):.1f}% {'menor' if diferenca_pess < 0 else 'maior'}. No otimista, seria {diferenca_otm:.1f}% maior.")

    # Usar quebras de linha duplas para Markdown
    st.markdown("\n\n".join(parecer))

def calculate_indicators_for_cultura(cultura, receitas_cultura, custos_cultura, anos, total_ativos_cultura):
    """Calcula indicadores financeiros para uma cultura específica."""
    # Verificar se os dados existem e não estão vazios
    if not receitas_cultura:
        return {}
    
    # Para DataFrame, usar .empty para verificar se está vazio
    if isinstance(custos_cultura, pd.DataFrame) and custos_cultura.empty:
        return {}
    elif isinstance(custos_cultura, dict) and not custos_cultura:
        return {}
    
    # Converter para listas se necessário
    if isinstance(receitas_cultura, dict):
        receitas = [receitas_cultura.get(ano, 0) for ano in anos]
    else:
        receitas = list(receitas_cultura)
    
    # Calcular custos totais por ano
    if isinstance(custos_cultura, pd.DataFrame) and not custos_cultura.empty:
        custos_totais = custos_cultura.sum(axis=0).tolist()
    elif isinstance(custos_cultura, dict):
        custos_totais = [custos_cultura.get(ano, 0) for ano in anos]
    else:
        custos_totais = [0] * len(anos)
    
    # Calcular lucro líquido
    lucro_liquido = [r - c for r, c in zip(receitas, custos_totais)]
    
    indicators = {}
    
    # 1. Margem Líquida (%)
    indicators["Margem Líquida (%)"] = [
        (l / r * 100) if r != 0 else 0 for l, r in zip(lucro_liquido, receitas)
    ]
    
    # 2. Retorno por Real Gasto
    indicators["Retorno por Real Gasto"] = [
        (l / c) if c != 0 else 0 for l, c in zip(lucro_liquido, custos_totais)
    ]
    
    # 3. Liquidez Operacional
    indicators["Liquidez Operacional"] = [
        (r / c) if c != 0 else 0 for r, c in zip(receitas, custos_totais)
    ]
    
    # 4. Custo por Receita (%)
    indicators["Custo por Receita (%)"] = [
        (c / r * 100) if r != 0 else 0 for c, r in zip(custos_totais, receitas)
    ]
    
    # 5. ROA (%)
    indicators["ROA (%)"] = [
        (l / total_ativos_cultura * 100) if total_ativos_cultura != 0 else 0 for l in lucro_liquido
    ]
    
    # 6. CAGR Receita (%)
    indicators["CAGR Receita (%)"] = calcular_cagr(receitas[0], receitas[-1], len(anos) - 1)
    
    # 7. CAGR Lucro Líquido (%)
    indicators["CAGR Lucro Líquido (%)"] = calcular_cagr(lucro_liquido[0], lucro_liquido[-1], len(anos) - 1)
    
    return indicators

def generate_financial_opinion_cultura(indicators_cultura, cultura, hectares_cultura):
    """Gera parecer financeiro específico para uma cultura."""
    st.markdown(f"#### 🌿 Parecer Financeiro - {cultura}")
    
    margem_media = np.mean(indicators_cultura["Margem Líquida (%)"])
    retorno_medio = np.mean(indicators_cultura["Retorno por Real Gasto"])
    liquidez_media = np.mean(indicators_cultura["Liquidez Operacional"])
    custo_receita_media = np.mean(indicators_cultura["Custo por Receita (%)"])
    roa_medio = np.mean(indicators_cultura["ROA (%)"])
    
    parecer = []
    
    # Análises específicas por cultura
    if margem_media < 10:
        parecer.append(f"• **Margem Líquida Baixa ({margem_media:.2f}%)**: A cultura {cultura} apresenta rentabilidade abaixo do ideal. Considere otimizar técnicas de cultivo ou renegociar preços de venda.")
    else:
        parecer.append(f"• **Margem Líquida Saudável ({margem_media:.2f}%)**: A cultura {cultura} apresenta boa rentabilidade. Mantenha as práticas atuais.")
    
    if retorno_medio < 0.2:
        parecer.append(f"• **Baixo Retorno por Real Gasto ({retorno_medio:.2f})**: Os investimentos em {cultura} estão gerando baixo retorno. Analise custos de insumos e produtividade.")
    else:
        parecer.append(f"• **Retorno Adequado por Real Gasto ({retorno_medio:.2f})**: Os investimentos em {cultura} estão gerando retorno satisfatório.")
    
    if liquidez_media < 1.5:
        parecer.append(f"• **Liquidez Operacional Baixa ({liquidez_media:.2f})**: A cultura {cultura} pode ter dificuldades para cobrir seus custos operacionais.")
    else:
        parecer.append(f"• **Liquidez Operacional Confortável ({liquidez_media:.2f})**: A cultura {cultura} tem boa capacidade de cobrir seus custos operacionais.")
    
    if custo_receita_media > 70:
        parecer.append(f"• **Alto Custo por Receita ({custo_receita_media:.2f}%)**: A cultura {cultura} tem custos elevados em relação à receita. Revise práticas agrícolas e custos de insumos.")
    else:
        parecer.append(f"• **Custo por Receita Controlado ({custo_receita_media:.2f}%)**: A cultura {cultura} apresenta boa gestão de custos.")
    
    if roa_medio < 5:
        parecer.append(f"• **ROA Baixo ({roa_medio:.2f}%)**: A eficiência dos ativos destinados à {cultura} está baixa. Considere melhorar a produtividade por hectare.")
    else:
        parecer.append(f"• **ROA Adequado ({roa_medio:.2f}%)**: Boa eficiência dos ativos destinados à cultura {cultura}.")
    
    # CAGR Analysis
    if indicators_cultura["CAGR Lucro Líquido (%)"] < 0:
        parecer.append(f"• **Crescimento Negativo ({indicators_cultura['CAGR Lucro Líquido (%)']:.2f}%)**: O lucro da cultura {cultura} está em declínio. Reavaliar viabilidade.")
    else:
        parecer.append(f"• **Crescimento Positivo ({indicators_cultura['CAGR Lucro Líquido (%)']:.2f}%)**: A cultura {cultura} apresenta crescimento sustentável.")
    
    # Área específica
    parecer.append(f"• **Área Cultivada**: {hectares_cultura:.1f} hectares representam {(hectares_cultura/st.session_state.get('hectares_total', hectares_cultura)*100):.1f}% da área total.")
    
    # Usar quebras de linha duplas para Markdown
    st.markdown("\n\n".join(parecer))

def main():
    display_indicator_explanation()

    # Carrega todos os dados necessários
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
        dre_data = session_data["dre_cenarios"][cenario_name]
        all_indicators[cenario_name] = calculate_indicators_for_scenario(cenario_name, dre_data, session_data)

    # Exibe as tabelas de indicadores GERAIS
    display_indicators_table(all_indicators, anos)

    # Exibe indicadores POR CULTURA E CENÁRIO
    all_indicators_cultura_cenarios = display_indicators_by_cultura(session_data)

    # NOVO: Exibe fluxo de caixa consolidado e por cultura
    fluxo_consolidado, fluxos_por_cultura = generate_fluxo_caixa_consolidado_e_culturas(session_data, all_indicators_cultura_cenarios)

    # Exibe o resumo financeiro
    display_financial_summary(session_data["dre_cenarios"], anos)

    # Gera e exibe as visualizações 
    generate_visualizations(session_data["dre_cenarios"], all_indicators, anos, nomes_cenarios, session_data)

    # Gera o parecer financeiro GERAL
    st.markdown("### 📝 Parecer Financeiro Consolidado")
    generate_financial_opinion(all_indicators, session_data)
    
    # Atualizar exportações para incluir dados por cultura e cenários
    generate_excel_export_with_cultura(all_indicators, session_data["dre_cenarios"], df_culturas_for_excel, nomes_cenarios, anos, all_indicators_cultura_cenarios, fluxo_consolidado, fluxos_por_cultura)

if __name__ == "__main__":
    main()