import streamlit as st
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from io import BytesIO

st.set_page_config(layout="wide", page_title="Fluxo de Caixa Projeção")
st.title("🎯 Fluxo de Caixa - Cenários (Projetado, Pessimista, Otimista)")

# Verificação de dados essenciais
if "plantios" not in st.session_state or not st.session_state["plantios"]:
    st.warning("Cadastre ao menos um plantio para gerar o fluxo de caixa.")
    st.stop()
if "fluxo_caixa" not in st.session_state:
    st.warning("Você precisa preencher as despesas antes de acessar esta página.")
    st.stop()

anos = [f"Ano {i+1}" for i in range(5)]
inflacoes = [st.session_state.get(f"inf_{i}", 4.0) for i in range(5)]
plantios = st.session_state["plantios"]
df_base_fluxo = st.session_state["fluxo_caixa"]

# Entradas do usuário para cenários
st.markdown("### 🔧 Ajustes de Cenário")

col1, col2 = st.columns(2)
with col1:
    pess_receita = st.slider("Pessimista: Receita - redução (%)", 0, 50, 15)
    pess_despesas = st.slider("Pessimista: Despesas - aumento (%)", 0, 50, 10)
with col2:
    otm_receita = st.slider("Otimista: Receita - aumento (%)", 0, 50, 10)
    otm_despesas = st.slider("Otimista: Despesas - redução (%)", 0, 50, 10)

# === CÁLCULO DA RECEITA ESTIMADA BASE ===
total_sacas = preco_total = hectares_total = 0

for p_data in plantios.values():
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

# Receita Estimada para o cenário base (projetado)
receita_base = []
for i in range(5):
    fator = np.prod([1 + inflacoes[j] / 100 for j in range(i + 1)])
    receita_base.append(hectares_total * media_receita_hectare * fator)

# CENÁRIOS DE RECEITA
receitas = {
    "Projetado": receita_base,
    "Pessimista": [r * (1 - pess_receita / 100) for r in receita_base],
    "Otimista": [r * (1 + otm_receita / 100) for r in receita_base]
}

# CENÁRIOS DE FLUXO DE DESPESAS (aplicando variação sobre df_base_fluxo)
def ajustar_despesas(df_base, ajuste_percentual):
    df = df_base.copy()
    for row in df.index:
        if row not in ["Receita Estimada", "Lucro Líquido", "Impostos Sobre Resultado"]:
            df.loc[row] = df.loc[row] * (1 + ajuste_percentual / 100)
    return df

fluxos = {
    "Projetado": df_base_fluxo.copy(),
    "Pessimista": ajustar_despesas(df_base_fluxo, pess_despesas),
    "Otimista": ajustar_despesas(df_base_fluxo, -otm_despesas)
}

# 1. Explicação do Cálculo das Receitas Futuras
with st.expander("🧾 Entenda como a receita futura é calculada por cultura"):
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

# === EXIBIÇÃO COMPARATIVA DE CENÁRIOS ===
st.markdown("### 📊 Análise de Cenários (Projetado, Pessimista, Otimista)")

abas = st.tabs(["📈 Projetado", "🔻 Pessimista", "🔺 Otimista"])
nomes_cenarios = ["Projetado", "Pessimista", "Otimista"]

# Função para formatar valores em Real Brasileiro
def format_brl(x):
    try:
        return f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return x

# === ESTILIZAÇÃO DAS TABELAS ===
def aplicar_estilo_fluxo(linha):
    if linha.name == "Receita Estimada":
        return ["background-color: #003366; color: white;" for _ in linha]
    elif linha.name == "Lucro Líquido":
        return ["background-color: #006400; color: white;" for _ in linha]
    else:
        return ["" for _ in linha]

def aplicar_estilo_dre(linha):
    if linha.name == "Receita":
        return ["background-color: #003366; color: white;" for _ in linha]
    elif linha.name in ["Margem de Contribuição", "Resultado Operacional", "Lucro Operacional", "Lucro Líquido"]:
        return ["background-color: #006400; color: white;" for _ in linha]
    else:
        return ["" for _ in linha]

def aplicar_estilo_retorno(linha):
    return ["background-color: #FF4040; color: white;" if x <= 0 else "" for x in linha]

def gerar_excel_download(df_fluxo, df_dre, df_retorno, nome_cenario):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Formatar Fluxo de Caixa e DRE com formato de moeda
        df_fluxo.to_excel(writer, sheet_name="Fluxo de Caixa")
        df_dre.to_excel(writer, sheet_name="DRE")
        # Formatar Retorno por Real Gasto como moeda
        workbook = writer.book
        worksheet = writer.sheets["Retorno por Real Gasto"] = workbook.add_worksheet("Retorno por Real Gasto")
        currency_format = workbook.add_format({'num_format': 'R$ #,##0.00'})
        for col_num, value in enumerate(df_retorno.columns.values):
            worksheet.write(0, col_num + 1, value)
        for row_num, value in enumerate(df_retorno.index.values):
            worksheet.write(row_num + 1, 0, value)
        for row_num, row_data in enumerate(df_retorno.values):
            for col_num, value in enumerate(row_data):
                worksheet.write(row_num + 1, col_num + 1, value, currency_format)
        writer.close()
    output.seek(0)

    st.download_button(
        label=f"⬇️ Baixar Excel - {nome_cenario}",
        data=output,
        file_name=f"cenario_{nome_cenario.lower()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

for aba, nome in zip(abas, nomes_cenarios):
    with aba:
        st.subheader(f"📊 Fluxo de Caixa - Cenário {nome}")
        df_fluxo = fluxos[nome].copy()
        df_fluxo.loc["Receita Estimada"] = receitas[nome]

        # --- Recalcular impostos com base coerente ao DRE ---
        df_fluxo.loc["Impostos Sobre Venda"] = df_fluxo.loc["Receita Estimada"] * 0.0485

        lucro_operacional = (
            df_fluxo.loc["Receita Estimada"]
            - df_fluxo.get("Impostos Sobre Venda", 0)
            - df_fluxo.get("Despesas Operacionais", 0)
            - df_fluxo.get("Despesas Administrativas", 0)
            - df_fluxo.get("Despesas RH", 0)
            - df_fluxo.get("Despesas Extra Operacional", 0)
        )

        df_fluxo.loc["Impostos Sobre Resultado"] = [
            lo * 0.15 if lo > 0 else 0 for lo in lucro_operacional
        ]

        # === DRE ===
        st.subheader(f"📘 DRE - Cenário {nome}")

        # Recupera informações de despesas corretamente da session_state
        df_despesas_info = pd.DataFrame(st.session_state.get("despesas", []))
        if not df_despesas_info.empty and "Categoria" in df_despesas_info.columns:
            df_despesas_info["Categoria"] = df_despesas_info["Categoria"].astype(str).str.strip()
        else:
            df_despesas_info = pd.DataFrame(columns=["Categoria", "Valor"])

        dre_calc = {}
        receita = receitas[nome]
        dre_calc["Receita"] = receita
        dre_calc["Impostos Sobre Venda"] = [r * 0.0485 for r in receita]

        def linha_despesa(cat):
            total = sum(df_despesas_info[df_despesas_info["Categoria"] == cat]["Valor"]) if not df_despesas_info.empty else 0
            return [
                total * np.prod([1 + inflacoes[j] / 100 for j in range(i + 1)]) *
                (1 + (pess_despesas if nome == "Pessimista" else (-otm_despesas if nome == "Otimista" else 0)) / 100)
                for i in range(5)
            ]

        dre_calc["Despesas Operacionais"] = linha_despesa("Operacional")
        dre_calc["Despesas Administrativas"] = linha_despesa("Administrativa")
        dre_calc["Despesas RH"] = linha_despesa("RH")
        extra_operacional = linha_despesa("Extra Operacional")

        if "emprestimos" in st.session_state:
            for emp in st.session_state["emprestimos"]:
                start_year_index = anos.index(emp["ano_inicial"])
                end_year_index = anos.index(emp["ano_final"])
                num_years = end_year_index - start_year_index + 1
                for i in range(start_year_index, min(start_year_index + min(emp["parcelas"], num_years), len(anos))):
                    extra_operacional[i] += emp["valor_parcela"] * (
                        1 + (pess_despesas if nome == "Pessimista" else (-otm_despesas if nome == "Otimista" else 0)) / 100
                    )

        dre_calc["Despesas Extra Operacional"] = extra_operacional
        dre_calc["Dividendos"] = linha_despesa("Dividendos")

        dre_calc["Margem de Contribuição"] = [
            receita[i] - dre_calc["Impostos Sobre Venda"][i] - dre_calc["Despesas Operacionais"][i]
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

        # Agora sim: Lucro Líquido no Fluxo vem do DRE
        df_fluxo.loc["Lucro Líquido"] = pd.Series(dre_calc["Lucro Líquido"], index=anos)

        ordem = ["Receita Estimada"] + [i for i in df_fluxo.index if i not in ["Receita Estimada", "Lucro Líquido"]] + ["Lucro Líquido"]
        df_fluxo = df_fluxo.loc[ordem]

        style_idx_fluxo = [
            {
                "selector": f"tbody tr:nth-child({df_fluxo.index.get_loc('Receita Estimada') + 1}) th",
                "props": [("background-color", "#003366"), ("color", "white")]
            },
            {
                "selector": f"tbody tr:nth-child({df_fluxo.index.get_loc('Lucro Líquido') + 1}) th",
                "props": [("background-color", "#006400"), ("color", "white")]
            }
        ]

        st.dataframe(
            df_fluxo.style
                .format(format_brl)
                .apply(aplicar_estilo_fluxo, axis=1)
                .set_table_styles(style_idx_fluxo, overwrite=False)
        )

        df_dre = pd.DataFrame(dre_calc, index=anos).T.loc[[ 
            "Receita", "Impostos Sobre Venda", "Despesas Operacionais",
            "Margem de Contribuição", "Despesas Administrativas", "Despesas RH",
            "Resultado Operacional", "Despesas Extra Operacional",
            "Lucro Operacional", "Impostos Sobre Resultado",
            "Dividendos", "Lucro Líquido"
        ]]

        style_idx_dre = []
        if "Receita" in df_dre.index:
            style_idx_dre.append({
                "selector": f"tbody tr:nth-child({df_dre.index.get_loc('Receita') + 1}) th",
                "props": [("background-color", "#003366"), ("color", "white")]
            })

        for linha_verde in ["Margem de Contribuição", "Resultado Operacional", "Lucro Operacional", "Lucro Líquido"]:
            if linha_verde in df_dre.index:
                style_idx_dre.append({
                    "selector": f"tbody tr:nth-child({df_dre.index.get_loc(linha_verde) + 1}) th",
                    "props": [("background-color", "#006400"), ("color", "white")]
                })

        st.dataframe(
            df_dre.style
                .format(format_brl)
                .apply(aplicar_estilo_dre, axis=1)
                .set_table_styles(style_idx_dre, overwrite=False),
            height=458
        )

        # === CÁLCULO DO RETORNO POR REAL GASTO ===
        st.subheader(f"📈 Retorno por Real Gasto - Cenário {nome}")
        despesas_totais = (
            df_dre.loc["Impostos Sobre Venda"]
            + df_dre.loc["Despesas Operacionais"]
            + df_dre.loc["Despesas Administrativas"]
            + df_dre.loc["Despesas RH"]
            + df_dre.loc["Despesas Extra Operacional"]
            + df_dre.loc["Dividendos"]
            + df_dre.loc["Impostos Sobre Resultado"]
        )
        retorno_por_real = [lucro / despesa if despesa > 0 else 0 for lucro, despesa in zip(df_dre.loc["Lucro Líquido"], despesas_totais)]
        
        df_retorno = pd.DataFrame({"Retorno por Real Gasto (R$)": retorno_por_real}, index=anos).T
        for i, retorno in enumerate(retorno_por_real):
            if retorno <= 0:
                st.warning(f"Ano {i+1}: Sem lucro líquido (retorno não positivo).")
            else:
                st.success(f"Ano {i+1}: Cada R\$ 1,00 gasto gera {format_brl(retorno)} de lucro líquido.")

        st.dataframe(
            df_retorno.style
                .format(format_brl)
                .apply(aplicar_estilo_retorno, axis=1),
            use_container_width=True
        )

        # 🔄 Adiciona os empréstimos ao df_despesas_info com Categoria 'Extra Operacional'
        emprestimos_detalhados = []
        if "emprestimos" in st.session_state:
            for i, emp in enumerate(st.session_state["emprestimos"], start=1):
                descricao = emp.get("objeto", f"Empréstimo {i}").strip() or f"Empréstimo {i}"
                total = emp["valor_parcela"] * min(emp["parcelas"], anos.index(emp["ano_final"]) - anos.index(emp["ano_inicial"]) + 1)
                emprestimos_detalhados.append({
                    "Descrição": f"{descricao} ({emp['ano_inicial']} a {emp['ano_final']})",
                    "Valor": total,
                    "Categoria": "Extra Operacional"
                })

        gerar_excel_download(df_fluxo, df_dre, df_retorno, nome)