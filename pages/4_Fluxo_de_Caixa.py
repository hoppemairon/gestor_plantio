import streamlit as st
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from io import BytesIO

from utils.session import carregar_configuracoes
from utils.dre import calcular_dre
carregar_configuracoes()

st.set_page_config(layout="wide", page_title="Fluxo de Caixa Projeção")
st.title("🎯 Fluxo de Caixa - Cenários (Projetado, Pessimista, Otimista)")

def main():

    # Verificação de dados essenciais
    if "plantios" not in st.session_state or not st.session_state["plantios"]:
        st.warning("Cadastre ao menos um plantio para gerar o fluxo de caixa.")
        st.stop()
    if "fluxo_caixa" not in st.session_state:
        st.warning("Você precisa preencher as despesas antes de acessar esta página.")
        st.stop()

    inflacao_padrao = 0.04
    anos = [f"Ano {i+1}" for i in range(5)]
    inflacoes = [st.session_state.get(f"inf_{i}", 4.0) for i in range(5)]
    plantios = st.session_state["plantios"]
    df_base_fluxo = st.session_state["fluxo_caixa"]

    with st.expander("🔧 Cenário e Inflação"):
        st.markdown("### 📈 Inflação Estimada por Ano")
        cols = st.columns(5)
        inflacoes = []

        for i, col in enumerate(cols):
            valor = st.session_state.get(f"inf_{i}", inflacao_padrao * 100)
            inflacoes.append(valor)
            with col:
                st.metric(f"Ano {i+1}", f"{valor:.2f}%")

        pess_receita = st.session_state["pess_receita"]
        pess_despesas = st.session_state["pess_despesas"]
        otm_receita = st.session_state["otm_receita"]
        otm_despesas = st.session_state["otm_despesas"]

        st.markdown("### 🔧 Parâmetros de Cenário Atuais")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("💸 Receita Pessimista", f"-{st.session_state.get('pess_receita', 15)}%")
        with col2:    
            st.metric("💰 Despesa Pessimista", f"+{st.session_state.get('pess_despesas', 10)}%")
        with col3:
            st.metric("💸 Receita Otimista", f"+{st.session_state.get('otm_receita', 10)}%")
        with col4:
            st.metric("💰 Despesa Otimista", f"-{st.session_state.get('otm_despesas', 10)}%")

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

    # Adicionar Receitas Extras
    receitas_extras = {"Operacional": [0] * 5, "Extra Operacional": [0] * 5}
    if "receitas_adicionais" in st.session_state:
        for receita in st.session_state["receitas_adicionais"].values():
            valor = receita["valor"]
            categoria = receita["categoria"]
            for ano in receita["anos_aplicacao"]:
                idx = anos.index(ano)
                if categoria == "Operacional":
                    fator = np.prod([1 + inflacoes[j] / 100 for j in range(idx + 1)])
                    receitas_extras["Operacional"][idx] += valor * fator
                else:
                    receitas_extras["Extra Operacional"][idx] += valor

    # CENÁRIOS DE RECEITA
    receitas = {
        "Projetado": [receita_base[i] + receitas_extras["Operacional"][i] for i in range(5)],
        "Pessimista": [(receita_base[i] + receitas_extras["Operacional"][i]) * (1 - pess_receita / 100) for i in range(5)],
        "Otimista": [(receita_base[i] + receitas_extras["Operacional"][i]) * (1 + otm_receita / 100) for i in range(5)]
    }

    # CENÁRIOS DE FLUXO DE DESPESAS
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

    col1, col2 = st.columns(2)

    with col1:
        with st.expander("🧾 Entenda como a receita futura é calculada por cultura"):
            st.markdown("""
            Para projetar as receitas futuras, consideramos a contribuição individual de cada cultura plantada e receitas adicionais, aplicando a taxa de inflação anual apenas às receitas operacionais.

            **Passos do Cálculo:**

            1.  **Receita Base por Cultura:** Para cada tipo de plantio que você cadastrou (ex: Milho, Soja), calculamos a receita esperada no ano base (sem inflação) multiplicando:
                *   `Hectares Plantados` (para aquela cultura)
                *   `Sacas por Hectare` (produtividade esperada para aquela cultura)
                *   `Preço por Saca` (preço de venda esperado para aquela cultura)

                Isso nos dá a **receita bruta inicial** que cada cultura contribui.

            2.  **Receitas Adicionais:** Incluímos receitas operacionais (com inflação) e extra operacionais (sem inflação) conforme cadastradas.

            3.  **Projeção com Inflação:** A receita base de cada cultura e receitas operacionais adicionais é projetada para os próximos 5 anos com inflação acumulada. Receitas extra operacionais não sofrem ajuste de inflação.

                *   **Exemplo (Ano 1):** `(Receita Base da Cultura + Receita Operacional) * (1 + Inflação Ano 1) + Receita Extra Operacional`
                *   **Exemplo (Ano 2):** `(Receita Base da Cultura + Receita Operacional) * (1 + Inflação Ano 1) * (1 + Inflação Ano 2) + Receita Extra Operacional`

            4.  **Receita Total Estimada:** A "Receita Estimada" no Fluxo de Caixa Consolidado é a soma das receitas projetadas de todas as culturas e receitas adicionais para cada ano.
            """)

    with col2:
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

    def format_brl(x):
        try:
            return f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except Exception:
            return x

    def aplicar_estilo_fluxo(linha):
        if linha.name == "Receita Estimada":
            return ["background-color: #003366; color: white;" for _ in linha]
        elif linha.name == "Lucro Líquido":
            return ["background-color: #006400; color: white;" for _ in linha]
        elif linha.name == "Receita Extra Operacional":
            return ["background-color: #0059b2; color: white;" for _ in linha]
        else:
            return ["" for _ in linha]

    def aplicar_estilo_dre(linha):
        if linha.name == "Receita":
            return ["background-color: #003366; color: white;" for _ in linha]
        elif linha.name in ["Margem de Contribuição", "Resultado Operacional", "Lucro Operacional", "Lucro Líquido"]:
            return ["background-color: #006400; color: white;" for _ in linha]
        elif linha.name == "Receita Extra Operacional":
            return ["background-color: #0059b2; color: white;" for _ in linha]
        else:
            return ["" for _ in linha]

    def aplicar_estilo_retorno(linha):
        return ["background-color: #FF4040; color: white;" if x <= 0 else "" for x in linha]

    def gerar_excel_download(df_fluxo, df_dre, df_retorno, resumo, nome_cenario):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_fluxo.to_excel(writer, sheet_name="Fluxo de Caixa")
            df_dre.to_excel(writer, sheet_name="DRE")
            resumo.to_excel(writer, sheet_name="Resumo Financeiro")
            df_retorno.to_excel(writer, sheet_name="Retorno por Real Gasto")
            workbook = writer.book
            currency_format = workbook.add_format({'num_format': 'R$ #,##0.00'})
            for sheet in writer.sheets.values():
                for col_num in range(len(df_fluxo.columns) + 1):
                    sheet.set_column(col_num, col_num, 15, currency_format)
        output.seek(0)
        st.download_button(
            label=f"⬇️ Baixar Excel - {nome_cenario}",
            data=output,
            file_name=f"cenario_{nome_cenario.lower()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"download_excel_{nome_cenario.lower()}"
        )

    for aba, nome in zip(abas, nomes_cenarios):
        with aba:
            st.subheader(f"📊 Fluxo de Caixa - Cenário {nome}")
            df_fluxo = fluxos[nome].copy()
            df_fluxo.loc["Receita Estimada"] = receitas[nome]
            df_fluxo.loc["Receita Extra Operacional"] = receitas_extras["Extra Operacional"]

            # Adicionar empréstimos ao fluxo de caixa (alinhado com DRE)
            emprestimos_por_ano = [0] * len(anos)
            if "emprestimos" in st.session_state:
                for emp in st.session_state["emprestimos"]:
                    try:
                        linha = f"Empréstimo: {emp['objeto']}"
                        if linha not in df_fluxo.index:
                            df_fluxo.loc[linha] = [0] * len(anos)
                        start_year_index = anos.index(emp["ano_inicial"])
                        end_year_index = anos.index(emp["ano_final"])
                        parcelas_restantes = emp["parcelas"]
                        for i in range(start_year_index, min(end_year_index + 1, len(anos))):
                            if parcelas_restantes > 0:
                                ajuste = (pess_despesas if nome == "Pessimista" else (-otm_despesas if nome == "Otimista" else 0)) / 100
                                valor_parcela = emp["valor_parcela"] * (1 + ajuste)
                                df_fluxo.at[linha, anos[i]] = valor_parcela  # Atribui em vez de somar
                                emprestimos_por_ano[i] = valor_parcela  # Define o valor para o DRE
                                parcelas_restantes -= 1
                    except (ValueError, KeyError):
                        st.warning(f"Empréstimo inválido: {emp.get('objeto', 'Desconhecido')}. Ignorando.")
                        continue

            ordem = ["Receita Estimada", "Receita Extra Operacional"] + [i for i in df_fluxo.index if i not in ["Receita Estimada", "Receita Extra Operacional"]]
            df_fluxo = df_fluxo.loc[ordem]

            style_idx_fluxo = [
                {
                    "selector": f"tbody tr:nth-child({df_fluxo.index.get_loc('Receita Estimada') + 1}) th",
                    "props": [("background-color", "#003366"), ("color", "white")]
                },
                {
                    "selector": f"tbody tr:nth-child({df_fluxo.index.get_loc('Receita Extra Operacional') + 1}) th",
                    "props": [("background-color", "#4682B4"), ("color", "white")]
                }
            ]

            st.dataframe(
                df_fluxo.style
                    .format(format_brl)
                    .apply(aplicar_estilo_fluxo, axis=1)
                    .set_table_styles(style_idx_fluxo, overwrite=False),
                use_container_width=True
            )

            st.subheader(f"📘 DRE - Cenário {nome}")

            df_despesas_info = pd.DataFrame(st.session_state.get("despesas", []))
            if not df_despesas_info.empty and "Categoria" in df_despesas_info.columns:
                df_despesas_info["Categoria"] = df_despesas_info["Categoria"].astype(str).str.strip()
            else:
                df_despesas_info = pd.DataFrame(columns=["Categoria", "Valor"])

            dre_calc = calcular_dre(
                nome, inflacoes, anos, hectares_total, total_sacas, preco_total,
                receitas, receitas_extras,
                df_despesas_info,
                st.session_state.get("emprestimos", []),
                pess_despesas, otm_despesas
            )

            # --- ADICIONE ESTAS LINHAS AQUI ---
            if "dre_cenarios" not in st.session_state:
                st.session_state["dre_cenarios"] = {}
            st.session_state["dre_cenarios"][nome] = dre_calc

            df_dre = pd.DataFrame(dre_calc, index=anos).T.loc[[ 
                "Receita", "Impostos Sobre Venda", "Despesas Operacionais",
                "Margem de Contribuição", "Despesas Administrativas", "Despesas RH",
                "Resultado Operacional", "Despesas Extra Operacional",
                "Lucro Operacional", "Impostos Sobre Resultado",
                "Receita Extra Operacional", "Dividendos", "Lucro Líquido"
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
                height=495
            )

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
                    st.success(f"Ano {i+1}: Cada R\\$ 1,00 gasto gera {format_brl(retorno)} de lucro líquido.")

            st.dataframe(
                df_retorno.style
                    .format(format_brl)
                    .apply(aplicar_estilo_retorno, axis=1),
                use_container_width=True
            )

            emprestimos_detalhados = []
            if "emprestimos" in st.session_state:
                for i, emp in enumerate(st.session_state["emprestimos"], start=1):
                    try:
                        descricao = emp.get("objeto", f"Empréstimo {i}").strip() or f"Empréstimo {i}"
                        start_year_index = anos.index(emp["ano_inicial"])
                        end_year_index = anos.index(emp["ano_final"])
                        num_years = end_year_index - start_year_index + 1
                        total = emp["valor_parcela"] * min(emp["parcelas"], num_years)
                        emprestimos_detalhados.append({
                            "Descrição": f"{descricao} ({emp['ano_inicial']} a {emp['ano_final']})",
                            "Valor": total,
                            "Categoria": "Extra Operacional"
                        })
                    except (ValueError, KeyError):
                        st.warning(f"Empréstimo inválido: {emp.get('objeto', 'Desconhecido')}. Ignorando.")
                        continue
            
            st.subheader(f"📌 Resumo Financeiro Anual - Cenário {nome}")

            resumo = pd.DataFrame({
                "Receita": df_dre.loc["Receita"],
                "Despesas Totais": despesas_totais,
                "Lucro Líquido": df_dre.loc["Lucro Líquido"]
            }, index=anos)

            st.dataframe(
                resumo.style.format(format_brl),
                use_container_width=True
            )

            gerar_excel_download(df_fluxo, df_dre, df_retorno, resumo, nome)

    # --- ADICIONE ESTAS LINHAS FORA DO LOOP, APÓS ELE TERMINAR ---
        st.session_state["receitas_cenarios"] = receitas
        st.session_state["inflacoes"] = inflacoes
        st.session_state["anos"] = anos

if __name__ == "__main__":
    main()