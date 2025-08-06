import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from dateutil.relativedelta import relativedelta
from io import BytesIO

st.set_page_config(layout="wide", page_title="Planejamento de Despesas")
st.title("Planejamento de Despesas")

def main():

    anos = [f"Ano {i+1}" for i in range(5)]
    inflacao_padrao = 0.04

    # --- Função para formatar valores em BRL ---
    def format_brl(valor):
        return f"R$ {valor:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")

    # --- Função para obter centros de custo disponíveis ---
    def obter_centros_custo():
        """Obtém lista de centros de custo baseados nas culturas cadastradas + Administrativo"""
        centros = ["Administrativo"]  # Sempre terá o centro administrativo
        
        # Adicionar culturas cadastradas como centros de custo
        if st.session_state.get('plantios'):
            culturas = set()
            for plantio in st.session_state['plantios'].values():
                if plantio.get('cultura'):
                    culturas.add(plantio['cultura'])
            centros.extend(sorted(list(culturas)))
        
        return centros

    # --- Função para calcular rateio administrativo ---
    def calcular_rateio_administrativo():
        """Calcula o percentual de rateio por cultura baseado na área plantada"""
        if not st.session_state.get('plantios'):
            return {}
        
        # Calcular área total e por cultura
        areas_por_cultura = {}
        area_total = 0
        
        for plantio in st.session_state['plantios'].values():
            cultura = plantio.get('cultura', '')
            hectares = plantio.get('hectares', 0)
            
            if cultura and hectares > 0:
                if cultura not in areas_por_cultura:
                    areas_por_cultura[cultura] = 0
                areas_por_cultura[cultura] += hectares
                area_total += hectares
        
        # Calcular percentuais de rateio
        rateio = {}
        if area_total > 0:
            for cultura, area in areas_por_cultura.items():
                rateio[cultura] = (area / area_total) * 100
        
        return rateio

    # --- Inicialização do st.session_state ---
    if "despesas" not in st.session_state:
        st.session_state["despesas"] = []

    if "emprestimos" not in st.session_state:
        st.session_state["emprestimos"] = []

    if "editing_expense_index" not in st.session_state:
        st.session_state["editing_expense_index"] = None

    if "editing_loan_index" not in st.session_state:
        st.session_state["editing_loan_index"] = None

    # --- INFLAÇÃO ---
    with st.expander("📈 Inflação Estimada por Ano"):
        cols = st.columns(5)
        inflacoes = []

        for i, col in enumerate(cols):
            valor = st.session_state.get(f"inf_{i}", inflacao_padrao * 100)
            inflacoes.append(valor)
            with col:
                st.metric(f"Ano {i+1}", f"{valor:.2f}%")

    # --- MODELOS DE EXCEL ---
    st.markdown("### 📥 Modelos de Excel para Preenchimento")

    col_mod1, col_mod2 = st.columns(2)

    with col_mod1:
        st.markdown("**📄 Modelo de Despesas**")
        modelo_despesas = pd.DataFrame({
            "Despesa": ["Energia", "Combustível"],
            "Valor": [10000, 5000],
            "Categoria": ["Operacional", "Operacional"],
            "Centro_Custo": ["Administrativo", "Soja"]
        })

        buffer_despesas = BytesIO()
        with pd.ExcelWriter(buffer_despesas, engine='xlsxwriter') as writer:
            modelo_despesas.to_excel(writer, index=False, sheet_name="Despesas")
        buffer_despesas.seek(0)
        st.download_button("⬇️ Baixar Modelo de Despesas", buffer_despesas, file_name="modelo_despesas.xlsx")

    with col_mod2:
        st.markdown("**📄 Modelo de Empréstimos**")
        modelo_emprestimos = pd.DataFrame({
            "banco": ["Banco X"],
            "valor_total": [100000],
            "objeto": ["Trator"],
            "encargos": [5.5],
            "parcelas": [5],
            "valor_parcela": [20000],
            "periodo": ["ANUAL"],
            "ano_inicial": ["Ano 1"],
            "ano_final": ["Ano 5"],
            "centro_custo": ["Soja"]
        })

        buffer_emprestimos = BytesIO()
        with pd.ExcelWriter(buffer_emprestimos, engine='xlsxwriter') as writer:
            modelo_emprestimos.to_excel(writer, index=False, sheet_name="Emprestimos")
        buffer_emprestimos.seek(0)
        st.download_button("⬇️ Baixar Modelo de Empréstimos", buffer_emprestimos, file_name="modelo_emprestimos.xlsx")

    # --- IMPORTAÇÃO DE EXCEL ---
    with st.expander("📤 Importar Despesas de Excel"):
        despesa_file = st.file_uploader("Upload do arquivo de despesas (.xlsx)", type=["xlsx"], key="upload_despesas")

        if despesa_file and not st.session_state.get("despesas_importadas_ok"):
            try:
                df_despesas = pd.read_excel(despesa_file)
                required_cols = {"Despesa", "Valor", "Categoria"}
                if not required_cols.issubset(df_despesas.columns):
                    st.error(f"Colunas obrigatórias: {required_cols}")
                else:
                    # Adicionar Centro_Custo se não existir
                    if "Centro_Custo" not in df_despesas.columns:
                        df_despesas["Centro_Custo"] = "Administrativo"
                    
                    novas = df_despesas[["Despesa", "Valor", "Categoria", "Centro_Custo"]].dropna().to_dict(orient="records")
                    st.session_state['despesas'].extend(novas)
                    st.session_state["despesas_importadas_ok"] = True
                    st.success(f"{len(novas)} despesas importadas com sucesso!")
                    st.rerun()
            except Exception as e:
                st.error(f"Erro ao ler arquivo: {e}")

    with st.expander("📤 Importar Empréstimos de Excel"):
        emprestimo_file = st.file_uploader("Upload do arquivo de empréstimos (.xlsx)", type=["xlsx"], key="upload_emprestimos")
        if emprestimo_file and not st.session_state.get("emprestimos_importados_ok"):
            try:
                df_emp = pd.read_excel(emprestimo_file)
                required_cols = {
                    "banco", "valor_total", "objeto", "encargos", "parcelas", "valor_parcela", "periodo", "ano_inicial", "ano_final"
                }
                if not required_cols.issubset(df_emp.columns):
                    st.error(f"Colunas obrigatórias: {required_cols}")
                else:
                    total = len(df_emp)
                    emprestimos_importados = []
                    progress_bar = st.progress(0, text="Iniciando importação de empréstimos...")

                    for i, (_, row) in enumerate(df_emp.iterrows()):
                        # Validação: ano_final >= ano_inicial
                        if anos.index(row["ano_final"]) < anos.index(row["ano_inicial"]):
                            st.error(f"Erro na linha {i+1}: Ano Final deve ser maior ou igual ao Ano Inicial.")
                            continue

                        # Adicionar centro_custo se não existir
                        centro_custo = row.get("centro_custo", "Administrativo")

                        novo_emp = {
                            "banco": row["banco"],
                            "valor_total": row["valor_total"],
                            "objeto": row["objeto"],
                            "encargos": row["encargos"],
                            "parcelas": int(row["parcelas"]),
                            "valor_parcela": row["valor_parcela"],
                            "periodo": row["periodo"],
                            "ano_inicial": row["ano_inicial"],
                            "ano_final": row["ano_final"],
                            "centro_custo": centro_custo
                        }

                        emprestimos_importados.append(novo_emp)
                        progress_bar.progress((i + 1) / total, text=f"Importando {i + 1} de {total} empréstimos...")

                    st.session_state["emprestimos"].extend(emprestimos_importados)
                    st.session_state["emprestimos_importados_ok"] = True
                    progress_bar.empty()
                    st.success(f"{total} empréstimos importados com sucesso!")
                    st.rerun()
            except Exception as e:
                st.error(f"Erro ao ler arquivo: {e}")

    # --- EXIBIR RATEIO ATUAL ---
    st.markdown("### 📊 Centro de Custos e Rateio")
    with st.expander("Ver Centros de Custo e Rateio Atual"):
        centros_disponiveis = obter_centros_custo()
        rateio = calcular_rateio_administrativo()
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("**🎯 Centros de Custo Disponíveis:**")
            for centro in centros_disponiveis:
                st.write(f"• {centro}")
        
        with col2:
            st.markdown("**📈 Rateio do Centro Administrativo:**")
            if rateio:
                for cultura, percentual in rateio.items():
                    st.write(f"• {cultura}: {percentual:.2f}%")
            else:
                st.write("• Nenhuma cultura cadastrada ainda")

    # --- ADICIONAR DESPESAS MANUALMENTE ---
    with st.expander("### 💰 Cadastro de Despesas"):

        is_editing = st.session_state['editing_expense_index'] is not None
        expense_to_edit = st.session_state['despesas'][st.session_state['editing_expense_index']] if is_editing else None

        with st.form("form_despesa", clear_on_submit=not is_editing):
            st.subheader(f"{'Editar' if is_editing else 'Adicionar'} Despesa")

            default_nome = expense_to_edit['Despesa'] if is_editing else ""
            default_valor = expense_to_edit['Valor'] if is_editing else 0.0
            categorias = ["Operacional", "RH", "Administrativa", "Extra Operacional", "Dividendos", "Impostos"]
            default_categoria_index = categorias.index(
                expense_to_edit['Categoria']) if is_editing and expense_to_edit['Categoria'] in categorias else 0

            col1, col2 = st.columns(2)
            with col1:
                nome = st.text_input("Nome da Despesa", value=default_nome, key="nome_despesa_input")
                valor = st.number_input("Valor Anual (R$)", min_value=0.0, step=100.0, value=default_valor, key="valor_despesa_input")
            
            with col2:
                categoria = st.selectbox("Categoria", categorias, index=default_categoria_index, key="categoria_select")
                
                # Centro de Custo
                centros_custo = obter_centros_custo()
                default_centro = expense_to_edit.get('Centro_Custo', 'Administrativo') if is_editing else 'Administrativo'
                default_centro_index = centros_custo.index(default_centro) if default_centro in centros_custo else 0
                centro_custo = st.selectbox("Centro de Custo", centros_custo, index=default_centro_index, key="centro_custo_select")

            col_buttons = st.columns([1, 1, 4])
            submit = False
            cancel = False
            with col_buttons[0]:
                submit = st.form_submit_button("Atualizar" if is_editing else "Adicionar")
            with col_buttons[1]:
                if is_editing:
                    cancel = st.form_submit_button("Cancelar Edição")

            if submit:
                if not nome or valor <= 0:
                    st.warning("Preencha todos os campos corretamente.")
                else:
                    nova_despesa = {
                        "Despesa": nome.strip(), 
                        "Valor": valor, 
                        "Categoria": categoria,
                        "Centro_Custo": centro_custo
                    }
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
    with st.expander("### 🏦 Cadastro de Empréstimos e Financiamentos"):

        editing_loan = st.session_state["editing_loan_index"] is not None
        emprestimo_to_edit = st.session_state["emprestimos"][st.session_state["editing_loan_index"]] if editing_loan else {}

        with st.form("form_emprestimo", clear_on_submit=not editing_loan):
            st.subheader(f"{'Editar' if editing_loan else 'Cadastrar'} Empréstimo")

            col1, col2 = st.columns(2)
            with col1:
                banco = st.text_input("Banco/Instituição", value=emprestimo_to_edit.get("banco", ""))
                objeto = st.text_input("Finalidade (ex: Trator, Sementes)", value=emprestimo_to_edit.get("objeto", ""))
                valor_total = st.number_input("Valor Total do Empréstimo", min_value=0.0, step=1000.0, format="%.2f",
                                            value=emprestimo_to_edit.get("valor_total", 0.0))
                encargos = st.number_input("Taxa de Juros (% ao ano)", min_value=0.0, step=0.1,
                                        value=float(emprestimo_to_edit.get("encargos", 0.0)))

            with col2:
                parcelas = st.number_input("Quantidade de Parcelas", min_value=1, step=1,
                                        value=emprestimo_to_edit.get("parcelas", 1))
                periodicidade = st.selectbox("Período de Pagamento", ["ANUAL", "SEMESTRAL", "MENSAL"],
                                            index=["ANUAL", "SEMESTRAL", "MENSAL"].index(
                                                emprestimo_to_edit.get("periodo", "ANUAL")))
                valor_parcela = st.number_input("Valor da Parcela (R$)", min_value=0.0, step=100.0, format="%.2f",
                                              value=emprestimo_to_edit.get("valor_parcela", 0.0))
                
                # Centro de Custo para Empréstimos
                centros_custo = obter_centros_custo()
                default_centro = emprestimo_to_edit.get('centro_custo', 'Administrativo')
                default_centro_index = centros_custo.index(default_centro) if default_centro in centros_custo else 0
                centro_custo_emp = st.selectbox("Centro de Custo", centros_custo, index=default_centro_index, key="centro_custo_emp_select")
                
            # Segunda linha para os anos
            col3, col4 = st.columns(2)
            with col3:
                ano_inicial = st.selectbox("Ano Inicial da Projeção", anos,
                                        index=anos.index(emprestimo_to_edit.get("ano_inicial", "Ano 1")))
            with col4:
                ano_final = st.selectbox("Ano Final da Projeção", anos,
                                        index=anos.index(emprestimo_to_edit.get("ano_final", "Ano 5")))

            # Validação: ano_final >= ano_inicial
            if anos.index(ano_final) < anos.index(ano_inicial):
                st.error("Ano Final deve ser maior ou igual ao Ano Inicial.")
            else:
                col_buttons = st.columns([1, 1, 3])
                enviar = False
                cancelar = False
                
                with col_buttons[0]:
                    enviar = st.form_submit_button("Atualizar" if editing_loan else "Cadastrar")
                with col_buttons[1]:
                    if editing_loan:
                        cancelar = st.form_submit_button("Cancelar")

                if enviar and valor_total > 0 and valor_parcela > 0 and banco.strip() and objeto.strip():
                    novo = {
                        "banco": banco.strip(),
                        "valor_total": valor_total,
                        "objeto": objeto.strip(),
                        "encargos": encargos,
                        "parcelas": parcelas,
                        "valor_parcela": valor_parcela,
                        "periodo": periodicidade,
                        "ano_inicial": ano_inicial,
                        "ano_final": ano_final,
                        "centro_custo": centro_custo_emp
                    }

                    if editing_loan:
                        st.session_state["emprestimos"][st.session_state["editing_loan_index"]] = novo
                        st.session_state["editing_loan_index"] = None
                    else:
                        st.session_state["emprestimos"].append(novo)
                    st.success("Empréstimo salvo com sucesso!")
                    st.rerun()
                elif enviar:
                    st.warning("Preencha todos os campos obrigatórios (Banco, Valor Total > 0, Valor da Parcela > 0 e Finalidade).")
                
                if editing_loan and cancelar:
                    st.session_state["editing_loan_index"] = None
                    st.rerun()

    # --- EXIBIÇÃO DE DESPESAS ---
    st.markdown("### Despesas Cadastradas")
    with st.expander("Despesas Cadastradas"):
        if not st.session_state['despesas']:
            st.info("Nenhuma despesa cadastrada.")
        else:
            for i, d in enumerate(st.session_state['despesas']):
                cols = st.columns([2, 2, 2, 2, 1, 1])
                cols[0].write(d["Despesa"])
                cols[1].write(format_brl(d["Valor"]))
                cols[2].write(d["Categoria"])
                cols[3].write(d.get("Centro_Custo", "Não definido"))
                if cols[4].button("Editar", key=f"edit_{i}"):
                    st.session_state['editing_expense_index'] = i
                    st.rerun()
                if cols[5].button("Excluir", key=f"del_{i}"):
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
                cols[1].write(e["objeto"])
                cols[2].write(format_brl(e["valor_total"]))
                cols[3].write(format_brl(e["valor_parcela"]))
                cols[4].write(e.get("centro_custo", "Não definido"))
                if cols[5].button("Editar", key=f"edit_loan_{i}"):
                    st.session_state["editing_loan_index"] = i
                    st.rerun()
                if cols[6].button("Excluir", key=f"del_loan_{i}"):
                    st.session_state["emprestimos"].pop(i)
                    st.rerun()

    # --- PROJEÇÃO ---
    st.markdown("---")
    st.markdown("### 📊 Projeção de Despesas com Inflação")

    tem_despesas = st.session_state.get('despesas') and isinstance(st.session_state['despesas'], list)
    tem_emprestimos = st.session_state.get('emprestimos') and isinstance(st.session_state['emprestimos'], list)

    # Initialize df_fluxo with empty DataFrame or proper default
    df_fluxo = pd.DataFrame()

    if not tem_despesas and not tem_emprestimos:
        st.info("Adicione despesas ou empréstimos para ver a projeção.")
    else:
        # --- PROJEÇÃO GERAL (como era antes) ---
        df_fluxo = pd.DataFrame(columns=anos)

        if tem_despesas:
            df_desp = pd.DataFrame(st.session_state['despesas'])

            if not df_desp.empty and "Despesa" in df_desp.columns and "Valor" in df_desp.columns:
                df_desp['Despesa_Normalized'] = df_desp['Despesa'].astype(str).str.strip()
                group = df_desp.groupby('Despesa_Normalized')['Valor'].sum()

                for i, ano in enumerate(anos):
                    fator = np.prod([1 + inflacoes[j] / 100 for j in range(i + 1)])
                    df_fluxo[ano] = group * fator if not group.empty else 0

        if tem_emprestimos:
            for emp in st.session_state['emprestimos']:
                linha = f"Empréstimo: {emp['objeto']}"
                if linha not in df_fluxo.index:
                    df_fluxo.loc[linha] = [0] * len(anos)
                start_year_index = anos.index(emp["ano_inicial"])
                end_year_index = anos.index(emp["ano_final"])
                num_years = end_year_index - start_year_index + 1
                for i in range(start_year_index, min(start_year_index + min(emp["parcelas"], num_years), len(anos))):
                    ano = anos[i]
                    df_fluxo.at[linha, ano] += emp["valor_parcela"]

        df_fluxo.index.name = "Despesa"
        
        # --- NOVA PROJEÇÃO POR CENTRO DE CUSTOS ---
        st.markdown("#### 📊 Projeção Geral")
        st.dataframe(df_fluxo.style.format(format_brl))
        
        # Calcular custos por cultura com rateio
        rateio_percentual = {}
        areas_por_cultura = {}
        area_total = 0
        
        if st.session_state.get('plantios'):
            for plantio in st.session_state['plantios'].values():
                cultura = plantio.get('cultura', '')
                hectares = plantio.get('hectares', 0)
                
                if cultura and hectares > 0:
                    if cultura not in areas_por_cultura:
                        areas_por_cultura[cultura] = 0
                    areas_por_cultura[cultura] += hectares
                    area_total += hectares
            
            if area_total > 0:
                for cultura, area in areas_por_cultura.items():
                    rateio_percentual[cultura] = area / area_total
        
        # Criar projeção por cultura
        if areas_por_cultura:
            st.markdown("#### 🌱 Projeção por Cultura (com Rateio Administrativo)")
            
            custos_por_cultura = {}
            for cultura in areas_por_cultura.keys():
                custos_por_cultura[cultura] = pd.DataFrame(columns=anos)
            
            # Processar despesas
            if tem_despesas:
                for despesa in st.session_state['despesas']:
                    nome_despesa = despesa['Despesa']
                    valor_base = despesa['Valor']
                    centro_custo = despesa.get('Centro_Custo', 'Administrativo')
                    
                    if centro_custo == 'Administrativo':
                        # Ratear para todas as culturas
                        for cultura, percentual in rateio_percentual.items():
                            valor_cultura = valor_base * percentual
                            linha_nome = f"{nome_despesa} (Rateio Adm.)"
                            
                            # Aplicar inflação por ano
                            valores_anos = []
                            for i, ano in enumerate(anos):
                                fator_inflacao = np.prod([1 + inflacoes[j] / 100 for j in range(i + 1)])
                                valores_anos.append(valor_cultura * fator_inflacao)
                            
                            if linha_nome not in custos_por_cultura[cultura].index:
                                custos_por_cultura[cultura].loc[linha_nome] = valores_anos
                            else:
                                custos_por_cultura[cultura].loc[linha_nome] += valores_anos
                    
                    elif centro_custo in custos_por_cultura:
                        # Despesa direta da cultura
                        valores_anos = []
                        for i, ano in enumerate(anos):
                            fator_inflacao = np.prod([1 + inflacoes[j] / 100 for j in range(i + 1)])
                            valores_anos.append(valor_base * fator_inflacao)
                        
                        custos_por_cultura[centro_custo].loc[nome_despesa] = valores_anos
            
            # Processar empréstimos
            if tem_emprestimos:
                for emp in st.session_state['emprestimos']:
                    nome_emprestimo = f"Empréstimo: {emp['objeto']}"
                    centro_custo = emp.get('centro_custo', 'Administrativo')
                    valor_parcela = emp.get('valor_parcela', 0)
                    
                    if centro_custo == 'Administrativo':
                        # Ratear para todas as culturas
                        for cultura, percentual in rateio_percentual.items():
                            valor_cultura = valor_parcela * percentual
                            linha_nome = f"{nome_emprestimo} (Rateio Adm.)"
                            
                            # Inicializar linha se não existir
                            if linha_nome not in custos_por_cultura[cultura].index:
                                custos_por_cultura[cultura].loc[linha_nome] = [0] * len(anos)
                            
                            # Aplicar parcelas nos anos apropriados
                            start_year_index = anos.index(emp.get("ano_inicial", "Ano 1"))
                            end_year_index = anos.index(emp.get("ano_final", "Ano 5"))
                            num_years = end_year_index - start_year_index + 1
                            parcelas = emp.get('parcelas', 1)
                            
                            for i in range(start_year_index, min(start_year_index + min(parcelas, num_years), len(anos))):
                                ano = anos[i]
                                custos_por_cultura[cultura].at[linha_nome, ano] += valor_cultura
                    
                    elif centro_custo in custos_por_cultura:
                        # Empréstimo direto da cultura
                        if nome_emprestimo not in custos_por_cultura[centro_custo].index:
                            custos_por_cultura[centro_custo].loc[nome_emprestimo] = [0] * len(anos)
                        
                        start_year_index = anos.index(emp.get("ano_inicial", "Ano 1"))
                        end_year_index = anos.index(emp.get("ano_final", "Ano 5"))
                        num_years = end_year_index - start_year_index + 1
                        parcelas = emp.get('parcelas', 1)
                        
                        for i in range(start_year_index, min(start_year_index + min(parcelas, num_years), len(anos))):
                            ano = anos[i]
                            custos_por_cultura[centro_custo].at[nome_emprestimo, ano] += valor_parcela
            
            # Exibir custos por cultura
            for cultura, df_cultura in custos_por_cultura.items():
                if not df_cultura.empty:
                    st.markdown(f"**🌿 {cultura}**")
                    df_cultura.index.name = "Item"
                    st.dataframe(df_cultura.style.format(format_brl))
                    
                    # Totais da cultura
                    totais_cultura = df_cultura.sum(axis=0)
                    st.markdown(f"**Total {cultura}:**")
                    cols_cultura = st.columns(len(anos))
                    for i, ano in enumerate(anos):
                        with cols_cultura[i]:
                            st.metric(ano, format_brl(totais_cultura.get(ano, 0)))
                    st.markdown("---")
            
            # Salvar dados para outras páginas
            st.session_state['custos_por_cultura'] = custos_por_cultura
            st.session_state['rateio_administrativo'] = rateio_percentual
        
        st.session_state['fluxo_caixa'] = df_fluxo

    # Only calculate totals if df_fluxo has data
    if not df_fluxo.empty:
        totais_por_ano = df_fluxo.sum(axis=0)
    else:
        # Create empty Series with anos as index
        totais_por_ano = pd.Series(0.0, index=anos)

    st.markdown("#### 💵 Total Geral de Despesas por Ano")
    cols_totais = st.columns(len(anos))
    for i, ano in enumerate(anos):
        with cols_totais[i]:
            # Safe access with default value of 0
            valor = totais_por_ano.get(ano, 0.0)
            st.metric(ano, format_brl(valor))

    # --- LIMPAR TUDO ---
    if st.button("Limpar Tudo", key="btn_clear_all"):
        if st.checkbox("Confirmar exclusão de todos os dados"):
            keys_to_clear = list(st.session_state.keys())
            for key in keys_to_clear:
                del st.session_state[key]
            st.session_state["despesas_importadas_ok"] = False
            st.session_state["emprestimos_importados_ok"] = False
            st.success("Todos os dados e uploads foram limpos!")
            st.rerun()
        else:
            st.warning("Marque a caixa de confirmação para limpar os dados.")

    # --- EXPORTAÇÃO ATUALIZADA ---
    st.markdown("---")
    st.markdown("### 📤 Exportar Dados")
    
    def criar_relatorio_excel():
        """Cria um arquivo Excel com fluxo de caixa geral e por cultura"""
        buffer = BytesIO()
        
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            # Aba 1: Fluxo de Caixa Geral
            if not df_fluxo.empty:
                df_export = df_fluxo.copy()
                df_export.to_excel(writer, sheet_name='Fluxo_Caixa_Geral', index=True)
            
            # Aba 2: Resumo de Totais por Ano
            if not totais_por_ano.empty:
                df_totais = pd.DataFrame({
                    'Ano': totais_por_ano.index,
                    'Total (R$)': totais_por_ano.values
                })
                df_totais.to_excel(writer, sheet_name='Totais_Geral', index=False)
            
            # Abas por Cultura
            if st.session_state.get('custos_por_cultura'):
                for cultura, df_cultura in st.session_state['custos_por_cultura'].items():
                    if not df_cultura.empty:
                        df_cultura.to_excel(writer, sheet_name=f'Custos_{cultura}', index=True)
                        
                        # Totais da cultura
                        totais_cultura = df_cultura.sum(axis=0)
                        df_total_cultura = pd.DataFrame({
                            'Ano': totais_cultura.index,
                            'Total (R$)': totais_cultura.values
                        })
                        df_total_cultura.to_excel(writer, sheet_name=f'Total_{cultura}', index=False)
            
            # Aba 3: Despesas Cadastradas
            if st.session_state.get('despesas'):
                df_despesas_export = pd.DataFrame(st.session_state['despesas'])
                df_despesas_export.to_excel(writer, sheet_name='Despesas_Cadastradas', index=False)
            
            # Aba 4: Empréstimos Cadastrados
            if st.session_state.get('emprestimos'):
                df_emprestimos_export = pd.DataFrame(st.session_state['emprestimos'])
                df_emprestimos_export.to_excel(writer, sheet_name='Emprestimos_Cadastrados', index=False)
            
            # Aba 5: Rateio Administrativo
            if st.session_state.get('rateio_administrativo'):
                df_rateio = pd.DataFrame([
                    {'Cultura': cultura, 'Percentual': f"{percentual*100:.2f}%", 'Area_ha': areas_por_cultura.get(cultura, 0)}
                    for cultura, percentual in st.session_state['rateio_administrativo'].items()
                ])
                df_rateio.to_excel(writer, sheet_name='Rateio_Administrativo', index=False)
            
            # Aba 6: Configurações (Inflação)
            df_config = pd.DataFrame({
                'Ano': anos,
                'Inflacao (%)': inflacoes
            })
            df_config.to_excel(writer, sheet_name='Configuracoes', index=False)
        
        buffer.seek(0)
        return buffer
    
    col_export1, col_export2 = st.columns([1, 3])
    
    with col_export1:
        if st.button("📊 Gerar Relatório Excel", type="primary"):
            if tem_despesas or tem_emprestimos:
                try:
                    excel_buffer = criar_relatorio_excel()
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = f"relatorio_despesas_completo_{timestamp}.xlsx"
                    
                    st.download_button(
                        label="⬇️ Baixar Relatório Excel",
                        data=excel_buffer,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_relatorio"
                    )
                    st.success("Relatório gerado com sucesso!")
                except Exception as e:
                    st.error(f"Erro ao gerar relatório: {e}")
            else:
                st.warning("Adicione despesas ou empréstimos antes de exportar.")
    
    with col_export2:
        st.info("📋 **O relatório inclui:**\n- Fluxo de caixa geral\n- Custos detalhados por cultura\n- Rateio administrativo\n- Despesas e empréstimos cadastrados\n- Configurações de inflação")

if __name__ == "__main__":
    main()