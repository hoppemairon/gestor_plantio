import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime

def criar_relatorio_ppt(all_indicators, all_dre_data, df_culturas_for_excel, nomes_cenarios, anos):
    """Cria uma apresentação PowerPoint editável com todos os dados das duas telas"""
    try:
        # Verificar e importar bibliotecas necessárias
        try:
            from pptx import Presentation
            from pptx.util import Inches, Pt
            from pptx.enum.text import PP_ALIGN
            from pptx.dml.color import RGBColor
        except ImportError as e:
            st.error(f"""
            ❌ **Biblioteca PowerPoint não encontrada!**
            
            **Para usar a exportação PPT, instale a biblioteca:**
            ```bash
            pip install python-pptx
            ```
            
            **Se você estiver usando conda:**
            ```bash
            conda install -c conda-forge python-pptx
            ```
            
            **Erro específico:** {str(e)}
            """)
            return None
        
        try:
            import matplotlib.pyplot as plt
            import matplotlib
            matplotlib.use('Agg')
        except ImportError as e:
            st.error(f"""
            ❌ **Biblioteca Matplotlib não encontrada!**
            
            **Para gráficos no PPT, instale:**
            ```bash
            pip install matplotlib
            ```
            
            **Erro:** {str(e)}
            """)
            return None
        
        # Criar apresentação
        prs = Presentation()
        
        # === SLIDE 1: ABERTURA ===
        slide_layout = prs.slide_layouts[0]  # Title slide
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        subtitle = slide.placeholders[1]
        
        title.text = "RELATÓRIO FINANCEIRO COMPLETO"
        subtitle.text = f"Gestor de Plantio - Fluxo de Caixa e Indicadores\nGerado em: {datetime.now().strftime('%d/%m/%Y')}"
        
        # === SLIDE 2: RECEITA POR CULTURA E DADOS BASE DO SISTEMA ===
        slide_layout = prs.slide_layouts[5]  # Blank slide
        slide = prs.slides.add_slide(slide_layout)
        
        # Título
        title_shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.1), Inches(9), Inches(0.6))
        title_frame = title_shape.text_frame
        title_p = title_frame.paragraphs[0]
        title_p.text = "RECEITA POR CULTURA (ANO BASE) E DADOS BASE DO SISTEMA"
        title_p.font.size = Pt(18)
        title_p.font.bold = True
        title_p.alignment = PP_ALIGN.CENTER
        
        # Tabela de culturas (lado esquerdo)
        rows_culturas = min(len(df_culturas_for_excel) + 1, 8)
        cols_culturas = 4
        table_culturas = slide.shapes.add_table(rows_culturas, cols_culturas, Inches(0.2), Inches(0.8), Inches(4.3), Inches(3)).table
        
        # Cabeçalho culturas
        headers_culturas = ['Cultura', 'Receita Total', 'Área (ha)', 'Receita/ha']
        for col, header in enumerate(headers_culturas):
            cell = table_culturas.cell(0, col)
            cell.text = header
            cell.text_frame.paragraphs[0].font.bold = True
            cell.text_frame.paragraphs[0].font.size = Pt(8)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(76, 175, 80)
            cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
        
        # Dados culturas
        for row, (_, cultura_data) in enumerate(df_culturas_for_excel.head(rows_culturas-1).iterrows(), 1):
            table_culturas.cell(row, 0).text = str(cultura_data['Cultura'])
            table_culturas.cell(row, 1).text = f"R$ {cultura_data['Receita Total']:,.0f}".replace(",", ".")
            table_culturas.cell(row, 2).text = f"{cultura_data['Área (ha)']:.1f}"
            table_culturas.cell(row, 3).text = f"R$ {cultura_data['Receita por ha']:,.0f}".replace(",", ".")
            for col in range(4):
                table_culturas.cell(row, col).text_frame.paragraphs[0].font.size = Pt(8)
        
        # Plantios cadastrados (lado direito, parte superior)
        if st.session_state.get('plantios'):
            subtitle1 = slide.shapes.add_textbox(Inches(4.8), Inches(0.8), Inches(4), Inches(0.4))
            subtitle1.text_frame.paragraphs[0].text = "PLANTIOS CADASTRADOS:"
            subtitle1.text_frame.paragraphs[0].font.bold = True
            subtitle1.text_frame.paragraphs[0].font.size = Pt(10)
            
            plantios_data = []
            for nome, dados in st.session_state['plantios'].items():
                plantios_data.append([
                    nome[:15],  # Truncar nome se muito longo
                    dados.get('cultura', '')[:10],
                    f"{dados.get('hectares', 0):.0f}",
                    f"R$ {dados.get('preco_saca', 0):.0f}"
                ])
            
            if plantios_data:
                rows_plantios = min(len(plantios_data) + 1, 6)
                cols_plantios = 4
                table_plantios = slide.shapes.add_table(rows_plantios, cols_plantios, Inches(4.8), Inches(1.2), Inches(4), Inches(1.5)).table
                
                # Cabeçalho plantios
                headers_plantios = ['Nome', 'Cultura', 'Ha', 'R$/Saca']
                for col, header in enumerate(headers_plantios):
                    cell = table_plantios.cell(0, col)
                    cell.text = header
                    cell.text_frame.paragraphs[0].font.bold = True
                    cell.text_frame.paragraphs[0].font.size = Pt(8)
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor(128, 128, 128)
                    cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
                
                # Dados plantios
                for row, plantio_data in enumerate(plantios_data[:rows_plantios-1], 1):
                    for col, value in enumerate(plantio_data):
                        table_plantios.cell(row, col).text = str(value)
                        table_plantios.cell(row, col).text_frame.paragraphs[0].font.size = Pt(7)
        
        # Despesas por categoria (lado direito, parte inferior)
        if st.session_state.get('despesas'):
            subtitle2 = slide.shapes.add_textbox(Inches(4.8), Inches(2.8), Inches(4), Inches(0.4))
            subtitle2.text_frame.paragraphs[0].text = "DESPESAS POR CATEGORIA:"
            subtitle2.text_frame.paragraphs[0].font.bold = True
            subtitle2.text_frame.paragraphs[0].font.size = Pt(10)
            
            df_despesas = pd.DataFrame(st.session_state['despesas'])
            despesas_resumo = df_despesas.groupby('Categoria')['Valor'].sum()
            
            despesas_data = []
            for categoria, valor in despesas_resumo.items():
                despesas_data.append([categoria[:15], f"R$ {valor:,.0f}".replace(",", ".")])
            
            if despesas_data:
                rows_despesas = min(len(despesas_data) + 1, 6)
                cols_despesas = 2
                table_despesas = slide.shapes.add_table(rows_despesas, cols_despesas, Inches(4.8), Inches(3.2), Inches(4), Inches(1.5)).table
                
                # Cabeçalho despesas
                headers_despesas = ['Categoria', 'Valor Total']
                for col, header in enumerate(headers_despesas):
                    cell = table_despesas.cell(0, col)
                    cell.text = header
                    cell.text_frame.paragraphs[0].font.bold = True
                    cell.text_frame.paragraphs[0].font.size = Pt(8)
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor(255, 165, 0)
                    cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
                
                # Dados despesas
                for row, (categoria, valor) in enumerate(despesas_data[:rows_despesas-1], 1):
                    table_despesas.cell(row, 0).text = categoria
                    table_despesas.cell(row, 1).text = valor
                    table_despesas.cell(row, 0).text_frame.paragraphs[0].font.size = Pt(7)
                    table_despesas.cell(row, 1).text_frame.paragraphs[0].font.size = Pt(7)
        
        # === SLIDES POR CENÁRIO (na ordem: Projetado, Pessimista, Otimista) ===
        for i, cenario in enumerate(nomes_cenarios):
            
            # SLIDE: DRE COMPLETO POR CENÁRIO
            slide_layout = prs.slide_layouts[5]  # Blank slide
            slide = prs.slides.add_slide(slide_layout)
            
            # Título
            title_shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.8))
            title_frame = title_shape.text_frame
            title_p = title_frame.paragraphs[0]
            title_p.text = f"DRE COMPLETO - CENÁRIO {cenario.upper()}"
            title_p.font.size = Pt(18)
            title_p.font.bold = True
            title_p.alignment = PP_ALIGN.CENTER
            
            # Tabela DRE completo
            dre_data = all_dre_data[cenario]
            dre_items = list(dre_data.keys())
            rows = len(dre_items) + 1
            cols = len(anos) + 1
            
            table = slide.shapes.add_table(rows, cols, Inches(0.5), Inches(1.2), Inches(8.5), Inches(5.5)).table
            
            # Cabeçalho
            table.cell(0, 0).text = "Item DRE"
            for col, ano in enumerate(anos, 1):
                cell = table.cell(0, col)
                cell.text = str(ano)
                cell.text_frame.paragraphs[0].font.bold = True
                cell.text_frame.paragraphs[0].font.size = Pt(9)
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(25, 25, 112)
                cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
            
            # Dados DRE
            for row, item in enumerate(dre_items, 1):
                table.cell(row, 0).text = item
                table.cell(row, 0).text_frame.paragraphs[0].font.bold = True
                table.cell(row, 0).text_frame.paragraphs[0].font.size = Pt(8)
                
                values = dre_data[item]
                for col, value in enumerate(values, 1):
                    table.cell(row, col).text = f"R$ {value:,.0f}".replace(",", ".")
                    table.cell(row, col).text_frame.paragraphs[0].font.size = Pt(8)
            
            # SLIDE: TODOS OS INDICADORES POR CENÁRIO
            slide_layout = prs.slide_layouts[5]  # Blank slide
            slide = prs.slides.add_slide(slide_layout)
            
            # Título
            title_shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.1), Inches(9), Inches(0.6))
            title_frame = title_shape.text_frame
            title_p = title_frame.paragraphs[0]
            title_p.text = f"TODOS OS INDICADORES - CENÁRIO {cenario.upper()}"
            title_p.font.size = Pt(18)
            title_p.font.bold = True
            title_p.alignment = PP_ALIGN.CENTER
            
            # Tabela com TODOS os indicadores
            indicators = all_indicators[cenario]
            indicator_list = list(indicators.items())
            
            # Dividir em duas tabelas
            mid_point = len(indicator_list) // 2
            first_half = indicator_list[:mid_point]
            second_half = indicator_list[mid_point:]
            
            # Primeira tabela
            rows1 = len(first_half) + 1
            cols = 6  # Indicador + 5 anos
            
            table1 = slide.shapes.add_table(rows1, cols, Inches(0.2), Inches(0.8), Inches(4.3), Inches(3)).table
            
            # Cabeçalho primeira tabela
            headers = ['Indicador', 'Ano 1', 'Ano 2', 'Ano 3', 'Ano 4', 'Ano 5']
            for col, header in enumerate(headers):
                cell = table1.cell(0, col)
                cell.text = header
                cell.text_frame.paragraphs[0].font.bold = True
                cell.text_frame.paragraphs[0].font.size = Pt(8)
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(68, 114, 196)
                cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
            
            # Dados primeira tabela
            for row, (indicator_name, values) in enumerate(first_half, 1):
                short_name = indicator_name[:18] + "..." if len(indicator_name) > 18 else indicator_name
                table1.cell(row, 0).text = short_name
                table1.cell(row, 0).text_frame.paragraphs[0].font.size = Pt(7)
                
                if isinstance(values, list):
                    for col, value in enumerate(values, 1):
                        if col < cols:
                            if 'R$' in indicator_name:
                                formatted_value = f"R$ {value:,.0f}".replace(",", ".")
                            elif '%' in indicator_name:
                                formatted_value = f"{value:.1f}%"
                            elif value == float('inf'):
                                formatted_value = "∞"
                            else:
                                formatted_value = f"{value:.2f}"
                            table1.cell(row, col).text = formatted_value
                            table1.cell(row, col).text_frame.paragraphs[0].font.size = Pt(7)
                else:
                    formatted_value = f"{values:.1f}%" if 'CAGR' in indicator_name else f"{values:.2f}"
                    table1.cell(row, 1).text = formatted_value
                    table1.cell(row, 1).text_frame.paragraphs[0].font.size = Pt(7)
                    for col in range(2, cols):
                        table1.cell(row, col).text = "-"
                        table1.cell(row, col).text_frame.paragraphs[0].font.size = Pt(7)
            
            # Segunda tabela
            if second_half:
                rows2 = len(second_half) + 1
                table2 = slide.shapes.add_table(rows2, cols, Inches(4.8), Inches(0.8), Inches(4.3), Inches(3)).table
                
                # Cabeçalho segunda tabela
                for col, header in enumerate(headers):
                    cell = table2.cell(0, col)
                    cell.text = header
                    cell.text_frame.paragraphs[0].font.bold = True
                    cell.text_frame.paragraphs[0].font.size = Pt(8)
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor(68, 114, 196)
                    cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
                
                # Dados segunda tabela
                for row, (indicator_name, values) in enumerate(second_half, 1):
                    short_name = indicator_name[:18] + "..." if len(indicator_name) > 18 else indicator_name
                    table2.cell(row, 0).text = short_name
                    table2.cell(row, 0).text_frame.paragraphs[0].font.size = Pt(7)
                    
                    if isinstance(values, list):
                        for col, value in enumerate(values, 1):
                            if col < cols:
                                if 'R$' in indicator_name:
                                    formatted_value = f"R$ {value:,.0f}".replace(",", ".")
                                elif '%' in indicator_name:
                                    formatted_value = f"{value:.1f}%"
                                elif value == float('inf'):
                                    formatted_value = "∞"
                                else:
                                    formatted_value = f"{value:.2f}"
                                table2.cell(row, col).text = formatted_value
                                table2.cell(row, col).text_frame.paragraphs[0].font.size = Pt(7)
                    else:
                        formatted_value = f"{values:.1f}%" if 'CAGR' in indicator_name else f"{values:.2f}"
                        table2.cell(row, 1).text = formatted_value
                        table2.cell(row, 1).text_frame.paragraphs[0].font.size = Pt(7)
                        for col in range(2, cols):
                            table2.cell(row, col).text = "-"
                            table2.cell(row, col).text_frame.paragraphs[0].font.size = Pt(7)
            
            # SLIDE: PARECER FINANCEIRO COMPLETO POR CENÁRIO
            slide_layout = prs.slide_layouts[5]  # Blank slide
            slide = prs.slides.add_slide(slide_layout)
            
            # Título
            title_shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.8))
            title_frame = title_shape.text_frame
            title_p = title_frame.paragraphs[0]
            title_p.text = f"PARECER FINANCEIRO - CENÁRIO {cenario.upper()}"
            title_p.font.size = Pt(18)
            title_p.font.bold = True
            title_p.alignment = PP_ALIGN.CENTER
            
            # Gerar parecer completo (replicando a lógica da função generate_financial_opinion)
            indicators = all_indicators[cenario]
            
            margem_media = np.mean(indicators["Margem Líquida (%)"])
            retorno_medio = np.mean(indicators["Retorno por Real Gasto"])
            liquidez_media = np.mean(indicators["Liquidez Operacional"])
            endividamento_medio = np.mean(indicators["Endividamento (%)"])
            produtividade_media = np.mean(indicators["Produtividade por Hectare (R$/ha)"])
            custo_receita_media = np.mean(indicators["Custo por Receita (%)"])
            dscr_values = [x for x in indicators["DSCR"] if x != float("inf")]
            dscr_medio = np.mean(dscr_values) if dscr_values else float("inf")
            break_even_media = np.mean(indicators["Break-Even Yield (sacas/ha)"])
            roa_medio = np.mean(indicators["ROA (%)"])
            
            parecer = []
            
            # Margem Líquida
            if margem_media < 10:
                parecer.append(f"• Margem Líquida Baixa ({margem_media:.2f}%): Rentabilidade abaixo do ideal. Considere renegociar preços com fornecedores ou investir em culturas de maior valor agregado.")
            else:
                parecer.append(f"• Margem Líquida Saudável ({margem_media:.2f}%): Boa rentabilidade. Monitore custos para manter a consistência.")

            # Retorno por Real Gasto
            if retorno_medio < 0.2:
                parecer.append(f"• Retorno por Real Gasto Baixo ({retorno_medio:.2f}): Gastos com baixo retorno. Avalie a redução de despesas operacionais ou otimize processos agrícolas.")
            else:
                parecer.append(f"• Retorno por Real Gasto Adequado ({retorno_medio:.2f}): Investimentos geram retorno satisfatório. Considere reinvestir em tecnologia para aumentar a produtividade.")

            # Liquidez Operacional
            if liquidez_media < 1.5:
                parecer.append(f"• Liquidez Operacional Baixa ({liquidez_media:.2f}): Risco de dificuldades para cobrir custos operacionais. Negocie prazos de pagamento ou busque linhas de crédito de curto prazo.")
            else:
                parecer.append(f"• Liquidez Operacional Confortável ({liquidez_media:.2f}): Boa capacidade de sustentar operações. Mantenha reservas para safras incertas.")

            # Endividamento
            if endividamento_medio > 30:
                parecer.append(f"• Alto Endividamento ({endividamento_medio:.2f}%): Dívidas elevadas. Priorize a quitação de empréstimos de alto custo ou renegocie taxas de juros.")
            else:
                parecer.append(f"• Endividamento Controlado ({endividamento_medio:.2f}%): Dívidas em nível gerenciável. Considere investimentos estratégicos, como expansão de área plantada.")

            # Custo por Receita
            if custo_receita_media > 70:
                parecer.append(f"• Custo por Receita Alto ({custo_receita_media:.2f}%): Custos operacionais consomem grande parte da receita. Analise insumos e processos para reduzir despesas.")
            else:
                parecer.append(f"• Custo por Receita Controlado ({custo_receita_media:.2f}%): Boa gestão de custos. Continue monitorando preços de insumos.")

            # DSCR
            if dscr_medio != float("inf") and dscr_medio < 1.25:
                parecer.append(f"• DSCR Baixo ({dscr_medio:.2f}): Risco de dificuldades no pagamento de dívidas. Considere reestruturar financiamentos ou aumentar a receita.")
            else:
                dscr_text = f"{dscr_medio:.2f}" if dscr_medio != float("inf") else "∞"
                parecer.append(f"• DSCR Adequado ({dscr_text}): Boa capacidade de cobrir dívidas. Mantenha o lucro operacional estável.")

            # ROA
            if roa_medio < 5:
                parecer.append(f"• ROA Baixo ({roa_medio:.2f}%): Baixa eficiência no uso de ativos. Avalie a venda de ativos ociosos ou investimentos em equipamentos mais produtivos.")
            else:
                parecer.append(f"• ROA Adequado ({roa_medio:.2f}%): Boa utilização dos ativos. Considere expansão controlada ou modernização.")

            # CAGR
            if indicators["CAGR Lucro Líquido (%)"] < 0:
                parecer.append(f"• Crescimento Negativo do Lucro ({indicators['CAGR Lucro Líquido (%)']:.2f}%): Lucro em queda. Revisar estratégias de custo, preço e produtividade.")
            else:
                parecer.append(f"• Crescimento do Lucro ({indicators['CAGR Lucro Líquido (%)']:.2f}%): Lucro em trajetória positiva. Considere reinvestir em áreas estratégicas.")

            # Texto do parecer
            parecer_text = slide.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(8.5), Inches(5.5))
            parecer_frame = parecer_text.text_frame
            parecer_frame.word_wrap = True
            
            for i, item in enumerate(parecer):
                if i == 0:
                    p = parecer_frame.paragraphs[0]
                else:
                    p = parecer_frame.add_paragraph()
                p.text = item
                p.font.size = Pt(11)
                p.space_after = Pt(4)
            
            # Inserir configurações apenas após o cenário Projetado
            if cenario == "Projetado":
                # === SLIDE: CONFIGURAÇÕES E PARÂMETROS ===
                slide_layout = prs.slide_layouts[5]  # Blank slide
                slide = prs.slides.add_slide(slide_layout)
                
                # Título
                title_shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.8))
                title_frame = title_shape.text_frame
                title_p = title_frame.paragraphs[0]
                title_p.text = "CONFIGURAÇÕES E PREMISSAS DO MODELO"
                title_p.font.size = Pt(18)
                title_p.font.bold = True
                title_p.alignment = PP_ALIGN.CENTER
                
                # Texto com configurações
                config_text = slide.shapes.add_textbox(Inches(1), Inches(1.2), Inches(8), Inches(5.5))
                config_frame = config_text.text_frame
                config_frame.word_wrap = True
                
                # Parâmetros de cenário
                config_p1 = config_frame.paragraphs[0]
                config_p1.text = "PARÂMETROS DE CENÁRIO:"
                config_p1.font.size = Pt(14)
                config_p1.font.bold = True
                
                config_info = [
                    f"• Receita Pessimista: -{st.session_state.get('pess_receita', 15)}%",
                    f"• Despesa Pessimista: +{st.session_state.get('pess_despesas', 10)}%",
                    f"• Receita Otimista: +{st.session_state.get('otm_receita', 10)}%",
                    f"• Despesa Otimista: -{st.session_state.get('otm_despesas', 10)}%",
                    "",
                    "INFLAÇÃO PROJETADA:",
                ]
                
                # Adicionar inflação por ano
                for i, ano in enumerate(anos):
                    inflacao = st.session_state.get('inflacoes', [4.0] * len(anos))[i]
                    config_info.append(f"• {ano}: {inflacao:.1f}%")
                
                # Adicionar resumo dos totais
                config_info.extend([
                    "",
                    "RESUMO GERAL:",
                    f"• Total de Hectares: {sum(p.get('hectares', 0) for p in st.session_state.get('plantios', {}).values()):.1f} ha",
                    f"• Número de Plantios: {len(st.session_state.get('plantios', {}))}",
                    f"• Número de Despesas: {len(st.session_state.get('despesas', []))}",
                    f"• Número de Empréstimos: {len(st.session_state.get('emprestimos', []))}"
                ])
                
                for info in config_info:
                    p = config_frame.add_paragraph()
                    p.text = info
                    p.font.size = Pt(11)
        
        # === SLIDE: GRÁFICO RECEITA vs LUCRO LÍQUIDO ===
        slide_layout = prs.slide_layouts[5]  # Blank slide
        slide = prs.slides.add_slide(slide_layout)
        
        # Título do slide
        title_shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.8))
        title_frame = title_shape.text_frame
        title_p = title_frame.paragraphs[0]
        title_p.text = "RECEITA vs LUCRO LÍQUIDO POR CENÁRIO"
        title_p.font.size = Pt(20)
        title_p.font.bold = True
        title_p.alignment = PP_ALIGN.CENTER
        
        # Criar gráfico
        try:
            fig, ax = plt.subplots(figsize=(12, 7))
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
        
            ax.set_xlabel('Anos', fontsize=12)
            ax.set_ylabel('Valores (R$)', fontsize=12)
            ax.set_xticks([p + width for p in x])
            ax.set_xticklabels(anos)
            ax.legend()
            ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'R$ {x/1e6:.1f}M'))
            plt.tight_layout()
            
            # Salvar gráfico
            img_buffer = BytesIO()
            plt.savefig(img_buffer, format='png', dpi=300, bbox_inches='tight')
            img_buffer.seek(0)
            plt.close()
            
            # Adicionar imagem ao slide
            slide.shapes.add_picture(img_buffer, Inches(1), Inches(1.3), Inches(8), Inches(5))
            
        except Exception as chart_error:
            st.warning(f"Aviso: Gráfico não pôde ser criado no PPT ({str(chart_error)})")
            text_shape = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(4))
            text_frame = text_shape.text_frame
            text_p = text_frame.paragraphs[0]
            text_p.text = "Gráfico de Receita vs Lucro Líquido\n(Dados disponíveis nas tabelas dos slides anteriores)"
            text_p.font.size = Pt(16)
            text_p.alignment = PP_ALIGN.CENTER
        
        # === SLIDE 12: AGRADECIMENTO ===
        slide_layout = prs.slide_layouts[5]  # Blank slide
        slide = prs.slides.add_slide(slide_layout)
        
        # Título de agradecimento
        thanks_shape = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(2))
        thanks_frame = thanks_shape.text_frame
        thanks_frame.word_wrap = True
        
        thanks_p1 = thanks_frame.paragraphs[0]
        thanks_p1.text = "OBRIGADO!"
        thanks_p1.font.size = Pt(36)
        thanks_p1.font.bold = True
        thanks_p1.alignment = PP_ALIGN.CENTER
        
        thanks_p2 = thanks_frame.add_paragraph()
        thanks_p2.text = "Relatório gerado pelo Sistema Gestor de Plantio"
        thanks_p2.font.size = Pt(16)
        thanks_p2.alignment = PP_ALIGN.CENTER
        
        thanks_p3 = thanks_frame.add_paragraph()
        thanks_p3.text = f"Data: {datetime.now().strftime('%d/%m/%Y às %H:%M')}"
        thanks_p3.font.size = Pt(12)
        thanks_p3.alignment = PP_ALIGN.CENTER
        
        # Salvar apresentação
        ppt_buffer = BytesIO()
        prs.save(ppt_buffer)
        ppt_buffer.seek(0)
        
        return ppt_buffer
        
    except Exception as e:
        st.error(f"""
        ❌ **Erro ao gerar PowerPoint:** {str(e)}
        
        **Possíveis soluções:**
        1. Verifique se instalou: `pip install python-pptx matplotlib`
        2. Reinicie o Streamlit após a instalação
        3. Se usando ambiente virtual, ative-o antes de instalar
        4. Tente: `pip install --upgrade python-pptx`
        """)
        return None
