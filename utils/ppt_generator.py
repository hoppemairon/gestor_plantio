import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime

def criar_relatorio_ppt(all_indicators, all_dre_data, df_culturas_for_excel, nomes_cenarios, anos, all_indicators_cultura_cenarios=None):
    """
    Gera um relatório em PowerPoint copiando EXATAMENTE tudo da aba Indicadores.
    """
    try:
        # Importações necessárias
        from pptx import Presentation
        from pptx.util import Inches, Pt
        from pptx.enum.text import PP_ALIGN
        from pptx.dml.color import RGBColor
        
        # Criar nova apresentação
        prs = Presentation()
        st.write("🔧 Iniciando geração PowerPoint baseado em 5_indicadores.py...")
        
        # Função para criar slide com DataFrame (igual ao Excel)
        def criar_slide_com_tabela(title_text, df):
            slide_layout = prs.slide_layouts[1]
            slide = prs.slides.add_slide(slide_layout)
            title = slide.shapes.title
            title.text = title_text
            
            # Remover placeholder de conteúdo
            if len(slide.placeholders) > 1:
                content_placeholder = slide.placeholders[1]
                sp = content_placeholder.element
                sp.getparent().remove(sp)
            
            # Criar tabela
            rows = len(df) + 1
            cols = len(df.columns)
            
            left = Inches(0.2)
            top = Inches(1.2)
            width = Inches(9.6)
            height = Inches(6)
            
            table = slide.shapes.add_table(rows, cols, left, top, width, height).table
            
            # Cabeçalho
            for j, col_name in enumerate(df.columns):
                cell = table.cell(0, j)
                cell.text = str(col_name)
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(68, 114, 196)
                cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
                cell.text_frame.paragraphs[0].font.bold = True
                cell.text_frame.paragraphs[0].font.size = Pt(9)
            
            # Dados
            for i, (idx, row) in enumerate(df.iterrows()):
                for j, value in enumerate(row):
                    cell = table.cell(i + 1, j)
                    
                    if pd.isna(value):
                        cell.text = ""
                    elif isinstance(value, (int, float)):
                        if "%" in str(df.columns[j]) or "Margem" in str(df.columns[j]):
                            cell.text = f"{value:.2f}%"
                        elif "R$" in str(df.columns[j]) or abs(value) >= 1000:
                            cell.text = f"R$ {value:,.0f}"
                        else:
                            cell.text = f"{value:.2f}"
                    else:
                        cell.text = str(value)
                    
                    cell.text_frame.paragraphs[0].font.size = Pt(8)
            
            return slide
        
        # SLIDE 1: ABERTURA
        slide_layout = prs.slide_layouts[0]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        subtitle = slide.placeholders[1]
        
        title.text = "Relatório Completo de Indicadores Financeiros"
        subtitle.text = f"Análise Consolidada e por Cultura | {anos[0]} - {anos[-1]}\nBaseado em 5_indicadores.py | {datetime.now().strftime('%d/%m/%Y')}"

        # PEGAR DADOS EXATOS DO SESSION_STATE (como em 5_indicadores.py)
        
        # 1. INDICADORES CONSOLIDADOS POR CENÁRIO
        st.write("📊 Criando slides de indicadores consolidados...")
        
        for cenario in nomes_cenarios:
            emoji = "📊" if cenario == "Projetado" else "📉" if cenario == "Pessimista" else "📈"
            
            # DRE Consolidado por cenário
            if cenario in all_dre_data:
                dre_df = pd.DataFrame(all_dre_data[cenario])
                dre_df.index = anos
                criar_slide_com_tabela(f"{emoji} DRE Consolidado - {cenario}", dre_df)
            
            # Indicadores Consolidados por cenário
            if cenario in all_indicators:
                # Pegar indicadores exatos como em 5_indicadores.py
                indicators_data = {}
                for key, values in all_indicators[cenario].items():
                    if isinstance(values, list) and len(values) == len(anos):
                        indicators_data[key] = values
                
                if indicators_data:
                    indicators_df = pd.DataFrame(indicators_data)
                    indicators_df.index = anos
                    criar_slide_com_tabela(f"{emoji} Indicadores Consolidados - {cenario}", indicators_df)
                
                # Parecer consolidado (texto igual ao 5_indicadores.py)
                slide_layout = prs.slide_layouts[1]
                slide = prs.slides.add_slide(slide_layout)
                title = slide.shapes.title
                content = slide.placeholders[1]
                
                title.text = f"{emoji} Parecer Consolidado - {cenario}"
                
                # COPIAR LÓGICA EXATA DO PARECER de 5_indicadores.py
                margem_media = np.mean(all_indicators[cenario]["Margem Líquida (%)"])
                retorno_medio = np.mean(all_indicators[cenario]["Retorno por Real Gasto"])
                liquidez_media = np.mean(all_indicators[cenario]["Liquidez Operacional"])
                roa_medio = np.mean(all_indicators[cenario]["ROA (%)"])
                
                parecer_items = []
                
                # Análise de Margem
                if margem_media < 10:
                    parecer_items.append("🔴 **MARGEM LÍQUIDA BAIXA**")
                    parecer_items.append(f"   Média: {margem_media:.1f}%")
                    parecer_items.append("   • Rentabilidade abaixo do ideal")
                    parecer_items.append("   • Necessário revisar custos")
                elif margem_media < 20:
                    parecer_items.append("⚠️ **MARGEM LÍQUIDA MODERADA**")
                    parecer_items.append(f"   Média: {margem_media:.1f}%")
                    parecer_items.append("   • Rentabilidade aceitável")
                else:
                    parecer_items.append("✅ **MARGEM LÍQUIDA EXCELENTE**")
                    parecer_items.append(f"   Média: {margem_media:.1f}%")
                    parecer_items.append("   • Excelente rentabilidade")
                
                parecer_items.append("")
                
                # Análise de Retorno
                if retorno_medio < 0.15:
                    parecer_items.append("🔴 **BAIXO RETORNO POR REAL GASTO**")
                    parecer_items.append(f"   Cada R$ 1,00 gasto retorna R$ {retorno_medio:.2f}")
                elif retorno_medio < 0.25:
                    parecer_items.append("⚠️ **RETORNO MODERADO**")
                    parecer_items.append(f"   Cada R$ 1,00 gasto retorna R$ {retorno_medio:.2f}")
                else:
                    parecer_items.append("✅ **EXCELENTE RETORNO**")
                    parecer_items.append(f"   Cada R$ 1,00 gasto retorna R$ {retorno_medio:.2f}")
                
                parecer_items.append("")
                
                # Análise de Liquidez
                if liquidez_media < 1.2:
                    parecer_items.append("🔴 **LIQUIDEZ CRÍTICA**")
                    parecer_items.append(f"   Índice: {liquidez_media:.2f}")
                    parecer_items.append("   • Risco de fluxo de caixa")
                elif liquidez_media < 1.8:
                    parecer_items.append("⚠️ **LIQUIDEZ MODERADA**")
                    parecer_items.append(f"   Índice: {liquidez_media:.2f}")
                else:
                    parecer_items.append("✅ **LIQUIDEZ CONFORTÁVEL**")
                    parecer_items.append(f"   Índice: {liquidez_media:.2f}")
                
                parecer_items.append("")
                
                # ROA
                parecer_items.append(f"📈 **ROA MÉDIO: {roa_medio:.1f}%**")
                
                # CAGR
                cagr_receita = all_indicators[cenario].get("CAGR Receita (%)", 0)
                cagr_lucro = all_indicators[cenario].get("CAGR Lucro Líquido (%)", 0)
                
                parecer_items.append(f"📊 **CRESCIMENTO:**")
                parecer_items.append(f"   • CAGR Receita: {cagr_receita:.1f}%/ano")
                parecer_items.append(f"   • CAGR Lucro: {cagr_lucro:.1f}%/ano")
                
                if cagr_lucro < 0:
                    parecer_items.append("   🔴 Tendência de queda")
                elif cagr_lucro < 5:
                    parecer_items.append("   ⚠️ Crescimento baixo")
                else:
                    parecer_items.append("   ✅ Crescimento saudável")
                
                content.text = "\n".join(parecer_items)

        # 2. INDICADORES POR CULTURA (copiando exato do session_state)
        st.write("🌱 Criando slides por cultura...")
        
        # Buscar dados por cultura do session_state
        dre_por_cultura_cenarios = st.session_state.get('dre_por_cultura_cenarios', {})
        indicadores_por_cultura_cenarios = st.session_state.get('indicadores_por_cultura_cenarios', {})
        
        # Para cada cenário
        for cenario in nomes_cenarios:
            emoji = "📊" if cenario == "Projetado" else "📉" if cenario == "Pessimista" else "📈"
            st.write(f"   Processando {cenario}...")
            
            # Para cada cultura
            if cenario in dre_por_cultura_cenarios:
                for cultura in dre_por_cultura_cenarios[cenario].keys():
                    st.write(f"     Cultura: {cultura}")
                    
                    # DRE por cultura
                    dre_cultura = dre_por_cultura_cenarios[cenario][cultura]
                    if dre_cultura:
                        dre_cultura_df = pd.DataFrame(dre_cultura)
                        dre_cultura_df.index = anos
                        criar_slide_com_tabela(f"{emoji} DRE {cultura} - {cenario}", dre_cultura_df)
                    
                    # Indicadores por cultura
                    if cenario in indicadores_por_cultura_cenarios and cultura in indicadores_por_cultura_cenarios[cenario]:
                        indicadores_cultura = indicadores_por_cultura_cenarios[cenario][cultura]
                        
                        # Filtrar apenas indicadores numéricos com listas
                        indicators_cultura_data = {}
                        for key, values in indicadores_cultura.items():
                            if isinstance(values, list) and len(values) == len(anos):
                                indicators_cultura_data[key] = values
                        
                        if indicators_cultura_data:
                            indicators_cultura_df = pd.DataFrame(indicators_cultura_data)
                            indicators_cultura_df.index = anos
                            criar_slide_com_tabela(f"{emoji} Indicadores {cultura} - {cenario}", indicators_cultura_df)
                        
                        # Parecer por cultura (igual ao 5_indicadores.py)
                        slide_layout = prs.slide_layouts[1]
                        slide = prs.slides.add_slide(slide_layout)
                        title = slide.shapes.title
                        content = slide.placeholders[1]
                        
                        title.text = f"{emoji} Parecer {cultura} - {cenario}"
                        
                        # PARECER DETALHADO POR CULTURA (copiando lógica do 5_indicadores.py)
                        try:
                            margem_cultura = np.mean(indicadores_cultura.get("Margem Líquida (%)", [0]))
                            retorno_cultura = np.mean(indicadores_cultura.get("Retorno por Real Gasto", [0]))
                            liquidez_cultura = np.mean(indicadores_cultura.get("Liquidez Operacional", [0]))
                            roa_cultura = np.mean(indicadores_cultura.get("ROA (%)", [0]))
                            
                            # Dados operacionais
                            plantios_cultura = [p for p in st.session_state.get('plantios', {}).values() if p.get('cultura') == cultura]
                            hectares_cultura = sum(p.get('hectares', 0) for p in plantios_cultura)
                            
                            receitas_por_cultura = st.session_state.get('receitas_por_cultura_cenarios', {}).get(cenario, {})
                            receita_total_cultura = sum(receitas_por_cultura.get(cultura, {}).get(str(ano), 0) for ano in anos) if cultura in receitas_por_cultura else 0
                            
                            lucro_total_cultura = receita_total_cultura * (margem_cultura / 100) if margem_cultura > 0 else 0
                            
                            parecer_cultura_items = []
                            parecer_cultura_items.append(f"ANÁLISE DETALHADA - {cultura.upper()}")
                            parecer_cultura_items.append("="*50)
                            parecer_cultura_items.append("")
                            
                            # Dados operacionais
                            parecer_cultura_items.append("📊 DADOS OPERACIONAIS:")
                            parecer_cultura_items.append(f"   • Área cultivada: {hectares_cultura:.1f} hectares")
                            parecer_cultura_items.append(f"   • Receita total ({len(anos)} anos): R$ {receita_total_cultura:,.0f}")
                            parecer_cultura_items.append(f"   • Lucro estimado: R$ {lucro_total_cultura:,.0f}")
                            
                            if hectares_cultura > 0:
                                parecer_cultura_items.append(f"   • Receita/hectare/ano: R$ {receita_total_cultura/(hectares_cultura*len(anos)):,.0f}")
                                parecer_cultura_items.append(f"   • Lucro/hectare/ano: R$ {lucro_total_cultura/(hectares_cultura*len(anos)):,.0f}")
                            
                            parecer_cultura_items.append("")
                            
                            # Análise de performance
                            parecer_cultura_items.append("💰 ANÁLISE DE RENTABILIDADE:")
                            if margem_cultura < 10:
                                parecer_cultura_items.append(f"🔴 Margem líquida BAIXA: {margem_cultura:.1f}%")
                                parecer_cultura_items.append("   • Ação urgente necessária")
                                parecer_cultura_items.append("   • Revisar custos e técnicas")
                            elif margem_cultura < 20:
                                parecer_cultura_items.append(f"⚠️ Margem líquida MODERADA: {margem_cultura:.1f}%")
                                parecer_cultura_items.append("   • Performance aceitável")
                                parecer_cultura_items.append("   • Buscar melhorias")
                            else:
                                parecer_cultura_items.append(f"✅ Margem líquida EXCELENTE: {margem_cultura:.1f}%")
                                parecer_cultura_items.append("   • Cultura muito rentável")
                                parecer_cultura_items.append("   • Considerar expansão")
                            
                            parecer_cultura_items.append("")
                            
                            # Eficiência
                            parecer_cultura_items.append("⚡ EFICIÊNCIA:")
                            parecer_cultura_items.append(f"   • Retorno/Real gasto: R$ {retorno_cultura:.2f}")
                            parecer_cultura_items.append(f"   • Liquidez operacional: {liquidez_cultura:.2f}")
                            parecer_cultura_items.append(f"   • ROA: {roa_cultura:.1f}%")
                            
                            # CAGR se disponível
                            cagr_receita_cultura = indicadores_cultura.get("CAGR Receita (%)", 0)
                            cagr_lucro_cultura = indicadores_cultura.get("CAGR Lucro Líquido (%)", 0)
                            
                            if cagr_lucro_cultura != 0:
                                parecer_cultura_items.append("")
                                parecer_cultura_items.append("📈 CRESCIMENTO:")
                                parecer_cultura_items.append(f"   • CAGR Receita: {cagr_receita_cultura:.1f}%/ano")
                                parecer_cultura_items.append(f"   • CAGR Lucro: {cagr_lucro_cultura:.1f}%/ano")
                                
                                if cagr_lucro_cultura < 0:
                                    parecer_cultura_items.append("   🔴 Tendência declinante")
                                elif cagr_lucro_cultura < 5:
                                    parecer_cultura_items.append("   ⚠️ Crescimento lento")
                                else:
                                    parecer_cultura_items.append("   ✅ Crescimento saudável")
                            
                            # Recomendação final
                            parecer_cultura_items.append("")
                            parecer_cultura_items.append("🎯 RECOMENDAÇÃO:")
                            
                            if margem_cultura >= 15 and retorno_cultura >= 0.2 and cagr_lucro_cultura >= 0:
                                parecer_cultura_items.append("   ✅ CULTURA ALTAMENTE RECOMENDADA")
                                parecer_cultura_items.append("   • Excelente performance geral")
                                parecer_cultura_items.append("   • Considerar aumento da área")
                            elif margem_cultura >= 10 and retorno_cultura >= 0.15:
                                parecer_cultura_items.append("   ✅ CULTURA RECOMENDADA")
                                parecer_cultura_items.append("   • Performance adequada")
                                parecer_cultura_items.append("   • Manter área atual")
                            else:
                                parecer_cultura_items.append("   ⚠️ CULTURA REQUER ATENÇÃO")
                                parecer_cultura_items.append("   • Baixa performance")
                                parecer_cultura_items.append("   • Revisar estratégia")
                            
                            content.text = "\n".join(parecer_cultura_items)
                            
                        except Exception as e:
                            content.text = f"Erro ao gerar parecer para {cultura}: {str(e)}"

        # 3. TABELA COMPARATIVA FINAL (como no final de 5_indicadores.py)
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        title.text = "📊 Resumo Comparativo Final"
        
        # Criar tabela comparativa de todos os cenários
        comparativo_data = []
        for cenario in nomes_cenarios:
            if cenario in all_indicators:
                receita_total = sum(all_dre_data[cenario]["Receita"])
                lucro_total = sum(all_dre_data[cenario]["Lucro Líquido"])
                margem_media = np.mean(all_indicators[cenario]["Margem Líquida (%)"])
                retorno_medio = np.mean(all_indicators[cenario]["Retorno por Real Gasto"])
                
                comparativo_data.append({
                    "Cenário": cenario,
                    f"Receita Total ({len(anos)} anos)": receita_total,
                    f"Lucro Total ({len(anos)} anos)": lucro_total,
                    "Margem Líquida Média (%)": margem_media,
                    "Retorno por Real Gasto": retorno_medio
                })
        
        if comparativo_data:
            comparativo_df = pd.DataFrame(comparativo_data)
            criar_slide_com_tabela("📊 Comparativo Final - Todos os Cenários", comparativo_df)

        # SLIDE FINAL: AGRADECIMENTO
        slide_layout = prs.slide_layouts[0]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        subtitle = slide.placeholders[1]
        
        title.text = "Relatório Completo Finalizado"
        subtitle.text = f"Baseado em 5_indicadores.py\nTodas as tabelas e pareceres incluídos\n{datetime.now().strftime('%d/%m/%Y às %H:%M')}"

        # Salvar apresentação
        output_ppt = BytesIO()
        prs.save(output_ppt)
        output_ppt.seek(0)
        
        st.write(f"✅ PowerPoint gerado com {len(prs.slides)} slides")
        st.write("📋 Conteúdo copiado exatamente do 5_indicadores.py")
        
        return output_ppt
        
    except Exception as e:
        st.error(f"❌ Erro ao gerar PowerPoint: {str(e)}")
        import traceback
        st.error(traceback.format_exc())
        return None