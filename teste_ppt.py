import pandas as pd
import streamlit as st
from utils.ppt_generator import criar_relatorio_ppt_completo

# Dados de exemplo para teste
anos = [2025, 2026, 2027]
nomes_cenarios = ["Projetado", "Pessimista", "Otimista"]

# DRE consolidado exemplo
all_dre_data = {
    "Projetado": {
        "Receita": [100000, 110000, 120000],
        "Impostos Sobre Venda": [5000, 5500, 6000],
        "Despesas Operacionais": [30000, 33000, 36000],
        "Despesas Administrativas": [15000, 16500, 18000],
        "Despesas RH": [20000, 22000, 24000],
        "Despesas Extra Operacional": [2000, 2200, 2400],
        "Dividendos": [5000, 5500, 6000],
        "Impostos Sobre Resultado": [3000, 3300, 3600],
        "Lucro Líquido": [20000, 22500, 25000]
    }
}

# Indicadores consolidados exemplo
all_indicators = {
    "Projetado": {
        "Margem Líquida (%)": [20, 20.45, 20.83],
        "Retorno por Real Gasto": [0.25, 0.27, 0.29],
        "Liquidez Operacional": [1.5, 1.6, 1.7],
        "ROA (%)": [15, 16, 17],
        "Produtividade por Hectare (R$/ha)": [5000, 5500, 6000],
        "Custo por Receita (%)": [65, 64, 63],
        "Endividamento (%)": [25, 23, 20],
        "DSCR": [2.0, 2.2, 2.5],
        "Break-Even Yield (sacas/ha)": [30, 28, 26],
        "CAGR Receita (%)": 10,
        "CAGR Lucro Líquido (%)": 12
    }
}

# DataFrame das culturas
df_culturas_for_excel = pd.DataFrame({
    "Cultura": ["Soja", "Milho"],
    "Hectares": [100, 80]
})

print("🧪 Iniciando teste da função PPT...")
print(f"Anos: {anos}")
print(f"Cenários: {nomes_cenarios}")
print(f"DRE disponível: {list(all_dre_data.keys())}")
print(f"Indicadores disponível: {list(all_indicators.keys())}")
print(f"Culturas: {df_culturas_for_excel['Cultura'].tolist()}")

# Teste básico da estrutura de dados
df_teste = pd.DataFrame(all_dre_data["Projetado"])
df_teste.index = [f"Ano {ano}" for ano in anos]
print("\n📊 Teste DataFrame original:")
print(df_teste.head())

df_transposto = df_teste.T
print("\n📊 Teste DataFrame transposto:")
print(df_transposto.head())
print(f"Índice transposto: {df_transposto.index.tolist()[:3]}...")
print(f"Colunas transpostas: {df_transposto.columns.tolist()}")

print("\n✅ Estrutura de dados parece estar correta!")
print("💡 Execute o teste no app principal para verificar a geração do PPT")
