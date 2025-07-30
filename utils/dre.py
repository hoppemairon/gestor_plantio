# utils/dre.py
import numpy as np
import pandas as pd

def calcular_dre(cenario, inflacoes, anos, hectares_total, total_sacas, preco_total, receitas, receitas_extras, despesas_info, emprestimos, pess_despesas, otm_despesas, fluxo_ajustado=None):
    if fluxo_ajustado is not None:
        fluxo = fluxo_ajustado
        # Verificação de colunas obrigatórias
        required_cols = [
            "Despesas Operacionais",
            "Despesas Administrativas",
            "Despesas RH",
            "Despesas Extra Operacional",
            "Impostos Sobre Venda",
            "Impostos Sobre Resultado",
            "Dividendos"
        ]
        if not all(col in fluxo for col in required_cols):
            raise ValueError("O fluxo ajustado fornecido está incompleto.")
        # Garantir uso dos valores do fluxo ajustado
        despesas_operacionais = fluxo["Despesas Operacionais"]
        despesas_administrativas = fluxo["Despesas Administrativas"]
        despesas_rh = fluxo["Despesas RH"]
        despesas_extra_operacional = fluxo["Despesas Extra Operacional"]
        impostos_sobre_venda = fluxo["Impostos Sobre Venda"]
        impostos_sobre_resultado = fluxo["Impostos Sobre Resultado"]
        dividendos = fluxo["Dividendos"]
    dre = {}
    receita = receitas[cenario]
    dre["Receita"] = receita
    # Se fluxo_ajustado foi passado, usar impostos_sobre_venda da variável, senão calcular
    if fluxo_ajustado is not None:
        dre["Impostos Sobre Venda"] = impostos_sobre_venda
        dre["Despesas Operacionais"] = despesas_operacionais
        dre["Despesas Administrativas"] = despesas_administrativas
        dre["Despesas RH"] = despesas_rh
        dre["Despesas Extra Operacional"] = despesas_extra_operacional
        dre["Dividendos"] = dividendos
        # Receita Extra Operacional segue a lógica anterior
        dre["Receita Extra Operacional"] = receitas_extras["Extra Operacional"]
        # Impostos Sobre Resultado: se existir no fluxo, usa, senão calcula
        if "Impostos Sobre Resultado" in fluxo:
            dre["Impostos Sobre Resultado"] = impostos_sobre_resultado
        else:
            dre["Impostos Sobre Resultado"] = [
                0 for _ in range(len(anos))
            ]
    else:
        dre["Impostos Sobre Venda"] = [r * 0.0485 for r in receita]

        def linha_despesa(cat):
            total = sum(despesas_info[despesas_info["Categoria"] == cat]["Valor"]) if not despesas_info.empty else 0
            ajuste = pess_despesas if cenario == "Pessimista" else (-otm_despesas if cenario == "Otimista" else 0)
            return [total * np.prod([1 + inflacoes[j] / 100 for j in range(i + 1)]) * (1 + ajuste / 100) for i in range(5)]

        dre["Despesas Operacionais"] = linha_despesa("Operacional")
        dre["Despesas Administrativas"] = linha_despesa("Administrativa")
        dre["Despesas RH"] = linha_despesa("RH")
        dre["Dividendos"] = linha_despesa("Dividendos")

        extra_operacional = [0] * 5
        for emp in emprestimos:
            try:
                start = anos.index(emp["ano_inicial"])
                end = anos.index(emp["ano_final"])
                parcelas = emp["parcelas"]
                ajuste = (pess_despesas if cenario == "Pessimista" else (-otm_despesas if cenario == "Otimista" else 0)) / 100
                for i in range(start, min(end + 1, 5)):
                    if parcelas > 0:
                        extra_operacional[i] += emp["valor_parcela"] * (1 + ajuste)
                        parcelas -= 1
            except:
                continue
        dre["Despesas Extra Operacional"] = extra_operacional
        dre["Receita Extra Operacional"] = receitas_extras["Extra Operacional"]

    # Cálculo das margens e resultados, sempre usando as variáveis corretas
    dre["Margem de Contribuição"] = [
        receita[i] - dre["Impostos Sobre Venda"][i] - dre["Despesas Operacionais"][i] for i in range(5)
    ]
    dre["Resultado Operacional"] = [
        dre["Margem de Contribuição"][i] - dre["Despesas Administrativas"][i] - dre["Despesas RH"][i] for i in range(5)
    ]
    dre["Lucro Operacional"] = [
        dre["Resultado Operacional"][i] - dre["Despesas Extra Operacional"][i] for i in range(5)
    ]
    # Impostos Sobre Resultado: se já definido (por fluxo), mantém, senão calcula aqui
    if fluxo_ajustado is not None and "Impostos Sobre Resultado" in fluxo:
        # Já está definido acima
        pass
    else:
        dre["Impostos Sobre Resultado"] = [
            dre["Lucro Operacional"][i] * 0.15 if dre["Lucro Operacional"][i] > 0 else 0 for i in range(5)
        ]
    dre["Lucro Líquido"] = [
        dre["Lucro Operacional"][i] - dre["Impostos Sobre Resultado"][i] - dre["Dividendos"][i] + dre["Receita Extra Operacional"][i] for i in range(5)
    ]
    return dre