"""
Conciliador de planilhas

Cruza três bases de um contrato (orçamento, lista de documentos e boletim de
medição) e gera um relatório em Excel com:
  - um resumo com os totais;
  - os documentos que aparecem em uma base e faltam em outra;
  - as divergências de folhas e de valores.

Uso:
    python conciliar.py
    python conciliar.py --orcamento meu_orcamento.xlsx --saida resultado.xlsx
"""

import argparse
from pathlib import Path

import pandas as pd
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

PASTA_DADOS = Path(__file__).parent / "dados"
TOLERANCIA_VALOR = 0.01  # diferença em reais aceita como "igual"


def ler_bases(caminho_orcamento, caminho_lista, caminho_medicao):
    orcamento = pd.read_excel(caminho_orcamento)
    lista = pd.read_excel(caminho_lista)
    medicao = pd.read_excel(caminho_medicao)

    # Padroniza o código para evitar erro por espaço ou letra minúscula
    for base in (orcamento, lista, medicao):
        base["codigo"] = base["codigo"].astype(str).str.strip().str.upper()

    return orcamento, lista, medicao


def cruzar_bases(orcamento, lista, medicao):
    """Junta as três bases pelo código do documento, mantendo todos os registros."""
    tabela = orcamento.merge(
        lista[["codigo", "folhas", "revisao", "status"]],
        on="codigo", how="outer", suffixes=("_orcamento", "_lista"), indicator="na_lista",
    )
    tabela["na_lista"] = tabela["na_lista"].map({"both": True, "left_only": False, "right_only": True})
    tabela["no_orcamento"] = tabela["titulo"].notna()

    tabela = tabela.merge(medicao[["codigo", "valor_medido", "boletim"]], on="codigo", how="left")
    tabela["medido"] = tabela["valor_medido"].notna()
    return tabela


def encontrar_problemas(tabela):
    problemas = []

    for _, linha in tabela.iterrows():
        codigo = linha["codigo"]

        if not linha["no_orcamento"]:
            problemas.append((codigo, "Sem correspondência", "Está na lista de documentos, mas não no orçamento"))
            continue

        if not linha["na_lista"]:
            problemas.append((codigo, "Sem correspondência", "Está no orçamento, mas não na lista de documentos"))

        if linha["na_lista"] and linha["folhas_orcamento"] != linha["folhas_lista"]:
            problemas.append((codigo, "Divergência de folhas",
                              f"Orçamento: {int(linha['folhas_orcamento'])} | Lista: {int(linha['folhas_lista'])}"))

        if linha["medido"] and abs(linha["valor_medido"] - linha["valor_orcado"]) > TOLERANCIA_VALOR:
            diferenca = linha["valor_medido"] - linha["valor_orcado"]
            problemas.append((codigo, "Divergência de valor",
                              f"Orçado: R$ {linha['valor_orcado']:,.2f} | Medido: R$ {linha['valor_medido']:,.2f} "
                              f"| Diferença: R$ {diferenca:,.2f}"))

        if linha["medido"] and linha.get("status") == "Cancelado":
            problemas.append((codigo, "Atenção", "Documento cancelado na lista, mas medido no boletim"))

    ordem = ["Sem correspondência", "Divergência de valor", "Divergência de folhas", "Atenção"]
    resultado = pd.DataFrame(problemas, columns=["codigo", "tipo", "detalhe"])
    resultado["tipo"] = pd.Categorical(resultado["tipo"], categories=ordem, ordered=True)
    return resultado.sort_values(["tipo", "codigo"]).reset_index(drop=True)


def montar_resumo(tabela, problemas):
    contagem = problemas["tipo"].value_counts()
    return pd.DataFrame([
        ("Documentos no orçamento", int(tabela["no_orcamento"].sum())),
        ("Documentos na lista", int(tabela["na_lista"].sum())),
        ("Documentos medidos", int(tabela["medido"].sum())),
        ("Sem correspondência", int(contagem.get("Sem correspondência", 0))),
        ("Divergências de folhas", int(contagem.get("Divergência de folhas", 0))),
        ("Divergências de valor", int(contagem.get("Divergência de valor", 0))),
        ("Pontos de atenção", int(contagem.get("Atenção", 0))),
    ], columns=["Indicador", "Quantidade"])


def salvar_relatorio(caminho, resumo, problemas, tabela):
    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        resumo.to_excel(writer, sheet_name="Resumo", index=False)
        problemas.to_excel(writer, sheet_name="Problemas", index=False)
        tabela.to_excel(writer, sheet_name="Base cruzada", index=False)

        cores = {
            "Sem correspondência": "F8D7DA",
            "Divergência de folhas": "FFF3CD",
            "Divergência de valor": "FFE5B4",
            "Atenção": "D1ECF1",
        }
        aba = writer.sheets["Problemas"]
        for linha in aba.iter_rows(min_row=2):
            cor = cores.get(linha[1].value)
            if cor:
                for celula in linha:
                    celula.fill = PatternFill("solid", start_color=cor)

        for aba in writer.sheets.values():
            for celula in aba[1]:
                celula.font = Font(bold=True)
            for i, coluna in enumerate(aba.columns, start=1):
                largura = max(len(str(c.value or "")) for c in coluna)
                aba.column_dimensions[get_column_letter(i)].width = min(largura + 2, 70)


def main():
    parser = argparse.ArgumentParser(description="Concilia orçamento, lista de documentos e boletim de medição.")
    parser.add_argument("--orcamento", default=PASTA_DADOS / "orcamento.xlsx")
    parser.add_argument("--lista", default=PASTA_DADOS / "lista_documentos.xlsx")
    parser.add_argument("--medicao", default=PASTA_DADOS / "boletim_medicao.xlsx")
    parser.add_argument("--saida", default="relatorio_conciliacao.xlsx")
    args = parser.parse_args()

    orcamento, lista, medicao = ler_bases(args.orcamento, args.lista, args.medicao)
    tabela = cruzar_bases(orcamento, lista, medicao)
    problemas = encontrar_problemas(tabela)
    resumo = montar_resumo(tabela, problemas)
    salvar_relatorio(args.saida, resumo, problemas, tabela)

    print(resumo.to_string(index=False))
    print(f"\nRelatório salvo em: {args.saida}")


if __name__ == "__main__":
    main()
