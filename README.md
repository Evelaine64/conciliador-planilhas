# 🔎 Conciliador de Planilhas

Script em Python que cruza três planilhas de um contrato (**orçamento**, **lista de documentos** e **boletim de medição**) e gera um relatório em Excel com tudo o que não bate entre elas.

## O problema

Em contratos de engenharia, o mesmo documento aparece em várias bases diferentes. Conferir à mão se o número de folhas, o valor medido e a existência de cada documento batem entre as bases é lento e fácil de errar, principalmente quando são centenas de documentos.

Este projeto é uma versão de exemplo, com **dados fictícios**, de um tipo de conciliação que faço no trabalho.

## O que o script encontra

| Tipo | Exemplo |
|---|---|
| **Sem correspondência** | Documento que está no orçamento, mas não na lista (ou o contrário) |
| **Divergência de valor** | Valor medido no boletim diferente do valor orçado |
| **Divergência de folhas** | Número de folhas diferente entre orçamento e lista |
| **Atenção** | Documento cancelado na lista, mas que foi medido |

## Resultado

O relatório `relatorio_conciliacao.xlsx` tem três abas:

- **Resumo** – totais de cada tipo de problema
- **Problemas** – lista de documentos com problema, colorida por tipo
- **Base cruzada** – as três bases juntas, para conferência

Exemplo com os dados fictícios:

```
              Indicador  Quantidade
Documentos no orçamento         300
    Documentos na lista         298
     Documentos medidos         210
    Sem correspondência          10
 Divergências de folhas          12
  Divergências de valor           8
      Pontos de atenção           6
```

## Como usar

```bash
pip install -r requirements.txt

# 1. Gerar as planilhas de exemplo (pasta dados/)
python gerar_dados_exemplo.py

# 2. Rodar a conciliação
python conciliar.py
```

Para usar com suas próprias planilhas:

```bash
python conciliar.py --orcamento orcamento.xlsx --lista lista.xlsx --medicao medicao.xlsx --saida resultado.xlsx
```

As planilhas precisam ter a coluna `codigo`. Veja os arquivos em `dados/` para o formato completo.

## Tecnologias

Python · Pandas · openpyxl

---
Feito por [Evelaine Pires](https://www.linkedin.com/in/evelaine01)
