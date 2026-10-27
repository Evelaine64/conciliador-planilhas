"""
Gera as três planilhas de exemplo usadas pelo conciliador.

Os dados são FICTÍCIOS: simulam um contrato de engenharia com documentos
que aparecem no orçamento, na lista de documentos e no boletim de medição.
Alguns erros são colocados de propósito para o conciliador encontrar.
"""

import random
from pathlib import Path

import pandas as pd

random.seed(42)
PASTA = Path(__file__).parent / "dados"
PASTA.mkdir(exist_ok=True)

DISCIPLINAS = ["Processo", "Mecânica", "Elétrica", "Instrumentação", "Civil"]
TIPOS = ["Memorial", "Desenho", "Lista", "Especificação", "Relatório"]
AUTORIZACOES = [f"AS-{n:03d}" for n in range(1, 6)]


def gerar_orcamento(qtd=300):
    linhas = []
    for i in range(1, qtd + 1):
        linhas.append({
            "codigo": f"DOC-{i:04d}",
            "titulo": f"{random.choice(TIPOS)} de {random.choice(DISCIPLINAS)} {i}",
            "disciplina": random.choice(DISCIPLINAS),
            "autorizacao": random.choice(AUTORIZACOES),
            "valor_orcado": round(random.uniform(800, 9000), 2),
            "folhas": random.randint(1, 40),
        })
    return pd.DataFrame(linhas)


def gerar_lista_documentos(orcamento):
    lista = orcamento[["codigo", "titulo", "folhas"]].copy()
    lista["revisao"] = [random.choice(["0", "A", "B", "1"]) for _ in range(len(lista))]
    lista["status"] = random.choices(["Emitido", "Em revisão", "Cancelado"], weights=[80, 17, 3], k=len(lista))

    # Erro 1: alguns documentos estão com número de folhas diferente do orçamento
    for idx in random.sample(range(len(lista)), 12):
        lista.loc[idx, "folhas"] += random.randint(1, 5)

    # Erro 2: alguns documentos sumiram da lista
    lista = lista.drop(random.sample(range(len(lista)), 6))

    # Erro 3: documentos que existem na lista mas não estavam no orçamento
    extras = pd.DataFrame([
        {"codigo": f"DOC-9{n:03d}", "titulo": "Documento não orçado", "folhas": 3,
         "revisao": "0", "status": "Emitido"} for n in range(1, 5)
    ])
    return pd.concat([lista, extras], ignore_index=True)


def gerar_medicao(orcamento):
    medidos = orcamento.sample(frac=0.7, random_state=1)[["codigo", "valor_orcado"]].copy()
    medidos = medidos.rename(columns={"valor_orcado": "valor_medido"})
    medidos["boletim"] = [f"BM-{random.randint(1, 4):02d}" for _ in range(len(medidos))]

    # Erro 4: alguns valores medidos diferentes do orçado
    for idx in random.sample(list(medidos.index), 8):
        medidos.loc[idx, "valor_medido"] = round(medidos.loc[idx, "valor_medido"] * random.uniform(1.05, 1.3), 2)

    return medidos.reset_index(drop=True)


if __name__ == "__main__":
    orcamento = gerar_orcamento()
    lista = gerar_lista_documentos(orcamento)
    medicao = gerar_medicao(orcamento)

    orcamento.to_excel(PASTA / "orcamento.xlsx", index=False)
    lista.to_excel(PASTA / "lista_documentos.xlsx", index=False)
    medicao.to_excel(PASTA / "boletim_medicao.xlsx", index=False)

    print(f"Planilhas de exemplo criadas em: {PASTA}")
