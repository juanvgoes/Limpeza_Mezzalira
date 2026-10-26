"""
Desafio CiTi - Etapa 1: do dado bagunçado ao dado útil
Metalúrgica Mezzalira

A ideia é seguir a lógica Bronze -> Prata -> Ouro:
  Bronze = cópia do original, sem mexer em nada
  Prata  = os dados já limpos
  Log    = tudo que eu corrigi (e por quê)
  Ouro   = faturamento por cidade e por mês

pra rodar: %run limpeza_mezzalira.py pedidos_metalurgica_mezzalira.xlsx pedidos_mezzalira_tratado.xlsx

"""
import sys
import pandas as pd

entrada = sys.argv[1] if len(sys.argv) > 1 else "pedidos_metalurgica_mezzalira.xlsx"
saida = sys.argv[2] if len(sys.argv) > 2 else "pedidos_mezzalira_tratado.xlsx"

#meses abreviados em português, porque o Python (e o Sheets) só entende "feb" e não "fev"
dicionario_meses = {"jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
                    "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12}

#como cada cidade deve ficar escrita no final. "rec" é a sigla que apareceu na planilha.
dicionario_cidades = {
    "recife": "Recife",
    "rec": "Recife",
    "caruaru": "Caruaru",
    "olinda": "Olinda",
    "petrolina": "Petrolina",
    "jaboatão dos guararapes": "Jaboatão dos Guararapes",
}

#aqui vou guardando cada correção pra montar a aba Log no final
log = []


def anotar(encontrei, fiz, linhas, justificativa=""):
    #junta os numeros das linhas num texto só (ex: "2, 6, 7")
    if len(linhas) == 0:
        texto_linhas = "-"
    else:
        texto_linhas = ", ".join(str(n) for n in linhas)
    log.append({
        "O que encontrei": encontrei,
        "O que fiz": fiz,
        "Linhas afetadas": texto_linhas,
        "Justificativa (se houve dúvida)": justificativa,
    })


#bronze: le tudo como veio, sem deixar o pandas "adivinhar" tipos
bronze = pd.read_excel(entrada, dtype=object)

#guarda o número da linha original (linha 1 = cabeçalho) pra conseguir citar no Log dps
prata = bronze.copy()
prata.insert(0, "Linha original", prata.index + 2)

colunas = ["Pedido", "Data", "Cliente", "Cidade", "Produto", "Valor (R$)"]


#linhas vazias
#linha vazia = todas as colunas vazias
linhas_vazias = prata[colunas].isna().all(axis=1)
anotar("Linhas completamente vazias no meio dos pedidos",
       "Removi essas linhas",
       prata.loc[linhas_vazias, "Linha original"].tolist())
prata = prata[~linhas_vazias]


#pedidos duplicados
#so é duplicado se TODAS as colunas forem iguais, só o numero do pedido não basta
duplicadas = prata.duplicated(subset=colunas, keep="first")

for i, linha in prata[duplicadas].iterrows():
    #procura de qual linha ela é copia
    primeira = prata[prata["Pedido"] == linha["Pedido"]].iloc[0]["Linha original"]
    anotar("Pedido " + str(linha["Pedido"]) + " aparece duas vezes (igual à linha " + str(primeira) + ")",
           "Apaguei a cópia e fiquei com a primeira ocorrência",
           [linha["Linha original"]],
           "As duas linhas são idênticas em todas as colunas, então tratei como "
           "o mesmo pedido lançado duas vezes. Se fosse venda nova, teria outro número.")

prata = prata[~duplicadas]


#datas em formatos diferentes
datas_novas = []
linhas_formato_iso = []   #formato 2026-01-01
linhas_formato_pt = []    #formato 01-fev-2026

for numero_linha, data in zip(prata["Linha original"], prata["Data"]):
    if hasattr(data, "year"):
        #já é data de verdade (o Excel guardou como data), não precisa mexer
        datas_novas.append(pd.Timestamp(data))
    else:
        texto = str(data).strip().lower()
        if texto[4] == "-":
            #formato 2026-01-01
            datas_novas.append(pd.Timestamp(texto))
            linhas_formato_iso.append(numero_linha)
        else:
            #formato 01-fev-2026, tem que trocar o mes em portugues pelo numero
            dia, mes, ano = texto.split("-")
            datas_novas.append(pd.Timestamp(int(ano), dicionario_meses[mes], int(dia)))
            linhas_formato_pt.append(numero_linha)

prata["Data"] = datas_novas

anotar("Datas gravadas como texto no padrão AAAA-MM-DD",
       "Converti para data de verdade",
       linhas_formato_iso)
anotar("Datas gravadas como texto no padrão DD-mês-AAAA (ex.: 01-fev-2026)",
       "Converti para data de verdade, lendo o mês em português",
       linhas_formato_pt,
       "O mês vem escrito abreviado (ex: fev), então não tem confusão entre dia e mês.")


#mesma cidade escrita de jeitos diferentes
cidades_novas = []
erradas = {}   #guarda as linhas de cada jeito errado de escrever

for numero_linha, cidade in zip(prata["Linha original"], prata["Cidade"]):
    certa = dicionario_cidades[cidade.strip().lower()]   #se aparecer cidade nova dá erro e eu vejo
    cidades_novas.append(certa)
    if cidade != certa:
        if cidade not in erradas:
            erradas[cidade] = []
        erradas[cidade].append(numero_linha)

prata["Cidade"] = cidades_novas

for escrita, linhas in erradas.items():
    justificativa = ""
    if escrita.strip().lower() == "rec":
        justificativa = ("Assumi que 'REC' é sigla de Recife, já que não existe outra cidade "
                         "com essa sigla na base. Não tenho como confirmar com a empresa.")
    anotar("Cidade escrita como '" + escrita + "'",
           "Troquei por '" + dicionario_cidades[escrita.strip().lower()] + "'",
           linhas,
           justificativa)


#valores escritos como texto
valores_novos = []
linhas_valor_texto = []

for numero_linha, valor in zip(prata["Linha original"], prata["Valor (R$)"]):
    if isinstance(valor, str):
        #"R$ 2.240,00" -> tira o R$ e o ponto, troca a virgula por ponto
        valor = valor.replace("R$", "").replace(".", "").replace(",", ".")
        linhas_valor_texto.append(numero_linha)
    valores_novos.append(float(valor))

prata["Valor (R$)"] = valores_novos

anotar("Valores digitados como texto (ex.: 'R$ 2.240,00')",
       "Tirei o 'R$' e o ponto de milhar, troquei vírgula por ponto e converti para número",
       linhas_valor_texto,
       "Segui o padrão brasileiro: ponto separa milhar e vírgula separa centavos.")


#ultimos ajustes da prata
prata["Pedido"] = prata["Pedido"].astype(int)
prata["Mês"] = prata["Data"].dt.strftime("%Y-%m")   #vai ajudar pro ouro
prata = prata.sort_values("Pedido").reset_index(drop=True)

#se algo der errado o código para ao invés de seguir
assert prata["Pedido"].is_unique, "ainda tem pedido repetido"
assert prata[colunas].notna().all().all(), "ainda tem campo vazio"
assert len(prata["Cidade"].unique()) == 5, "número de cidades diferente do esperado"

anotar("Conferência final",
       "A base foi de " + str(len(bronze)) + " linhas de dados para " + str(len(prata)),
       [],
       "Saíram " + str(linhas_vazias.sum()) + " linhas vazias e " + str(duplicadas.sum()) + " duplicadas.")


#ouro: faturamento por cidade e mês
#margins=True cria a linha e a coluna de total
ouro = prata.pivot_table(index="Cidade", columns="Mês", values="Valor (R$)",
                         aggfunc="sum", fill_value=0, margins=True, margins_name="Total")


#salva no excel
with pd.ExcelWriter(saida, engine="openpyxl") as arquivo:
    bronze.to_excel(arquivo, sheet_name="Bronze", index=False)
    prata.to_excel(arquivo, sheet_name="Prata", index=False)
    pd.DataFrame(log).to_excel(arquivo, sheet_name="Log", index=False)
    ouro.to_excel(arquivo, sheet_name="Ouro")

    #deixa a coluna de data no formato dd/mm/aaaa (a coluna C da prata)
    for celula in arquivo.sheets["Prata"]["C"][1:]:
        celula.number_format = "DD/MM/YYYY"


print(pd.DataFrame(log).to_string(index=False))
print()
print(ouro)
