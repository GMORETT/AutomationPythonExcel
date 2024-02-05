import xlwings as xw
import os

def obter_nomes_tabelas(planilha, nome_aba):
    # Abre a planilha
    wb = xw.Book(planilha)
    sheet = wb.sheets[nome_aba]

    # Obtém todos os nomes de tabela na aba especificada
    nomes_tabelas = []
    for tabela in sheet.api.ListObjects:
        nomes_tabelas.append(tabela.Name)

    return nomes_tabelas

if __name__ == "__main__":
    # Caminho completo para a planilha Excel
    caminho_planilha = r'\\Usuarios$\gmorett\Documents\planilha.xlsx'
    aba_leitura = 'aba1'  # Substitua pelo nome da aba onde está a tabela

    nomes_tabelas = obter_nomes_tabelas(caminho_planilha, aba_leitura)
    print("Nomes das tabelas na aba {}: {}".format(aba_leitura, nomes_tabelas))
