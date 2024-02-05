from datetime import datetime, timedelta
import os
import json
import xlwings as xw


def extrair_numero_data(caminho_planilha, nome_aba_numero, celula_numero, nome_aba_data, celula_data):
    workbook = xw.Book(caminho_planilha)
    numero_aba = workbook.sheets[nome_aba_numero]
    numero = numero_aba.range(celula_numero).value

    data_aba = workbook.sheets[nome_aba_data]
    data = data_aba.range(celula_data).value

    workbook.close()
    return numero, data.date() if data else None



def ler_arquivo_json(caminho_arquivo):
    with open(caminho_arquivo, 'r') as arquivo:
        dados = json.load(arquivo)

    for objeto in dados:
        if isinstance(objeto.get('data'), str):
            try:
                objeto['data'] = datetime.strptime(objeto['data'], '%Y-%m-%d %H:%M:%S')
            except ValueError:
                print(f"Aviso: Não foi possível converter a data '{objeto['data']}' para objeto de data.")

    return dados


def imprimir_objetos_por_data(dados, data, nome_arquivo):
    objetos_do_dia = [objeto for objeto in dados if objeto.get('data').date() == data]
    if objetos_do_dia:
        print(f"Objetos registrados em {data} no arquivo {nome_arquivo}:")
        for objeto in objetos_do_dia:
            print(objeto)
    else:
        print(f"Nenhum objeto encontrado para a data {data} no arquivo {nome_arquivo}.")


def adicionar_dados_a_planilha(caminho_planilha, dados_por_aba, configuracoes_abas):
    workbook = xw.Book(caminho_planilha)

    dados_em_lote_por_aba = {aba_nome: [] for aba_nome in dados_por_aba.keys()}

    for aba_nome, dados_aba in dados_por_aba.items():
        valor_limite = configuracoes_abas[aba_nome]['valor_limite']
        for objeto in dados_aba:
            data = objeto['data']
            valores = objeto['Valores']
            dia = data.day
            mes = data.month
            ano = data.year
            valor_objeto = float(valores)
            status = "VERDADEIRO" if valor_objeto >= valor_limite else "FALSO"
            dados_em_lote_por_aba[aba_nome].append([data, valores, dia, mes, ano, status, 0])

    for aba_nome, dados_em_lote in dados_em_lote_por_aba.items():
        aba = workbook.sheets[aba_nome]
        ultima_linha = aba.range('A1').end('down').row
        if dados_em_lote:
            aba.range(f'A{ultima_linha+1}').value = dados_em_lote

    workbook.save()
    workbook.close()



caminho_planilha = r'\\Usuarios\gabrielmorett\Documents\planilha.xlsx'
celula_numero = 'D2'
celula_data = 'C2'
nome_aba_numero = 'Relatório Diário'
nome_aba_data = 'Relatório Diário'

pasta_contendo_jsons = r'\\Usuarios$\gabrielmorett\Documents\pastajsons'
configuracoes_abas = {
    'aba1': {'arquivo': 'arquivojson1.json', 'valor_limite': 20},
    'aba2': {'arquivo': 'arquivojson2.json', 'valor_limite': 30}
        #Adicione mais elementos ao dicionario caso queira especificar mais abas e arquivos com seus respectivos valores limites
}

numero, data = extrair_numero_data(caminho_planilha, nome_aba_numero, celula_numero, nome_aba_data, celula_data)

dados_por_aba = {}

if numero == 1:
    dias_processar = 1
elif numero == 2:
    dias_processar = 3
elif numero == 3:
    dias_processar = 4
else:
    print("Número inválido.")

datas_a_processar = [data + timedelta(days=i) for i in range(dias_processar)] #Para processar mais datas de uma vez

dados_por_aba = {}

for aba_nome, config in configuracoes_abas.items():
    nome_arquivo = config['arquivo']
    caminho_arquivo_json = os.path.join(pasta_contendo_jsons, nome_arquivo)
    dados = ler_arquivo_json(caminho_arquivo_json)
    imprimir_objetos_por_data(dados, data, nome_arquivo)

    objetos_filtrados = [objeto for objeto in dados if objeto.get('data').date() in datas_a_processar]

    if len(objetos_filtrados) >= config['valor_limite']:
        dados_por_aba[aba_nome] = objetos_filtrados

adicionar_dados_a_planilha(caminho_planilha, dados_por_aba, configuracoes_abas)

