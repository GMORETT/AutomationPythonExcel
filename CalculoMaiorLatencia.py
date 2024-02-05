from collections import OrderedDict
import os
import pandas as pd
import xlwings as xw
from datetime import timedelta
import json

def extrair_valor_celula_excel(caminho_do_arquivo, nome_planilha, linha, coluna):
    try:
        planilha = pd.read_excel(caminho_do_arquivo, sheet_name=nome_planilha, engine='openpyxl')

        valor_celula = planilha.at[linha, coluna]
        return valor_celula
    except KeyError:
        print(f"Célula ({linha}, {coluna}) não encontrada na planilha '{nome_planilha}'.")
        print(f"Cabeçalhos da planilha: {list(planilha.columns)}")
        print(f"Índices da planilha: {list(planilha.index)}")
        return None
    except FileNotFoundError:
        print(f"Arquivo '{caminho_do_arquivo}' não encontrado.")
        return None


caminho_do_arquivo = r'\\Usuarios\gabrielmorett\Documents\Planinlha Relatorio.xlsx'
nome_planilha_desejada = 'Relatório Diário'
linha_desejada = 0
coluna_desejada = 'Data'

valor_celula = extrair_valor_celula_excel(caminho_do_arquivo, nome_planilha_desejada, linha_desejada, coluna_desejada)

if valor_celula is not None:
    print(
        f"Valor da célula ({linha_desejada}, {coluna_desejada}) na planilha '{nome_planilha_desejada}': {valor_celula}")

    data_celula = pd.to_datetime(valor_celula)

    minha_variavel = data_celula

def carregar_dados(caminho_json):
    with open(caminho_json, 'r') as file:
        data = json.load(file)

    df = pd.DataFrame(data)
    df['Valores'] = pd.to_numeric(df['Valores'], errors='coerce')
    df['data'] = pd.to_datetime(df['data'])
    df = df.sort_values(by=['data', 'Circuito'])
    return df

def obter_data():
    data_desejada = data_celula

    try:
        return pd.to_datetime(data_desejada)
    except ValueError:
        raise ValueError('Formato de data inválido. Utilize o formato YYYY-MM-DD.')


def encontrar_maior_intervalo(df_filtrado, valor_limite):
    max_duracao_sequencia = timedelta()
    duracao_sequencia_atual = timedelta()
    horarios_sequencia = []
    circuito_sequencia = None
    valores_sequencia = []

    horarios_max_duracao = []
    valores_max_duracao = []
    nome_circuito_max_duracao = None

    for indice, linha in df_filtrado.iterrows():
        if linha['Valores'] >= valor_limite:
            if duracao_sequencia_atual == timedelta():
                horario_inicial = linha['data']
                circuito_sequencia = linha['Circuito']
            horarios_sequencia.append(linha['data'])
            valores_sequencia.append(linha['Valores'])

            if len(horarios_sequencia) >= 2:
                intervalo_tempo = horarios_sequencia[-1] - horarios_sequencia[-2]
                duracao_sequencia_atual += intervalo_tempo
            else:
                duracao_sequencia_atual = timedelta()  # Reinicia a contagem se houver apenas um horário

        else:
            if len(horarios_sequencia) >= 1:
                horario_final = linha['data']
                duracao_sequencia = horario_final - horario_inicial
                if duracao_sequencia >= max_duracao_sequencia:
                    max_duracao_sequencia = duracao_sequencia
                    horarios_max_duracao = horarios_sequencia.copy()
                    valores_max_duracao = valores_sequencia.copy()
                    nome_circuito_max_duracao = circuito_sequencia

            duracao_sequencia_atual = timedelta()
            horarios_sequencia = []
            valores_sequencia = []
            circuito_sequencia = None

    # Verifica se a sequência atual é o maior intervalo
    if len(horarios_sequencia) >= 1:
        horario_final = df_filtrado.iloc[-1]['data']
        duracao_sequencia = horario_final - horario_inicial
        if duracao_sequencia >= max_duracao_sequencia:
            max_duracao_sequencia = duracao_sequencia
            horarios_max_duracao = horarios_sequencia.copy()
            valores_max_duracao = valores_sequencia.copy()
            nome_circuito_max_duracao = circuito_sequencia

    return max_duracao_sequencia, horarios_max_duracao, valores_max_duracao, nome_circuito_max_duracao




def escrever_resultado_excel(sheet, linha, coluna, valor):
    sheet.range(f'{coluna}{linha}').value = valor

def imprimir_resultados(sheet, linha, nome_circuito_max_duracao, valor_limite, max_duracao_sequencia):
    if max_duracao_sequencia > timedelta(minutes=0):
        print(f'Maior intervalo de tempo em que os objetos ficaram com valores >= {valor_limite} para o circuito {nome_circuito_max_duracao}: {max_duracao_sequencia}')
        escrever_resultado_excel(sheet, linha, 'J', str(max_duracao_sequencia).split()[2])
    else:
        print(f'Nenhum intervalo encontrado para o circuito {nome_circuito_max_duracao}.')
        escrever_resultado_excel(sheet, linha, 'J', '00:00:00')


def main():
    caminho_arquivo_excel = r'\\Usuarios\gabrielmorett\Documents\Planinlha Relatorio.xlsx'

    if not os.path.exists(caminho_arquivo_excel):
        raise FileNotFoundError(f'O arquivo Excel especificado não existe: {caminho_arquivo_excel}')

    wb = xw.Book(caminho_arquivo_excel)
    sheet = wb.sheets.active

    # Extrair o número da célula da planilha
    numero_analise = int(sheet.range('D2').value)

    # Verificar se o número extraído é válido
    if numero_analise <= 0:
        raise ValueError('O número extraído da célula da planilha não é válido.')

    # Dicionário que mapeia nomes de arquivos para valores limite específicos com ordem
    valores_limites = OrderedDict({
        'arquivojson1.json': 30,
        'arquivojson2.json': 40,
    
        # Adicione mais linhas conforme necessário
    })

    linha = 5  # Ajuste conforme necessário
    for arquivo_json, valor_limite in valores_limites.items():
        print(f'\nProcessando arquivo: {arquivo_json}')

        diretorio_json = r'\\Usuarios$\gabrielmorett\Documents\pastajsons'
        caminho_completo = os.path.join(diretorio_json, arquivo_json)

        try:
            df = carregar_dados(caminho_completo)
            data_desejada = obter_data()

            # Definir o período de análise com base no número extraído da célula
            datas_analise = [data_desejada + timedelta(days=i) for i in range(numero_analise)]

            # DataFrame para armazenar os dados de todos os dias
            df_periodo = pd.DataFrame()

            for data_desejada in datas_analise:
                df_filtrado = df[df['data'].dt.date == data_desejada.date()]

                if not df_filtrado.empty:
                    df_periodo = pd.concat([df_periodo, df_filtrado])

            # Calcular o maior intervalo para todo o período
            max_duracao_sequencia, _, _, nome_circuito_max_duracao = encontrar_maior_intervalo(df_periodo, valor_limite)
            imprimir_resultados(sheet, linha, nome_circuito_max_duracao, valor_limite, max_duracao_sequencia)
            linha += 1

        except pd.errors.EmptyDataError:
            print(f'O arquivo {arquivo_json} está vazio ou não contém dados válidos.')

        except Exception as e:
            print(f'Ocorreu um erro ao processar o arquivo {arquivo_json}: {e}')

    wb.save()
    wb.close()


if __name__ == "__main__":
    main()