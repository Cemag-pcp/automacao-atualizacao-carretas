import gspread
from google.oauth2 import service_account
import pandas as pd
from gspread_dataframe import set_with_dataframe

def configuracoes_iniciais():
    #Configuração inicial
    # service_account_info = ["GOOGLE_SERVICE_ACCOUNT"]
    scope = ['https://www.googleapis.com/auth/spreadsheets',
            "https://www.googleapis.com/auth/drive"]
    

    credentials = service_account.Credentials.from_service_account_file(r'C:\Users\Engine\robo_atualizacao_carretas\automacao-atualizacao-carretas\credentials.json', scopes=scope)

    client = gspread.authorize(credentials)
    
    return client

def conectar_planilha_apontamento():
    """
    Retorna as celulas e codigos atualizados
    """

    #Configurações iniciais
    client = configuracoes_iniciais()

    #ID planilha apontamento
    sheet_id_apontamento = '1x26yfwoF7peeb59yJuJuxCQNlqjCjh65NYS1RIrC0Zc'

    #Abrindo a planilha pelo ID
    
    sh_apontamento_montagem = client.open_by_key(sheet_id_apontamento)

    wks_apontamento = sh_apontamento_montagem.worksheet('RQ PCP 002-000 (APONTAMENTO MONTAGEM)')
    list_montagem = wks_apontamento.get_all_values()

    #Tratando o df para pegar as colunas necessárias
    itens_montagem = pd.DataFrame(list_montagem)
    itens_montagem.columns = itens_montagem.iloc[4]
    itens_montagem = itens_montagem.drop(index=[0,1,2,3,4])

    itens_montagem = itens_montagem[['Código','Célula']]
    itens_montagem = itens_montagem[
        itens_montagem['Célula'].notna() & (itens_montagem['Célula'] != '') &
        itens_montagem['Código'].notna() & (itens_montagem['Código'] != '')
    ]
    df_celuas_codigo_atualizados = itens_montagem.groupby(['Código'],as_index=False).last().reset_index()

    return df_celuas_codigo_atualizados

def armazenar_base_atualizada_planilha(df):

    # df = pd.read_excel(r"C:\Users\TIDEV\Documents\planilhas_auxiliares\carretas_final.xlsx")

    client = configuracoes_iniciais()

    sheet_id = "1A67y-gk0P5qW_jDaxL4B9I-wP9wDM6mjJ91BMrzGWHw"

    sh_base_conjuntos = client.open_by_key(sheet_id)

    # Selecionar a aba
    worksheet = sh_base_conjuntos.worksheet("BASE ATUALIZADA")

    # Limpa a planilha Inteira
    worksheet.clear()
    # Substituir tudo com o DataFrame
    set_with_dataframe(worksheet, df)  # df é seu DataFrame pandas

    print("BASE ATUALIZADA COM SUCESSO!")

def armazenar_carretas_pe(df):

    # df = pd.read_excel(r"C:\Users\TIDEV\Documents\planilhas_auxiliares\carretas_final.xlsx")

    df = df.fillna("")

    client = configuracoes_iniciais()

    sheet_id = "1A67y-gk0P5qW_jDaxL4B9I-wP9wDM6mjJ91BMrzGWHw"

    sh_base_conjuntos = client.open_by_key(sheet_id)

    # Selecionar a aba
    worksheet = sh_base_conjuntos.worksheet("BASE ATUALIZADA")

    # Limpa a planilha Inteira
    # worksheet.clear()
    # Substituir tudo com o DataFrame
    # set_with_dataframe(worksheet, df)  # df é seu DataFrame pandas

    # 4. Encontrar a última linha preenchida
    ultima_linha = len(worksheet.get_all_values())

    # 5. Inserir DataFrame a partir da próxima linha
    worksheet.insert_rows(df.values.tolist(), row=ultima_linha+1)

    print("BASE ATUALIZADA COM SUCESSO!")

def base_felipe():
    """
        Compara o retorno da API com a base_felipe para ver oq ta faltando ser explodido
    """
    client = configuracoes_iniciais()

    #ID planilha
    sheet_id = '1n2J6n_VxOsVxY5ikjJeDGva7oHTUJOlzadFfUbJnaSE'

    sh = client.open_by_key(sheet_id)
    #worksheet_name
    wks = sh.worksheet('BASE')
    list1 = wks.get_all_values()

    itens = pd.DataFrame(list1)
    itens.columns = itens.iloc[0]
    itens = itens.drop(index=0) 

    itens['carreta'] = itens['carreta'].str.strip()
    itens = itens['carreta'].drop_duplicates()
    # print(itens)

    carretas = itens.to_list()

    print(len(carretas))

    return carretas

# armazenar_base_atualizada_planilha()
# armazenar_carretas_pe()