import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time
import pandas as pd
from bs4 import BeautifulSoup
from tratamento_plan import tratar_df_final
from conectar_planilha import armazenar_base_atualizada_planilha, armazenar_carretas_pe, puxar_pecas_aba
from utils import fechar_todas_abas
from collections import defaultdict
import random
import re
from verificar_chrome import *



def puxando_carretas():
    # lista = ['77997944', 'CBH7-2E FO SS RS/T M24', '66826075', '84494370', '77182604', '78542451', '80415918', '79382990', '75712112', '032513', '82169949', '95673995', '90338803', 'CBHM5000 GR CS RD P650(R) M17', '81557346', 'CBH5 FO SC RD MM P750(I) M21', '83311673', '71361057', '75238296', 'FTC4300R-1E SS RD BB M24', '77645951', 'FTC6500-1E SS T BB P750(I) M22', '91487379', '84914709', 'CBHM3500 SC RD BE M17', 'FTC6500-1E SS T R20 BB M22', '77978081', '58561424', '75445188', '90866402', '93252310', 'FTC4300R-1E CS RD P750(I) M24', '88991624', '76262608', '94246796', 'CBHM3500 SS RS P750(R) M17', '030383', 'FA4 FB SS RS A90 M22', '96725973', '402795', 'FTC6500 CS RS/RS M22', 'FTD5000 SS T M22', 'CBH5 FO SS RD MM P750(I) M21', '75451478', 'CBHM5000 GR SC RD P650(R) M17', 'FA5 SS T M23', '83004134', 'CBH5 UG SC T M21', '75464267', '78243844', 'FA2A SS T M23', '75438833', 'FTC4300R SS RS/RS BB M24', 'CBH6R FO SC T P750(R) M21', 'CBHM5000 GR SC RD MM P750(R) M17', 'FTC4300R-1E SS RD BB P750(I) M24', 'CBH6 FO SS RD MM P750(I) M22', '78185870', '81473731', 'CBHM5000 GR SS RD BE M17', 'CBH5 UG SS T P700(R) M21', 'CHASSI FTC4300 SS M22', 'CBHM5000 GR CS RD P750(I) M17', 'FA4 SS RS A90 P750(I) M23', 'CBHM5000 GR SC RD BE M17', '93253429', '81589725', 'CBHM6000 UG SS RD M17', '030378', '94047858', 'CBH7 FO SS T MM P700(R) M21', 'FTC6500 SS RS/RS P750(I) M22', '81358950', '94246581', '84753591', 'F4 SS RS/RS A90 CB M23', 'CBHM5000 CA SC RD ABA MM M17', '75509234', '88561046', '93306979', 'CBH6R FO SC RD TTUG P750(I) M21', '77183194', 'CHASSI FTC4300-1E SS T BB M22', '91487397', '622309M25', 'CBH5 FO SC T MM M21', 'CBH7 FO SS T MM M21', '84798396', 'FTC4300R-1E SS T M24', '95930536', '76037344', 'FTC6500 SS RS/RS BB P750(I) M22+', '79397053', 'F6 SS RS/RS A60 MT M22', '93923931', '036860', 'CBH6 FO SS RDC P750(I) M22', 'CBHM6000-2E SS RS/RD P750(I) M17', '86195534', 'CBHM4500 SS RD P650(R) M17']

    # print(len(lista))
    
    link = "https://cemag.innovaro.com.br/api/publica/v1/tabelas/listarProdutos"

    response = requests.get(link)

    if response.status_code == 200:
        dados = response.json()
    else:
        print('erro')

    # carretas_totais = []
    carretas_base_completo = []
    codigo_carretas = []
    carretas_com_cores = []
    codigo_cores = ['VJ', 'VM', 'AN', 'LC', 'LJ', 'AM','AV','CO']
    #adicionar tudo em uma lista menos os fora de linha
    #aplicar regex para achar codigo de cor e removê-lo


    for produto in dados['produtos']:
        # carretas_totais.append(str(produto.get('codigo')).strip())
        # if (produto.get('cor') == 'Laranja' or produto.get('cor') == '') and "fora de linha" not in produto.get('nome').lower():
        if "fora de linha" not in produto.get('nome').lower() and "fora de linha" not in produto.get('codigo').lower() and "+" not in produto.get('codigo'):
            cor_codigo_numerico = str(produto.get('codigo'))[-2:]

            try:
                cor_t1 = str(produto.get('codigo')).split(' ')[-1] # Cor no final do codigo
                cor_t2 = str(produto.get('codigo')).split(' ')[-2] # Cor na penultima posicao do codigo
                cor_t3 = str(produto.get('codigo')).split(' ')[-3] # Cor na antepenultima posicao do codigo
            except IndexError as e:
                cor_t2 = 'Não tem cor nessa posicao'
                cor_t3 = 'Não tem cor nessa posicao'

            if any(cor in codigo_cores for cor in [cor_t1, cor_t2, cor_t3, cor_codigo_numerico]):
                carretas_com_cores.append(str(produto.get('codigo')).strip())
            else:
                codigo_carretas.append(str(produto.get('codigo')).strip())
            carreta = {
                'chave':str(produto.get('chave')).strip(),
                'codigo': str(produto.get('codigo')).strip()
            }
            #adicionando em uma lista so de chaves
            
            #adicionando em uma lista de dicionario que contem chave e codigo
            carretas_base_completo.append(carreta)
            

    
    # print(len(carretas_base))
    print("codigos sem cor: ", len(codigo_carretas))
    print("codigos com cor: ",len(carretas_com_cores))
    return codigo_carretas,carretas_com_cores, carretas_base_completo

def rodar_automacao(carretas):
    # Inicializa o navegador
    lista_df = []
    while len(carretas) > 0:
        try:
            driver = webdriver.Chrome()
        except:
            chrome_driver_path = verificar_chrome_driver()
            driver = webdriver.Chrome(chrome_driver_path)

        # Acessa o sistema
        driver.get("http://192.168.3.141/sistema")
        # driver.get("https://hcemag.innovaro.com.br/sistema")
        # driver.maximize_window() # Maximiza a tela
        time.sleep(2)

        # Faz login
        driver.find_element(By.ID, "username").send_keys("FILIPE")
        driver.find_element(By.ID, "password").send_keys("3470Filipe13")
        # driver.find_element(By.ID, "username").send_keys("ti.prod")
        # driver.find_element(By.ID, "password").send_keys("cem@1610")
        driver.find_element(By.ID, "submit-login").click()
        time.sleep(5)

        # fechar_todas_abas(driver)
        
        busca_carretas = carretas[:100]

        resultado = ";".join(busca_carretas)
        resultado += ";"
        
        # Navega nos menus
        driver.find_element(By.ID, "bt_1892603865").click()
        time.sleep(2)
        driver.find_element(By.XPATH, "//span[text()='Produção']").click()
        time.sleep(1)
        driver.find_element(By.XPATH, "//span[text()='Consultas gerenciais']").click()
        time.sleep(1)
        driver.find_element(By.XPATH, "//span[text()='Processos e composição de recursos com custos (BOM) CEMAG']").click()
        time.sleep(3)

        # input('tete')
        # Fechar menus
        driver.find_element(By.ID, "bt_1892603865").click()
        time.sleep(2)
        driver.find_element(By.XPATH, "//span[text()='Consultas gerenciais']").click()
        time.sleep(1)
        driver.find_element(By.XPATH, "//span[text()='Produção']").click()
        time.sleep(1)
        driver.find_element(By.ID, "bt_1892603865").click()
        time.sleep(2)


        # Acessa o iframe onde está o campo de pesquisa
        driver.switch_to.frame("inlineFrameTabId1")

        # Cola o valor no campo de recursos
        input_recurso = driver.find_element(By.NAME, 'recursos')
        input_recurso.clear()
        time.sleep(1)
        # input_recurso.send_keys("CBHM5000 GR SS RD P750(I) M17;CBHM5000 CA SC RD MM M17")
        input_recurso.send_keys(resultado)
        print("✅ Valor colado no campo de recursos.")
        time.sleep(3)

        # Sai do iframe para clicar no botão "Executar"
        driver.switch_to.default_content()

        # 🔁 Loop até a mensagem "Realizando a consulta..." aparecer
        print("⏳ Esperando a mensagem de execução aparecer...")

        cont_vezes_sem_aparecer_msg = 0
        cont_msg_waitbox = 0
        while True:
            try:
                # Primeiro clique no botão "Executar"
                botao_executar = driver.find_element(By.XPATH, "//span[contains(., 'Executar')]")
                ActionChains(driver).move_to_element(botao_executar).click().perform()
                print("🔄 Primeiro clique no botão 'Executar'")
                time.sleep(10)

                while True:
                    # Verifica se a mensagem apareceu
                    mensagem = driver.find_elements(By.ID, "waitMessageBox")
                    if mensagem:
                        # print("✅ Mensagem detectada:", mensagem[0].text)
                        if cont_msg_waitbox == 0:
                            print(f"⚠️ Mensagem ainda está aparecendo: {mensagem[0].text}")
                        cont_msg_waitbox+=1
                    else:
                        print("✅ Mensagem desapareceu")
                        break
                    
                    time.sleep(2)

                # Acessa o iframe para clicar no botão "Todos"
                driver.switch_to.frame("inlineFrameTabId1")
                botao_todos = driver.find_element(By.XPATH, "//span[contains(., 'Todos')]")
                ActionChains(driver).move_to_element(botao_todos).click().perform()
                print("🟢 Botão 'Todos' clicado")
                time.sleep(1)
                driver.switch_to.default_content()

                # Segundo clique no botão "Executar"
                print('2 clique')
                botao_executar = driver.find_element(By.XPATH, "//span[contains(., 'Executar')]")
                ActionChains(driver).move_to_element(botao_executar).click().perform()
                print("✅ Segundo clique no botão 'Executar' enviado")

            except Exception as e:
                print(f"❌ Erro ao clicar: {e}")

            time.sleep(2)

            # Verifica se a mensagem apareceu
            if cont_vezes_sem_aparecer_msg < 5:
                cont_vezes_sem_aparecer_msg += 1
                mensagem = driver.find_elements(By.ID, "content_statusMessageBox")
                if mensagem:
                    print("✅ Mensagem detectada:", mensagem[0].text)
                    break

                print("⚠️ Mensagem ainda não apareceu. Tentando novamente...\n")
                time.sleep(2)
            else:
                print('Mensagem não apareceu, continuando o processo...')

        cont = 0
        while True:
            # Verifica se a mensagem apareceu
            mensagem = driver.find_elements(By.ID, "content_statusMessageBox")
            if mensagem:
                # print("✅ Mensagem detectada:", mensagem[0].text)
                if cont == 0:
                    print(f"⚠️ Mensagem ainda está aparecendo: {mensagem[0].text}")
                cont+=1
            else:
                print("✅ Mensagem desapareceu")
                break
            
            time.sleep(2)

        cont_msg_progresso = 0
        while True:
            # Verifica se a mensagem apareceu
            mensagem = driver.find_elements(By.ID, "progressMessageBox")
            if mensagem:
                # print("✅ Mensagem detectada:", mensagem[0].text)
                if cont_msg_progresso == 0:
                    print(f"⚠️ Mensagem progresso: {mensagem[0].text}")
                cont_msg_progresso+=1
            else:
                print("✅ Mensagem desapareceu")
                break
            
            time.sleep(2)

        # # 🟢 Aguarda a tabela ser carregada após a mensagem
        # time.sleep(60)

        # Volta para o iframe para acessar a tabela
        driver.switch_to.frame("inlineFrameTabId1")

        # Extrai as linhas da tabela localizada em //*[@id="lid-0"]/tbody
        # linhas = driver.find_elements(By.XPATH, '//*[@id="lid-0"]/tbody/tr')

        # print('Percorrendo linhas...')
        # dados = []
        # for linha in linhas:
        #     colunas = linha.find_elements(By.TAG_NAME, 'td')
        #     if colunas:
        #         dados.append([coluna.text.strip() for coluna in colunas])

        

        # 1. Extrai o HTML da tabela de uma vez
        html_tabela = driver.find_element(By.XPATH, '//*[@id="lid-0"]/tbody').get_attribute("outerHTML")

        # 2. Processa com BeautifulSoup (muito mais rápido)
        soup = BeautifulSoup(html_tabela, 'html.parser')
        linhas = soup.find_all('tr')

        dados = []
        for linha in linhas:
            colunas = linha.find_all('td')
            if colunas:
                dados.append([coluna.get_text(strip=True) for coluna in colunas])

        print('Linhas extraídas:', len(dados))


        # Exibe os dados coletados
        print("📋 Dados extraídos")
        # for linha in dados:
        #     print(linha)

        #Retira os elementos já buscado
        carretas = carretas[100:]
        # Salva os dados em Excel
        df = pd.DataFrame(dados)
        lista_df.append(df)
        
        print("✅ Dados adicionados em uma lista de dataframes")

        time.sleep(2)

        driver.quit()


        
    print('Finalização das explosões')
    # Concatena todos em um único DataFrame
    df_final = pd.concat(lista_df, ignore_index=True)
    df_final.columns = ['Recurso','Qtd.','TOTAL','Und.','Observação','Depósito Origem','Depósito Destino','Custo']
    
    # df_final.to_excel("resultado_innovaro.xlsx", index=False)

    print(f'Tamanho da lista de df: {len(lista_df)}')
    # print("✅ Dados salvos no arquivo 'resultado_innovaro.xlsx'")

    return df_final

def main():

    inicio = time.time()

    #Pegando as carretas atualizadas da api
    codigo_carretas, carretas_com_cores, carretas_base_completo = puxando_carretas()

    # print(codigo_carretas)

    # Valores que já existem nos codigos sem cor
    resultado = {}
    for item in codigo_carretas:
        correspondentes = [c for c in carretas_com_cores if c.startswith(item)]
        if correspondentes:
            resultado[item] = correspondentes

    # Removendo os valores que já existem sem cor na lista de codigo_carretas
    
    for items in resultado.values():
        for i in items:
            try:
                carretas_com_cores.remove(i)
            except ValueError as e:
                print(f'Codigo {i} já removido')

    print(len(carretas_com_cores))
    # print(carretas_com_cores)

    codigo_cores = ['VJ', 'VM', 'AN', 'LC', 'LJ', 'AM','AV','CO']

    grupos = defaultdict(list)

    # Criando grupos com codigos com prefixo repetido
    # CBHM6000 CA SC T MM P700(R) M21: ['CBHM6000 CA SC T MM P700(R) M21 AV']
    # CBHM3500 SS RS BE MM M17: ['CBHM3500 SS RS BE MM M17 VM']
    # 623064: ['623064AV', '623064AN', '623064CO', '623064LC', '623064VJ', '623064VM']
    # 623063: ['623063AN', '623063AV', '623063CO', '623063LC', '623063VJ', '623063VM']
    # 623016: ['623016AV', '623016AN', '623016VJ']
    # 623019: ['623019AN', '623019LC', '623019VJ', '623019VM']
    # CBHM6000 CA SC RD P650(R) M21: ['CBHM6000 CA SC RD P650(R) M21 VJ']
    # CBHM6000 CA SC RD P750(R) M21: ['CBHM6000 CA SC RD P750(R) M21 AV']
    # CBH7 FO SC RD MM P700(R) M21: ['CBH7 FO SC RD MM P700(R) M21 VM', 'CBH7 FO SC RD MM P700(R) M21 AV']

    for item in carretas_com_cores:
        cor_encontrada = None
        for cor in codigo_cores:
            if f" {cor} " in f" {item} " or item.endswith(cor):
                cor_encontrada = cor
                # remover a cor do item para criar a base
                base = item.replace(cor, "").strip()
                grupos[base].append(item)
                break

        if not cor_encontrada:
            grupos["OUTROS"].append(item)

    # Sorteando valores por cada grupo(chave)
    sorteados = []
    for chave, itens in grupos.items():
        item_sorteado = random.choice(itens)
        sorteados.append(item_sorteado)

    # print(sorteados)
    lista_completa_produtos = codigo_carretas + sorteados

    lista_set = set(lista_completa_produtos)

    print(len(lista_set))

    #Elemento com PE
    lista_elementos_diferentes_pe = []

    for elemento in lista_set:
        if 'PE' in elemento:
            lista_elementos_diferentes_pe.append(elemento)

    
    lista_set_sem_pe = [x for x in lista_set if x not in lista_elementos_diferentes_pe]


    lista_pesquisa_bom_chaves = []
    lista_pesquisa_bom_chaves_pe = []
    # lista de codigos e chaves para adicionar ao dataframe
    lista_completa_chaves_codigos = []


    for item in carretas_base_completo:
        if item['codigo'] in lista_set_sem_pe:
            lista_pesquisa_bom_chaves.append(item['chave'])
        if item['codigo'] in lista_elementos_diferentes_pe:
            lista_pesquisa_bom_chaves_pe.append(item['chave'])
        if item['codigo'] in lista_set_sem_pe or item['codigo'] in lista_elementos_diferentes_pe:
            lista_completa_chaves_codigos.append({
                    'codigo': re.sub(r"(LC|LH|VM|VJ|AN|AV)$", "", item['codigo']).strip(), 
                    'chave': item['chave']
                })

    print(lista_completa_chaves_codigos)

    print(len(lista_pesquisa_bom_chaves))
    print(len(list(set(lista_pesquisa_bom_chaves))))

    lista_sem_duplicadas_chaves_bom = list(set(lista_pesquisa_bom_chaves))

    # return 'oi'
    #verificar tudo oq em carretas_com_cores que tenha o mesmo sufixo de codigo_carretas

    # busca_carretas = lista_completa_produtos[:100]
    #Explodindo as carretas de 100 em 100 e concatenando-as
    # lista_sem_duplicadas_chaves_bom = ['79051673']
    df_final = rodar_automacao(lista_sem_duplicadas_chaves_bom)

    fim = time.time()

    duracao_segundos = fim - inicio
    duracao_minutos = duracao_segundos / 60

    # Salvar o que veio do innovaro
    df_final.to_excel(r"C:\Users\Engine\planilhas_auxiliares\resultado_innovaro.xlsx", index=False)
    # df_final.to_excel(r"C:\Users\TIDEV\planilha_atualizacao_carretas\resultado_innovaro.xlsx", index=False)
    #Tratamento da planilha
    df_tratado = tratar_df_final(df_final)

    dataframe_chaves_codigos = pd.DataFrame(lista_completa_chaves_codigos)

    df_combinado_normal = pd.merge(df_tratado, dataframe_chaves_codigos, left_on='CARRETA', right_on='codigo', how='left')

    armazenar_base_atualizada_planilha(df_combinado_normal)

    # Rodando carretas com PE
    df_final_pe = rodar_automacao(lista_pesquisa_bom_chaves_pe)

    print(f"Tempo de execução da explosão das carretas: {duracao_minutos:.2f} minutos")

    # Salvar o que veio do innovaro
    df_final_pe.to_excel(r"C:\Users\Engine\planilhas_auxiliares\resultado_innovaro_pe.xlsx", index=False)
    #Tratamento da planilha
    df_tratado_pe = tratar_df_final(df_final_pe)

    df_combinado_pe = pd.merge(df_tratado_pe, dataframe_chaves_codigos, left_on='CARRETA', right_on='codigo', how='left')

    armazenar_carretas_pe(df_combinado_pe)

    # Puxando as peças da aba BASE PEÇAS e colando na aba BASE ATUALIZADA
    puxar_pecas_aba()

    fim_total = time.time()

    duracao_segundos = fim_total - inicio
    duracao_minutos = duracao_segundos / 60

    print(f"Tempo de execução da aplicação completa: {duracao_minutos:.2f} minutos")




if __name__ == "__main__":
    main()
