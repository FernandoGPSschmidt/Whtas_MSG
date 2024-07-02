import streamlit as st
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.common.keys import Keys
import pandas as pd
import warnings
import datetime as dt
import urllib.parse
from selenium.common.exceptions import TimeoutException  # Corrigido: importando TimeoutException



def count(x):
    while True:
        st.write(f"Seu valor é {x}")
        time.sleep(1)
        x += 1
        
# Função para clicar em elementos padrão
def click(campo):
    time.sleep(3)
    element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, campo)))
    element.click()

#realiza login no portal 
def login():
    driver.get("https://portal.gpssa.com.br/gps/portal.aspx")
    escrever("/html/body/form/div[1]/div/div[1]/div/div/div/div/div[4]/div/div[1]/div/div/div[2]/div[1]/div/div/input","fernando.schmidt")
    escrever("/html/body/form/div[1]/div/div[1]/div/div/div/div/div[4]/div/div[1]/div/div/div[3]/div[1]/div/div/input", "GPS@032024")
    click("/html/body/form/div[1]/div/div[1]/div/div/div/div/div[4]/div/div[2]/div/div/a/span/span/span[2]")
    dois_click("/html/body/div[1]/div[2]/div[4]/div/div/a[6]/span/span/span[2]")
    dois_click("/html/body/form/div[2]/div[3]/div/div/div/div/div/div[2]/div/div[1]/table[15]/tbody/tr/td/div/span")

from datetime import datetime, timedelta #atualizar linha de codigo
from calendar import monthrange
def escrever(campo, texto):
    time.sleep(3)
    try:
        # Aguardar até que o elemento de entrada de texto esteja presente
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, campo))
        )

        # Limpar o campo de entrada, se necessário
        element.clear()

        # Enviar as teclas desejadas para o campo de entrada
        element.send_keys(texto)

    except TimeoutException:
        # Se o elemento não estiver presente, imprimir uma mensagem
        print("Timeout: Elemento de entrada não encontrado. Ignorando e continuando.")

        # Não é necessário fechar o navegador aqui, pois queremos continuar o loop

def acessar_relatorio(element):
    #tela para acessar o campo de atendimento do portal
    iframe_busc = driver.find_element(By.XPATH,element)                                       
    iframe_busc.click()
    from_trat =  driver.switch_to.frame(iframe_busc)


def dois_click(campo):
    try:
        # Aguardar um pouco para que você possa ver o resultado
        time.sleep(5)
        element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, campo)))
        actions = ActionChains(driver)
        actions.double_click(element)
        actions.perform()
    except TimeoutException:
        # Se o elemento não for encontrado dentro do tempo de espera, imprimir uma mensagem
        print("Timeout: Elemento não encontrado. Ignorando e continuando.")

def escrever_e_click(campo, texto):
    try:
        # Aguardar até que o elemento de entrada de texto esteja presente
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, campo))
        )

        # Limpar o campo de entrada, se necessário
        element.clear()
        time.sleep(3)
        for _ in range(8):
            element.send_keys(Keys.ARROW_LEFT)
            time.sleep(0.7)
        element.send_keys()
        for char in texto:
            element.send_keys(char)
            time.sleep(1)
        # Enviar as teclas desejadas para o campo de entrada
        time.sleep(3)
        element.send_keys(Keys.ENTER)


    except TimeoutException:
        # Se o elemento não estiver presente, imprimir uma mensagem
        print("Timeout: Elemento de entrada não encontrado. Ignorando e continuando.")
def escrever(campo, texto):
    time.sleep(3)
    try:
        # Aguardar até que o elemento de entrada de texto esteja presente
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, campo))
        )

        # Limpar o campo de entrada, se necessário
        element.clear()

        # Enviar as teclas desejadas para o campo de entrada
        element.send_keys(texto)

    except TimeoutException:
        # Se o elemento não estiver presente, imprimir uma mensagem
        print("Timeout: Elemento de entrada não encontrado. Ignorando e continuando.")


st.title("Envio de Mensagens Automático Whatsapp")




with st.form(key="add_form"):

    mensagem = st.text_area("Mensagem a ser enviada.")


    arquivo = st.file_uploader("Escolha um arquivo XLSX", type=["xlsx"])
    
    data_ex = {
        "COLABORADOR": [
            "12345678910 - COLABORDOR DA SILVA",
            "12345678910 - COLABORDOR DA SILVA",
           "12345678910 - COLABORDOR DA SILVA",
        ],
        "CPF": ["12345678910", "12345678910", "12345678910"],
        "STATUS": ["Enviado", "Enviado", "Enviado"],
        "NÃO ASSINADA": ["4", "4", "4"]
    }
    st.write("Siga esse padrão de tabela ⬇")
    df_exemplo = pd.DataFrame(data_ex)
    st.table(df_exemplo)


    submit_button = st.form_submit_button(label="Adicionar")
    
    if submit_button:
     
            df_upload =  pd.read_excel(arquivo)
            st.table(df_upload)
            row_index = 0  # Índice da linha desejada
            max_rows = len(df_upload)

            #recebe o diretorio da pasta que possui os pdf´s
            options = webdriver.EdgeOptions()
            options.add_argument("--headless")
            driver = webdriver.Edge()
            # Maximiza a janela do navegador
            driver.maximize_window()


            #realiza a def usuario para logar no portal com as credenciais do Fernando
            login()
            #aguarda um pouco para o portal iniciar
            time.sleep(5)
            #clica em Gestão de pessoas
            dois_click("//*[text()='Gestão de Pessoas']")
            #clica em documentos e contratos
            dois_click("//*[text()='4. Documentos e Contratos']")
            #clica no Card de colaborador
            dois_click("//*[text()='4.1 - Card - Gestão do Colaborador']")
            #acessando o relatório do portal 
            time.sleep(10)
            acessar_relatorio("/html/body/form/div[4]/div/div/div/div/div[3]/div/div[2]/div[2]/div[1]/div/iframe")
            print("ok iframe acessado")
            #df_upload =  pd.read_excel(arquivo)
            while row_index <= max_rows :
                
                try:
                    df_upload = df_upload
                    
                    cpf_valor = df_upload.loc[row_index, "CPF"]
                    cpf_str = int(cpf_valor)
                    nome = str(df_upload.loc[row_index, "COLABORADOR"])

                    # Divida o nome completo pelo espaço e selecione o primeiro elemento (que é o primeiro nome)
                    first_name = nome.split(" ", 1)[0]
                    # Imprima o primeiro nome extraído
            

                    status = df_upload.loc[row_index, "STATUS"]
                    if status != "Enviado":

                        print("--------------INICIANDO AUTOMAÇÃO--------------")
                        data_atual =  dt.datetime.today() 
                        data_limite = data_atual + dt.timedelta(days=10)
                        data_limte_format = data_limite.strftime('%d/%m/%Y')
                        print(cpf_valor)
                        
                        ##Aqui se inicia o Loop, Percorrendo todos os arquivos da posta e os lançando
                        
                        #foca para a primeira aba do navegador
                        browser_tabs = driver.window_handles
                        driver.switch_to.window(browser_tabs[0])
                        #acessa o iframe do Card}
                        acessar_relatorio("/html/body/form/div[4]/div/div/div/div/div[3]/div/div[2]/div[2]/div[1]/div/iframe")
                        #pesquisando pelo nome da pessoa no Card
                        escrever("//*[@ID='txtPesquisa-inputEl']", cpf_str)
                        #Pesquisando
                        click("//*[@ID='btnPesquisar-btnWrap']")
                        time.sleep(8)
                        #clicando no campo de Lupa
                        click("/html/body/div[1]/div[2]/div[3]/div/div[2]/table/tbody/tr/td[1]/div/div")
                        time.sleep(8)
                        #mudando o foco para a aba que acabamos de abrir
                        browser_tabs = driver.window_handles
                        driver.switch_to.window(browser_tabs[1])
                        xpath_elemento = "/html/body/div[1]/div/div/div/div/div/div/div/div/div/div/div/div/div/div/header/div/div[4]/div/div[2]/a[2]"
                        elemento_link = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath_elemento)))
                        # Obtém o valor do atributo href (o link)
                        inicio_link = elemento_link.get_attribute('href')
                    
                        # Obtém o valor do atributo href (o link)
                        link = f'{inicio_link}?text=Prezado(a)%2C%20{first_name}%0AEspero%20que%20essa%20mensagem%20o(a)%20encontre%20bem.%0AN%C3%A3o%20identificamos%20sua%20resposta%20a%20nossa%20Entrevista%20de%20Desligamento%20e%20gostar%C3%ADamos%20de%20refor%C3%A7ar%20a%20import%C3%A2ncia%20desse%20retorno%20para%20nossa%20constante%20melhora%20em%20nossos%20processos%20internos.%0AEncaminhamos%20abaixo%2C%20o%20link%20para%20que%20acesse%20e%20responda.%0A%20%0AAgradecemos%20a%20sua%20colabora%C3%A7%C3%A3o!%0A%0Ahttps%3A%2F%2Fapp.smartsheet.com%2Fb%2Fform%2F3c1bac2add974db1aa8b4482f02a6323'
                    
                        
                        # Imprime o link
                        driver.get(link)
                        time.sleep(1)
                        click('/html/body/div[1]/div[1]/div[2]/div/section/div/div/div/div[2]/div[1]/a/span')
                        click('/html/body/div[1]/div[1]/div[2]/div/section/div/div/div/div[3]/div/div/h4[2]/a/span')
                    
                        time.sleep(15)
                        click('/html/body/div[1]/div/div/div[2]/div[4]/div/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span')
                        
                        # imagem = "C:\\Users\\fernando.galves\\gpssa.com.br\\Dashboard_Brian_Silva - Documentos\\05 - AUTOMAÇÕES\\Envio Folha de Ponto\\Automação Mensagem Whatsapp sobre folha ponto\\aplicativo_GPSvc_ Menu Ponto 2.pdf"
                        # print("Anexando imagem:", imagem)
                        # attachment_icon = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[@title="Anexar"]')))
                        # attachment_icon.click()
                        
                        # # Escolhe a opção de imagem
                        # image_option = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')))
                        # image_option.send_keys(imagem)
                        # click('/html/body/div[1]/div/div/div[2]/div[2]/div[2]/span/div/div/div/div[2]/div/div[2]/div[2]/div/div/span')
                        time.sleep(3)
                        df_upload.loc[row_index, 'STATUS'] = 'Enviado'
                        df_upload.to_excel("arquivo_atualizado.xlsx", index=False)  # Salva o DataFrame atualizado de volta no arquivo Excel
                        driver.close()

                    row_index += 1
                except Exception as e:
                    df_upload.loc[row_index, 'STATUS'] = 'Não Enviado'
                    st.write(f'Erro: {e}')
                    df_upload.to_excel("arquivo_atualizado.xlsx", index=False)  # Salva o DataFrame atualizado de volta no arquivo Excel
                    row_index += 1
                    driver.close()