import os
import pandas as pd
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys

usuar = pd.read_csv('C:\\Users\\ti.lcs\\Desktop\\AutomationBots\\usuario.txt')
txt = usuar.usuario
user = str(txt[0])
n1 = int(txt[1])
n2 = int(txt[2])
link = str(txt[3])
x = True
cont = 0
subcont = 0

def files_path(path):
    return [os.path.join(p, file) for p, _, files in os.walk(os.path.abspath(path)) for file in files]

chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument(r"user-data-dir=C:\\Users\\" + user + "\\AppData\\Local\\Google\\Chrome\\User Data\\" + str(usuar.registro[4]))
driver = webdriver.Chrome(executable_path=r'./chromedriver.exe', options=chrome_options)
driver.get("https://" + link + "/?cmp=006&mi=TrvApprEmplSub")
time.sleep(12)

try:
    driver.find_element_by_xpath('//*[@id="trvappremplsub_1_QuickFilterControl_Input_input"]').click
except:
    input("Por favor efetue o login na plataforma\numa vez finalizado pressione ENTER")

while True:
    try:
        #Tenta identificar um novo aquivo inserino na pasta Entrada
        arquivos = files_path('Cadastro Representantes/Entrada')
        planilha = str(arquivos[0])
        df = pd.read_excel(planilha)
        try:
            nomes = df.nome
        except AttributeError:
            try:
                nomes = df.NOME_COMPLETO
            except:
                nomes = df.NOME
        try:
            representantes = df.representante
        except AttributeError:
            representantes = df.REPRESENTANTE

        contNome = 0
        cadastro = []
        nomeErros = []
        reprErros =[]
        execTime = time.perf_counter()
        data_e_hora_atuais = datetime.now()
        #Caminho e nome do arquivo
        date = data_e_hora_atuais.strftime('Cadastro Representantes/logData/log-%d-%m-%y-%H-%M.xlsx')
        mes = data_e_hora_atuais.strftime("%m")
        #Caminho e nome do arquivo
        erros = data_e_hora_atuais.strftime("Cadastro Representantes/Erros/erro-%d-%m-%y-%H-%M.xlsx")

        planilha = planilha.split("\\")
        
        try:
            driver.find_element_by_xpath('//*[@id="trvappremplsub_1_QuickFilterControl_Input_input"]').click
        except:
            input("Por favor efetue o login na plataforma\numa vez finalizado pressione ENTER")

        
#Essa Primeira se????o for??a uma entrada errada para gerar os elementos HTML do bot??o collapse para o auxilio de futuros erros

        #Pressionar Novo
        driver.find_element_by_xpath('//*[@id="trvappremplsub_1_SystemDefinedNewButton"]').click()
        time.sleep(2)
        #Inserir nome colaborador InputNome
        driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_DelegatingWorker_DirPerson_FK_Name_input"]').send_keys("CLAUDIONOR LOURENCO JUNIOR")
        time.sleep(2)
        #Pressionar Selecionar
        try:
            driver.find_element_by_xpath('//*[@id="HcmWorkerLookUp_' + str(cont + 2) + '_OK"]').click()
        except:
            try:
                driver.find_element_by_xpath('//*[@id="HcmWorkerLookUp_' + str(subcont + 2) + '_OK"]').click()
            except:
                subcont += 1
                driver.find_element_by_xpath('//*[@id="HcmWorkerLookUp_' + str(subcont + 2) + '_OK"]').click()

        time.sleep(2)
        #Inserir nome representante InputRepresentante
        driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_editDelegateUser_input"]').send_keys("karen")
        #Inserir Data Inicio
        driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_DateFrom_input"]').send_keys("01/" + mes + "/2021")
        driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_DateFrom_input"]').send_keys(Keys.TAB)
        #Inserir Data Fim
        driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_DateTo_input"]').send_keys("31/12/2099")
        driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_DateTo_input"]').send_keys(Keys.TAB)
        time.sleep(1)
        #Pressionar Salvar
        driver.find_element_by_xpath('//*[@id="trvappremplsub_1_SystemDefinedSaveButton"]').click()
        time.sleep(2)

        #Identificador do id do elemento collapse
        try:
            while x:
                try:
                    driver.find_element_by_xpath('//*[@id="trvappremplsub_' + str(n2) + '_' + str(n1) + '"]/div[1]/button[1]').click()
                    print("n1: " + str(n1) + "\nn2: " + str(n2))
                    x = False
                except:
                    n1 += 1
                    if n1 == 100:
                        n1 = 19
                        n2 = 2
        finally:
            #Pressionar Delete
            driver.find_element_by_xpath('//*[@id="trvappremplsub_1_SystemDefinedDeleteButton"]').click()
            cont += 1
            subcont += 1
        
#Fim da Primeira se????o
        
#Loop de execu????o dos registros

        while (contNome<nomes.size):
            #Pressionar Novo
            driver.find_element_by_xpath('//*[@id="trvappremplsub_1_SystemDefinedNewButton"]').click()
            time.sleep(2)
            #Inserir nome colaborador InputNome
            driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_DelegatingWorker_DirPerson_FK_Name_input"]').send_keys(nomes[contNome].strip().upper())
            time.sleep(2)
            #Pressionar Selecionar
            try:
                driver.find_element_by_xpath('//*[@id="HcmWorkerLookUp_' + str(cont + 2) + '_OK"]').click()
            except:
                try:
                    driver.find_element_by_xpath('//*[@id="HcmWorkerLookUp_' + str(subcont + 2) + '_OK"]').click()
                except:
                    subcont += 1
                    driver.find_element_by_xpath('//*[@id="HcmWorkerLookUp_' + str(subcont + 2) + '_OK"]').click()

            time.sleep(2)
            #Inserir nome representante InputRepresentante
            driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_editDelegateUser_input"]').send_keys(representantes[contNome].strip().lower())
            #Inserir Data Inicio
            driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_DateFrom_input"]').send_keys("01/" + mes + "/2021")
            driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_DateFrom_input"]').send_keys(Keys.TAB)
            #Inserir Data Fim
            driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_DateTo_input"]').send_keys("31/12/2099")
            driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_DateTo_input"]').send_keys(Keys.TAB)
            time.sleep(1)
            #Salvar
            driver.find_element_by_xpath('//*[@id="trvappremplsub_1_SystemDefinedSaveButton"]').click()
            time.sleep(2)

            try:
                #Tenta Pressionar bot??o collapse erros
                driver.find_element_by_xpath('//*[@id="trvappremplsub_' + str(n2) + '_' + str(n1) + '"]/div[1]/button[1]').click()
                print("Cadastro do " + nomes[contNome] + " como representante " + representantes[contNome] + " fracassou!")
                cadastro.append("NAO")
                nomeErros.append(nomes[contNome])
                reprErros.append(representantes[contNome])
                #Deleta a lina do colaborador
                driver.find_element_by_xpath('//*[@id="trvappremplsub_1_SystemDefinedDeleteButton"]').click()             
            except:
                try:
                    #Tenta pressionar botao op????es
                    driver.find_element_by_xpath('//*[@id="trvappremplsub_1_SystemDefinedOptions_button"]').click()
                    print("Cadastro do " + nomes[contNome] + " efetuado, como representante " + representantes[contNome])
                    cadastro.append("SIM")
                    driver.find_element_by_xpath('//*[@id="trvappremplsub_1_SystemDefinedOptions_button"]').click()

                except:
                    subcont += 1
                    driver.find_element_by_xpath('//*[@id="SysBoxForm_' + str(subcont + 2) + '_Close"]').click()
                    print("Cadastro do " + nomes[contNome] + " como representante " + representantes[contNome] + " fracassou!")
                    cadastro.append("NAO")
                    nomeErros.append(nomes[contNome])
                    reprErros.append(representantes[contNome])
                    #Deleta a linha do colaborador
                    driver.find_element_by_xpath('//*[@id="trvappremplsub_1_SystemDefinedDeleteButton"]').click()
            contNome += 1
            cont += 1
            subcont += 1

        #Gera DataFrame para cria????o do log
        df1 = pd.DataFrame({'nome': nomes, 'representante': representantes, 'cadastrado': cadastro})
        writer = pd.ExcelWriter(date, engine='xlsxwriter')
        df1.to_excel(writer, sheet_name='Sheet1', index=False)
        writer.close()
        if len(nomeErros) != 0:
            #Gera DataFrame para cria????o do Erro
            df2 = pd.DataFrame({'nome': nomeErros, 'representante': reprErros})
            writer2 = pd.ExcelWriter(erros, engine='xlsxwriter')
            df2.to_excel(writer2, sheet_name='Sheet1', index=False)
            writer2.close()
            print("Log de erros: " + erros + " criado")
        print('Criado o log da opera????o: ' + date)
        print('Remo????o do Excel: '+ arquivos[0])
        os.remove(str(arquivos[0]))
        arquivos.pop(0)
        endTime = time.perf_counter() - execTime
        print("Processo finalizado com sucesso!\na Execu????o da planilha demorou " + str(endTime) + " segundos")

    except:
        #Pressiona o bot??o op????es para impedir a desconex??o por inatividade
        driver.find_element_by_xpath('//*[@id="trvappremplsub_1_SystemDefinedOptions_button"]').click()
        print("Aguardando planilha . . .")
        logs = files_path("Cadastro Representantes/logData")

        #For para deletar logs com mais de dois 2 dias de cria????o
        for lista in range(len(logs)):
            planilha = logs[lista].split("\\")
            planilha = planilha[len(planilha)-1]

            dia = int(planilha[4] + planilha[5])
            mes = int(planilha[7] + planilha[8])
            ano = int("20" + planilha[10] + planilha[11])

            data_arquivo = datetime(ano, mes, dia)

            data_e_hora_atuais = datetime.now()

            diferen??a = data_e_hora_atuais - data_arquivo
            #Caso queira alterar a quantidade de dias mudar o n??mero
            if diferen??a.days > 2:
                os.remove(str(logs[lista]))
        time.sleep(1)
        driver.find_element_by_xpath('//*[@id="trvappremplsub_1_SystemDefinedOptions_button"]').click()

    time.sleep(30)

#Fim do loop