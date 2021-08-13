import os
import pandas as pd
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys

user = pd.read_csv('usuario.txt')
txt = user.usuario
user = str(txt[0])
n1 = int(txt[1])
n2 = int(txt[2])
x = True
cont = 0
subcont = 0

def files_path(path):
    return [os.path.join(p, file) for p, _, files in os.walk(os.path.abspath(path)) for file in files]

chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument(r"user-data-dir=C:\\Users\\" + user + "\\AppData\\Local\\Google\\Chrome\\User Data\\Default")
driver = webdriver.Chrome(executable_path=r'./chromedriver.exe', options=chrome_options)
#driver.get("https://ensp-365-axprod.operations.dynamics.com/?cmp=006&mi=TrvApprEmplSub")
driver.get("https://enesa-uat.sandbox.operations.dynamics.com/?cmp=006&mi=TrvApprEmplSub")
time.sleep(12)

try:
    driver.find_element_by_xpath('//*[@id="trvappremplsub_1_QuickFilterControl_Input_input"]').click
except:
    input("Por favor efetue o login na plataforma\numa vez finalizado pressione ENTER")

while True:
    try:
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
        execTime = time.perf_counter()
        data_e_hora_atuais = datetime.now()
        date = data_e_hora_atuais.strftime('Cadastro Representantes/logData/log-%d-%m-%y-%H-%M.xlsx')
        mes = data_e_hora_atuais.strftime("%m")

        planilha = planilha.split("\\")

        try:
            driver.find_element_by_xpath('//*[@id="trvappremplsub_1_QuickFilterControl_Input_input"]').click
        except:
            input("Por favor efetue o login na plataforma\numa vez finalizado pressione ENTER")

        #Novo
        driver.find_element_by_xpath('//*[@id="trvappremplsub_1_SystemDefinedNewButton"]').click()
        time.sleep(2)
        #InputNome
        driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_DelegatingWorker_DirPerson_FK_Name_input"]').send_keys("CLAUDIONOR LOURENCO JUNIOR")
        time.sleep(2)
        #Selecionar
        try:
            driver.find_element_by_xpath('//*[@id="HcmWorkerLookUp_' + str(cont + 2) + '_OK"]').click()
        except:
            try:
                driver.find_element_by_xpath('//*[@id="HcmWorkerLookUp_' + str(subcont + 2) + '_OK"]').click()
            except:
                subcont += 1
                driver.find_element_by_xpath('//*[@id="HcmWorkerLookUp_' + str(subcont + 2) + '_OK"]').click()

        time.sleep(2)
        #InputRepresentante
        driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_editDelegateUser_input"]').send_keys("karen")
        #DataIni
        driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_DateFrom_input"]').send_keys("01/" + mes + "/2021")
        driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_DateFrom_input"]').send_keys(Keys.TAB)
        #DataFim
        driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_DateTo_input"]').send_keys("31/12/2099")
        driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_DateTo_input"]').send_keys(Keys.TAB)
        time.sleep(1)
        #Salvar
        driver.find_element_by_xpath('//*[@id="trvappremplsub_1_SystemDefinedSaveButton"]').click()
        time.sleep(2)

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
            driver.find_element_by_xpath('//*[@id="trvappremplsub_1_SystemDefinedDeleteButton"]').click()
            cont += 1
            subcont += 1
        
        while (contNome<nomes.size):
            #Novo
            driver.find_element_by_xpath('//*[@id="trvappremplsub_1_SystemDefinedNewButton"]').click()
            time.sleep(2)
            #InputNome
            driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_DelegatingWorker_DirPerson_FK_Name_input"]').send_keys(nomes[contNome])
            time.sleep(2)
            #Selecionar
            try:
                driver.find_element_by_xpath('//*[@id="HcmWorkerLookUp_' + str(cont + 2) + '_OK"]').click()
            except:
                try:
                    driver.find_element_by_xpath('//*[@id="HcmWorkerLookUp_' + str(subcont + 2) + '_OK"]').click()
                except:
                    subcont += 1
                    driver.find_element_by_xpath('//*[@id="HcmWorkerLookUp_' + str(subcont + 2) + '_OK"]').click()

            time.sleep(2)
            #InputRepresentante
            driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_editDelegateUser_input"]').send_keys(representantes[contNome])
            #DataIni
            driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_DateFrom_input"]').send_keys("01/" + mes + "/2021")
            driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_DateFrom_input"]').send_keys(Keys.TAB)
            #DataFim
            driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_DateTo_input"]').send_keys("31/12/2099")
            driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_DateTo_input"]').send_keys(Keys.TAB)
            time.sleep(1)
            #Salvar
            driver.find_element_by_xpath('//*[@id="trvappremplsub_1_SystemDefinedSaveButton"]').click()
            time.sleep(2)

            try:
                driver.find_element_by_xpath('//*[@id="trvappremplsub_' + str(n2) + '_' + str(n1) + '"]/div[1]/button[1]').click()
                print('tenta')
                print("Cadastro do " + nomes[contNome] + " como representante " + representantes[contNome] + " fracassou!")
                cadastro.append("NAO")
                driver.find_element_by_xpath('//*[@id="trvappremplsub_1_SystemDefinedDeleteButton"]').click()             
            except:
                try:
                    driver.find_element_by_xpath('//*[@id="trvappremplsub_1_SystemDefinedOptions_button"]').click()
                    print("Cadastro do " + nomes[contNome] + " efetuado, como representante " + representantes[contNome])
                    cadastro.append("SIM")

                except:
                    subcont += 1
                    driver.find_element_by_xpath('//*[@id="SysBoxForm_' + str(subcont + 2) + '_Close"]').click()
                    print("Cadastro do " + nomes[contNome] + " como representante " + representantes[contNome] + " fracassou!")
                    cadastro.append("NAO")
                    driver.find_element_by_xpath('//*[@id="trvappremplsub_1_SystemDefinedDeleteButton"]').click()
            contNome += 1
            cont += 1
            subcont += 1

        df1 = pd.DataFrame({'nome': nomes, 'representante': representantes, 'cadastrado': cadastro})
        writer = pd.ExcelWriter(date, engine='xlsxwriter')
        df1.to_excel(writer, sheet_name='Sheet1', index=False)
        writer.close()
        print('Criado o log da operação: ' + date)
        print('Remoção do Excel: '+ arquivos[0])
        os.remove(str(arquivos[0]))
        arquivos.pop(0)
        endTime = time.perf_counter() - execTime
        print("Processo finalizado com sucesso!\na Execução da planilha demorou " + str(endTime) + " segundos")

    except:
        print("Aguardando planilha . . .")

    time.sleep(30)
