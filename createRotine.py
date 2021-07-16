import os
import pandas as pd
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys


user = "gustavo.boteon"
cont = 0
subcont = 0

chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument(r"user-data-dir=C:\\Users\\" + user + "\\AppData\\Local\\Google\\Chrome\\User Data\\Default")
driver = webdriver.Chrome(executable_path=r'./chromedriver.exe', options=chrome_options)
#driver.get("https://ensp-365-axprod.operations.dynamics.com/?cmp=001&mi=DefaultDashboard")
driver.get("https://enesa-uat.sandbox.operations.dynamics.com/?cmp=001&mi=TrvApprEmplSub")
time.sleep(12)

while True:
    try:
        planilha = "Cadastro Representantes/Entrada/inserirRepr.xlsx"
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
        writer = pd.ExcelWriter(date, engine='xlsxwriter')

        planilha = planilha.split("\\")

        try:
            driver.find_element_by_xpath('//*[@id="trvappremplsub_1_QuickFilterControl_Input_input"]').click
        except:
            input("Por favor efetue o login na plataforma\numa vez finalizado pressione ENTER")

        while (contNome<nomes.size):
            #Novo
            #driver.find_element_by_xpath('/html/body/div[2]/div[1]/div[5]/div/form[2]/div[2]/div/div[3]/button[1]').click()
            driver.find_element_by_xpath('//*[@id="trvappremplsub_1_SystemDefinedNewButton"]').click()
            time.sleep(2)
            #InputNome
            #driver.find_element_by_xpath('//*[@id="trvappremplsub_2_TrvAppEmplSub_DelegatingWorker_DirPerson_FK_Name_input"]').send_keys(nomeFunc)
            driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_DelegatingWorker_PersonnelNumber_input"]').send_keys(Keys.CONTROL, "a")
            time.sleep(1)
            driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_DelegatingWorker_PersonnelNumber_input"]').send_keys(Keys.DELETE)
            driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_DelegatingWorker_DirPerson_FK_Name_input"]').send_keys(Keys.CONTROL, "a")
            time.sleep(1)
            driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_DelegatingWorker_DirPerson_FK_Name_input"]').send_keys(Keys.RETURN)
            driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_DelegatingWorker_DirPerson_FK_Name_input"]').send_keys(nomes[contNome])
            time.sleep(2)
            #Selecionar
            try:
                driver.find_element_by_xpath('//*[@id="trvappremplsub_1_19"]/div[1]/button[2]').click
            finally:
                try:
                    driver.find_element_by_xpath('//*[@id="HcmWorkerLookUp_' + str(cont + 2) + '_OK"]').click()

                except:
                    try:
                        driver.find_element_by_xpath('//*[@id="HcmWorkerLookUp_' + str(subcont + 2) + '_OK"]').click()
                    except:
                        subcont += 1

            time.sleep(1)
            #InputRepresentante
            #driver.find_element_by_xpath('//*[@id="trvappremplsub_2_TrvAppEmplSub_editDelegateUser_input"]').send_keys(nomeRep)
            driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_editDelegateUser_input"]').send_keys(Keys.CONTROL, "a")
            driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_editDelegateUser_input"]').send_keys(Keys.RETURN)
            driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_editDelegateUser_input"]').send_keys(representantes[contNome])
            try:
                driver.find_element_by_xpath('//*[@id="trvappremplsub_1_19"]/div[1]/button[2]').click()
                print("Cadastro do " + nomes[cont-1] + " como representante " + representantes[cont-1] + " fracassou!")
                cadastro.pop()
                cadastro.append("NAO")
            except :
                print('prosseguir')
            finally:
                #DataIni
                #driver.find_element_by_xpath('//*[@id="trvappremplsub_2_TrvAppEmplSub_DateFrom_input"]').send_keys("01/02/2021")
                driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_DateFrom_input"]').send_keys("01/" + mes + "/2021")
                driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_DateFrom_input"]').send_keys(Keys.TAB)
                #DataFim
                #driver.find_element_by_xpath('//*[@id="trvappremplsub_2_TrvAppEmplSub_DateTo_input"]').send_keys("31/12/2099")
                driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_DateTo_input"]').send_keys("31/12/2099")
                driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_DateTo_input"]').send_keys(Keys.TAB)
                time.sleep(1)
                #Salvar
                driver.find_element_by_xpath('//*[@id="trvappremplsub_1_SystemDefinedSaveButton"]').click()
                time.sleep(1)
                try:
                    driver.find_element_by_xpath('//*[@id="trvappremplsub_1_SystemDefinedOptions_button"]').click()
                    print("Cadastro do " + nomes[contNome] + " efetuado, como representante " + representantes[contNome])
                    cadastro.append("SIM")
                except:
                    subcont += 1
                    driver.find_element_by_xpath('//*[@id="SysBoxForm_' + str(subcont + 2) + '_Close"]').click()
                    driver.find_element_by_xpath('//*[@id="trvappremplsub_1_SystemDefinedDeleteButton"]').click()
                    print("Cadastro do " + nomes[contNome] + " como representante " + representantes[contNome] + " fracassou!")
                    cadastro.append("NAO")
                contNome += 1
                cont += 1
                subcont += 1

        df1 = pd.DataFrame({'nome': nomes, 'representante': representantes, 'cadastrado': cadastro})
        writer = pd.ExcelWriter(date, engine='xlsxwriter')
        df1.to_excel(writer, sheet_name='Sheet1', index=False)
        writer.close()
        os.remove('Cadastro Representantes/Entrada/inserirRepr.xlsx')
        endTime = time.perf_counter() - execTime
        print("Processo finalizado com sucesso!\na Execução da planilha demorou " + str(endTime) + " segundos")

    except:
        print("Aguardando planilha . . .")

    time.sleep(30)
