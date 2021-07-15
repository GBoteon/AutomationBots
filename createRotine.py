import os
import pandas as pd
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys

data_e_hora_atuais = datetime.now()
date = data_e_hora_atuais.strftime('logData/log-%d-%m-%y-%H-%M.xlsx')
mes = data_e_hora_atuais.strftime("%m")
user = "gustavo.boteon"

chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument(r"user-data-dir=C:\\Users\\" + user + "\\AppData\\Local\\Google\\Chrome\\User Data\\Default")
driver = webdriver.Chrome(executable_path=r'./chromedriver.exe', options=chrome_options)
#driver.get("https://ensp-365-axprod.operations.dynamics.com/?cmp=001&mi=DefaultDashboard")
driver.get("https://enesa-uat.sandbox.operations.dynamics.com/?cmp=001&mi=TrvApprEmplSub")
time.sleep(12)

while True:
    try:
        execTime = time.perf_counter()
        cadastro = []
        writer = pd.ExcelWriter(date, engine='xlsxwriter')
        planilha = "inserirRepr.xlsx"
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

        planilha = planilha.split("\\")

        cont = 0
        subcont = 0

        try:
            driver.find_element_by_xpath('//*[@id="trvappremplsub_1_QuickFilterControl_Input_input"]').click
        except:
            input("Por favor efetue o login na plataforma\numa vez finalizado pressione ENTER")

        while (cont<nomes.size):
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
            driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_DelegatingWorker_DirPerson_FK_Name_input"]').send_keys(nomes[cont])
            time.sleep(2)
            #Selecionar
            try:
                driver.find_element_by_xpath('//*[@id="HcmWorkerLookUp_' + str(cont + 2) + '_OK"]').click()

            except:
                driver.find_element_by_xpath('//*[@id="HcmWorkerLookUp_' + str(subcont + 2) + '_OK"]').click()

            time.sleep(1)
            #InputRepresentante
            #driver.find_element_by_xpath('//*[@id="trvappremplsub_2_TrvAppEmplSub_editDelegateUser_input"]').send_keys(nomeRep)
            driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_editDelegateUser_input"]').send_keys(Keys.CONTROL, "a")
            driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_editDelegateUser_input"]').send_keys(Keys.RETURN)
            driver.find_element_by_xpath('//*[@id="trvappremplsub_1_TrvAppEmplSub_editDelegateUser_input"]').send_keys(representantes[cont])
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
                print("Cadastro do " + nomes[cont] + " efetuado, como representante " + representantes[cont])
                cadastro.append("SIM")
            except:
                subcont += 1
                driver.find_element_by_xpath('//*[@id="SysBoxForm_5_Close"]').click()
                driver.find_element_by_xpath('//*[@id="trvappremplsub_1_SystemDefinedDeleteButton"]').click()
                print("Cadastro do " + nomes[cont] + " como representante " + representantes[cont] + " fracassou!")
                cadastro.append("NAO")
            cont += 1
            subcont += 1

        df1 = pd.DataFrame({'nome': nomes, 'representante': representantes, 'cadastrado': cadastro})
        writer = pd.ExcelWriter(date, engine='xlsxwriter')
        df1.to_excel(writer, sheet_name='Sheet1', index=False)
        writer.close()
        os.remove('inserirRepr.xlsx')
        endTime = time.perf_counter() - execTime
        print("Processo finalizado com sucesso!\na Execução da planilha demorou " + str(endTime) + " segundos")

    except:
        print("Aguadando planilha . . .")
    
    time.sleep(10)
