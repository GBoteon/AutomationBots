import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

df = pd.read_excel('teste func.xlsx')
nomes = df.nome

nomeRep = 'eliana.subires'
#nomeFunc = input("Insira o nome do colaborador: ")
#nomeRep = input("Insira o user o representante: ")
cont = 0

chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument(r"user-data-dir=C:\Users\gustavo.boteon\AppData\Local\Google\Chrome\User Data\Bot")
driver = webdriver.Chrome(executable_path=r'./chromedriver.exe', options=chrome_options)
#driver.get("https://ensp-365-axprod.operations.dynamics.com/?cmp=001&mi=DefaultDashboard")
driver.get("https://enesa-uat.sandbox.operations.dynamics.com/?cmp=001&mi=DefaultDashboard")
print("pausa")
time.sleep(12)
print("volta")
#Favoritos
#driver.find_element_by_xpath('/html/body/div[2]/div[1]/div[4]/div/div/div[2]/div[2]').click()
driver.find_element_by_xpath('/html/body/div[2]/div/div[5]/div/div/div[2]/div[2]').click()
time.sleep(1)
#Representantes
#driver.find_element_by_xpath('/html/body/div[2]/div[1]/div[4]/div/div[1]/div[2]/div[3]').click()
driver.find_element_by_xpath('/html/body/div[2]/div/div[5]/div/div[1]/div[2]/div[3]').click()
print("pausa")
time.sleep(5)
print("volta")
while (cont<3):
    #Novo
    #driver.find_element_by_xpath('/html/body/div[2]/div[1]/div[5]/div/form[2]/div[2]/div/div[3]/button[1]').click()
    driver.find_element_by_xpath('/html/body/div[2]/div/div[6]/div/form[2]/div[2]/div/div[3]/button[1]').click()
    print("pausa")
    time.sleep(2)
    print("volta")
    #InputNome
    #driver.find_element_by_xpath('//*[@id="trvappremplsub_2_TrvAppEmplSub_DelegatingWorker_DirPerson_FK_Name_input"]').send_keys(nomeFunc)
    driver.find_element_by_xpath('//*[@id="trvappremplsub_2_TrvAppEmplSub_DelegatingWorker_DirPerson_FK_Name_input"]').send_keys(nomes[cont])
    print("pausa")
    time.sleep(2)
    print("volta")
    #Selecionar
    #driver.find_element_by_xpath('//*[@id="HcmWorkerLookUp_3_OK"]').click()
    driver.find_element_by_xpath('//*[@id="HcmWorkerLookUp_' + str(cont + 3) + '_OK"]').click()
    time.sleep(1)
    #InputRepresentante
    #driver.find_element_by_xpath('//*[@id="trvappremplsub_2_TrvAppEmplSub_editDelegateUser_input"]').send_keys(nomeRep)
    driver.find_element_by_xpath('//*[@id="trvappremplsub_2_TrvAppEmplSub_editDelegateUser_input"]').send_keys(nomeRep)
    #DataIni
    #driver.find_element_by_xpath('//*[@id="trvappremplsub_2_TrvAppEmplSub_DateFrom_input"]').send_keys("01/02/2021")
    driver.find_element_by_xpath('//*[@id="trvappremplsub_2_TrvAppEmplSub_DateFrom_input"]').send_keys("01/02/2021")
    driver.find_element_by_xpath('//*[@id="trvappremplsub_2_TrvAppEmplSub_DateFrom_input"]').send_keys(Keys.TAB)
    #DataFim
    #driver.find_element_by_xpath('//*[@id="trvappremplsub_2_TrvAppEmplSub_DateTo_input"]').send_keys("31/12/2099")
    driver.find_element_by_xpath('//*[@id="trvappremplsub_2_TrvAppEmplSub_DateTo_input"]').send_keys("31/12/2099")
    driver.find_element_by_xpath('//*[@id="trvappremplsub_2_TrvAppEmplSub_DateTo_input"]').send_keys(Keys.TAB)
    print("pausa")
    time.sleep(1)
    print("volta")
    #Salvar
    driver.find_element_by_xpath('//*[@id="trvappremplsub_2_SystemDefinedSaveButton"]').click()
    print("pausa")
    time.sleep(1)
    print("volta")
    cont += 1
