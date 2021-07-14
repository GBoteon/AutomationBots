import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys

print("Bot para cadastro de Representantes")
user = str(input("Digite seu user:\n"))
planilha = str(input("Insira o caminho da planilha:\n"))
df = pd.read_excel(planilha)
try:
    nomes = df.nome
    coluna = "nome"
except AttributeError:
    try:
        nomes = df.NOME_COMPLETO
        coluna = "NOME_COMPLETO"
    except:
        nomes = df.NOME
        coluna = "NOME"

val = False
tempo = 18 + 7 + (7*nomes.size)
planilha = planilha.split("\\")

print("\nEsse bot fara o cadastro dos colaboradores na planilha *" + planilha[len(planilha)-1] + "* possicionados na coluna *" + coluna + "*,\natualmente existem " + str(nomes.size) + " colaboradores a serem adicionados e a operação ira demorar aproximadamente "+ str(tempo) + " segundos\n")
menu = int(input("Qual será o representante dos colaboradores da planinha\n1.Eliana Subires dos Santos\n2.Jennifer Cruz Asunção\n3.Juliana Cecilia Saiki Dutra\n4.Daniele Gomes Souza\n5.Priscilla Tavares\n6.Bruna Correia de Souza\n7.Gilglezia Maria  de Souza Oliveira\nDigite o Numero: "))
if menu == 1:
    nomeRep = "eliana.subires"
elif menu == 2:
    nomeRep = "jennifer.assuncao"
elif menu == 3:
    nomeRep = "juliana.dutra"
elif menu == 4:
    nomeRep = "daniele.souza"
elif menu == 5:
    nomeRep = "priscilla.tavares"
elif menu == 6:
    nomeRep = "bruna.souza"
elif menu == 7:
    nomeRep = "gilglezia.oliveira"

cont = 0
execTime = time.perf_counter()

chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument(r"user-data-dir=C:\\Users\\" + user + "\\AppData\\Local\\Google\\Chrome\\User Data\\Default")
driver = webdriver.Chrome(executable_path=r'./chromedriver.exe', options=chrome_options)
#driver.get("https://ensp-365-axprod.operations.dynamics.com/?cmp=001&mi=DefaultDashboard")
driver.get("https://enesa-uat.sandbox.operations.dynamics.com/?cmp=001&mi=DefaultDashboard")
time.sleep(12)
try:
    #Favoritos
    #driver.find_element_by_xpath('/html/body/div[2]/div[1]/div[4]/div/div/div[2]/div[2]').click()
    driver.find_element_by_xpath('/html/body/div[2]/div/div[5]/div/div/div[2]/div[2]').click()
    time.sleep(1)
    #Representantes
    #driver.find_element_by_xpath('/html/body/div[2]/div[1]/div[4]/div/div[1]/div[2]/div[3]').click()
    driver.find_element_by_xpath('/html/body/div[2]/div/div[5]/div/div[1]/div[2]/div[3]').click()
    time.sleep(5)

except:
    print("Faça o login na plataforma para prosseguir")
    val = input("Pressione ENTER para prosseguir")
    #Favoritos
    #driver.find_element_by_xpath('/html/body/div[2]/div[1]/div[4]/div/div/div[2]/div[2]').click()
    driver.find_element_by_xpath('//*[@id="navPaneFavoritesID"]').click()
    time.sleep(1)
    #Representantes
    #driver.find_element_by_xpath('/html/body/div[2]/div[1]/div[4]/div/div[1]/div[2]/div[3]').click()
    driver.find_element_by_xpath('//*[@id="mainPane"]/div[5]/div/div[1]/div[2]/div[3]').click()
    time.sleep(5)

finally:
    while (cont<nomes.size):
        #Novo
        #driver.find_element_by_xpath('/html/body/div[2]/div[1]/div[5]/div/form[2]/div[2]/div/div[3]/button[1]').click()
        driver.find_element_by_xpath('/html/body/div[2]/div/div[6]/div/form[2]/div[2]/div/div[3]/button[1]').click()
        time.sleep(2)
        #InputNome
        #driver.find_element_by_xpath('//*[@id="trvappremplsub_2_TrvAppEmplSub_DelegatingWorker_DirPerson_FK_Name_input"]').send_keys(nomeFunc)
        driver.find_element_by_xpath('//*[@id="trvappremplsub_2_TrvAppEmplSub_DelegatingWorker_DirPerson_FK_Name_input"]').send_keys(nomes[cont])
        time.sleep(2)
        #Selecionar
        try:
            driver.find_element_by_xpath('//*[@id="HcmWorkerLookUp_' + str(cont + 3) + '_OK"]').click()
        except:
            driver.find_element_by_xpath('//*[@id="HcmWorkerLookUp_' + str(cont + 5) + '_OK"]').click()

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
        time.sleep(1)
        #Salvar
        driver.find_element_by_xpath('//*[@id="trvappremplsub_2_SystemDefinedSaveButton"]').click()
        time.sleep(1)
        cont += 1
        
endTime = time.perf_counter() - execTime
print("Processo finalizado com sucesso!\na Executção do progama demorou " + str(endTime) + " segundos")
val = input("Pressione ENTER para finalizar o progama")
driver.quit()
