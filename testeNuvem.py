import os
import pandas as pd
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys

usuar = pd.read_csv('usuario.txt')
user = str(usuar.usuario[0])
df2 = pd.read_excel(r'C:\Users\gustavo.boteon\OneDrive - ENESA Engenharia S.A\PowerApps\inserirRepr - Copia - Copia - Copia.xlsx')
try:
    nomesNuvem = df2.nome
except AttributeError:
    try:
        nomesNuvem = df2.NOME_COMPLETO
    except:
        nomesNuvem = df2.NOME
try:
    representantesNuvem = df2.representante
except AttributeError:
    representantesNuvem = df2.REPRESENTANTE

if usuar.registro[0] != len(nomesNuvem):
    for cont in range(len(nomesNuvem)):
        print(nomesNuvem[cont] + " Cadastrado")
        time.sleep(0.5)

    data_e_hora_atuais = datetime.now()
    cloud = data_e_hora_atuais.strftime('Cadastro Representantes/Entrada/cloud-%d-%m-%y-%H-%M.xlsx')
    mes = data_e_hora_atuais.strftime("%m")
    df2 = pd.DataFrame({'nome': nomesNuvem[usuar.registro[0]:len(nomesNuvem)-1], 'representante': representantesNuvem[usuar.registro[0]:len(nomesNuvem)-1]})
    writer = pd.ExcelWriter(cloud, engine='xlsxwriter')
    df2.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.close()


    print(usuar.registro[0])
    print(len(nomesNuvem))
    usuar['registro'] = len(nomesNuvem)

    usuar.to_csv('usuario.txt', index=False)

    print(usuar.registro[0])