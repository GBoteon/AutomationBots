import os
import pandas as pd
import time
from datetime import datetime

user = pd.read_csv('usuario.txt')
txt = user.usuario
user = str(txt[0])

while True:
    usuar = pd.read_csv('usuario.txt')
    df = pd.read_excel(r'C:\Users\\' + user + '\OneDrive - ENESA Engenharia S.A\Documentos - Cadastro de Representantes\Entrada 001\Modelo inserir cadastro.xlsx')
    df2 = pd.read_excel(r'C:\Users\\' + user + '\OneDrive - ENESA Engenharia S.A\Documentos - Cadastro de Representantes\Entrada 006\Modelo inserir cadastro.xlsx')
    registro = int(usuar.registro[0])
    registro2 = int(usuar.registro[1])

    try:
        nomesNuvem = df.nome
        nomesNuvem2 = df2.nome
    except AttributeError:
        try:
            nomesNuvem = df.NOME_COMPLETO
            nomesNuvem2 = df2.NOME_COMPLETO
        except:
            nomesNuvem = df.NOME
            nomesNuvem2 = df2.NOME
    try:
        representantesNuvem = df.representante
        representantesNuvem2 = df2.representante
    except AttributeError:
        representantesNuvem = df2.REPRESENTANTE
        representantesNuvem2 = df2.REPRESENTANTE

    sizelist = len(nomesNuvem)
    sizelist2 = len(nomesNuvem2)

    if registro != sizelist:
        data_e_hora_atuais = datetime.now()
        cloud = data_e_hora_atuais.strftime('C:\\Users\\' + user + '\\Desktop\\UAT - 001 Cadastro de Representantes\\Cadastro Representantes\\Entrada\\cloud001-%d-%m-%y-%H-%M.xlsx')
        cloud4 = data_e_hora_atuais.strftime('C:\\Users\\' + user + '\\Desktop\\AX - 001 Cadastro de Representantes\\Cadastro Representantes\\Entrada\\cloud001-%d-%m-%y-%H-%M.xlsx')
        df = pd.DataFrame({'nome': nomesNuvem[registro:sizelist], 'representante': representantesNuvem[registro:sizelist]})
        df4 = pd.DataFrame({'nome': nomesNuvem[registro:sizelist], 'representante': representantesNuvem[registro:sizelist]})
        writer = pd.ExcelWriter(cloud, engine='xlsxwriter')
        writer4 = pd.ExcelWriter(cloud4, engine='xlsxwriter')
        df.to_excel(writer, sheet_name='Sheet1', index=False)
        df4.to_excel(writer4, sheet_name='Sheet1', index=False)
        writer.close()
        writer4.close()
        
    elif registro2 != sizelist2:
        data_e_hora_atuais = datetime.now()
        cloud2 = data_e_hora_atuais.strftime('C:\\Users\\' + user + '\\Desktop\\UAT - 006 Cadastro de Representantes\\Cadastro Representantes\\Entrada\\cloud006-%d-%m-%y-%H-%M.xlsx')
        cloud3 = data_e_hora_atuais.strftime('C:\\Users\\' + user + '\\Desktop\\AX - 006 Cadastro de Representantes\\Cadastro Representantes\\Entrada\\cloud006-%d-%m-%y-%H-%M.xlsx')
        df2 = pd.DataFrame({'nome': nomesNuvem2[registro2:sizelist2], 'representante': representantesNuvem2[registro2:sizelist2]})
        df3 = pd.DataFrame({'nome': nomesNuvem2[registro2:sizelist2], 'representante': representantesNuvem2[registro2:sizelist2]})
        writer2 = pd.ExcelWriter(cloud2, engine='xlsxwriter')
        writer3 = pd.ExcelWriter(cloud3, engine='xlsxwriter')
        df2.to_excel(writer2, sheet_name='Sheet1', index=False)
        df3.to_excel(writer3, sheet_name='Sheet1', index=False)
        writer2.close()
        writer3.close()
        
    usuar.registro = [len(nomesNuvem), len(nomesNuvem2), "enesa-uat.sandbox.operations.dynamics.com", "ensp-365-axprod.operations.dynamics.com"]
    usuar.to_csv('usuario.txt', index=False)

    print("Aguardando inserção de dados...")
    time.sleep(60)
    