import os
import pandas as pd
import time
from datetime import datetime

data_e_hora_atuais = datetime.now()
date = data_e_hora_atuais.strftime('Cadastro Representantes/logData/log-%d-%m-%y-%H-%M.xlsx')
erros = data_e_hora_atuais.strftime("Cadastro Representantes/Erros/erro-%d-%m-%y-%H-%M.xlsx")

nomes = []

print("Nomes: " + str(len(nomes)))

nomes = ["a", "s", "f", "f"]
representantes = ["a", "s", "f", "f"]
nomeErros = ["a", "s", "f", "f"]
reprErros = ["a", "s", "f", "f"]

'''
df1 = pd.DataFrame({'nome': nomes, 'representante': representantes})
df2 = pd.DataFrame({'nome': nomeErros, 'representante': reprErros})
writer = pd.ExcelWriter(date, engine='xlsxwriter')
writer2 = pd.ExcelWriter(erros, engine='xlsxwriter')
df1.to_excel(writer, sheet_name='Sheet1', index=False)
df2.to_excel(writer2, sheet_name='Sheet1', index=False)
writer.close()
writer2.close()
'''