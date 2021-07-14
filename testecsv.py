import time
from datetime import datetime
import pandas as pd


planilha = str(input("Insira o caminho da planilha:\n"))
df = pd.read_excel(planilha)
nomes = df.nome
print(nomes.size)
planilha = planilha.split("\\")
print(planilha[len(planilha)-1])

#data_e_hora_atuais = datetime.now()
#data_e_hora_em_texto = data_e_hora_atuais.strftime("%d/%m/%Y")

#print(data_e_hora_em_texto)