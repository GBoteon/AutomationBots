import pandas as pd

df = pd.read_excel('teste func.xlsx')
nomes = df.nome
print(nomes[0])
