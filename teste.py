import pandas as pd

user = pd.read_csv('usuario.txt')
user = user.usuario
user = str(user[0])
print(user)
