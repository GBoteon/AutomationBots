import os

def files_path(path):
    return [os.path.join(p, file) for p, _, files in os.walk(os.path.abspath(path)) for file in files]


Arquivos = files_path('Cadastro Representantes/Entrada')

print(Arquivos[0])
os.remove(str(Arquivos[0]))
Arquivos.pop(0)
print(Arquivos[0])
