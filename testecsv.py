import time
from datetime import datetime
import pandas as pd




data_e_hora_atuais = datetime.now()
data_e_hora_em_texto = data_e_hora_atuais.strftime("%m")

print(data_e_hora_em_texto)