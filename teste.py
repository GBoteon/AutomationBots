import os
import pandas as pd
import time
from datetime import datetime

data_e_hora_atuais = datetime.now()

data = datetime((2021),(8),(1))
data = data.strftime("%d")

print(data_e_hora_atuais)
print(data)