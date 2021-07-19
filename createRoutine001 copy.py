import os
import pandas as pd
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys

user = pd.read_csv('usuario.txt')
user = user.usuario
user = str(user[0])
cont = 0
subcont = 0

def files_path(path):
    return [os.path.join(p, file) for p, _, files in os.walk(os.path.abspath(path)) for file in files]

chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument(r"user-data-dir=C:\\Users\\" + user + "\\AppData\\Local\\Google\\Chrome\\User Data\\Default")
driver = webdriver.Chrome(executable_path=r'./chromedriver.exe', options=chrome_options)
#driver.get("https://ensp-365-axprod.operations.dynamics.com/?cmp=001&mi=DefaultDashboard")
driver.get("https://enesa-uat.sandbox.operations.dynamics.com/?cmp=001&mi=TrvApprEmplSub")
time.sleep(12)

while True:
    input("enter")
    driver.find_element_by_xpath('//*[@id="trvappremplsub_1_19"]/div[1]/button[2]').click
