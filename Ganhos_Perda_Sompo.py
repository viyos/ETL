import time
from selenium.webdriver.common.by import By
from datetime import datetime, timedelta
import os
import shutil
import zipfile
from datetime import date, timedelta
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import pandas as pd
import mysql.connector
from sqlalchemy import create_engine

login='ADMINISTRADOR'
senha='branco'

db_data = 'mysql+mysqldb://' + 'user' + ':' + 'password' + '@' + 'IP' + 'library' + '?charset=utf8mb4'
engine = create_engine(db_data)

def Remove_Arquivo(path_arq):
	if os.path.exists(path_arq):
		os.remove(path_arq)

def bot_cb(lista):
    for data in lista:
        Remove_Arquivo(r'Relatorio Perda Mercado Analitico.xls')
        Remove_Arquivo(r'Relatorio Ganho Mercado Analitico.xls')
        edge_options = webdriver.EdgeOptions()
        ser = Service("C:\\Users\\victory\\OneDrive - HDI SEGUROS SA\\Área de Trabalho\\msedgedriver.exe")    
        driver = webdriver.Edge(service=ser)
        driver.maximize_window()
        driver.get("https://centralbonus/")
        time.sleep(2)
        loginbtn = driver.find_element(By.CSS_SELECTOR, "#txtUsuario")
        loginbtn.send_keys(login)
        passbtn = driver.find_element(By.CSS_SELECTOR, "#txtSenha")
        passbtn.send_keys(senha)
        enter_btn=driver.find_element(By.CSS_SELECTOR, "#btnLogar")
        enter_btn.click()
        time.sleep(2)
        driver.get("https://centralbonus/")
        #Tempo de espera para o carregamento da página
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable(('xpath','/html/body/form/div[3]/div[4]/div/div[1]/span/input')))
        time.sleep(2)
        #Limpa a barra da data inicial
        driver.find_element(By.XPATH, "/html/body/form/div[3]/div[4]/div/div[1]/span/input").clear()
        #Envia a data inicial
        driver.find_element(By.XPATH, "/html/body/form/div[3]/div[4]/div/div[1]/span/input").send_keys(data)
        #Limpa a barra da data final
        driver.find_element(By.XPATH, "/html/body/form/div[3]/div[4]/div/div[3]/span/input").clear()
        #Envia a data final
        driver.find_element(By.XPATH, "/html/body/form/div[3]/div[4]/div/div[3]/span/input").send_keys(data)
        #Pesquisa pelo relatorio
        pesquisar_btn = driver.find_element(By.CSS_SELECTOR,'#MainContent_btnPesquisar')
        pesquisar_btn.click()
        time.sleep(2)
        download_btn=driver.find_element(By.CSS_SELECTOR,"#MainContent_btnExcelAnalitico")
        download_btn.click()
        data = data.replace('/','-')
        time.sleep(15)
        src_ganho = r"Relatorio Ganho Mercado Analitico.xls"
        src_ganho_csv = r"Relatorio Ganho Mercado Analitico.csv"
        try:
            WebDriverWait(driver, 5).until(EC.alert_is_present())
            driver.switch_to.alert.accept()
        except TimeoutException:
            try:
                read_excel_ganhos = pd.read_html(src_ganho)
                df_ganhos = read_excel_ganhos[0]
                df_ganhos = df_ganhos[:]
            except ValueError:
                pass
            try:
                df_ganhos[['CodSeguradora', 'Seguradora','None']] = df_ganhos['Seguradora'].str.split(' ', expand=True)
                del df_ganhos['None'] #Depois disso enviar para o MySQL
            except ValueError:
                df_ganhos[['CodSeguradora', 'Seguradora']] = df_ganhos['Seguradora'].str.split(' ', expand=True)
            except:
                pass
            try:
                df_ganhos.to_sql('ganhos_sompo' , engine, if_exists='append', index=False,method='multi')
                df_ganhos.to_csv(src_ganho_csv, index=False,sep=';',encoding='utf-8-sig')
                new_name_ganho = "Relatorio Ganho Mercado Analitico_" + data
                dst_ganho = rf"\\hdist02\departamentos\ProdutoAuto\Analytics_Digital\Victor Yoshida\Relatorio_Ganhos_Sompo\{new_name_ganho}.csv"
                shutil.move(src_ganho_csv, dst_ganho)
            except:
                pass
        driver.get("https://centralbonus/")
        #Tempo de espera para o carregamento da página
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable(('xpath','/html/body/form/div[3]/div[4]/div/div[1]/span/input')))
        time.sleep(2)
        #Limpa a barra da data inicial
        driver.find_element(By.XPATH, "/html/body/form/div[3]/div[4]/div/div[1]/span/input").clear()
        #Envia a data inicial
        driver.find_element(By.XPATH, "/html/body/form/div[3]/div[4]/div/div[1]/span/input").send_keys(data)
        #Limpa a barra da data final
        driver.find_element(By.XPATH, "/html/body/form/div[3]/div[4]/div/div[3]/span/input").clear()
        #Envia a data final
        driver.find_element(By.XPATH, "/html/body/form/div[3]/div[4]/div/div[3]/span/input").send_keys(data)
        #Pesquisa pelo relatorio
        pesquisar_btn = driver.find_element(By.CSS_SELECTOR,'#MainContent_btnPesquisar')
        pesquisar_btn.click()
        time.sleep(2)
        download_btn=driver.find_element(By.CSS_SELECTOR,"#MainContent_btnExcelAnalitico")
        download_btn.click()
        data = data.replace('/','-')
        time.sleep(10)
        src_perda = r"Relatorio Perda Mercado Analitico.xls"
        src_perda_csv = r"Relatorio Perda Mercado Analitico.csv"
        try:
            WebDriverWait(driver, 5).until(EC.alert_is_present())
            driver.switch_to.alert.accept()
        except TimeoutException:
            try:
                read_excel_perdas = pd.read_html(src_perda)
                df_perdas = read_excel_perdas[0]
                df_perdas = df_perdas[:]
            except ValueError:
                pass
            try:
                df_perdas[['CodSeguradora', 'Seguradora','None']] = df_perdas['Seguradora'].str.split(' ', expand=True)
                del df_perdas['None'] #Depois disso enviar para o MySQL
            except ValueError:
                df_perdas[['CodSeguradora', 'Seguradora']] = df_perdas['Seguradora'].str.split(' ', expand=True)
            except:
                pass
            try:  
                df_perdas.to_sql('perdas_sompo' , engine, if_exists='append', index=False,method='multi')
                df_perdas.to_csv(src_perda_csv, index=False,sep=';',encoding='utf-8-sig')   
                new_name_perda = "Relatorio Perda Mercado Analitico_" + data
                dst_perda = rf"Relatorio_Perdas_Sompo\{new_name_perda}.csv"
                shutil.move(src_perda_csv, dst_perda)
            except:
                pass
        driver.close()

def lista_datas():

     start_dt = datetime.today() + timedelta(days=-2) 
     end_dt = datetime.today() + timedelta(days=-2) 

     # difference between current and previous date
     delta = timedelta(days=1)

     # store the dates between two dates in a list
     dates = []

     while start_dt <= end_dt:
         # add current date to list by converting  it to iso format
         dates.append(start_dt.strftime('%d/%m/%Y'))
         #30/09/2023
         # increment start date by timedelta
         start_dt += delta

     return dates

if __name__=="__main__":
    lista = lista_datas()
    bot_cb(lista)
    #print(lista)