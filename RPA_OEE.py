#bibliotecas
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from webdriver_manager.microsoft import EdgeChromiumDriverManager
import time
import pyautogui as pag
import urllib.request
import os
import pyperclip 

pag.PAUSE = 1.5
pag.FAILSAFE = True

def is_connected():
    try:
        urllib.request.urlopen('https://autowerk.com/performance-portal/oee', timeout=1)
        return True
    except:
        return False
      
def abrir_navegador():
    #abrir o navegador e entrar no portal
    driver = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()))
    driver.get("https://autowerk.com/performance-portal/oee")
    driver.maximize_window()
    time.sleep(4.0)
    return driver

def clicar_login(driver):  
    time.sleep(1)
    pag.moveTo() #usuario
    time.sleep(0.5)
    pag.click() #campo de login
    pag.typewrite('')
    pag.press('tab') #campo de senha
    pag.typewrite('')
    pag.press('enter')   
    
def clicar_oee(driver):
    time.sleep(1)
    pag.click(x=17, y=301) #menu oee
    pag.click(x=32, y=364) #data de ontem

def clicar_turnos(driver):
    time.sleep(0.5)
    pag.click() #turno
    pag.click() #1T
    pag.click() #2T
    pag.click() #3T
    
def salvar_df(driver):
    def excluir_arquivo_turno(caminho_arquivo):
        if os.path.exists(caminho_arquivo):
            os.remove(caminho_arquivo)
            print(f"Arquivo {caminho_arquivo} excluído com sucesso.")
        else:
            print(f"O arquivo {caminho_arquivo} não existe.")  
        
    pag.click() #clicar no logo do excel
    time.sleep(15)
    #excel
    
    import warnings #ignorar avisos
    warnings.filterwarnings("ignore")

    import pandas as pd
    #abrindo a planilha salva e colocando a coluna do turno 
    oee = pd.read_excel(r'C:\Users\user\Downloads\OEE.xlsx', sheet_name = 'Performance')
    oee.insert(0, "TURNO", '1T', True)

    oee2 = pd.read_excel(r'C:\Users\user\Downloads\OEE(1).xlsx', sheet_name = 'Performance')
    oee2.insert(0, "TURNO", '2T', True)

    oee3 = pd.read_excel(r'C:\Users\user\Downloads\OEE(2).xlsx', sheet_name = 'Performance')
    oee3.insert(0, "TURNO", '3T', True)
    
    #concatenando as 3 tabelas
    oee_atualizada = pd.concat([oee, oee2, oee3])
    oee_atualizada = oee_atualizada.reset_index(drop=True)
    oee_group = oee.groupby(list(oee.columns))
    
    #data de ontem para inserir na nova coluna
    from datetime import date, timedelta
    ontem = (date.today() - timedelta(1))

    data = ontem.strftime('%Y-%m-%d')
    oee_atualizada.insert(0, "DATA", data, True)

    #criando e extraindo numero da semana em uma nova coluna
    oee_atualizada['DATA'] = pd.to_datetime(oee_atualizada['DATA'])
    semana = oee_atualizada['DATA'].dt.isocalendar().week
    oee_atualizada.insert(1, "CW", semana, True)

    #colocando as colunas LINHA e PN
    oee_atualizada.insert(3, "LINHA", '', True)
    oee_atualizada.insert(4, "PN", '', True)
    
    #colocar a data no formato br
    oee_atualizada['DATA'] = ontem.strftime('%d/%m/%Y')
    
    #exportando para o excel o dataframe atualizado
    caminho = r'W:\localserver\OEE\Controle de Performance\Produção AW.xlsx'
    with pd.ExcelWriter(caminho,engine="openpyxl", mode='a', if_sheet_exists = 'overlay') as writer:  
        oee_atualizada.to_excel(writer, sheet_name='DADOS OEE', index=False, 
                                startrow=writer.sheets['DADOS OEE'].max_row, header=False)  
    
    #excluir os arquivos dos turnos 
    caminho_arquivo = (r'C:\Users\user\Downloads\OEE.xlsx')
    excluir_arquivo_turno(caminho_arquivo)
    
    caminho_arquivo = (r'C:\Users\user\Downloads\OEE.xlsx')
    excluir_arquivo_turno(caminho_arquivo)
    
    caminho_arquivo = (r'C:\Users\user\Downloads\OEE.xlsx')
    excluir_arquivo_turno(caminho_arquivo)
    
#fechar guia
def fechar_navegador(driver):
    time.sleep(0.5)
    driver.quit()
    
# Abrir navegador e entrar no portal
is_connected()
oee = abrir_navegador()
clicar_login(oee)
clicar_oee(oee)
clicar_turnos(oee)
salvar_df(oee)
fechar_navegador(oee)
