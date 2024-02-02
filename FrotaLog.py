from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.alert import Alert

import json
import os
import time
from Sheets_Python import Sheets_Python

class FrotaLog():
    def __init__(self):
        self.downloads_path = os.path.join(os.getcwd(),"Downloads_Selenium")
        self.driver_path = os.path.join(os.getcwd(),'assets\\chromedriver.exe')
        
        self.login = ''
        self.senha = ''

        
        self.options = webdriver.ChromeOptions()
        self.appState = {
            "recentDestinations": [
                {
                    "id": "Save as PDF",
                    "origin": "local"
                }
            ],
            "selectedDestinationId": "Save as PDF",
            "version": 2
        } 
        self.prefs = {'printing.print_preview_sticky_settings.appState': json.dumps(self.appState),
                'savefile.default_directory':self.downloads_path,
                'download.prompt_for_download': False,
                'download.directory_upgrade': True,
                'plugins.always_open_pdf_externally': True,
                'download.default_directory' : self.downloads_path}
        self.options.add_experimental_option("prefs",self.prefs)
        self.options.add_argument('--kiosk-printing')
    
    def take_print_of_route(self, coordinates,id_sheets_to_search,range_sheets_to_search):
        to_search = Sheets_Python().coleta_dados_sheets(id_sheets=id_sheets_to_search,range_sheets=range_sheets_to_search)
        print(to_search)

        chrome = webdriver.Chrome(service= Service(self.driver_path), options=self.options)

        chrome.get('https://www.frotalog.com.br/MBServerO/login.do') 
        while True:
            try:
                alert = chrome.switch_to.alert
                alert.accept()
                break
            except:
                pass
        chrome.maximize_window()
        WebDriverWait(chrome, 200).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="userName"]')))
        chrome.find_element('xpath', '//*[@id="userName"]').send_keys(self.login)
        WebDriverWait(chrome, 200).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="box"]//input[@name="password"]')))
        chrome.find_element('xpath', '//*[@id="box"]//input[@name="password"]').send_keys(self.senha)
        WebDriverWait(chrome, 200).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="box"]/div[1]//input[@type="image"]')))
        chrome.find_element('xpath', '//*[@id="box"]/div[1]//input[@type="image"]').click()
        
        try:
            for x in range(len(to_search[0])):
                WebDriverWait(chrome, 200).until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'main'))) #By.NAME
                WebDriverWait(chrome, 200).until(EC.element_to_be_clickable((By.XPATH, '//span[text()="Selecione a empresa"]/..')))
                chrome.find_element('xpath', '//span[text()="Selecione a empresa"]/..').click()
                WebDriverWait(chrome, 200).until(EC.element_to_be_clickable((By.XPATH, '//label[text()=" ECOELÉTRICA ENGENHARIA"]/input')))
                chrome.find_element('xpath', '//label[text()=" ECOELÉTRICA ENGENHARIA"]/input').click()
                WebDriverWait(chrome, 200).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="showAlerts"]')))
                chrome.find_element('xpath', '//*[@id="showAlerts"]').click()
            
                START_DATE = f'{to_search[1][x]} 00:00'
                END_DATE = f'{to_search[1][x]} 23:00'
                PLACA = to_search[0][x]
                print(START_DATE)
                print(END_DATE)
                WebDriverWait(chrome, 200).until(EC.element_to_be_clickable((By.XPATH, '//a[@href="#rebuildRoute"]')))
                chrome.find_element('xpath', '//a[@href="#rebuildRoute"]').click()
                WebDriverWait(chrome, 200).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="beginDate"]')))
                chrome.find_element('xpath', '//*[@id="beginDate"]').clear()
                chrome.find_element('xpath', '//*[@id="beginDate"]').send_keys(START_DATE)
                
                WebDriverWait(chrome, 200).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="endDate"]')))
                chrome.find_element('xpath', '//*[@id="endDate"]').clear()
                chrome.find_element('xpath', '//*[@id="endDate"]').send_keys(END_DATE)
                
                WebDriverWait(chrome, 200).until(EC.element_to_be_clickable((By.XPATH, '//div[@id="rebuildRoute"]//input[@placeholder="Digite para filtrar, ou ESC para limpar"]')))
                chrome.find_element('xpath', '//div[@id="rebuildRoute"]//input[@placeholder="Digite para filtrar, ou ESC para limpar"]').clear()
                chrome.find_element('xpath', '//div[@id="rebuildRoute"]//input[@placeholder="Digite para filtrar, ou ESC para limpar"]').send_keys(PLACA)
                WebDriverWait(chrome, 200).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="rebuildRoute"]//input[@name="tipoRec" and @value="2"]')))
                chrome.find_element('xpath', '//*[@id="rebuildRoute"]//input[@name="tipoRec" and @value="2"]').click()
                WebDriverWait(chrome, 200).until(EC.element_to_be_clickable((By.XPATH, '//div[@id="rebuildRoute"]//input[@value="Pesquisar"]')))
                chrome.find_element('xpath', '//div[@id="rebuildRoute"]//input[@value="Pesquisar"]').click()
                WebDriverWait(chrome, 200).until(EC.element_to_be_clickable((By.XPATH, '//img[contains(@src,"Fim")]')))
                WebDriverWait(chrome, 200).until(EC.element_to_be_clickable((By.XPATH, '//img[contains(@src,"Inicio")]')))
                WebDriverWait(chrome, 200).until(EC.element_to_be_clickable((By.XPATH, '//a[@href="#radiusSrch"]')))
                chrome.find_element('xpath', '//a[@href="#radiusSrch"]').click()
                
                
                WebDriverWait(chrome, 200).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="radiusSrch-coord-check"]')))
                chrome.find_element('xpath', '//*[@id="radiusSrch-coord-check"]').click()
                WebDriverWait(chrome, 200).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="radiusSrch"]//input[@placeholder="Entre com o endereço"]')))
                chrome.find_element('xpath', '//*[@id="radiusSrch"]//input[@placeholder="Entre com o endereço"]').clear()
                chrome.find_element('xpath', '//*[@id="radiusSrch"]//input[@placeholder="Entre com o endereço"]').send_keys(coordinates)
                WebDriverWait(chrome, 200).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="map_canvas"]//img[contains(@src,"marker-icon")]')))
                
                chrome.implicitly_wait(10)
                time.sleep(3)
                if not os.path.exists(os.path.join(os.getcwd(),PLACA)):
                    os.mkdir(os.path.join(os.getcwd(),PLACA))
                chrome.save_screenshot(os.path.join(os.getcwd(),f'{PLACA}\\{START_DATE[:10].replace("/","_")}_COORDENADAS.png'))
                WebDriverWait(chrome, 200).until(EC.element_to_be_clickable((By.XPATH, '//a[@href="#rebuildRoute"]')))
                chrome.find_element('xpath', '//a[@href="#rebuildRoute"]').click()
                WebDriverWait(chrome, 200).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="map_canvas"]//img[contains(@src,"marker-icon")]')))
                time.sleep(1)
                chrome.save_screenshot(os.path.join(os.getcwd(),f'{PLACA}\\{START_DATE[:10].replace("/","_")}_PLACA.png'))
                
                chrome.refresh()
        except Exception as err:
            print(err)
            input('Wait!\nLook on the page for information about the error')
            
            
if __name__ == "__main__":
    FrotaLog().take_print_of_route(coordinates='-7.896365729917244, -40.09166381256962',id_sheets_to_search='1xyyV66wzTxbkOcMKMxtOpHfO0lzN9a_IYptBKXEiy6I',range_sheets_to_search='To_Search!A2:B')