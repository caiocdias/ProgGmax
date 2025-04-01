from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time
import sys


class Navegador_Correlatos:
    def __init__(self):
        self.driver = webdriver.Firefox()

    # Método para fazer login no GServ Web
    def loginGserv(self, valores_login):
        try:
            self.driver.get(url=str(valores_login['link']))
            self.driver.maximize_window()
            time.sleep(1)
            self.driver.find_element(By.ID, "txtUser").send_keys(valores_login['matricula'])
            self.driver.find_element(By.ID, "txtSenha").send_keys(valores_login['senha'])
            time.sleep(1)
            self.driver.find_element(By.ID, "Button1").click()
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'PanelBarItem39')))
            self.driver.find_element(By.ID, "PanelBarItem39").click()
            return True
        except:
            return False

    def buscarNS(self, ns):
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'RadTextBox1')))
        self.driver.find_element(By.ID, "RadTextBox1").click
        time.sleep(1)
        self.driver.find_element(By.ID, "RadTextBox1").send_keys(Keys.CONTROL, "A")
        time.sleep(1)
        self.driver.find_element(By.ID, "RadTextBox1").send_keys(str(ns))
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'Button17')))
        self.driver.find_element(By.ID, "Button17").click()
        WebDriverWait(self.driver, 30).until(
            EC.invisibility_of_element_located((By.ID, '___Form1_AjaxLoadingMainAjaxPanel')),
            "Erro ao buscar NS, carregamento estourou o tempo limite.")
        try:
            self.driver.find_element(By.XPATH,
                                     "/html/body/form/div[3]/div/div[2]/div[2]/div[3]/div[2]/div[1]/table/tbody/tr/td/div")
            return False  # NS Não está cadastrada
        except:
            return True  # NS Está cadastrada

    def cadastrarNS(self, nome, cod_local):
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'Button13')))
        self.driver.find_element(By.ID, "Button13").click()
        WebDriverWait(self.driver, 10).until(EC.frame_to_be_available_and_switch_to_it(
            (By.XPATH, '/html/body/form/div[1]/table/tbody/tr[2]/td[2]/iframe')))
        time.sleep(1)
        WebDriverWait(self.driver, 10).until(
            EC.invisibility_of_element_located((By.CLASS_NAME, 'rwWindowContent rwExternalContent rwLoading')))

        # Box nome cliente
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'RadTextBox2')))
        self.driver.find_element(By.ID, "RadTextBox2").send_keys(str(nome))

        # Box local
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'ComboBox3_Input')))
        self.driver.find_element(By.ID, "ComboBox3_Input").send_keys(str(cod_local))
        time.sleep(1)
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, 'rcbHovered')))
        self.driver.find_element(By.CLASS_NAME, "rcbHovered").click()

        # Salvando
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'Button11')))
        self.driver.find_element(By.ID, "Button11").click()
        time.sleep(1)
        WebDriverWait(self.driver, 30).until(
            EC.invisibility_of_element_located((By.ID, '___Form1_AjaxLoadingMainAjaxPanel')))
        self.driver.switch_to.default_content()
        WebDriverWait(self.driver, 30).until(
            EC.invisibility_of_element_located((By.ID, '___Form1_AjaxLoadingMainAjaxPanel')))

    def cadastrarCorrelato(self, contrato, cod_serv, cod_acao, base, matricula_resp, dat_exec, obs):
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'Button16')))
        self.driver.find_element(By.ID, "Button16").click()
        WebDriverWait(self.driver, 10).until(EC.frame_to_be_available_and_switch_to_it(
            (By.XPATH, '/html/body/form/div[1]/table/tbody/tr[2]/td[2]/iframe')))
        time.sleep(1)
        WebDriverWait(self.driver, 10).until(
            EC.invisibility_of_element_located((By.CLASS_NAME, 'rwWindowContent rwExternalContent rwLoading')))

        # Box contrato
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'ComboBox9_Input')))
        self.driver.find_element(By.ID, "ComboBox9_Input").send_keys(str(contrato))
        time.sleep(1)
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, 'rcbHovered')))
        self.driver.find_element(By.CLASS_NAME, "rcbHovered").click()
        time.sleep(1)

        # Box Serviço
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'ComboBox6_Input')))
        self.driver.find_element(By.ID, "ComboBox6_Input").send_keys(str(cod_serv))
        time.sleep(1)
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, 'rcbHovered')))
        self.driver.find_element(By.CLASS_NAME, "rcbHovered").click()
        time.sleep(1)

        # Box Acao
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'ComboBox7_Input')))
        self.driver.find_element(By.ID, "ComboBox7_Input").send_keys(str(cod_acao))
        time.sleep(1)
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, 'rcbHovered')))
        self.driver.find_element(By.CLASS_NAME, "rcbHovered").click()

        while self.driver.find_element(By.ID, "RadTextBox21").get_attribute("value") == "":
            time.sleep(1)

        # Box base
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'ComboBox11_Input')))
        self.driver.find_element(By.ID, "ComboBox11_Input").send_keys(str(base))
        time.sleep(1)
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, 'rcbHovered')))
        self.driver.find_element(By.CLASS_NAME, "rcbHovered").click()

        # Box Responsavel
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'ComboBox10_Input')))
        self.driver.find_element(By.ID, "ComboBox10_Input").send_keys(str(matricula_resp))
        time.sleep(1)
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, 'rcbHovered')))
        self.driver.find_element(By.CLASS_NAME, "rcbHovered").click()
        time.sleep(1)

        # Box data de exec
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'DatePicker4_dateInput')))
        self.driver.find_element(By.ID, "DatePicker4_dateInput").send_keys(str(dat_exec))

        # Box obs
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'RadTextBox12')))
        self.driver.find_element(By.ID, "RadTextBox12").send_keys(str(obs))

        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'Button11')))
        self.driver.find_element(By.ID, "Button11").click()
        time.sleep(1)
        WebDriverWait(self.driver, 30).until(
            EC.invisibility_of_element_located((By.ID, '___Form1_AjaxLoadingMainAjaxPanel')))
        self.driver.switch_to.default_content()
        WebDriverWait(self.driver, 30).until(
            EC.invisibility_of_element_located((By.ID, '___Form1_AjaxLoadingMainAjaxPanel')))

    def cadastroReq(self, req, item, qtd_us, cod_acao, ns):
        time.sleep(3)
        try:
            i = 1
            while True:
                texto_serv = self.driver.find_element(By.XPATH,
                                                      "/html/body/form/div[3]/div/div[2]/div[3]/div[1]/div[2]/div[1]/table/tbody/tr[{}]/td[4]/span".format(
                                                          i)).text
                texto_serv = texto_serv[0:3]
                if texto_serv == cod_acao:
                    break
                else:
                    i = i + 1
        except:
            sys.exit  # ação nao existe
        self.driver.execute_script("arguments[0].scrollIntoView();", self.driver.find_element(By.XPATH,
                                                                                              "/html/body/form/div[3]/div/div[2]/div[3]/div[1]/div[2]/div[1]/table/tbody/tr[{}]/td[4]/span".format(
                                                                                                  i)))
        self.driver.find_element(By.XPATH,
                                 "/html/body/form/div[3]/div/div[2]/div[3]/div[1]/div[2]/div[1]/table/tbody/tr[{}]/td[4]/span".format(
                                     i)).click()
        WebDriverWait(self.driver, 30).until(
            EC.invisibility_of_element_located((By.ID, '___Form1_AjaxLoadingMainAjaxPanel')),
            "___Form1_AjaxLoadingMainAjaxPanel, timeout de carregamento.")
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'Button18')))
        self.driver.find_element(By.ID, "Button18").click()
        time.sleep(2)
        WebDriverWait(self.driver, 10).until(
            EC.invisibility_of_element_located((By.CLASS_NAME, 'rwWindowContent rwExternalContent rwLoading')))
        WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.ID, 'Grid3_ctl00_ctl02_ctl02_TB_GridColumn330')))
        self.driver.find_element(By.ID, "Grid3_ctl00_ctl02_ctl02_TB_GridColumn330").send_keys(req)
        self.driver.find_element(By.ID, "Grid3_ctl00_ctl02_ctl02_TB_GridColumn331").send_keys(item)
        self.driver.find_element(By.ID, "Grid3_ctl00_ctl02_ctl02_TB_GridColumn332").send_keys(qtd_us)
        time.sleep(1)
        self.driver.find_element(By.ID, "Grid3_ctl00_ctl02_ctl02_PerformInsertButton").click()
        time.sleep(1)
        WebDriverWait(self.driver, 30).until(
            EC.invisibility_of_element_located((By.ID, '___Form1_AjaxLoadingMainAjaxPanel')))