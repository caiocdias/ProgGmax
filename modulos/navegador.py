from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
import time
from selenium.common.exceptions import ElementClickInterceptedException, StaleElementReferenceException

class Navegador:
    def __init__(self):
        self.driver = webdriver.Firefox()

    def _set_masked_date(self, input_id: str, date_str: str, blur_locator=None, timeout=10, tries=2):
        """
        Preenche campos de data com máscara (Telerik RadDatePicker/RadDateInput) de forma robusta.
        1) Tenta usar a API JS do Telerik ($find(...).set_selectedDate / set_value).
        2) Se não houver, faz fallback para digitação controlada (CONTROL+A, DELETE, digitação char a char).
        3) Dispara 'change'/'blur' e valida o valor final; reitera se necessário.
        """
        wait = WebDriverWait(self.driver, timeout)
        el = wait.until(EC.element_to_be_clickable((By.ID, input_id)))
        self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)

        js = r"""
            (function(inputId, value){
                function dispatchAll(elem){
                    try {
                        elem.dispatchEvent(new Event('input', {bubbles:true}));
                        elem.dispatchEvent(new Event('change', {bubbles:true}));
                        elem.dispatchEvent(new Event('blur', {bubbles:true}));
                    } catch(e) {}
                }

                var el = document.getElementById(inputId);
                if (!el) return null;

                // Tenta Telerik: pegar o RadDatePicker e/ou RadDateInput
                try {
                    if (window.$find) {
                        var pickerId = inputId.replace('_dateInput', '');
                        var picker = $find(pickerId);
                        if (picker && picker.set_selectedDate) {
                            // value dd/MM/yyyy
                            var p = value.split('/');
                            var d = new Date(parseInt(p[2],10), parseInt(p[1],10)-1, parseInt(p[0],10));
                            picker.set_selectedDate(d);
                            var di = picker.get_dateInput && picker.get_dateInput();
                            if (di && di.set_value) {
                                di.set_value(value);
                                // alguns skins exigem raise_change
                                if (di.raise_change) di.raise_change();
                            } else {
                                el.value = value;
                                dispatchAll(el);
                            }
                            return el.value;
                        }
                        // Ou RadDateInput diretamente
                        var ri = $find(inputId);
                        if (ri && ri.set_value) {
                            ri.set_value(value);
                            if (ri.raise_change) ri.raise_change();
                            return el.value;
                        }
                    }
                } catch(e) {}

                // Fallback: setar direto + eventos
                el.value = value;
                dispatchAll(el);
                return el.value;
            })(arguments[0], arguments[1]);
        """

        for _ in range(tries):
            # 1) tenta via JS/Telerik
            got = self.driver.execute_script(js, input_id, date_str)
            time.sleep(0.2)

            # blur opcional (p.ex. clicar num label que valida o campo)
            if blur_locator:
                try:
                    self.driver.find_element(*blur_locator).click()
                except Exception:
                    pass

            val = self.driver.find_element(By.ID, input_id).get_attribute("value") or ""
            if val.strip() == date_str:
                return

            # 2) fallback: digitação controlada
            el = self.driver.find_element(By.ID, input_id)
            el.click()
            el.send_keys(Keys.CONTROL, 'a')
            el.send_keys(Keys.DELETE)
            time.sleep(0.1)
            for ch in date_str:
                el.send_keys(ch)
                time.sleep(0.05)  # dá tempo para a máscara
            el.send_keys(Keys.TAB)
            time.sleep(0.2)

            if blur_locator:
                try:
                    self.driver.find_element(*blur_locator).click()
                except Exception:
                    pass

            val = self.driver.find_element(By.ID, input_id).get_attribute("value") or ""
            if val.strip() == date_str:
                return

        raise RuntimeError(f"Falha ao setar data em {input_id}: obtive '{val}' (esperado '{date_str}')")

    # Método para fazer login no GServ Web
    def loginGserv(self, valores_login):
        try:
            self.driver.get(url=str(valores_login['link']))
            self.driver.maximize_window()
            time.sleep(1)
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'txtUser')))
            self.driver.find_element(By.ID, "txtUser").send_keys(valores_login['matricula'])
            self.driver.find_element(By.ID, "txtSenha").send_keys(valores_login['senha'])
            time.sleep(1)
            self.driver.find_element(By.ID, "Button1").click()
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'PanelBarItem38')))
            self.driver.find_element(By.ID, "PanelBarItem38").click()
            return True
        except:
            return False

    def buscarNS(self, numNS):
        time.sleep(2)
        WebDriverWait(self.driver, 10).until(EC.visibility_of_element_located(
            (By.ID, 'RadTextBox1')),
                                             "Erro ao buscar NS, box da NS nao localizado.")
        self.driver.find_element(By.ID, "RadTextBox1").click()

        element = self.driver.find_element(By.ID, "RadTextBox1")
        actions = ActionChains(self.driver)
        actions.double_click(element).perform()

        time.sleep(1)
        self.driver.find_element(By.ID, "RadTextBox1").send_keys(
            numNS)
        time.sleep(1)
        self.driver.find_element(By.XPATH, "/html/body/form/div[3]/div/div[2]/div[1]/div[2]/button").click()
        time.sleep(1)
        WebDriverWait(self.driver, 30).until(
            EC.invisibility_of_element_located((By.ID, '___Form1_AjaxLoadingMainAjaxPanel')),
            "Erro ao buscar NS, carregamento estourou o tempo limite.")
        try:
            self.driver.find_element(By.XPATH,
                                     "/html/body/form/div[3]/div/div[2]/div[1]/div[3]/div/div[1]/table/tbody/tr/td[2]/div")
            return False  # NS Não está cadastrada
        except:
            return True  # NS Está cadastrada

    def buscarServ(self, serv=str, med=str):
        try:
            i = 1
            while True:
                texto_serv = self.driver.find_element(By.XPATH,
                                                      "/html/body/form/div[3]/div/div[2]/div[2]/div/div[2]/div[1]/table/tbody/tr[{}]/td[4]/span".format(
                                                          i)).text
                texto_serv = texto_serv[0:4]
                num_med = self.driver.find_element(By.XPATH,
                                                   "/html/body/form/div[3]/div/div[2]/div[2]/div/div[2]/div[1]/table/tbody/tr[{}]/td[6]".format(
                                                       i)).text
                if texto_serv == serv and num_med == med:
                    return i
                else:
                    i = i + 1
        except:
            return 0  # Serviço não existe

    def buscarAcao(self, serv_index=int, acao=str):
        wait = WebDriverWait(self.driver, 30)
        overlay = (By.ID, '___Form1_AjaxLoadingMainAjaxPanel')

        # 1) garanta que não há overlay antes de interagir
        wait.until(EC.invisibility_of_element_located(overlay))

        # 2) use o <td> como alvo principal (mais confiável que o <span>)
        td_locator = (By.XPATH,
                      f"/html/body/form/div[3]/div/div[2]/div[2]/div/div[2]/div[1]/table/tbody/tr[{serv_index}]/td[4]")

        # 3) traga para a viewport no centro, depois espere ficar clicável
        td_el = wait.until(EC.presence_of_element_located(td_locator))
        self.driver.execute_script("arguments[0].scrollIntoView({block:'center', inline:'nearest'});", td_el)
        td_el = wait.until(EC.element_to_be_clickable(td_locator))

        # 4) clique com fallback para JS (evita interceptação)
        try:
            td_el.click()
        except (ElementClickInterceptedException, StaleElementReferenceException):
            self.driver.execute_script("arguments[0].click();", td_el)

        # 5) agora sim espere o overlay sumir por causa do carregamento pós-clique
        wait.until(EC.invisibility_of_element_located(overlay))
        try:
            i = 1
            while True:
                texto_acao = self.driver.find_element(By.XPATH,
                                                      "/html/body/form/div[3]/div/div[2]/div[3]/div[1]/div[2]/div[1]/table/tbody/tr[{}]/td[3]".format(
                                                          i)).text
                texto_acao = texto_acao[0:3]

                if texto_acao == acao:
                    return i
                else:
                    i = i + 1
        except:
            return 0

    def cadastrarNS(self, nom_cliente, serv_exp, mercado, cod_local, dat_compromisso, telefone, celular, abscissa,
                    ordenada):
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'Button13')),
                                             "Botao de cadastrar NS não foi encontrado a tempo.")
        time.sleep(1)
        self.driver.find_element(By.ID, "Button13").click()
        WebDriverWait(self.driver, 10).until(EC.frame_to_be_available_and_switch_to_it(
            (By.XPATH, '/html/body/form/div[1]/table/tbody/tr[2]/td[2]/iframe')), "Iframe não carregou a tempo.")

        # Box cliente
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'RadTextBox8')),
                                             "Box de nome do cliente não foi encontrado a tempo.")
        if (len(nom_cliente) < 5):
            nom_cliente = nom_cliente + " ####"
        self.driver.find_element(By.ID, "RadTextBox8").send_keys(nom_cliente)

        # Box serv_exp
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'ComboBox3_Input')),
                                             "Box de tipo de serv_exp não foi encontrado a tempo.")
        self.driver.find_element(By.ID, "ComboBox3_Input").send_keys(serv_exp)
        time.sleep(1)
        WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'rcbHovered')),
                                             "Dropdown de serv_exp selecionado não foi encontrado a tempo.")
        self.driver.find_element(By.CLASS_NAME, "rcbHovered").click()

        # Box mercado
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'ComboBox1_Input')),
                                             "Box de mercado não foi encontrado a tempo.")
        self.driver.find_element(By.ID, "ComboBox1_Input").send_keys(mercado)
        time.sleep(1)
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, 'rcbHovered')),
                                             "Dropdown de mercado selecionado não foi encontrado a tempo.")
        self.driver.find_element(By.CLASS_NAME, "rcbHovered").click()

        # Box local
        WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, 'ComboBox2_Input')),
                                             "Box de local não foi encontrado a tempo.")
        self.driver.find_element(By.ID, "ComboBox2_Input").send_keys(cod_local)
        time.sleep(1)
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, 'rcbHovered')),
                                             "Dropdown de local selecionado não foi encontrado a tempo.")
        self.driver.find_element(By.CLASS_NAME, "rcbHovered").click()

        # Box dat_compromisso
        if dat_compromisso != "01/01/1900":
            self._set_masked_date("DatePicker1_dateInput", dat_compromisso)

        # Box telefone
        if telefone != "":
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'RadTextBox15')))
            self.driver.find_element(By.ID, "RadTextBox15").send_keys(telefone)

        # Box celular
        if celular != "":
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'RadTextBox17')))
            self.driver.find_element(By.ID, "RadTextBox17").send_keys(celular)

        # Box abscissa
        if abscissa != "":
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'RadTextBox13')))
            self.driver.find_element(By.ID, "RadTextBox13").send_keys(abscissa)

        # Box ordenada
        if ordenada != "":
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'RadTextBox14')))
            self.driver.find_element(By.ID, "RadTextBox14").send_keys(ordenada)
        # Salvando
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'Button11')),
                                             "Botao de salvar não foi encontrado a tempo.")
        self.driver.find_element(By.ID, "Button11").click()
        time.sleep(1)
        """
        WebDriverWait(self.driver, 30).until(
            EC.invisibility_of_element_located((By.ID, '___Form1_AjaxLoadingMainAjaxPanel')),
            "___Form1_AjaxLoadingMainAjaxPanel, timeout de carregamento.")"""
        self.driver.switch_to.default_content()
        WebDriverWait(self.driver, 30).until(
            EC.invisibility_of_element_located((By.ID, '___Form1_AjaxLoadingMainAjaxPanel')),
            "___Form1_AjaxLoadingMainAjaxPanel, timeout de carregamento.")

    def cadastrarServ(self, contrato, cod_servico, num_med, dat_recebimento, base, prazo):
        WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, 'Button16')))
        self.driver.find_element(By.ID, "Button16").click()
        WebDriverWait(self.driver, 10).until(EC.frame_to_be_available_and_switch_to_it(
            (By.XPATH, '/html/body/form/div[1]/table/tbody/tr[2]/td[2]/iframe')), "Iframe não carregou a tempo.")
        WebDriverWait(self.driver, 10).until(
            EC.invisibility_of_element_located((By.CLASS_NAME, 'rwWindowContent rwExternalContent rwLoading')),
            "Timeout de carregamento interno do serviço.")

        # Box contrato
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'ComboBox1_Input')),
                                             "Box de contrato não foi encontrado a tempo.")
        self.driver.find_element(By.ID, "ComboBox1_Input").send_keys(str(contrato))
        time.sleep(1)
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, 'rcbHovered')),
                                             "Dropdown de contrato selecionado não foi encontrado a tempo.")
        self.driver.find_element(By.CLASS_NAME, "rcbHovered").click()
        time.sleep(1)

        # Box cod_serv
        time.sleep(1)
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'ComboBox2_Input')),
                                             "Box do código de serviço não foi encontrado a tempo.")
        self.driver.find_element(By.ID, "ComboBox2_Input").send_keys(cod_servico)
        time.sleep(1)
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, 'rcbItem')),
                                             "Dropdown de serviço selecionado não foi encontrado a tempo.")
        self.driver.find_element(By.CLASS_NAME, "rcbItem").click()

        # Aguardando prazo preencher
        while self.driver.find_element(By.ID, "RadTextBox9").get_attribute("value") == "":
            time.sleep(1)

        # Box med_sap
        self.driver.find_element(By.ID, "Label31").click()
        time.sleep(1)
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'RadTextBox8')),
                                             "Box de número da medida não foi encontrado a tempo.")
        self.driver.find_element(By.ID, "RadTextBox8").send_keys(num_med)

        # Box dat_recebimento
        if dat_recebimento != "01/01/1900":
            self._set_masked_date("DatePicker1_dateInput", dat_recebimento, blur_locator=(By.ID, 'Label32'))
            WebDriverWait(self.driver, 30).until(
                EC.invisibility_of_element_located((By.ID, '___Form1_AjaxLoadingMainAjaxPanel')))

            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'Label32')),
                                                 "Label de data de recebimento não foi encontrado a tempo.")
            self.driver.find_element(By.ID, 'Label32').click()
            time.sleep(1)
            WebDriverWait(self.driver, 30).until(
                EC.invisibility_of_element_located((By.ID, '___Form1_AjaxLoadingMainAjaxPanel')),
                "___Form1_AjaxLoadingMainAjaxPanel, timeout de carregamento.")

        # Box prazo
        if prazo != "":
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'RadTextBox9')),
                                                 "Box da data de recebimento não foi encontrado a tempo.")
            self.driver.find_element(By.ID, "RadTextBox9").click()

            element = self.driver.find_element(By.ID, "RadTextBox9")
            actions = ActionChains(self.driver)
            actions.double_click(element).perform()

            time.sleep(1)
            self.driver.find_element(By.ID, "RadTextBox9").send_keys(prazo)
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'Label34')),
                                                 "Label de data de recebimento não foi encontrado a tempo.")
            self.driver.find_element(By.ID, 'Label34').click()
            time.sleep(1)
            WebDriverWait(self.driver, 30).until(
                EC.invisibility_of_element_located((By.ID, '___Form1_AjaxLoadingMainAjaxPanel')),
                "___Form1_AjaxLoadingMainAjaxPanel, timeout de carregamento.")

        # Box base
        time.sleep(1)
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'ComboBox3_Input')),
                                             "Box de base não foi encontrado a tempo.")
        self.driver.find_element(By.ID, "ComboBox3_Input").send_keys(base)
        time.sleep(1)
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, 'rcbHovered')),
                                             "Dropdown de base selecionada não foi encontrado a tempo.")
        self.driver.find_element(By.CLASS_NAME, "rcbHovered").click()

        # Salvando
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'Button11')))
        self.driver.find_element(By.ID, "Button11").click()
        time.sleep(3)
        self.driver.switch_to.default_content()
        WebDriverWait(self.driver, 30).until(
            EC.invisibility_of_element_located((By.ID, '___Form1_AjaxLoadingMainAjaxPanel')))

    def cadastrarAcoes(self, serv_index=int, cod_acao=str, matricula_resp=str, dat_conclusao=str, prox_acao=str,
                       obs=str):
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH,
                                                                         '/html/body/form/div[3]/div/div[2]/div[2]/div/div[2]/div[1]/table/tbody/tr[{}]/td[4]/span'.format(
                                                                             serv_index))))
        self.driver.execute_script("arguments[0].scrollIntoView({block: 'nearest'});",
                                   self.driver.find_element(By.XPATH,
                                                            "/html/body/form/div[3]/div/div[2]/div[2]/div/div[2]/div[1]/table/tbody/tr[{}]/td[4]/span".format(
                                                                serv_index)))
        time.sleep(1)
        self.driver.find_element(By.XPATH,
                                 "/html/body/form/div[3]/div/div[2]/div[2]/div/div[2]/div[1]/table/tbody/tr[{}]/td[4]/span".format(
                                     serv_index)).click()
        WebDriverWait(self.driver, 30).until(
            EC.invisibility_of_element_located((By.ID, '___Form1_AjaxLoadingMainAjaxPanel')))
        self.driver.execute_script(
            "arguments[0].scrollIntoView({block: 'nearest', behavior: 'instant'});",
            self.driver.find_element(By.XPATH, "/html/body/form/div[3]/div/div[2]/div[3]/div[1]/div[1]/button")
        )
        time.sleep(1)
        WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/form/div[3]/div/div[2]/div[3]/div[1]/div[1]/button')))
        self.driver.find_element(By.XPATH, "/html/body/form/div[3]/div/div[2]/div[3]/div[1]/div[1]/button").click()
        WebDriverWait(self.driver, 10).until(EC.frame_to_be_available_and_switch_to_it(
            (By.XPATH, '/html/body/form/div[1]/table/tbody/tr[2]/td[2]/iframe')))

        # Box cod_acao
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, "ComboBox7_Input")))
        self.driver.find_element(By.ID, "ComboBox7_Input").send_keys(cod_acao)
        time.sleep(1)
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, 'rcbHovered')))
        self.driver.find_element(By.CLASS_NAME, "rcbHovered").click()
        time.sleep(0.5)
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'Label42')))
        self.driver.find_element(By.ID, "Label42").click()

        time.sleep(0.5)
        # Aguardando preencher tprev
        while self.driver.find_element(By.ID, "DatePicker6_dateInput").get_attribute("value") == "":
            time.sleep(1)

        # Box responsavel
        if matricula_resp != "":
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'ComboBox8_Input')))
            self.driver.find_element(By.ID, 'ComboBox8_Input').send_keys(matricula_resp)
            time.sleep(1)
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, 'rcbHovered')))
            self.driver.find_element(By.CLASS_NAME, "rcbHovered").click()
        time.sleep(1)

        # Se necessario, concluir a ação e selecionar a proxima
        if dat_conclusao != "":
            self._set_masked_date("DatePicker7_dateInput", dat_conclusao)
            self._set_masked_date("DatePicker8_dateInput", dat_conclusao)

            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'RadTextBox19')))
            self.driver.find_element(By.ID, "RadTextBox19").send_keys("1000")
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'ComboBox9_Input')))
            self.driver.find_element(By.ID, "ComboBox9_Input").send_keys(prox_acao)
            time.sleep(1)
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, 'rcbHovered')))
            self.driver.find_element(By.CLASS_NAME, "rcbHovered").click()

        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'RadTextBox21')))
        self.driver.find_element(By.ID, "RadTextBox21").send_keys(obs)

        # Salvando
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'Button11')))
        self.driver.find_element(By.ID, "Button11").click()
        time.sleep(1)
        self.driver.switch_to.default_content()
        WebDriverWait(self.driver, 30).until(
            EC.invisibility_of_element_located((By.ID, '___Form1_AjaxLoadingMainAjaxPanel')))

    def trocarRespAcao(self, acao_index=int, matricula_resp=str):
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH,
                                                                         '/html/body/form/div[3]/div/div[2]/div[3]/div[1]/div[2]/div[1]/table/tbody/tr[{}]/td[1]/input'.format(
                                                                             acao_index))))
        self.driver.execute_script("arguments[0].scrollIntoView();", self.driver.find_element(By.XPATH,
                                                                                              "/html/body/form/div[3]/div/div[2]/div[3]/div[1]/div[2]/div[1]/table/tbody/tr[{}]/td[1]/input".format(
                                                                                                  acao_index)))
        time.sleep(1)
        self.driver.find_element(By.XPATH,
                                 "/html/body/form/div[3]/div/div[2]/div[3]/div[1]/div[2]/div[1]/table/tbody/tr[{}]/td[1]/input".format(
                                     acao_index)).click()
        WebDriverWait(self.driver, 10).until(EC.frame_to_be_available_and_switch_to_it(
            (By.XPATH, '/html/body/form/div[1]/table/tbody/tr[2]/td[2]/iframe')))
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'ComboBox8_Input')))

        # Trocando responsavel
        self.driver.find_element(By.ID, "ComboBox8_Input").send_keys(Keys.CONTROL, "A")
        self.driver.find_element(By.ID, "ComboBox8_Input").send_keys(matricula_resp)
        time.sleep(1)
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, 'rcbHovered')))
        self.driver.find_element(By.CLASS_NAME, "rcbHovered").click()
        time.sleep(1)
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'Button11')))
        self.driver.find_element(By.ID, "Button11").click()
        time.sleep(1)
        self.driver.switch_to.default_content()
        WebDriverWait(self.driver, 30).until(
            EC.invisibility_of_element_located((By.ID, '___Form1_AjaxLoadingMainAjaxPanel')))

    def trocarData(self, dat_recebimento, serv_index):
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH,
                                                                         "/html/body/form/div[3]/div/div[2]/div[3]/div/div[2]/div[1]/table/tbody/tr[{}]/td[1]/input".format(
                                                                             serv_index))))
        self.driver.execute_script("arguments[0].scrollIntoView();", self.driver.find_element(By.XPATH,
                                                                                              "/html/body/form/div[3]/div/div[2]/div[3]/div/div[2]/div[1]/table/tbody/tr[{}]/td[1]/input".format(
                                                                                                  serv_index)))
        time.sleep(1)
        self.driver.find_element(By.XPATH,
                                 "/html/body/form/div[3]/div/div[2]/div[3]/div/div[2]/div[1]/table/tbody/tr[{}]/td[1]/input".format(
                                     serv_index)).click()
        time.sleep(1)
        WebDriverWait(self.driver, 10).until(EC.frame_to_be_available_and_switch_to_it(
            (By.XPATH, '/html/body/form/div[1]/table/tbody/tr[2]/td[2]/iframe')))
        time.sleep(1)
        WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/form/div[3]/div/div[1]/div/div/div/div[12]/div/div')))
        self.driver.find_element(By.XPATH, "/html/body/form/div[3]/div/div[1]/div/div/div/div[12]/div/div").click()
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(
            (By.XPATH, '/html/body/form/div[3]/div/div[1]/div/div/div/div[12]/div/div/input[1]')))
        self.driver.find_element(By.XPATH,
                                 "/html/body/form/div[3]/div/div[1]/div/div/div/div[12]/div/div/input[1]").send_keys(
            dat_recebimento)
        time.sleep(1)
        WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/form/div[3]/div/div[1]/div/div/div/div[14]/span[2]')))
        self.driver.find_element(By.XPATH, "/html/body/form/div[3]/div/div[1]/div/div/div/div[14]/span[2]").click()
        time.sleep(1)
        WebDriverWait(self.driver, 30).until(
            EC.invisibility_of_element_located((By.ID, '___Form1_AjaxLoadingMainAjaxPanel')),
            "___Form1_AjaxLoadingMainAjaxPanel, timeout de carregamento.")
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(
            (By.XPATH, '/html/body/form/div[3]/div/div[1]/div/div/div/div[14]/span[2]/input[1]')))
        time.sleep(1)

        WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/form/div[3]/div/div[1]/div/div/div/div[4]/button')))
        self.driver.find_element(By.XPATH, "/html/body/form/div[3]/div/div[1]/div/div/div/div[4]/button").click()
        time.sleep(1)
        self.driver.switch_to.default_content()
        WebDriverWait(self.driver, 30).until(
            EC.invisibility_of_element_located((By.ID, '___Form1_AjaxLoadingMainAjaxPanel')))
        time.sleep(1)

    def concluir_acao(self, acao_index, matricula_resp, dat_ini, dat_conc, us_prj, us_top, us_geo, us_int, prox_acao,
                      obs):
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH,
                                                                         f"/html/body/form/div[3]/div/div[2]/div[4]/div[1]/div[2]/div[1]/table/tbody/tr[{acao_index}]/td[1]/input")))
        self.driver.execute_script("arguments[0].scrollIntoView();", self.driver.find_element(By.XPATH,
                                                                                              f"/html/body/form/div[3]/div/div[2]/div[4]/div[1]/div[2]/div[1]/table/tbody/tr[{acao_index}]/td[1]/input"))
        time.sleep(1)
        self.driver.find_element(By.XPATH,
                                 f"/html/body/form/div[3]/div/div[2]/div[4]/div[1]/div[2]/div[1]/table/tbody/tr[{acao_index}]/td[1]/input").click()
        time.sleep(1)
        WebDriverWait(self.driver, 10).until(EC.frame_to_be_available_and_switch_to_it(
            (By.XPATH, '/html/body/form/div[1]/table/tbody/tr[2]/td[2]/iframe')))
        time.sleep(1)

        # Box de Responsavel
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'ComboBox8_Input')))
        self.driver.find_element(By.ID, 'ComboBox8_Input').send_keys(Keys.CONTROL, 'A')
        time.sleep(0.5)
        self.driver.find_element(By.ID, 'ComboBox8_Input').send_keys(matricula_resp)
        time.sleep(1)
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, 'rcbHovered')))
        self.driver.find_element(By.CLASS_NAME, 'rcbHovered').click()
        time.sleep(1)

        # Box data de inicio
        if dat_ini != "":
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'DatePicker7_dateInput')))
            self.driver.find_element(By.ID, 'DatePicker7_dateInput').send_keys(Keys.CONTROL, 'A')
            time.sleep(0.5)
            self.driver.find_element(By.ID, 'DatePicker7_dateInput').send_keys(dat_ini)

        # Box data de conclusao
        if dat_conc != "":
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'DatePicker8_dateInput')))
            self.driver.find_element(By.ID, 'DatePicker8_dateInput').send_keys(Keys.CONTROL, 'A')
            time.sleep(0.5)
            self.driver.find_element(By.ID, 'DatePicker8_dateInput').send_keys(dat_conc)

        # Box %exec
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'RadTextBox19')))
        self.driver.find_element(By.ID, 'RadTextBox19').send_keys(Keys.CONTROL, 'A')
        time.sleep(0.5)
        self.driver.find_element(By.ID, 'RadTextBox19').send_keys("100")

        # Box US_PRJ
        if us_prj != '':
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'RadTextBox1')))
            self.driver.find_element(By.ID, 'RadTextBox1').send_keys(Keys.CONTROL, 'A')
            self.driver.find_element(By.ID, 'RadTextBox1').send_keys(us_prj)

        # Box US_TOP
        if us_top != '':
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'RadTextBox2')))
            self.driver.find_element(By.ID, 'RadTextBox2').send_keys(Keys.CONTROL, 'A')
            self.driver.find_element(By.ID, 'RadTextBox2').send_keys(us_top)

        # Box US_GEO
        if us_geo != '':
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'RadTextBox3')))
            self.driver.find_element(By.ID, 'RadTextBox3').send_keys(Keys.CONTROL, 'A')
            self.driver.find_element(By.ID, 'RadTextBox3').send_keys(us_geo)

        # Box US_INT
        if us_int != '':
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'RadTextBox4')))
            self.driver.find_element(By.ID, 'RadTextBox4').send_keys(Keys.CONTROL, 'A')
            self.driver.find_element(By.ID, 'RadTextBox4').send_keys(us_int)

        # Box proxacao
        if prox_acao != '':
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'ComboBox9_Input')))
            self.driver.find_element(By.ID, 'ComboBox9_Input').send_keys(Keys.CONTROL, 'A')
            time.sleep(0.5)
            self.driver.find_element(By.ID, 'ComboBox9_Input').send_keys(prox_acao)
            time.sleep(1)
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, 'rcbHovered')))
            self.driver.find_element(By.CLASS_NAME, 'rcbHovered').click()
            time.sleep(1)

        if obs != '':
            WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'RadTextBox21')))
            self.driver.find_element(By.ID, 'RadTextBox21').send_keys(Keys.CONTROL, 'A')
            self.driver.find_element(By.ID, 'RadTextBox21').send_keys(obs)

        time.sleep(1)
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'Button11')))
        self.driver.find_element(By.ID, "Button11").click()
        time.sleep(1)
        WebDriverWait(self.driver, 30).until(
            EC.invisibility_of_element_located((By.ID, '___Form1_AjaxLoadingMainAjaxPanel')))
        self.driver.switch_to.default_content()
        WebDriverWait(self.driver, 30).until(
            EC.invisibility_of_element_located((By.ID, '___Form1_AjaxLoadingMainAjaxPanel')))
        time.sleep(1)

    def lancarRequisicao(self, acao_index, req_num, item, qtd_us):
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH,
                                                                         f"/html/body/form/div[3]/div/div[2]/div[4]/div[1]/div[2]/div[1]/table/tbody/tr[{acao_index}]/td[3]/span")))
        self.driver.execute_script("arguments[0].scrollIntoView();", self.driver.find_element(By.XPATH,
                                                                                              f"/html/body/form/div[3]/div/div[2]/div[4]/div[1]/div[2]/div[1]/table/tbody/tr[{acao_index}]/td[3]/span"))
        time.sleep(1)
        self.driver.find_element(By.XPATH,
                                 f"/html/body/form/div[3]/div/div[2]/div[4]/div[1]/div[2]/div[1]/table/tbody/tr[{acao_index}]/td[3]/span").click()
        WebDriverWait(self.driver, 30).until(
            EC.invisibility_of_element_located((By.ID, '___Form1_AjaxLoadingMainAjaxPanel')),
            "___Form1_AjaxLoadingMainAjaxPanel, timeout de carregamento.")
        self.driver.execute_script("arguments[0].scrollIntoView();", self.driver.find_element(By.ID, 'Button19'))
        self.driver.find_element(By.ID, 'Button19').click()
        WebDriverWait(self.driver, 30).until(
            EC.invisibility_of_element_located((By.ID, '___Form1_AjaxLoadingMainAjaxPanel')),
            "___Form1_AjaxLoadingMainAjaxPanel, timeout de carregamento.")
        WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.ID, 'Grid4_ctl00_ctl02_ctl02_TB_GridColumn360')))
        self.driver.find_element(By.ID, 'Grid4_ctl00_ctl02_ctl02_TB_GridColumn360').send_keys(req_num)
        self.driver.find_element(By.ID, 'Grid4_ctl00_ctl02_ctl02_TB_GridColumn361').send_keys(item)
        self.driver.find_element(By.ID, 'Grid4_ctl00_ctl02_ctl02_TB_GridColumn362').send_keys(qtd_us)
        time.sleep(0.5)
        self.driver.find_element(By.ID, 'Grid4_ctl00_ctl02_ctl02_PerformInsertButton').click()
        time.sleep(1)
        WebDriverWait(self.driver, 30).until(
            EC.invisibility_of_element_located((By.ID, '___Form1_AjaxLoadingMainAjaxPanel')))
        time.sleep(1)

    def trocarObs(self, acao_index, text_obs):
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH,
                                                                         f"/html/body/form/div[3]/div/div[2]/div[4]/div[1]/div[2]/div[1]/table/tbody/tr[{acao_index}]/td[1]/input")))
        self.driver.execute_script("arguments[0].scrollIntoView();", self.driver.find_element(By.XPATH,
                                                                                              f"/html/body/form/div[3]/div/div[2]/div[4]/div[1]/div[2]/div[1]/table/tbody/tr[{acao_index}]/td[1]/input"))
        time.sleep(1)
        self.driver.find_element(By.XPATH,
                                 f"/html/body/form/div[3]/div/div[2]/div[4]/div[1]/div[2]/div[1]/table/tbody/tr[{acao_index}]/td[1]/input").click()
        WebDriverWait(self.driver, 10).until(EC.frame_to_be_available_and_switch_to_it(
            (By.XPATH, '/html/body/form/div[1]/table/tbody/tr[2]/td[2]/iframe')))
        time.sleep(1)
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'RadTextBox21')))
        self.driver.find_element(By.ID, 'RadTextBox21').send_keys(Keys.CONTROL, 'A')
        self.driver.find_element(By.ID, 'RadTextBox21').send_keys(text_obs)

        time.sleep(1)
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID, 'Button11')))
        self.driver.find_element(By.ID, "Button11").click()
        time.sleep(1)
        WebDriverWait(self.driver, 30).until(
            EC.invisibility_of_element_located((By.ID, '___Form1_AjaxLoadingMainAjaxPanel')))
        self.driver.switch_to.default_content()
        WebDriverWait(self.driver, 30).until(
            EC.invisibility_of_element_located((By.ID, '___Form1_AjaxLoadingMainAjaxPanel')))
        time.sleep(1)