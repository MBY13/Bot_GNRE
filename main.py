from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException, ElementClickInterceptedException
import time
import json

class TabelaExtractor:
    def __init__(self, driver):
        self.driver = driver

    def extract_table_data(self):
        """Método para extrair dados da tabela usando JavaScript e filtrar linhas não relevantes."""
        try:
            table_data = self.driver.execute_script("""
                let rows = document.querySelectorAll('#tblContent tr');
                let data = [];
                rows.forEach(row => {
                    let cells = row.querySelectorAll('td');
                    let rowData = [];
                    cells.forEach(cell => {
                        rowData.push(cell.innerText.trim());
                    });
                    // Verifica se a linha contém dados relevantes
                    if (rowData.length > 0 && !rowData[0].includes('Registros') && rowData[0] !== '') {
                        data.push(rowData);
                    }
                });
                return data;
            """)
            return table_data
        except Exception as e:
            print(f"Erro ao extrair dados via JavaScript: {e}")
            return []

class JsonExporter:
    @staticmethod
    def save_to_json(data, filename):
        """Método para salvar os dados em um arquivo .json."""
        try:
            with open(filename, 'w', encoding='utf-8') as file:
                json.dump(data, file, indent=4, ensure_ascii=False)
            print(f"Dados salvos em {filename}")
        except Exception as e:
            print(f"Erro ao salvar os dados em {filename}: {e}")

class GNREBot:
    def __init__(self):
        self.service = Service()
        self.options = webdriver.ChromeOptions()
        self.driver = webdriver.Chrome(service=self.service, options=self.options)
        self.data = {
            "Receitas": {},
            "Detalhamento das Receitas": {},
            "Produtos": {},
            "Documentos de Origem": {},
            "Campos Adicionais": {}
        }
        self.extractor = TabelaExtractor(self.driver)

    def acessar_pagina(self, url):
        self.driver.get(url)

    def clicar_e_processar_links(self):
        links_text = [
            'Receitas',
            'Detalhamento das Receitas',
            'Produtos',
            'Documentos de Origem',
            'Campos Adicionais'
        ]

        try:
            for link in self.driver.find_elements(By.TAG_NAME, 'a'):
                text = link.text.strip()
                if text in links_text:
                    try:
                        self.fechar_dialog_se_existir()  
                        self.switch_to_default_content()
                        self.fechar_dialog_se_existir()
                        link.click()
                        self.processar_estados(text)
                        self.clicar_botao_cancelar_e_fechar()
                    except Exception as e:
                        print(f"Erro ao clicar no link '{text}': {e}")
        except NoSuchElementException as e:
            print(f"Links não encontrados: {e}")

    def processar_estados(self, categoria):
        try:
            select_element = Select(self.driver.find_element(By.ID, 'cmbUF'))
            for Estado in select_element.options:
                estado_nome = Estado.text.strip()
                if estado_nome != 'Todas as Receitas':
                    self._processar_estado(Estado, estado_nome, categoria)
        except NoSuchElementException as e:
            print(f"Elemento 'cmbUF' não encontrado: {e}")

    def _processar_estado(self, Estado, estado_nome, categoria):
        try:
            time.sleep(5)
            Estado.click()
            self._clicar_botao_imprimir()
            self._extrair_e_salvar_dados(estado_nome, categoria)
        except Exception as e:
            print(f"Erro ao processar estado {estado_nome} na categoria {categoria}: {e}")

    def _clicar_botao_imprimir(self):
        try:
            button_imprimir = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='Visualizar impressão']]"))
            )
            self.driver.execute_script("arguments[0].click();", button_imprimir)
        except (TimeoutException, ElementClickInterceptedException) as e:
            print(f"Erro ao clicar no botão 'Visualizar impressão': {e}")

    def _extrair_e_salvar_dados(self, estado_nome, categoria):
        try:
            table = WebDriverWait(self.driver, 20).until(
                EC.visibility_of_element_located((By.ID, 'tblContent'))
            )
            print(f"Tabela encontrada para o estado {estado_nome} na categoria {categoria}")
            time.sleep(5)
            table_data = self.extractor.extract_table_data()
            if table_data:
                self.data[categoria][estado_nome] = table_data
            else:
                print(f"Nenhum dado relevante encontrado na tabela para o estado {estado_nome} na categoria {categoria}.")
            time.sleep(2)
        except TimeoutException as e:
            print(f"Tabela não carregou a tempo: {e}")
        except NoSuchElementException as e:
            print(f"Tabela não encontrada: {e}")

    def clicar_botao_cancelar_e_fechar(self):
        try:
            # Clicar no botão "Cancelar"
            cancelar_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()=' Cancelar ']]"))
            )
            self.driver.execute_script("arguments[0].scrollIntoView(true);", cancelar_button)  # Certifica-se de que o botão está visível
            self.driver.execute_script("arguments[0].click();", cancelar_button)
            print("Botão 'Cancelar' clicado com sucesso.")
            time.sleep(2)

            # Clicar no botão "Fechar"
            fechar_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='Fechar']]"))
            )
            self.driver.execute_script("arguments[0].scrollIntoView(true);", fechar_button)  # Certifica-se de que o botão está visível
            self.driver.execute_script("arguments[0].click();", fechar_button)
            print("Botão 'Fechar' clicado com sucesso.")
            time.sleep(2)

        except TimeoutException as e:
            print(f"Erro ao clicar nos botões 'Cancelar' ou 'Fechar': {e}")
        except NoSuchElementException as e:
            print(f"Botões 'Cancelar' ou 'Fechar' não encontrados: {e}")
        except ElementClickInterceptedException as e:
            print(f"Elemento bloqueado ao tentar clicar: {e}")

    def salvar_dados(self, filename):
        JsonExporter.save_to_json(self.data, filename)

    def fechar(self):
        self.driver.quit()
            
    def fechar_dialog_se_existir(self):
        try:
            # Tenta fechar qualquer diálogo de sobreposição
            fechar_button = WebDriverWait(self.driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='Fechar']]"))
            )
            self.driver.execute_script("arguments[0].click();", fechar_button)
            print("Janela de diálogo fechada com sucesso.")
            time.sleep(1)
        except (TimeoutException, NoSuchElementException):
            print("Nenhum diálogo de sobreposição detectado.")

    def switch_to_iframe_if_exists(self):
        try:
            iframe = WebDriverWait(self.driver, 5).until(
                EC.presence_of_element_located((By.ID, "frameImp"))
            )
            self.driver.switch_to.frame(iframe)
            print("Mudança para o iframe bem-sucedida.")
        except (TimeoutException, NoSuchElementException):
            print("Nenhum iframe detectado, permanecendo no conteúdo principal.")

    def switch_to_default_content(self):
        self.driver.switch_to.default_content()
        print("Voltando para o conteúdo principal.")

if __name__ == "__main__":
    bot = GNREBot()
    bot.acessar_pagina('https://www.gnre.pe.gov.br:444/gnre/portal/consultarTabelas.jsp')
    bot.clicar_e_processar_links()
    bot.salvar_dados('tabela_dados.json')
    bot.fechar()