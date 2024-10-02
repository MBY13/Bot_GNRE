from unidecode import unidecode
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

    def normalize_text(self, text):
        """Remove acentos e troca 'ç' por 'c'."""
        return unidecode(text)

    def extract_table_data(self, estado_nome):
        """Método para extrair dados da tabela com cabeçalhos das colunas e adicionar o estado como coluna."""
        try:
            table_data = self.driver.execute_script("""
                let rows = document.querySelectorAll('#tblContent tr');
                let data = [];
                let headers = [];
                
                // Extrair cabeçalhos da tabela
                let headerRow = rows[0];
                headerRow.querySelectorAll('th').forEach(header => {
                    headers.push(header.innerText.trim());
                });
                
                // Adiciona "Estado" como coluna
                headers.push('Estado');
                
                // Extrair dados das linhas da tabela
                rows.forEach((row, rowIndex) => {
                    if (rowIndex > 0) {  // Ignora a linha de cabeçalhos
                        let cells = row.querySelectorAll('td');
                        let rowData = {};
                        cells.forEach((cell, cellIndex) => {
                            rowData[headers[cellIndex]] = cell.innerText.trim();  // Associa o valor à coluna correspondente
                        });
                        // Adiciona o estado à linha
                        rowData['Estado'] = arguments[0];
                        
                        // Verifica se a linha contém dados relevantes
                        if (Object.keys(rowData).length > 0 && !rowData[headers[0]].includes('Registros') && rowData[headers[0]] !== '') {
                            data.push(rowData);
                        }
                    }
                });
                return data;
            """, estado_nome)  # Passa o nome do estado como argumento para o script JavaScript
            
            # Normaliza os cabeçalhos e os dados extraídos
            normalized_headers = [self.normalize_text(header) for header in table_data[0].keys()]
            normalized_data = [
                {self.normalize_text(k): self.normalize_text(v) if isinstance(v, str) else v for k, v in row.items()}
                for row in table_data
            ]
            
            return normalized_data
        except Exception as e:
            print(f"Erro ao extrair dados via JavaScript: {e}")
            return []

# O restante do código GNREBot permanece o mesmo

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
            # "Receitas": {},
            "Detalhamento das Receitas": {},
            # "Produtos": {},
            "Documentos de Origem": {},
            "Campos Adicionais": {}
        }
        self.extractor = TabelaExtractor(self.driver)

    def acessar_pagina(self, url):
        self.driver.get(url)

    def clicar_e_processar_links(self):
        links_text = [
            # 'Receitas',
            'Detalhamento das Receitas',
            # 'Produtos',
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
            table_data = self.extractor.extract_table_data(estado_nome)  # Passa o nome do estado
            if table_data:
                # Verifica se a categoria existe e é uma lista, caso contrário, cria uma lista vazia
                if not isinstance(self.data[categoria], list):
                    self.data[categoria] = []
                # Adiciona os dados extraídos na lista
                self.data[categoria].extend(table_data)
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