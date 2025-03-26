import pandas as pd, re
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

linhaInicial = 1
quantidadeLinhasParaLer = 149   
colunasComDados = 'A:K'

driver_path = r'C:\Users\rlvju\Área de Trabalho\automacaoDeCilindros\chromedriver-win64\chromedriver.exe'
caminhoPlanilhCilindros = r'C:\Users\rlvju\Área de Trabalho\automacaoDeCilindros\Pedidos.xlsx'
caminhoPlanilhaColetas = r'C:\Users\rlvju\Área de Trabalho\automacaoDeCilindros\Coletas.xlsx'

df = pd.read_excel(caminhoPlanilhCilindros, header=0, skiprows=range(1, linhaInicial), nrows=quantidadeLinhasParaLer, usecols=colunasComDados)
df_coletas = pd.DataFrame(columns=['CPF', 'Código de Coleta', 'Código de Rastreio'])

options = webdriver.ChromeOptions()
options.add_experimental_option("debuggerAddress", "localhost:9222")

service = Service(driver_path)
driver = webdriver.Chrome(service=service, options=options)

def extrair_numero_complemento(endereco):
    match = re.search(r',\s*(\d*)\s*(.*)', endereco)
    if match:
        numero = match.group(1)
        complemento = match.group(2)
        if not numero:
            complemento = match.group(0).split(',', 1)[1].strip()
            numero = None
    else:
        numero = None
        complemento = None
    return numero, complemento

def verificar_coleta():
    try:
        WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.NAME, 'remIdentificacao')))
        return True
    except:
        return False
    
def copiar_codigo(xpath):
    codigo_elemento = driver.find_element(By.XPATH, xpath)
    return codigo_elemento.text.strip()

for _, linha in df.iterrows():
    if linha.isnull().any():
        break

    nome = linha['NOME']
    cpf = linha['CPF']
    rua = linha['RUA']
    bairro = linha['BAIRRO']
    cep = linha['CEP']
    cidade = linha['CIDADE']
    uf = linha['UF']
    ddd = ['11']
    telefone = ['30039030']
    email = linha['EMAIL']
    pedido = linha['PEDIDO']

    numero, complemento = extrair_numero_complemento(rua)

    selecaoServico = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.NAME, 'servico'))
    )
    selectServico = Select(selecaoServico)
    selectServico.select_by_index(2)

    driver.find_element(By.NAME, 'cepOrigem').send_keys(cep)

    botaoOk = driver.find_element(By.XPATH, "//input[@value='Ok']")
    botaoOk.click()

    if not verificar_coleta():
        df_coletas = df_coletas._append({
            'CPF': cpf,
            'Código de Coleta': 0,
            'Código de Rastreio': 0
        }, ignore_index=True)
        df_coletas.to_excel(caminhoPlanilhaColetas, index=False)
        driver.find_element(By.NAME, 'cepOrigem').clear()
        continue
    else:
        selecaoCartaoDePostagem = driver.find_element(By.NAME, 'cartao')
        selectcCartaoDePostagem = Select(selecaoCartaoDePostagem)
        selectcCartaoDePostagem.select_by_index(1)

        driver.find_element(By.NAME, 'remIdentificacao').send_keys(cpf)
        driver.find_element(By.NAME, 'remNome').send_keys(nome)
        driver.find_element(By.NAME, 'remBairro').send_keys(bairro)
        driver.find_element(By.NAME, 'remLogradouro').send_keys(rua)

        if numero:
            driver.find_element(By.NAME, 'remNumero').send_keys(numero)
        if complemento:
            driver.find_element(By.NAME, 'remComplemento').send_keys(complemento)
        
        driver.find_element(By.NAME, 'remDDD').send_keys(ddd)
        driver.find_element(By.NAME, 'remTelefone').send_keys(telefone)
        driver.find_element(By.NAME, 'remEmail').send_keys(email)
        driver.find_element(By.NAME, 'objControleCliente').send_keys(pedido)
        driver.find_element(By.NAME, 'objDescricao').send_keys('Cilindro Refil de CO2 60 Litros Sodastream')
        driver.find_element(By.NAME, 'objValorDeclarado1').send_keys('320')
        driver.find_element(By.NAME, 'objValorDeclarado2').send_keys('00')
        driver.find_element(By.NAME, 'destIdentificacao').send_keys('01490698006689')

        botaoEnviar = driver.find_element(By.XPATH, "//input[@value='Enviar Solicitação']")
        botaoEnviar.click()

        alert = WebDriverWait(driver, 10).until(EC.alert_is_present())
        alert.accept()

        WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//b[contains(text(), 'Pedido de coleta realizado com sucesso')]"))
        )
        codigo_coleta = copiar_codigo("//b[string-length(text()) = 9]")
        codigo_rastreio = copiar_codigo("//td[b[contains(text(), 'QS')]]/b")

        df_coletas = df_coletas._append({
            'CPF': cpf,
            'Código de Coleta': codigo_coleta,
            'Código de Rastreio': codigo_rastreio
        }, ignore_index=True)

        df_coletas.to_excel(caminhoPlanilhaColetas, index=False)

        botaoNovaColeta = driver.find_element(By.XPATH, "//input[@value='Nova solicitação']")
        botaoNovaColeta.click()