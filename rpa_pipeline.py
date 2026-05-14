from selenium.webdriver.support import expected_conditions as ec
from SolutionPacket.Solution_telegram import TelegramSend
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from SolutionPacket.Solution_email import Smtp
from SolutionPacket.Solution_bank import Bank
from selenium.webdriver.common.by import By
from selenium import webdriver
from datetime import datetime
from dotenv import load_dotenv
from time import sleep
import pandas as pd
import openpyxl
import logging
import shutil
import pyotp
import xlrd
import calendar
import locale
import time
import os

# ── Logging ──────────────────────────────────────────────────────────────────
os.makedirs("logs", exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(f"logs/garantia_{datetime.today().strftime('%Y%m%d')}.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)
log = logging.getLogger(__name__)

# ── Env ───────────────────────────────────────────────────────────────────────
load_dotenv()

TOKEN          = os.getenv("TELEGRAM_TOKEN")
CHAT_ID        = os.getenv("TELEGRAM_CHAT_ID")
CHAT_ID_RPA    = os.getenv("TELEGRAM_CHAT_ID_RPA")
TOTP_KEY       = os.getenv("TOTP_KEY")
PASTA_BASE     = os.getenv("PASTA_BASE")
PASTA_VALIDADA = os.getenv("PASTA_VALIDADA")
PASTA_QUALIDADE= os.getenv("PASTA_QUALIDADE")

# ── Locale / Datas ────────────────────────────────────────────────────────────
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
ano      = datetime.today().strftime('%Y')
mes      = int(datetime.today().strftime('%m'))
nome_mes = calendar.month_name[mes]

# ── Telegram ──────────────────────────────────────────────────────────────────
tele_cliente = TelegramSend('Qualidade_Garantia: ')
tele         = TelegramSend('Qualidade_Garantia: ')
email_geral  = Smtp('Garantia')
assunto      = 'Acompanhamento Qualidade Garantia'

# ── Pastas por ano/mês ────────────────────────────────────────────────────────
meses = [
    "01 - Janeiro", "02 - Fevereiro", "03 - Março", "04 - Abril",
    "05 - Maio", "06 - Junho", "07 - Julho", "08 - Agosto",
    "09 - Setembro", "10 - Outubro", "11 - Novembro", "12 - Dezembro"
]
if not os.path.isdir(PASTA_BASE):
    os.makedirs(PASTA_BASE, exist_ok=True)
    for mes_item in meses:
        os.makedirs(os.path.join(PASTA_BASE, mes_item), exist_ok=True)


# ── Helpers ───────────────────────────────────────────────────────────────────
def parse_data(valor):
    if not valor or str(valor).strip() in ('None', '', 'nan'):
        return datetime.strptime('01/01/2001', '%d/%m/%Y').date()
    for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y'):
        try:
            return datetime.strptime(str(valor).strip(), fmt).date()
        except:
            pass
    return datetime.strptime('01/01/2001', '%d/%m/%Y').date()


def aguardar_download_concluir(diretorio, timeout=60):
    inicio = time.time()
    while time.time() - inicio < timeout:
        pendentes = [f for f in os.listdir(diretorio) if f.endswith('.crdownload') or f.endswith('.tmp')]
        if not pendentes:
            return True
        sleep(1)
    log.warning("Timeout aguardando download")
    return False


def obter_celulas_por_status(navegador, status):
    celulas = navegador.find_elements(By.XPATH,
        '//*[@id="header-table-page"]/groupui-table-cell[.//groupui-tag]')
    indices = []
    for idx, celula in enumerate(celulas):
        try:
            tag = celula.find_element(By.XPATH, './/groupui-tag')
            if tag.text.strip().lower() == status.lower():
                indices.append(idx)
        except:
            pass
    return indices


def processar_item(navegador, wait_sdd, diretorio_download, idx):
    try:
        celulas = navegador.find_elements(By.XPATH,
            '//*[@id="header-table-page"]/groupui-table-cell[.//groupui-tag]')
        celula = celulas[idx]
        navegador.execute_script("arguments[0].scrollIntoView(true);", celula)
        sleep(1)
        celula.click()
        sleep(3)

        btn_documents = wait_sdd.until(
            ec.element_to_be_clickable((By.XPATH,
                '//*[@id="content-unit-header"]/div/groupui-accordion/span/groupui-headline'))
        )
        btn_documents.click()
        sleep(2)

        links_pendentes = navegador.find_elements(By.XPATH,
            '//*[@id="content-unit-header"]/div/groupui-accordion/groupui-list/ul/li/div/groupui-link[@variant="primary"]')

        if not links_pendentes:
            log.info(f"idx {idx+1}: nenhum link pendente")
            return False

        link  = links_pendentes[0]
        titulo = link.get_attribute('title') or f'arquivo_{idx+1}'
        arquivos_antes = set(os.listdir(diretorio_download))
        navegador.execute_script("arguments[0].scrollIntoView(true);", link)
        sleep(0.5)
        link.click()
        log.info(f"idx {idx+1}: baixando '{titulo}'")

        aguardar_download_concluir(diretorio_download, timeout=60)
        sleep(3)

        novos = set(os.listdir(diretorio_download)) - arquivos_antes
        log.info(f"idx {idx+1}: baixados {novos}")
        return True

    except Exception as erro:
        log.error(f"Erro ao processar idx {idx+1}: {erro}")
        return False


# ── Extração Excel → SQL Server ───────────────────────────────────────────────
def extrair_2(arquivo_, chave):
    arquivo_xlsx = arquivo_.replace('.xls', '_temp.xlsx') if arquivo_.endswith('.xls') else arquivo_
    if arquivo_ != arquivo_xlsx:
        shutil.copy2(arquivo_, arquivo_xlsx)

    try:
        df = pd.read_excel(arquivo_xlsx, engine='openpyxl')
    except Exception:
        df = pd.read_excel(arquivo_, engine='xlrd')

    df_sem_nan = df.fillna('None')
    indice_pagina = df_sem_nan[df_sem_nan.astype(str).apply(
        lambda x: 'Número de Operação' in ' '.join(x), axis=1)].index[0]
    df_colunas = df_sem_nan.iloc[indice_pagina + 1:]

    mes_criacao = ''
    try:
        wb = openpyxl.load_workbook(arquivo_xlsx, data_only=True)
        sheet = wb.active
        for row in sheet.iter_rows():
            if row[0].value == chave:
                mes_criacao = row[1].value
                break
    except Exception:
        try:
            wb = xlrd.open_workbook(arquivo_)
            sheet = wb.sheet_by_index(0)
            for row_idx in range(sheet.nrows):
                if sheet.cell_value(row_idx, 0) == chave:
                    mes_criacao = sheet.cell_value(row_idx, 1)
        except Exception as e:
            log.error(f"Erro ao ler mes_criacao: {e}")

    if arquivo_ != arquivo_xlsx and os.path.exists(arquivo_xlsx):
        os.remove(arquivo_xlsx)

    data_execucao = datetime.today().strftime('%d/%m/%Y')

    for _, row in df_colunas.iterrows():
        try:
            insert = f"""INSERT INTO [dbo].[Garantia_Volks]
               ([Num_Operacao],[Cod_Peca_Montada_Original],[Cod_Peca_Montada_Nao_Ordenado],
                [Descricao_Pecas],[Op_mao_de_obra],[Descricao],[Cod_Dano],[KDNR],
                [Num_Peca_Modulo],[Cod_Reparacao],[Qtd_Pecas_Desmontadas],[Qtd_Pecas_Montadas],
                [Fator_Tecnico],[VIN],[Modelo_Vendas],[Data_Producao],[Data_Recep_Reparo],
                [Data_Reg_Inst],[Quilometros_Milhas],[Fator_Custos],[Chave_Estorno],
                [Tipo_Estorno_Garantia],[Unidades_Tempo],[Custos_Material],[Custos_mao_de_obra],
                [Custos_adicionais],[Valor_Debitado],[Valor_Reclamacao_Material],[Preco_Compra],
                [Num_Concessionaria],[Num_Reclamacao],[Tipo_Reclamacao],
                [Chave_identificacao_reclamacao],[mes_criacao],[data_execucao])
            VALUES
               ('{row[0]}','{row[1]}','{row[2]}','{row[3]}','{row[4]}','{row[5]}',
                '{str(row[6]).replace("-","")}','{str(row[7]).replace("-","")}',
                '{str(row[8]).replace("-","")}','{str(row[9]).replace("-","")}',
                '{str(row[10]).replace("-","")}','{str(row[11]).replace("-","")}',
                '{str(row[12]).replace("-","")}','{row[13]}','{row[14]}',
                '{parse_data(row[15])}','{parse_data(row[16])}','{parse_data(row[17])}',
                '{str(row[18]).replace("-","")}','{str(row[19]).replace("-","")}',
                '{row[20]}','{row[21]}','{str(row[22]).replace("-","")}',
                '{float(str(row[23]).replace(",","."))}','{float(str(row[24]).replace(",","."))}',
                '{float(str(row[25]).replace(",","."))}','{float(str(row[26]).replace(",","."))}',
                '{float(str(row[27]).replace(",","."))}','{str(row[28]).replace("-","")}',
                '{row[30]}','{str(row[31]).replace("-","")}','{str(row[32]).replace("-","")}',
                '{row[33]}','{str(mes_criacao)}','{str(data_execucao)}')"""
            cursor.execute(insert)
            cursor.commit()
            log.info("Insert realizado")
        except Exception as er:
            log.error(f"Erro insert: {er}")


# ── Conexões ──────────────────────────────────────────────────────────────────
rpa_cred = Bank('Credenciais')
cursor1  = rpa_cred.bank_connection(
    os.getenv("BANK_CRED_USER"), os.getenv("BANK_CRED_PASS"),
    os.getenv("BANK_CRED_HOST"), os.getenv("BANK_CRED_DB"))

rpa_29 = Bank('rpa')
cursor = rpa_29.bank_connection(
    os.getenv("BANK_RPA_USER"), os.getenv("BANK_RPA_PASS"),
    os.getenv("BANK_RPA_HOST"), os.getenv("BANK_RPA_DB"))

pasta = "Arquivos"
sleep(3)

# ── Main ──────────────────────────────────────────────────────────────────────
try:
    diretorio_atual   = os.getcwd()
    diretorio_download = os.path.join(diretorio_atual, pasta)
    options = Options()
    options.add_experimental_option("prefs", {
        "download.default_directory": diretorio_download,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "plugins.always_open_pdf_externally": True,
        "plugins.plugins_disabled": ["Chrome PDF Viewer"]
    })

    navegador = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    navegador.maximize_window()
    navegador.implicitly_wait(4)

    logins = cursor1.execute(
        "SELECT usuario FROM [datadriven_paranoa].[dbo].[dw_view_acessos_rpa] "
        "WHERE nome LIKE '%Garantia Volkswagen%' AND cliente = 'Qualidade'").fetchall()[0]
    senhas = cursor1.execute(
        "SELECT senha FROM [datadriven_paranoa].[dbo].[dw_view_acessos_rpa] "
        "WHERE nome LIKE '%Volkswagen%' AND cliente = 'Qualidade' "
        "AND senha IS NOT NULL AND LTRIM(RTRIM(senha)) <> ''").fetchall()[0]
    login = senha = ''
    for login_, senha_ in zip(logins, senhas):
        login = login_.strip()
        senha = senha_.strip()
    if not senha:
        tele_cliente.telegram_bot("Qualidade Garantia - Senha expirada. Favor atualizar", TOKEN, CHAT_ID_RPA)
        navegador.close()

    wait = WebDriverWait(navegador, 10)

    def get_totp(secret):
        return pyotp.TOTP(secret).now()

    cod = get_totp(TOTP_KEY)

    try:
        navegador.get(
            r"https://lso.volkswagen.de/lona/eai/b2b9ext/pages/login.html"
            r"?url=/one-kbp/content/pt/kbp_private/kbpprivate_systems/kbpprivate_systempage_16248.jsp%3Fbrand%3Dnull")
        navegador.find_element(By.ID, 'contentForm:profileIdInput').send_keys(login)
        sleep(2)
        navegador.find_element(By.ID, 'contentForm:passwordInput').send_keys(senha)
        sleep(2)
        navegador.find_element(By.ID, 'contentForm:passwordLoginAction').click()
        sleep(2)
    except Exception as erro_gr:
        log.error(f"Erro login: {erro_gr}")
        screenshot_path = os.path.join(os.getcwd(), "erro_Garantia.png")
        navegador.save_screenshot(screenshot_path)
        tele_cliente.telegram_bot_image('Qualidade Garantia - erro de login', TOKEN, CHAT_ID_RPA, screenshot_path)

    sleep(3)
    navegador.find_element(By.XPATH, '//*[@id="otp"]').send_keys(cod)
    sleep(1)
    navegador.find_element(By.XPATH, '//*[@id="panel-login"]/div/div[2]/form/div/button').click()
    sleep(3)
    navegador.find_element(By.XPATH, '//*[@id="opg-nav-list"]/li[5]/a').click()
    sleep(2)
    navegador.find_element(By.XPATH, '//*[@id="system-dropdown"]/div[2]/ul[1]/li[1]/a').click()
    sleep(3)

    # ── SDD Loop ──────────────────────────────────────────────────────────────
    try:
        while True:
            navegador.get("https://lso.volkswagen.de/sdd")
            sleep(5)
            wait_sdd = WebDriverWait(navegador, 20)

            indices_seen = obter_celulas_por_status(navegador, 'seen')
            processou_seen = False
            for idx in indices_seen:
                if processar_item(navegador, wait_sdd, diretorio_download, idx):
                    processou_seen = True
                    break
            if processou_seen:
                continue

            indices_new = obter_celulas_por_status(navegador, 'new')
            if not indices_new:
                log.info("Nenhum arquivo novo e/ou pendente. Finalizando.")
                tele.telegram_bot("Qualidade Garantia - Nenhum arquivo novo e/ou pendente. Finalizando.", TOKEN, CHAT_ID)
                break

            log.info(f"{len(indices_new)} New encontrados. Processando primeiro.")
            processar_item(navegador, wait_sdd, diretorio_download, indices_new[0])

    except Exception as erro_sdd:
        log.error(f"Erro no SDD: {erro_sdd}")
        tele_cliente.telegram_bot(f'Qualidade Garantia - Erro no portal SDD: {erro_sdd}', TOKEN, CHAT_ID_RPA)

    sleep(3)
    navegador.quit()

except Exception as e:
    diretorio_atual = os.getcwd()
    log.error(f"Erro geral: {e}")
    if 'no such element' not in str(e).lower() and 'unable to locate' not in str(e).lower():
        tele_cliente.telegram_bot(f'Qualidade Garantia - Erro {str(e)}', TOKEN, CHAT_ID_RPA)

# ── Mover e Processar ─────────────────────────────────────────────────────────
os.chdir(os.path.join(diretorio_atual, "Arquivos"))
for arquivo in os.listdir():
    if (arquivo.endswith('.xls') and 'G ' not in arquivo) or arquivo.endswith('.pdf'):
        shutil.move(
            os.path.join(diretorio_atual, "Arquivos", arquivo),
            os.path.join(PASTA_VALIDADA, arquivo))

os.chdir(PASTA_VALIDADA)
total_processados = 0
for arquivo_ in os.listdir():
    if (arquivo_.endswith('.xls') and 'G ' not in arquivo_) or arquivo_.endswith('.pdf'):
        if arquivo_.endswith('.xls'):
            extrair_2(arquivo_, "Mês de Criação")
            total_processados += 1
        pasta_destino = os.path.join(
            PASTA_QUALIDADE, str(ano),
            f"{str(mes).zfill(2)} - {nome_mes.capitalize()}")
        shutil.move(os.path.join(PASTA_VALIDADA, arquivo_), os.path.join(pasta_destino, arquivo_))

if total_processados > 0:
    try:
        tele.telegram_bot(
            f"Robô Garantia - Foram localizadas {total_processados} garantia(s), já disponíveis no BI.",
            TOKEN, CHAT_ID)
    except Exception as erro:
        tele.telegram_bot(f"Erro na extração da garantia: {erro}", TOKEN, CHAT_ID_RPA)
