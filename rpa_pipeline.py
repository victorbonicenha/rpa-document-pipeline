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
from time import sleep
import pandas as pd
import openpyxl
import shutil
import pyotp
import xlrd
import calendar
import locale
import time
import os

#--------- Config de Email, caso seja necessario envio de emails novamente --------# 
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
caminho_foto = os.path.join(os.getcwd(), 'logo.jpg')
email_geral = Smtp('Garantia')
assunto = 'Acompanhamento Qualidade Garantia'

#----------Telegram usado no codigo Old --------------#
tele_cliente = TelegramSend('Qualidade_Garantia: ')
tele = TelegramSend('Qualidade_Garantia: ')
token = '898143240:AAHuspXl8VHSHYucIghHvqwrtr5X_J4uGcA'
chat_id = '-4544977407'
chat_id_rpa = '-921014429'

#----------Telegram do bom de senha -----------------#
#ele_cliente = TelegramSend('Qualidade_Garantia: ')
#tele = TelegramSend('Qualidade_Garantia: ')
#token='7648994612:AAH0WXuxbShdvnY36M3CuYnedJVW2q3oZd0'
#chat_id='7815571743'
#chat_id_rpa='7815571743'

#----------- Variaveis diretórios e tempo/data -------#
ano = datetime.today().strftime('%Y')
mes = int(datetime.today().strftime('%m'))
nome_mes = calendar.month_name[mes]
data_atual = datetime.now().strftime("%d-%m-%Y")
pasta_base = rf'Q:\Docs Qualidade\02. Restrito\01 - CONTROLES QUALIDADE INDUSTRIAL\03 - GARANTIA\01 - VW\02 - NOTAS DE DÉBITOS\{str(ano)}'

meses = [
    "01 - Janeiro", "02 - Fevereiro", "03 - Março", "04 - Abril",
    "05 - Maio", "06 - Junho", "07 - Julho", "08 - Agosto",
    "09 - Setembro", "10 - Outubro", "11 - Novembro", "12 - Dezembro"
]

if not os.path.isdir(pasta_base):
    os.makedirs(pasta_base, exist_ok=True)
    for mes_item in meses:
        caminho_mes = os.path.join(pasta_base, mes_item)
        os.makedirs(caminho_mes, exist_ok=True)


#--------------- Função de Extração de dados dos arquivos excel baixado e inserção no banco de dados----#
def extrair_2(arquivo_, chave):
    arquivo_xlsx = arquivo_.replace('.xls', '_temp.xlsx') if arquivo_.endswith('.xls') else arquivo_
    if arquivo_ != arquivo_xlsx:
        shutil.copy2(arquivo_, arquivo_xlsx)

    try:
        df = pd.read_excel(arquivo_xlsx, engine='openpyxl')
    except Exception:
        df = pd.read_excel(arquivo_, engine='xlrd')

    df_sem_nan = df.fillna('None')
    indice_pagina = df_sem_nan[df_sem_nan.astype(str).apply(lambda x: 'Número de Operação'
                                                                      in ' '.join(x), axis=1)].index[0]
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
            print(f'Qualidade Garantia - erro ao ler mes_criacao: {e}')

    if arquivo_ != arquivo_xlsx and os.path.exists(arquivo_xlsx):
        os.remove(arquivo_xlsx)

    data_execucao = datetime.today().strftime('%d/%m/%Y')

    for _, row in df_colunas.iterrows():
        num_op = row[0]
        cod_peca_mon = row[1]
        cod_peca_mon_Nord = row[2]
        descricao_pecas = row[3]
        op_mao_obra = row[4]
        descricao = row[5]
        cod_dano_ = row[6]
        cod_dano = str(cod_dano_).replace("-", "")
        kdnr_ = row[7]
        kdnr = str(kdnr_).replace("-", "")
        num_peca_mod_ = row[8]
        num_peca_mod = str(num_peca_mod_).replace("-", "")
        cod_rep_ = row[9]
        cod_rep = str(cod_rep_).replace("-", "")
        qtde_pec_desm_ = row[10]
        qtde_pec_desm = str(qtde_pec_desm_).replace("-", "")
        qtde_pec_mon_ = row[11]
        qtde_pec_mon = str(qtde_pec_mon_).replace("-", "")
        fator_tec_ = row[12]
        fator_tec = str(fator_tec_).replace("-", "")
        vin = row[13]
        mod_vendas = row[14]
        data_prod_string = row[15]
        data_rec_rep_string = row[16]
        data_reg_inst_string = row[17]
        km_milhas_ = row[18]
        km_milhas = str(km_milhas_).replace("-", "")
        fat_custo_add_ = row[19]
        fat_custo_add = str(fat_custo_add_).replace("-", "")
        chave_estorno = row[20]
        tipo_estorno = row[21]
        unid_temp_ = row[22]
        unid_temp = str(unid_temp_).replace("-", "")
        custo_mat_ = row[23]
        custo_mat = str(custo_mat_).replace("-", "")
        custo_mao_obra_ = row[24]
        custo_mao_obra = str(custo_mao_obra_).replace("-", "")
        custo_add_ = row[25]
        custo_add = str(custo_add_).replace("-", "")
        val_debitado_ = row[26]
        val_debitado = str(val_debitado_).replace("-", "")
        val_reclam_mat_ = row[27]
        val_reclam_mat = str(val_reclam_mat_).replace("-", "")
        preco_compra_ = row[28]
        preco_compra = str(preco_compra_).replace("-", "")
        num_concessionaria = row[30]
        num_reclam_ = row[31]
        num_reclam = str(num_reclam_).replace("-", "")
        tipo_reclam_ = row[32]
        tipo_reclam = str(tipo_reclam_).replace("-", "")
        chave_id_reclam = row[33]

        try:
            data_prod = datetime.strptime(str(data_prod_string), "%Y-%m-%d").strftime("%d-%m-%Y")
            data_prod_formatada = datetime.strptime(data_prod, "%d-%m-%Y").date()
        except Exception as erro_dt:
            print(erro_dt)
            data_prod_formatada = datetime.strptime('01/01/2001', '%d/%m/%Y').date()
        try:
            data_rec_rep = datetime.strptime(str(data_rec_rep_string), "%Y-%m-%d").strftime("%d-%m-%Y")
            data_rec_rep_formatada = datetime.strptime(data_rec_rep, "%d-%m-%Y").date()
        except Exception as erro_dt:
            print(erro_dt)
            data_rec_rep_formatada = datetime.strptime('01/01/2001', '%d/%m/%Y').date()
        try:
            data_reg_inst = datetime.strptime(str(data_reg_inst_string), "%Y-%m-%d").strftime("%d-%m-%Y")
            data_reg_inst_formatada = datetime.strptime(data_reg_inst, "%d-%m-%Y").date()
        except Exception as erro_dt:
            print(erro_dt)
            data_reg_inst_formatada = datetime.strptime('01/01/2001', '%d/%01/2001', '%d/%m/%Y').date()

        try:
            sleep(0.05)
            insert = f"""INSERT INTO [dbo].[Garantia_Volks]
               ([Num_Operacao]
               ,[Cod_Peca_Montada_Original]
               ,[Cod_Peca_Montada_Nao_Ordenado]
               ,[Descricao_Pecas]
               ,[Op_mao_de_obra]
               ,[Descricao]
               ,[Cod_Dano]
               ,[KDNR]
               ,[Num_Peca_Modulo]
               ,[Cod_Reparacao]
               ,[Qtd_Pecas_Desmontadas]
               ,[Qtd_Pecas_Montadas]
               ,[Fator_Tecnico]
               ,[VIN]
               ,[Modelo_Vendas]
               ,[Data_Producao]
               ,[Data_Recep_Reparo]
               ,[Data_Reg_Inst]
               ,[Quilometros_Milhas]
               ,[Fator_Custos]
               ,[Chave_Estorno]
               ,[Tipo_Estorno_Garantia]
               ,[Unidades_Tempo]
               ,[Custos_Material]
               ,[Custos_mao_de_obra]
               ,[Custos_adicionais]
               ,[Valor_Debitado]
               ,[Valor_Reclamacao_Material]
               ,[Preco_Compra]
               ,[Num_Concessionaria]
               ,[Num_Reclamacao]
               ,[Tipo_Reclamacao]
               ,[Chave_identificacao_reclamacao]
               ,[mes_criacao]
               ,[data_execucao])
         VALUES
            ( '{num_op}', '{cod_peca_mon}', '{cod_peca_mon_Nord}', '{descricao_pecas}', '{op_mao_obra}', '{descricao}',
             '{cod_dano}', '{kdnr}', '{num_peca_mod}', '{cod_rep}', '{qtde_pec_desm}', '{qtde_pec_mon}', '{fator_tec}',
              '{vin}', '{mod_vendas}', '{data_prod_formatada}', '{data_rec_rep_formatada}', '{data_reg_inst_formatada}',
               '{km_milhas}', '{fat_custo_add}', '{chave_estorno}', '{tipo_estorno}', '{unid_temp}',
                '{float(str(custo_mat).replace(",", "."))}',
                 '{float(str(custo_mao_obra).replace(",", "."))}',
                  '{float(str(custo_add).replace(",", "."))}',
                   '{float(str(val_debitado).replace(",", "."))}',
                   '{float(str(val_reclam_mat).replace(",", "."))}', '{preco_compra}',
                    '{num_concessionaria}', '{num_reclam}', '{tipo_reclam}', '{chave_id_reclam}',
                     '{str(mes_criacao)}', '{str(data_execucao)}')"""
            cursor.execute(insert)
            cursor.commit()
            print('Qualidade Garantia - insert')
        except Exception as er:
            print('Qualidade Garantia - erro insert')
            print('Qualidade Garantia - ' + str(er))


#--------------Função visando esperar o término do Download de fato e garantir que nao quebre ou falte algum-------#

def aguardar_download_concluir(diretorio, timeout=60):
    inicio = time.time()
    while time.time() - inicio < timeout:
        pendentes = [f for f in os.listdir(diretorio) if f.endswith('.crdownload') or f.endswith('.tmp')]
        if not pendentes:
            return True
        sleep(1)
    print("Qualidade Garantia - timeout aguardando download")
    return False


#----------------- Função de verificação se tem status "NEWS" ou "SEEN" que indicam itens novos e itens que ainda tem Downloads nao feito --------#
def obter_celulas_por_status(navegador, status):
    """Retorna indices das celulas com o status informado (new ou seen)"""
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

#------------- Função de processamento, Download dos arquivos cujo estejam dentro do "NEW" ou "SEEN" mostrado acima-------------#
def processar_item(navegador, wait_sdd, diretorio_download, idx):
    """
    Abre o item pelo indice, verifica links primary (faltam baixar) e baixa um por vez.
    Retorna True se baixou pelo menos 1 arquivo, False se nao havia primary.
    """
    try:
        celulas = navegador.find_elements(By.XPATH,
            '//*[@id="header-table-page"]/groupui-table-cell[.//groupui-tag]')
        celula = celulas[idx]
        navegador.execute_script("arguments[0].scrollIntoView(true);", celula)
        sleep(1)
        celula.click()
        sleep(3)

        # Abre accordion Documents
        btn_documents = wait_sdd.until(
            ec.element_to_be_clickable((By.XPATH,
                '//*[@id="content-unit-header"]/div/groupui-accordion/span/groupui-headline'))
        )
        btn_documents.click()
        sleep(2)

        # Verifica se tem links primary (faltam baixar)
        links_pendentes = navegador.find_elements(By.XPATH,
            '//*[@id="content-unit-header"]/div/groupui-accordion/groupui-list/ul/li/div/groupui-link[@variant="primary"]')

        if not links_pendentes:
            print(f'Qualidade Garantia - idx {idx+1}: nenhum link pendente')
            return False

        # Baixa apenas o primeiro primary (pagina recarrega depois do click)
        link = links_pendentes[0]
        titulo = link.get_attribute('title') or f'arquivo_{idx+1}'
        arquivos_antes = set(os.listdir(diretorio_download))
        navegador.execute_script("arguments[0].scrollIntoView(true);", link)
        sleep(0.5)
        link.click()
        print(f'Qualidade Garantia - idx {idx+1}: baixando "{titulo}"')

        aguardar_download_concluir(diretorio_download, timeout=60)
        sleep(3)

        arquivos_depois = set(os.listdir(diretorio_download))
        novos = arquivos_depois - arquivos_antes
        print(f'Qualidade Garantia - idx {idx+1}: baixados {novos}')
        return True

    except Exception as erro:
        print(f'Qualidade Garantia - erro ao processar idx {idx+1}: {erro}')
        return False


#-------- Algumas configs basicas do banco de dados, navegador e caminho que Robo irá percorrer até os downloads ------------#
rpa_cred = Bank('Credenciais')
cursor1 = rpa_cred.bank_connection('user_rpa', '%*4Us5z$', '186.193.228.29', 'datadriven_paranoa')

rpa_29 = Bank('rpa')
cursor = rpa_29.bank_connection('elipse', 'E#lipse#365#ic', '192.168.0.30', 'ELIPSE_E3')

pasta = "Arquivos"

sleep(3)

try:
    diretorio_atual = os.getcwd()
    diretorio_download = diretorio_atual + f'\\{pasta}'
    options = Options()
    options.add_experimental_option("prefs", {
        "download.default_directory": diretorio_download,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "plugins.always_open_pdf_externally": True,
        "plugins.plugins_disabled": ["Chrome PDF Viewer"]
    })

    servico = Service(ChromeDriverManager().install())
    navegador = webdriver.Chrome(service=servico, options=options)
    navegador.maximize_window()
    navegador.implicitly_wait(4)

    logins = cursor1.execute(
        "SELECT usuario FROM [datadriven_paranoa].[dbo].[dw_view_acessos_rpa] where nome like '%Garantia Volkswagen%' and cliente = 'Qualidade'").fetchall()[0]
    senhas = cursor1.execute(
        "SELECT senha FROM [datadriven_paranoa].[dbo].[dw_view_acessos_rpa] WHERE nome LIKE '%Volkswagen%' AND cliente = 'Qualidade' AND senha IS NOT NULL AND LTRIM(RTRIM(senha)) <> ''").fetchall()[0]
    login = ''
    senha = ''

    try:
        for login_, senha_ in zip(logins, senhas):
            login = login_.strip()
            senha = senha_.strip()
        if not senha:
            tele_cliente.telegram_bot("Qualidade Garantia - Senha expirada. Favor atualizar", token, chat_id_rpa)
            navegador.close()
    except Exception as e:
        print('Qualidade Garantia - Senha expirada. Favor atualizar')

    wait = WebDriverWait(navegador, 10)
    navegador.maximize_window()
    key = 'GOVJKEZXVTVFZ7RJTR2HTILDXTUWFO2K'

#---------- Função para puxar o codigo gerado de authenticação -----------# 
    def get_totp(secret):
        totp = pyotp.TOTP(secret)
        return totp.now()

    cod = get_totp(key)

    try:
        navegador.get(
            r"https://lso.volkswagen.de/lona/eai/b2b9ext/pages/login.html?url=/one-kbp/content/pt/kbp_private/kbpprivate_systems/kbpprivate_systempage_16248.jsp%3Fbrand%3Dnull")
        navegador.find_element(By.ID, 'contentForm:profileIdInput').send_keys(login)
        sleep(2)
        navegador.find_element(By.ID, 'contentForm:passwordInput').send_keys(senha)
        sleep(2)
        navegador.find_element(By.ID, 'contentForm:passwordLoginAction').click()
        sleep(2)
    except Exception as erro_gr:
        print(erro_gr)
        screenshot_path = os.path.join(os.getcwd(), "erro_Garantia.png")
        navegador.save_screenshot(screenshot_path)
        tele_cliente.telegram_bot_image('Qualidade Garantia - erro de login', token, chat_id_rpa, screenshot_path)

    sleep(3)
    navegador.find_element(By.XPATH, '//*[@id="otp"]').send_keys(cod)
    sleep(1)
    navegador.find_element(By.XPATH, '//*[@id="panel-login"]/div/div[2]/form/div/button').click()
    sleep(3)

    navegador.find_element(By.XPATH, '//*[@id="opg-nav-list"]/li[5]/a').click()
    sleep(2)
    navegador.find_element(By.XPATH, '//*[@id="system-dropdown"]/div[2]/ul[1]/li[1]/a').click()
    sleep(3)

# ========== PORTAL SDD - DOWNLOAD (Parte nova => 14/05/26)==========
#------------ Baiscamente Loop de verificação de Itens e garantindo que todos os novos e pendentes sejam baixados-------------#
    try:
        while True:
            # Recarrega o SDD a cada iteracao (apos cada download volta pra tela anterior)
            navegador.get("https://lso.volkswagen.de/sdd")
            sleep(5)
            wait_sdd = WebDriverWait(navegador, 20)

            # 1. Prioridade: Seen com links pendentes
            indices_seen = obter_celulas_por_status(navegador, 'seen')
            processou_seen = False

            for idx in indices_seen:
                resultado = processar_item(navegador, wait_sdd, diretorio_download, idx)
                if resultado:
                    processou_seen = True
                    break  # Volta ao inicio do while para re-verificar

            if processou_seen:
                continue

            # 2. Sem Seen pendente: procura New
            indices_new = obter_celulas_por_status(navegador, 'new')

            if not indices_new:
                msg = 'Nenhum arquivo novo e/ou pendente. Finalizando.'
                tele.telegram_bot(f"{msg}", token, chat_id)
                break

            print(f'Qualidade Garantia - {len(indices_new)} New encontrados. Processando primeiro.')
            processar_item(navegador, wait_sdd, diretorio_download, indices_new[0])
            # Volta ao inicio do while — o New virou Seen e sera tratado na proxima iteracao

    except Exception as erro_sdd:
        print(f'Qualidade Garantia - erro no SDD: {erro_sdd}')
        tele_cliente.telegram_bot(f'Qualidade Garantia - Erro no portal SDD: {erro_sdd}', token, chat_id_rpa)

    # ========== FIM SDD ==========
#--------- Tratativas de Erro ----------#
    try:
        tabela_screenshot = navegador.find_elements(By.ID, 'overview')
        if tabela_screenshot:
            prints = tabela_screenshot[0].screenshot_as_png
            with open(diretorio_atual + f"\\{pasta}\\tabela.png", "wb") as f:
                f.write(prints)
    except Exception as erro_el:
        print(erro_el)

    sleep(3)
    navegador.quit()

except Exception as e:
    diretorio_atual = os.getcwd()
    if 'no such element' in str(e).lower() or 'unable to locate' in str(e).lower():
        sleep(60)
        print(f'Qualidade Garantia - Sem novos arquivos no portal: {str(e)}')
    else:
        tele_cliente.telegram_bot(f'Qualidade Garantia - Erro {str(e)}', token, chat_id_rpa)

# ========== MOVER E PROCESSAR ARQUIVOS ==========
#------------ Parte de mover o arquivo da pasta "BASE" para pasta "DESTINO" e fazendo uma contagem e/ou checagem de quantos itens foram nesse processo e avisando telegram com erro ou sucesso----------# 
os.chdir(os.path.join(diretorio_atual, "Arquivos"))
todos_arquivos_pasta = os.listdir()
if len(todos_arquivos_pasta) > 0:
    for arquivo in todos_arquivos_pasta:
        if (arquivo.endswith('.xls') and 'G ' not in arquivo) or arquivo.endswith('.pdf'):
            pastaqualidade_garantia = r'Q:\Docs Qualidade\02. Restrito\01 - CONTROLES QUALIDADE INDUSTRIAL\03 - GARANTIA\01 - VW\02 - NOTAS DE DÉBITOS\Garantia_Validada'
            shutil.move(diretorio_atual + r"\Arquivos" + f"\\{arquivo}", f"{pastaqualidade_garantia}\\{arquivo}")

pasta_validado = r'Q:\Docs Qualidade\02. Restrito\01 - CONTROLES QUALIDADE INDUSTRIAL\03 - GARANTIA\01 - VW\02 - NOTAS DE DÉBITOS\Garantia_Validada'
os.chdir(pasta_validado)
todos_arquivos = os.listdir()
total_processados = 0
if len(todos_arquivos) > 0:
    for arquivo_ in todos_arquivos:
        if (arquivo_.endswith('.xls') and 'G ' not in arquivo_) or arquivo_.endswith('.pdf'):
            chave_procurada = "Mês de Criação"
            if arquivo_.endswith('.xls'):
                extrair_2(arquivo_, chave_procurada)
                total_processados += 1

            pastaqualidade = rf'Q:\Docs Qualidade\02. Restrito\01 - CONTROLES QUALIDADE INDUSTRIAL\03 - GARANTIA\01 - VW\02 - NOTAS DE DÉBITOS\{str(ano)}\{str(mes).zfill(2)} - {nome_mes.capitalize()}'
            shutil.move(pasta_validado + f"\\{arquivo_}", f"{pastaqualidade}\\{arquivo_}")

if total_processados > 0:
    try:
        tele.telegram_bot(f"Robô Garantia - Foram localizadas {total_processados} garantia(s), já disponíveis no BI.", token, chat_id)
    except Exception as erro:
        tele.telegram_bot(f"Erro na extração da garantia: {erro}", token, chat_id_rpa)
