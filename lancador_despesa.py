from playwright.sync_api import sync_playwright
import time
import pandas as pd
from dotenv import load_dotenv
import os
load_dotenv()

df = pd.read_excel('dados.xlsx', sheet_name='importar')
df['Data Competencia'] = df['Data Competencia'].dt.strftime("%d/%m/%Y")
df['Data Vencimento'] = df['Data Vencimento'].dt.strftime("%d/%m/%Y")

lancamentos = []



def login(condominio):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()
        page.goto('', timeout=0)
        page.click('text= Aceitar todas')
        page.fill('#UsuarioDcEmail', os.getenv("USER_NAME"))
        page.fill('#UsuarioDcSenha', os.getenv("PASSWORD"))
        page.click('text= Entrar')
        page.click('text= Roderjan Condomínios')
        page.click('text = Entendi e Aceito')
        time.sleep(2)
        page.goto('', timeout=0)
        page.select_option('#CondominioId', str(condominio))
        context.storage_state(path="state.json")
        print('login realizado com sucesso')
        context.close()
        browser.close()


def lancador(numero_lancamentos):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(storage_state="state.json")
        page = context.new_page()   
        page.goto('', timeout=0)
        for linhas in numero_lancamentos:
            locator = df.loc[df['Documento'] == linhas].reset_index(drop=True)
            page.click('text= Adicionar Despesa', timeout=0)
            page.click('#s2id_LancamentoFornecedorId')
            page.fill('#select2-drop > div > input', locator['Fornecedor'][0])
            page.keyboard.press('Enter')
            page.fill('#LancamentoNmLancamento', locator['Historico'][0])
            page.click('#s2id_LancamentoPlanoContaId')
            page.fill('#select2-drop > div > input', locator['Plano Conta'][0])
            page.keyboard.press('Enter')
            page.fill('#LancamentoNrDocumento', str(locator['Documento'][0]))
            #competência
            page.locator('#LancamentoDtCompetencia').press('Control+A')
            page.locator('#LancamentoDtCompetencia').press('Backspace')
            page.locator('#LancamentoDtCompetencia').type(str(locator['Data Competencia'][0]))
            #vencimento
            page.locator('#LancamentoDtVencimento').press('Control+A')
            page.locator('#LancamentoDtVencimento').press('Backspace')
            page.locator('#LancamentoDtVencimento').type(str(locator['Data Vencimento'][0]))
            page.click('#FormLancamento > div.panel > div > div:nth-child(1) > div:nth-child(2) > fieldset > div:nth-child(6) > div > div:nth-child(6) > div > div > div', timeout=0)
            for index, row in locator.iterrows(): 
                page.fill('#EditNmItem', row['Historico'])
                page.click('#s2id_EditPlanoContaId')
                page.fill('#select2-drop > div > input', row['Plano Conta'])
                page.keyboard.press('Enter')
                page.fill('#EditVlItem', str(row['VALORCONTACONTABIL']))
                page.click('#btnIncluirItem')
            page.click('#FormLancamento > div.panel > div > div:nth-child(2) > div.col-md-4.pull-right > button', timeout=0)
            print('Lançamento: ' + str(linhas)+ ' realizado.')
        print('Finalizado!!!')
        context.close()
        browser.close()

login()
lancador(lancamentos)