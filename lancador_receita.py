from playwright.sync_api import sync_playwright
import pandas as pd
from dotenv import load_dotenv
import os
load_dotenv()


df = pd.read_excel('dados.xlsx', sheet_name='importar')
df.data = df.data.dt.strftime("%d/%m/%Y")


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
        page.wait_for_timeout(10000)
        page.goto('', timeout=0)
        page.select_option('#CondominioId', str(condominio))
        context.storage_state(path="state.json")
        print('login realizado com sucesso')
        context.close()
        browser.close()



def lancador_receita():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(storage_state="state.json")
        page = context.new_page()   
        page.goto('', timeout=0)     
        for index, linha in df.iterrows():
            page.click('text= Adicionar Receita', timeout=0)
            page.click('#s2id_LancamentoTpOrigem')
            page.fill('#select2-drop > div > input', 'Outros')
            page.keyboard.press('Enter')
            #Descricao
            page.fill('#LancamentoNmLancamento', linha.historico)
            #Valor
            page.fill('#LancamentoVlLancamento', str(linha.valor))
            #Conta
            page.click('#s2id_LancamentoContaBancariaId')
            page.fill('#select2-drop > div > input', linha.conta)
            page.keyboard.press('Enter')
            #Plano de contas
            page.click('#s2id_LancamentoPlanoContaId')
            page.fill('#select2-drop > div > input', linha.plano_conta)
            page.keyboard.press('Enter')
            #Competencia
            page.locator('#LancamentoDtCompetencia').press('Control+A')
            page.locator('#LancamentoDtCompetencia').press('Backspace')
            page.locator('#LancamentoDtCompetencia').type(str(linha.data))
            #Vencimento
            page.locator('#LancamentoDtVencimento').press('Control+A')
            page.locator('#LancamentoDtVencimento').press('Backspace')
            page.locator('#LancamentoDtVencimento').type(str(linha.data))
            #salvar
            page.click('#FormLancamento > div.panel > div > div:nth-child(2) > div.col-md-4.pull-right > button')
            #
            print('Conta: ' + linha.conta +' '+ 'Lancamento: ' + linha.historico + ' Salvo')
            page.wait_for_timeout(6000)
        print('Finalizado')
    
login()
lancador_receita()


 