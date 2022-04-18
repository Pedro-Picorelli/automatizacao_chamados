import time
import pandas as pd
from selenium import webdriver
import os

#fecha o popup de diligência
def fechar_popup():
    time.sleep(1)
    navegador.find_element(by='xpath', value='//*[@id="ModalChamadosEmAtraso"]/div/div/div[1]/button').click()
    #fecha pop up no OK
    #navegador.find_element_by_xpath('//*[@id="ModalNotificaChamados"]').click()

#realiza o login no site da netview
def login():
    navegador.find_element(by='id', value='Login').send_keys('#####')
    navegador.find_element(by='id', value='Senha').send_keys('*****')
    navegador.find_element(by='xpath', value='//*[@id="frmLogin"]/div[4]/button').click()


#realiza os filtros da planilha
def filtro_planilha():
    navegador.find_element(by='name', value='codStatus').send_keys('Aguardando validação')
    navegador.find_element(by='xpath', value='//*[@id="dtInicio"]').send_keys('01/01/2022')
    navegador.find_element(by='xpath', value='//*[@id="dtFim"]').send_keys('31/12/2022')
    navegador.find_element(by='xpath', value='//*[@id="btnFiltrar"]').click()

#INICIO
#abre o navegador
navegador = webdriver.Edge()
navegador.maximize_window()
navegador.get('https://suporte.netviewinformatica.com.br/Suporte/Home')
login()
fechar_popup()
filtro_planilha()
fechar_popup()

#clica para exportar
navegador.find_element(by='xpath', value='//*[@id="btnExcel"]/a').click()

#pega a planilha e transforma em um datafreme
time.sleep(3)
df = pd.read_csv("C:\\Users\pedro\Downloads\Chamados.csv", encoding="UTF-8", sep=";", usecols=['N°', 'Sistema', 'Assunto', 'Data Update', 'Tipo Chamado'])
df.insert(5, 'Ultima interação', '')
df.insert(6, 'DT_LIMITE_FC', '',)
df.insert(7, 'Foi nossa SIM/NÃO', '',)
df.insert(8, 'Responsável', '')

i = 0
#abrir OS
for N in df["N°"]:

    navegador.get('https://suporte.netviewinformatica.com.br/Suporte/Details/' + str(N))
    fechar_popup()

    #pegar informações da OS
    #data da ultmima interação
    dataweb = navegador.find_element(by='xpath', value='//*[@id="containerHistoricoChamado"]/div[1]/div/ul/div[1]/div/li/div/div[3]/span')
    data = pd.to_datetime(dataweb.text)
    dataformatada = data.strftime('%d/%m/%Y')
    df.iloc[i, 5] = dataformatada
    #data limite para validar/interagir
    datalim = pd.date_range(start=data, periods=2, freq='4D')
    dataformatada = pd.to_datetime(datalim[1])
    dataformatada = dataformatada.strftime('%d/%m/%Y')
    df.iloc[i, 6] = dataformatada

    #pega o usuário que interagiu por ultimo
    ul = navegador.find_element(by='xpath', value='//*[@id="containerHistoricoChamado"]/div[1]/div/ul/div/div/li/div/div[1]/span')
    interacao = ul.text.lower().strip()
    if interacao.find('account_circle'):
        df.iloc[i, 7] = 'não'
    else:
        interacao = interacao.replace('account_circle', '')
        df.iloc[i, 7] = 'sim'
        df.iloc[i, 8] = ul.text.replace('account_circle', '').strip()

    i += 1

df.to_excel('Chamados2.xlsx', index=False)
navegador.close()
os.remove("C:\\Users\pedro\Downloads\Chamados.csv")
os.startfile('Chamados2.xlsx')
