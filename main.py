from time import sleep
import pandas as pd
from selenium import webdriver
import os
def main():
    navegador = webdriver.Edge()
    abrir_site(navegador)
    login(navegador)
    df = montar_df(navegador)
    df.to_excel('ChamadosFC.xlsx', index=False)
    os.remove("C:\\Users\pedro\Downloads\Chamados.csv")
    #Planilha do FF
    '''abrir_site(navegador)
    loginFF(navegador)
    filtro_planilha(navegador)
    df = cria_df()
    incluir_informacoes(df, navegador)
    df.to_excel('ChamadosFF.xlsx', index=False)
    os.remove("C:\\Users\pedro\Downloads\Chamados.csv")
    os.startfile('ChamadosFF.xlsx')'''
    os.startfile('ChamadosFC.xlsx')
    navegador.close()
def abrir_site(navegador):
    navegador.maximize_window()
    navegador.get('https://suporte.netviewinformatica.com.br/Suporte/Home')
#realiza o login no site da netview
def login(navegador):
    navegador.find_element(by='id', value='Login').send_keys('login@ferreiraechagas.com.br')
    navegador.find_element(by='id', value='Senha').send_keys('*****')
    navegador.find_element(by='xpath', value='//*[@id="frmLogin"]/div[4]/button').click()
    fechar_popup(navegador)
def loginFF(navegador):
    navegador.find_element(by='id', value='Login').send_keys('loginFF@ferreiraechagas.com.br')
    navegador.find_element(by='id', value='Senha').send_keys('*****')
    navegador.find_element(by='xpath', value='//*[@id="frmLogin"]/div[4]/button').click()
    fechar_popup(navegador)
#fecha o popup de diligência
def fechar_popup(navegador):
    sleep(1.5)
    navegador.find_element(by='xpath', value='//*[@id="ModalChamadosEmAtraso"]/div/div/div[1]/button').click()
def montar_df(navegador):
    filtro_planilha(navegador)
    df = cria_df()
    incluir_informacoes(df, navegador)
    return df
#realiza os filtros da planilha
def filtro_planilha(navegador):
    navegador.find_element(by='name', value='codStatus').send_keys('Aguardando validação')
    navegador.find_element(by='xpath', value='//*[@id="dtInicio"]').send_keys('01/01/2022')
    navegador.find_element(by='xpath', value='//*[@id="dtFim"]').send_keys('31/12/2022')
    navegador.find_element(by='xpath', value='//*[@id="btnFiltrar"]').click()
    fechar_popup(navegador)
    # clica para exportar
    navegador.find_element(by='xpath', value='//*[@id="btnExcel"]/a').click()
def cria_df():
    sleep(3)
    df = pd.read_csv("C:\\Users\pedro\Downloads\Chamados.csv",
                     encoding="UTF-8", sep=";", usecols=['N°', 'Sistema', 'Assunto', 'Data Update', 'Tipo Chamado'])
    df.insert(5, 'Ultima interação', '')
    df.insert(6, 'DT_LIMITE_FC', '')
    df.insert(7, 'Interação é nossa?', '')
    df.insert(8, 'Responsável', '')
    return df
def incluir_informacoes(df, navegador):
    i = 0
    # abrir OS
    for N in df["N°"]:
        navegador.get('https://suporte.netviewinformatica.com.br/Suporte/Details/' + str(N))
        fechar_popup(navegador)
        # pegar informações da OS
        preencher_datas(i, df, navegador)
        # pega o usuário que interagiu por ultimo
        ul = navegador.find_element(by='xpath',
                                    value='//*[@id="containerHistoricoChamado"]/div[1]/div/ul/div/div/li/div/div[1]/span')
        interacao = ul.text.lower().strip()
        if interacao.find('account_circle'):
            df.iloc[i, 7] = 'Não'
        else:
            interacao = interacao.replace('account_circle', '')
            df.iloc[i, 7] = 'Sim'
            df.iloc[i, 8] = ul.text.replace('account_circle', '').strip()
        i += 1
def preencher_datas(i, df, navegador):
    # data da ultmima interação
    data = navegador.find_element(by='xpath',
                                  value='//*[@id="containerHistoricoChamado"]/div[1]/div/ul/div[1]/div/li/div/div[3]/span')
    data = pd.to_datetime(data.text)
    df.iloc[i, 5] = data_formatada(data)
    # data limite para validar/interagir
    datalim = pd.date_range(start=data, periods=2, freq='4D')
    df.iloc[i, 6] = data_formatada(datalim[1])
def data_formatada(data):
    data = data.strftime('%d/%m/%Y')
    return data
main()
