from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd
import time
navegador = webdriver.Chrome()
#importar a base de dados
tabela = pd.read_excel('buscas.xlsx')



#CONSTRUINDO FUNÇÕES DE PESQUISAS

def busca_google_shopping(navegador, produto, termos_banidos, preco_min, preco_max):
    # abrindo o navegador
    navegador.get('https://www.google.com.br/')

    # tratando as informaçoes vindo da tabela
    produto = produto.lower()
    termos_banidos = termos_banidos.lower()
    lista_termos_produtos = produto.split(" ")
    lista_termos_banidos = termos_banidos.split(" ")

    # pesquisando o produto e clicando na aba shopping
    navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(
        produto, Keys.ENTER)
    elementos = navegador.find_elements(By.CLASS_NAME, 'hdtb-mitem')
    for item in elementos:
        if "Shopping" in item.text:
            item.click()
            break

            # pegando a lista de resultados e pegando as informações dele, Nome, preço e link
    lista_resultados = navegador.find_elements(By.CLASS_NAME, 'sh-dgr__grid-result')
    lista_ofertas = []
    for resultado in lista_resultados:
        nomes = resultado.find_element(By.CLASS_NAME, 'Xjkr3b').text
        nomes = nomes.lower()
        # fazendo o tratamento dos nomes, verificando se tem informações não desejadas
        tem_termos_banidos = False
        for palavras in lista_termos_banidos:
            if palavras in nomes:
                tem_termos_banidos = True

        tem_todos_termos = True
        for palavras in lista_termos_produtos:
            if palavras not in nomes:
                tem_todos_termos = False
        # verificando o nome
        if not tem_termos_banidos and tem_todos_termos:
            # fazendo o tratamento do preço
            try:
                precos = resultado.find_element(By.CLASS_NAME, 'a8Pemb').text
                precos = precos.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
                precos = float(precos)
                preco_min = float(preco_min)
                preco_max = float(preco_max)
                # Verificando se o preco esta dentro do preco minimo e maximo
                if preco_min <= precos <= preco_max:
                    elemento_link = resultado.find_element(By.CLASS_NAME, 'aULzUe')
                    elemento_pai = elemento_link.find_element(By.XPATH,
                                                              '..')  # '..' ele pega o elemento anterior a ele, ou seja o elemento pai
                    link = elemento_pai.get_attribute('href')
                    lista_ofertas.append((nomes, precos, link))
            except:
                continue

    return lista_ofertas


def busca_buscape(navegador, produto, termos_banidos, preco_min, preco_max):
    # abrindo o navegador
    navegador.get('https://www.buscape.com.br/')

    # tratando as informaçoes vindo da tabela
    produto = produto.lower()
    termos_banidos = termos_banidos.lower()
    lista_termos_produtos = produto.split(" ")
    lista_termos_banidos = termos_banidos.split(" ")
    preco_min = float(preco_min)
    preco_max = float(preco_max)

    # pesquisando o produto e clicando na aba shopping
    navegador.find_element(By.XPATH,
                           '//*[@id="new-header"]/div[1]/div/div/div[3]/div/div/div[2]/div/div[1]/input').send_keys(
        produto, Keys.ENTER)
    time.sleep(5)
    # pegando a lista de resultados e pegando as informações dele, Nome, preço e link
    lista_resultados = navegador.find_elements(By.CLASS_NAME, 'SearchCard_ProductCard_Inner__7JhKb')
    lista_ofertas = []
    for resultado in lista_resultados:
        nome = resultado.find_element(By.CLASS_NAME, 'Text_MobileLabelXs__rr7ZF ').text
        nome = nome.lower()
        preco = resultado.find_element(By.CLASS_NAME, 'Text_MobileHeadingS__XS_Au').text
        link = resultado.get_attribute('href')

        # fazendo o tratamento dos nomes, verificando se tem informações não desejadas
        tem_termos_banidos = False
        for palavras in lista_termos_banidos:
            if palavras in nome:
                tem_termos_banidos = True

        tem_todos_termos = True
        for palavras in lista_termos_produtos:
            if palavras not in nome:
                tem_todos_termos = False
        # verificando o nome
        if not tem_termos_banidos and tem_todos_termos:
            preco = preco.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
            preco = float(preco)
            if preco_min <= preco <= preco_max:
                lista_ofertas.append((nome, preco, link))
    return lista_ofertas


#JUNTANDO TODOS OS RESULTADOS EM UMA TABELA

tabela_ofertas = pd.DataFrame()

for linhas in tabela.index:
    produto = tabela.loc[linhas, "Nome"]
    termos_banidos = tabela.loc[linhas, "Termos banidos"]
    preco_min = tabela.loc[linhas, "Preço mínimo"]
    preco_max = tabela.loc[linhas, "Preço máximo"]
    lista_ofertas_google = busca_google_shopping(navegador, produto, termos_banidos, preco_min, preco_max)
    if lista_ofertas_google:
        tabela_google = pd.DataFrame(lista_ofertas_google, columns=['Produto', 'Preço', 'Link'])
        tabela_ofertas = pd.concat([tabela_ofertas, tabela_google], ignore_index=True)
    else:
        tabela_google = None
    lista_ofertas_buscape = busca_buscape(navegador, produto, termos_banidos, preco_min, preco_max)
    if lista_ofertas_buscape:
        tabela_buscape = pd.DataFrame(lista_ofertas_buscape, columns=['Produto', 'Preço', 'Link'])
        tabela_ofertas = pd.concat([tabela_ofertas, tabela_buscape], ignore_index=True)
    else:
        tabela_buscape = None



#EXPORTANDO A TABELA COM TODOS OS RESULTADOS PARA O EXCEL

tabela_ofertas.to_excel("Ofestas.xlsx", index=False)

#ENVIANDO O E-MAIL

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

if len(tabela_ofertas.index) > 0:
    # 1- STARTAR O SERVIDOR SMTP
    host = 'smtp.gmail.com'
    port = '587'
    login = 'maykon.rubens@gmail.com'
    senha = 'mwcxukhhwfmbkdjl'

    # Dando start no servidor
    server = smtplib.SMTP(host, port)
    server.ehlo()
    server.starttls()
    server.login(login, senha)

    # 2- CONSTRUIR O EMAIL TIPO MIME
    corpo = f'''<b><p>Prezados,</p></b>
    <b><p>Encontramos alguns produtos em ofertas com a faixa de preço desejada, segue a tabela com detalhes</p></b>
    {tabela_ofertas.to_html(index=False)}
    <b><p>Qualquer dúvida estou a disposição.</p></b>
    '''
    email_msg = MIMEMultipart()
    email_msg['From'] = login
    email_msg['To'] = 'maykon.devpython@gmail.com'
    email_msg['Subject'] = 'Produto(s) Encontrado(s) na faixa de preço desejada'
    email_msg.attach(MIMEText(corpo, 'html'))  # serve para anexar o corpo no email.

    # 4- ENVIAR O EMAIL TIPO MIME NO SERVIDOR SMTP
    server.sendmail(email_msg['From'], email_msg['To'], email_msg.as_string())

    server.quit()
    print('E-mail Enviado.')

navegador.quit()