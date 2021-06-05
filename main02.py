"""
desc: script destinado a receber dados de uma planilha e enviar para o email alertas sobre limites de prazos
Auth: 333333333
ultima modificação: 03/06/2021
"""

# ------ imports -------
import pandas as pd
from time import sleep
import smtplib


# ------- DEFS ----------


def disparar_email_pelo_Gmail(login, para, topico, msg):
    conexao = smtplib.SMTP('smtp.gmail.com', 587)
    conexao.starttls()
    conexao.login(login[0], login[1])
    conexao.sendmail(login, para, f'Subject:{topico}\n\n{msg}')
    conexao.quit()
    print("Email enviado com sucesso!")


# ----- inicio ---------
# listas a serem usadas



contas_venc2 = list()
contas_venc = list()
index1 = 0

login = ['erasmoks.junior@gmail.com', 'Er@smo17']
para = 'juniorerasmo02@hotmail.com'
topico = 'Alerta de Vencimento'

colunasBF = pd.read_excel("ContasaPagar.xlsx", usecols="B,F")
x = pd.read_excel("ContasaPagar.xlsx", usecols="F")

for index, item in enumerate(colunasBF['Prazo']):  # iterar sobre a coluna "Prazo"
    if item == "Vence Hoje":  # identificar contas que vencem hoje
        contas_venc.append(index)
for venc in contas_venc:
    for item in colunasBF['Descrição']:
        if index1 == venc:
            contas_venc2.append(item)
        index1 += 1
c = 0

while True:
    try:
        for i in x['Prazo']:
            if i == "Vence Hoje":
                for k in contas_venc2:
                    msg = f'Existe uma conta para pagar hoje de {contas_venc2}, verifique a tabela de Contas a Pagar !\n\n' \
                    f'Att fulano'
                    disparar_email_pelo_Gmail(login, para, topico, msg)
            elif i == "Vencido":
                msg = f'Existe uma conta vencida, verifique a tabela de Contas a Pagar !\n\nAtt fulano'
                disparar_email_pelo_Gmail(login, para, topico, msg)
            c += 1
        print("Retorno com a verificação em 12 horas")
        sleep(43200)  # 43200 seg são 12 horas

    except:
        print(" ----- erro ----- ")
