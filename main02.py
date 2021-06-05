"""
desc: script destinado a receber dados de uma planilha e enviar para o email alertas sobre limites de prazos
Auth: treinee Admfin
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
# -------- Configs -----

login = ['email_remetente', 'senha']
para = 'email_Destinatário'
topico = 'Alerta de Vencimento'

# -------------------------------------------------------
colunasBF = pd.read_excel("ContasaPagar.xlsx", usecols="B,F") # faz uma verredura na coluna B e F em busca de atrasos
x = pd.read_excel("ContasaPagar.xlsx", usecols="F")
# -------- Looping de verificação continua --------------
while True:
    try:
        for i in x['Prazo']:
            print("Vericando atrasos")
            if i == "Vence Hoje":
                print("Existe uma conta para vencer hoje")
                msg = f'Existe uma conta para pagar HOJE!, verifique a tabela de Contas a Pagar !\n\n' \
                f'Att fulano'
                print("Disparando Email de alerta")
                disparar_email_pelo_Gmail(login, para, topico, msg)
            elif i == "Vencido":
                print("Existe uma conta Vencida!")                
                msg = f'Existe uma conta vencida, verifique a tabela de Contas a Pagar !\n\nAtt fulano'
                print("Disparando Email de alerta")
                disparar_email_pelo_Gmail(login, para, topico, msg)
        print("Retorno com a verificação em 12 horas")
        sleep(43200)  # 43200 seg são 12 horas

    except:
        print(" ----- erro ----- ")
