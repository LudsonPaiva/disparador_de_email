
import win32com.client as client
import pandas as pd
import datetime as dt

# ler o arquivo em excell
tabela = pd.read_excel('mailing_e-mails.xlsx')
# print(tabela)

# ver como os dados estao chegando
tabela.info()

hoje = dt.datetime.now()
# print(hoje)

tabela_devedores = tabela.loc[tabela['STATUS'] == 'PENDENTE']
# print(tabela_devedores)
tabela_devedores = tabela_devedores.loc[tabela_devedores['DATA PREVISTA PARA PAGAMENTO'] < hoje]
print(tabela_devedores)

# criando e-mail atravez do outlook
outlook = client.Dispatch('Outlook.Application')
emissor = outlook.session.Accounts['lpaiva@proteste.org.br']
# mensagem = outlook.CreateItem(0)
# mensagem.Display()

# dados do e-mail
# mensagem.To = 'ludsonpaiva@yahoo.com.br'
# mensagem.Subject = 'Bem Vindo a Proteste'
# mensagem.Body = """

# MENSAGEM

# """

# tocar o emissor
# mensagem._oleobj_.Invoke(*(64209,0,8,0,emissor))
# mensagem.Save()
# mensagem.Send()

# criar uma lista para cada linha
dados = tabela_devedores[['CPF', 'VALOR EM ABERTO', 'DATA PREVISTA PARA PAGAMENTO', 'E-MAIL', 'NF']].values.tolist()

for dado in dados:
    cliente = dado[0]
    destinatario = dado[3]
    nf = dado[4]
    prazo = dado[2]
    prazo = prazo.strftime('%d/%m/%Y')
    valor = dado[0]

    mensagem = outlook.CreateItem(0)
    # mensagem.Display()

    # dados do e-mail
    mensagem.To = destinatario
    mensagem.Subject = 'Bem Vindo a Proteste'
    mensagem.Body = f"""

    Prezado Cliente {cliente},

    Bem vindo a PROTESTE.
    Seu código de cliente é {nf}, o valor de sua assinatura é R$ {valor:.2f} e a próxima parcela será cobrada em: {prazo}.

    Qualquer coisa é só ligar.

    Att.
    Equipe de atendimento.

    """

    # tocar o emissor
    mensagem._oleobj_.Invoke(*(64209, 0, 8, 0, emissor))
    mensagem.Save()
    mensagem.Send()