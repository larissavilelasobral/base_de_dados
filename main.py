import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from email.mime.base import MIMEBase
from email import encoders

from segredos import senha, email_de_origem, email_de_destino
import pandas as pd

tabela_vendas = pd.read_excel('Vendas.xlsx')

# faturamento por loja
pd.set_option('display.max_columns', None)
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)
print('-' * 50)

# quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)
print('-' * 50)

# ticket medio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
print(ticket_medio)

# enviar um email com o relatório
EMAIL_ADDRESS = email_de_origem
EMAIL_PASSWORD = senha

msg = MIMEMultipart()
msg['Subject'] = 'Relatoria de Vendas Mensal'
msg['From'] = email_de_origem
msg['To'] = email_de_destino
msg.attach(MIMEText(f''' 
<h2>Bom dia!</h2>

<h2>Faturamento:</h2>
{faturamento.to_html()}

<h2>Quantidade Vendida:</h2>
{quantidade.to_html()}

<h2>Ticket Medio dos Produtos em cada Loja:</h2>
{ticket_medio.to_html()}

Obrigado!
''', 'html'))

cam_arquivo = 'C:\\Users\\96086\\PycharmProjects\\pythonProject02\\Vendas.xlsx'
attchment = open(cam_arquivo, 'rb')

att = MIMEBase('application', 'octet-stream')
att.set_payload(attchment.read())
encoders.encode_base64(att)

att.add_header('Content-Disposition', f'attachment; filename=Vendas.xlsx')
attchment.close()
msg.attach(att)

with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
    smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
    smtp.send_message(msg)