import email
import smtplib
import ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd

tabela = pd.read_excel(r'emails.xlsx')
print("Lista de emails a serem enviados:", tabela)

nome = tabela['nome']

for lista in tabela.email:

    mail_content = f'''teste {nome:}'''

    # The mail addresses and password
    sender_address = 'email'
    sender_pass = 'senha'
    receiver_address = lista

    # Setup the MIME
    message = MIMEMultipart()
    message['From'] = sender_address
    message['To'] = receiver_address
    # The subject line
    message['Subject'] = 'Isso é apenas um email de teste.'
    message.attach(MIMEText(mail_content, 'plain'))

    # Create SMTP session for sending the mail
    session = smtplib.SMTP('smtp.gmail.com', 587)  # use gmail with port
    session.starttls()  # enable security
    # login with mail_id and password
    session.login(sender_address, sender_pass)
    text = message.as_string()
    session.sendmail(sender_address, receiver_address, text)
session.quit()

print('Emails enviados com sucesso!')
