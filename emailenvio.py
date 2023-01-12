#Envio de e-mail com documentos - Automatico#

#Packages Necessarios 
import smtplib
import email.message
import csv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

#Listas
arquivos = ["qtdcrtdia.xlsx","qtdcrtmes.xlsx"] #Arquivos
lista = [] #Lista de destinatarios do .txt

#Leitura do .txt de remetendes
def ler_txt():
    with open('emails.txt') as file:
        while line := file.readline():
            lista.append(line)

#Envio de E-mail
def enviar_email():
    ler_txt() #chama a funcao de leitura de txt
    for emails in lista: #envia o e-mail um a um para cada destinatario 
        corpo_email = """
        <p>Ola,</p>
        <p></p>
        <p>Em anexo o relatorio de container por usuario.</p>
        <p></p>
        <p>Atenciosamente,</p>
        <p></p>
        <p>Relatorio de envio automatico, nao responder.</p>
        """ # cria o corpo de email em html
        msg = MIMEMultipart() #chama da funcao de e-mail
        msg['Subject'] = "Relatorio de Container" #Assunto
        msg['From'] = 'destinaraio@teste.com' #Remetende
        msg['To'] = emails #Destinatario de acordo com a lista
        password = 'senha' #Senha do e-mail
        msg.attach(MIMEText(corpo_email,'html')) #Anexa o corpo de e-mail

        for arquivo in arquivos: #Anexa arquivo um por um
            attachement = open(arquivo, 'rb')
            att = MIMEBase('application', 'octet-strem')
            att.set_payload(attachement.read())
            encoders.encode_base64(att)
            att.add_header('Content-Disposition',f'attachement; filename= '+str(arquivo))
            attachement.close()
            msg.attach(att)       

        #Configuracoes de servidor de e-mail SMPTP
        s = smtplib.SMTP('smtp.teste.com: 111') #smtp
        s.starttls() #inicia o servico
        s.login(msg['From'], password) #faz o login
        s.sendmail(msg['From'], [msg['To']], msg.as_string()) #envia o e-mail
        s.quit() #encerra o servico
        print('Email enviado para:'+str(emails)) # printa no terminal a confirmação de envio para o destinatario


# Mauricio Lascoski -