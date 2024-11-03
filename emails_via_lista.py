import smtplib
import ssl
import csv
import mimetypes
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# Informações do remetente e destinatário
email_remetente = "email_de_origem@gmail.com"
senha = "xxx xxx xxx xxx"  # Use a senha de aplicativo gerada para e-mails do google


# Caminho do arquivo CSV
caminho_csv = "C:\\Users\\Gabriel\\Downloads\\pdf\\dados.csv"  # Substitua pelo caminho do seu arquivo CSV

# Ler o arquivo CSV
with open(caminho_csv, newline='', encoding='utf-8') as csvfile:
    leitor_csv = csv.reader(csvfile, delimiter=';')
    next(leitor_csv)  # Pular o cabeçalho, se houver
    
    for linha in leitor_csv:
        # Extrair dados de cada coluna
        cnpj = linha[0]
        nome_empresa = linha[1]
        email_destinatario = linha[2]
        caminho_arquivo = linha[3]
        
        # Verificar se o e-mail está vazio
        if not email_destinatario:
            print(f"E-mail não encontrado para a empresa {nome_empresa}. Pulando.")
            continue
        
        # Criar o conteúdo do e-mail
        assunto = "Assunto do E-mail"  # Você pode personalizar o assunto
        corpo_html = f"""
        <html>
          <body>
            <p>Prezado(a) {nome_empresa},</p>
            <p>Segue em anexo o documento referente ao CNPJ {cnpj}.</p>
            <p>Atenciosamente,</p>
            <p>Sua Empresa</p>
          </body>
        </html>
        """
        
        # Configurar a mensagem
        mensagem = MIMEMultipart()
        mensagem["From"] = email_remetente
        mensagem["To"] = email_destinatario
        mensagem["Subject"] = assunto
        
        # Anexar o corpo do e-mail
        #parte_texto = MIMEText(f"Prezado(a) {nome_empresa},\n\nSegue em anexo o documento referente ao CNPJ {cnpj}.\n\nAtenciosamente,\nSua Empresa", "plain")
        parte_html = MIMEText(corpo_html, "html")
        mensagem.attach(parte_html)
        
        # Verificar se o arquivo existe
        if os.path.isfile(caminho_arquivo):
            # Determinar o tipo MIME
            tipo_mime, _ = mimetypes.guess_type(caminho_arquivo)
            if tipo_mime is None:
                tipo_mime = "application/octet-stream"
            tipo_principal, subtipo = tipo_mime.split("/", 1)
            
            with open(caminho_arquivo, "rb") as anexo:
                parte_anexo = MIMEBase(tipo_principal, subtipo)
                parte_anexo.set_payload(anexo.read())
            encoders.encode_base64(parte_anexo)
            parte_anexo.add_header(
                "Content-Disposition",
                f"attachment; filename= {os.path.basename(caminho_arquivo)}",
            )
            mensagem.attach(parte_anexo)
        else:
            print(f"Arquivo não encontrado: {caminho_arquivo}. Pulando o envio para {email_destinatario}.")
            continue  # Pular para o próximo destinatário
        
        # Conectar ao servidor SMTP do Gmail e enviar o e-mail
        try:
            contexto = ssl.create_default_context()
            with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=contexto) as servidor:
                servidor.login(email_remetente, senha)
                servidor.sendmail(email_remetente, email_destinatario, mensagem.as_string())
            print(f"E-mail enviado com sucesso para {email_destinatario}!")
        except Exception as e:
            print(f"Ocorreu um erro ao enviar o e-mail para {email_destinatario}: {e}")
