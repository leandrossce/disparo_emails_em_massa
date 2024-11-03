import smtplib
import ssl
import csv
import mimetypes
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage  # Import necessário para imagens
from email.mime.base import MIMEBase
from email import encoders

# Informações do remetente e destinatário
email_remetente = "e-mail-de-origem@gmail.com"
senha = "xxxxxxxxxxxxx"  # Use a senha de aplicativo gerada

# Caminho do arquivo CSV
caminho_csv = "C:\\Users\\Gabriel\\Downloads\\pdf\\dados.csv"  # Substitua pelo caminho do seu arquivo CSV

# Caminho da imagem a ser exibida no corpo do e-mail
caminho_imagem = "imagem.jpg"  # Substitua pelo caminho da sua imagem

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
        caminho_imagem = linha[4]

        # Verificar se o e-mail está vazio
        if not email_destinatario:
            print(f"E-mail não encontrado para a empresa {nome_empresa}. Pulando.")
            continue

        # Criar a mensagem MIMEMultipart
        mensagem = MIMEMultipart("related")  # Usamos "related" para incluir imagens inline
        mensagem["From"] = email_remetente
        mensagem["To"] = email_destinatario
        mensagem["Subject"] = "Assunto do E-mail"

        # Criar a parte alternativa para texto e HTML
        mensagem_alternativa = MIMEMultipart("alternative")
        mensagem.attach(mensagem_alternativa)

        # Corpo em texto simples
        '''
        texto_plain = f"""Prezado(a) {nome_empresa},

Segue em anexo o documento referente ao CNPJ {cnpj}.

Atenciosamente,
Sua Empresa
"""
        
        parte_texto = MIMEText(texto_plain, "plain")
        mensagem_alternativa.attach(parte_texto)
        '''
        # Verificar se a imagem existe
        if os.path.isfile(caminho_imagem):
            # Ler a imagem em modo binário
            with open(caminho_imagem, "rb") as img:
                mime_image = MIMEImage(img.read())
                # Definir o Content-ID da imagem
                mime_image.add_header("Content-ID", "<imagem_corpo>")
                mime_image.add_header("Content-Disposition", "inline", filename=os.path.basename(caminho_imagem))
            # Anexar a imagem à mensagem principal
            mensagem.attach(mime_image)
        else:
            print(f"Imagem não encontrada: {caminho_imagem}. O e-mail será enviado sem a imagem.")
            # Você pode optar por continuar ou pular o envio neste caso

        # Corpo em HTML, referenciando a imagem pelo Content-ID
        corpo_html = f"""
        <html>
          <body>
            <p>Prezado(a) {nome_empresa},</p>
            <p>Segue em anexo o documento referente ao CNPJ {cnpj}.</p>
            <img src="cid:imagem_corpo" alt="Imagem" style="width:300px;height:auto;">
            <p>Atenciosamente,</p>
            <p>Sua Empresa</p>
          </body>
        </html>
        """
        parte_html = MIMEText(corpo_html, "html")
        mensagem_alternativa.attach(parte_html)

        # Anexar o arquivo especificado na coluna D
        if os.path.isfile(caminho_arquivo):
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
            with smtplib.SMTP("smtp.gmail.com", 587) as servidor: #with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=contexto) as servidor:
                servidor.starttls(context=contexto)
                servidor.login(email_remetente, senha)

                servidor.sendmail(email_remetente, email_destinatario, mensagem.as_string())
            print(f"E-mail enviado com sucesso para {email_destinatario}!")
        except Exception as e:
            print(f"Ocorreu um erro ao enviar o e-mail para {email_destinatario}: {e}")
