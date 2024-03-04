import fs
import pandas as pd
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

global counttryaccess
counttryaccess = 0

def readfile_controle_bases(counttryaccess):
    # Definir as informações de conexão SMB
    servidor = '192.168.0.100'
    port = '445'
    server_name = 'server'
    user = 'relatorio'
    pwd = 'MY_PWD'
    compartilhamento = 'Dados/FOLDER/FOLDER'
    caminho_arquivo = 'FILE.xlsx'

    try:
        # Abrir o sistema de arquivos remoto
        smb_fs = fs.open_fs(f'smb://{user}:{pwd}@{servidor}:{port}/{compartilhamento}?direct-tcp=True')

        # Abrir o arquivo Excel como um objeto de arquivo
        with smb_fs.open(caminho_arquivo, 'rb') as FILE_excel:
            # Ler o arquivo Excel usando o pandas
            df = pd.read_excel(FILE_excel)
        return df

    except Exception as e:
        print(f"Erro ao acessar o diretório remoto: {e}")
        counttryaccess += 1
        if counttryaccess < 5:
            readfile_controle_bases(counttryaccess)

def send_email(df):
    df['Validade'] = pd.to_datetime(df['Validade'], format='%Y-%m-%d', errors='coerce')

    data_atual = datetime.now()

    # Filtrar os dados como no reader.py
    result = df[df['Validade'] <= data_atual]
    result_orderby = result.sort_values(by='Solicitante', ascending=True)
    #print(result_orderby[['Chamado','Solicitante','Base', 'Validade']])

    # Preparar o corpo do e-mail em formato HTML
    html_content = "<h1>Relatório de Bases Vencidas</h1>"
    html_content += "<p>Segue abaixo relatório com as bases vencidadas.</p>"
    html_content += result_orderby[['Chamado','Solicitante','Base', 'Validade']].to_html(index=False)

    # Configurar e enviar o e-mail
    sender_email = "report-noreply@gmail.com"
    receiver_email = "n2@gmail.com"
    password = "MY_PWD_EMAIL"

    message = MIMEMultipart("alternative")
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = "Relatório de Bases venciadas - Homolog/Testes"

    # Adicionar o corpo do e-mail em formato HTML
    message.attach(MIMEText(html_content, "html"))

    # Conectar-se ao servidor SMTP e enviar o e-mail
    with smtplib.SMTP("smtp.GMAIL.COM", 587) as server:
        #server.starttls()
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, message.as_string())

    #print("E-mail enviado com sucesso!")

if __name__ == "__main__":
    #Apenas envias as segunda-feiras
    if datetime.now().weekday() == 0:
        send_email(readfile_controle_bases(counttryaccess))