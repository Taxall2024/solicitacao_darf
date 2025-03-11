import os
import pickle
import base64
import streamlit as st
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
import datetime
from email.header import Header

SCOPES = ["https://www.googleapis.com/auth/gmail.send"]
#CREDENTIALS_PATH = 'credentials.json'
TOKEN_PATH = "token.pickle"

def get_credentials():
  
    if os.path.exists(TOKEN_PATH):
        with open(TOKEN_PATH, "rb") as token:
            creds = pickle.load(token)
        if creds and creds.valid:
            return creds  
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())  
            with open(TOKEN_PATH, "wb") as token:
                pickle.dump(creds, token)
            return creds

    # Caso contrário, peça autenticação
    google_secrets = st.secrets["google_oauth"]

    credentials_info = {
        "web": {
            "client_id": google_secrets["client_id"],
            "project_id": google_secrets["project_id"],
            "auth_uri": google_secrets["auth_uri"],
            "token_uri": google_secrets["token_uri"],
            "auth_provider_x509_cert_url": google_secrets["auth_provider_x509_cert_url"],
            "client_secret": google_secrets["client_secret"],
            "redirect_uris": google_secrets["redirect_uris"],
        }
    }

    flow = InstalledAppFlow.from_client_config(credentials_info, SCOPES)
    creds = flow.run_local_server(port=8080)

    # Salva o token para evitar reautenticação repetitiva
    with open("token.pickle", "wb") as token:
        pickle.dump(creds, token)

    return creds


# Lista de empresas fornecidas
empresas = [
    "3EX", "Alfa e Omega", "ANGELO SOM", "ALFASEG", "APPSERVICE", "ATIVA SYSTEM", "BRASFORT ADM", "BRASFORT SEG", 
    "CARISE", "CIENGE ENGENHARIA", "CIENGE TELECOM", "CILIA", "CRETA", "DIGIX", "ESPAÇO MARKETING", 
    "EURO SEGURANÇA", "EURO SERV", "EUROSEG", "FLAVIA", "G&E", "G9", "GLOBAL SERVIÇOS", "GLOBALTECH", 
    "GRUPO 5 ESTRELAS", "GSI SERVIÇOS", "INCRO", "JAUA", "JAVA", "LEMARC", "LIMPORT MACAÉ", "LPT GESTÃO", 
    "M5 SEGURANÇA", "MASTROS SEGURANÇA", "MASTROS SERVIÇOS", "MÁXIMA", "MAXTEC", "MUNDO DIGITAL", 
    "ORCALI LIMPEZA", "ORCALI SEGURANÇA", "ORCALI SERVIÇOS", "PALMACEA", "PM LOCAÇÕES", "PMT", "PORTLIMP", 
    "QUALITY MAX", "RIO GRANDENSE", "SOLLO SERVIÇOS", "TRIUNFO SEGURANÇA", "TRIUNFO SERVIÇOS", 
    "TRIUNFO SERVIÇOS", "ULTRA", "UNINEURO NEUROCIRURGIA", "UNISERV", "VISAN", "WGA BIO", "WGA SEGURANÇA"
]

# Obter o mês e ano atual
hoje = datetime.datetime.now()
competencia = (hoje.replace(day=1) - datetime.timedelta(days=1)).strftime("%m/%Y")

def autenticar_gmail():
    creds = None


    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)
    else:
        creds = get_credentials()

    # Garante que o token ainda é válido
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            creds = get_credentials()

    # Autenticando com Gmail API
    service = build("gmail", "v1", credentials=creds)

    st.write("Autenticado com sucesso no Gmail API!")

    return creds

def enviar_email(destinatario, assunto, corpo_email, nome_do_arquivo_pdf=None):
    creds = autenticar_gmail()
    if not creds:
        st.error("Falha na autenticação do Gmail.")
        return
    
    service = build("gmail", "v1", credentials=creds)
    
    msg = MIMEMultipart()
    # Codificando o assunto corretamente
    msg["Subject"] = Header(assunto, "utf-8").encode()
    msg["From"] = "bruno.cardoso@taxall.com.br, negocios@taxall.com.br"
    msg["To"] = f"{destinatario}, negocios@taxall.com.br"

    # Converta a imagem para base64 e a insira no corpo do e-mail
    with open("bruno_cardoso.jpg", "rb") as img_file:
        img_data = img_file.read()
        img_base64 = base64.b64encode(img_data).decode()

    corpo_html = f"""
    <html>
        <head></head>
        <body style="font-family: Tahoma; font-size: 14px;">
            <p>Prezados,</p>
            <p>Esperamos que estejam bem.</p>
            <p>Para darmos continuidade aos trabalhos da competência {competencia}, solicitamos a gentileza de nos encaminharem os seguintes documentos:</p>
            <ul>
                <li>{'</li><li>'.join(documentos)}</li>
            </ul>
            <p>Caso haja alguma dúvida ou necessidade de esclarecimento adicional, estamos à disposição para auxiliá-los.</p>
            <p>Agradecemos pela colaboração e ficamos no aguardo dos documentos.</p>
            <p>Atenciosamente,</p>
            <p><img src="data:image/jpg;base64,{img_base64}" alt="Assinatura" /></p>
        </body>
    </html>
    """
    
    msg.attach(MIMEText(corpo_html, "html", _charset="utf-8"))
    
    if nome_do_arquivo_pdf and os.path.exists(nome_do_arquivo_pdf):
        with open(nome_do_arquivo_pdf, "rb") as f:
            arquivo_pdf = f.read()
            arquivo_anexado = MIMEApplication(arquivo_pdf, _subtype="pdf")
            arquivo_anexado.add_header("Content-Disposition", f"attachment; filename={os.path.basename(nome_do_arquivo_pdf)}")
            msg.attach(arquivo_anexado)

    raw_message = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    message = {"raw": raw_message}

    service.users().messages().send(userId="me", body=message).execute()
    st.success("Email enviado com sucesso!")

st.title("Formulário de Solicitação de Documentos")
empresa = st.selectbox("Selecione a Empresa", empresas)
email_cliente = st.text_input("Email do Cliente")
documentos = st.multiselect("Documentos Necessários", [
    "DARF - INSS", "DARF - IRRF", "DARF - IRPJ", "DARF - CSLL", "Planilha de Retenções", "Notas Fiscais", "DARF - PIS", "DARF - COFINS", "Folha de Pagamento", 
    "Apuração de PIS e COFINS"
])

if st.button("Enviar Solicitação"):
    if empresa and email_cliente and documentos:
        corpo_email = f"""
        Prezados, boa tarde!
        
        Para realização dos trabalhos da competência {competencia}, solicitamos o envio dos arquivos:
        
        • {', '.join(documentos)}
        
        Qualquer dúvida estamos à disposição.
        
        Atenciosamente,
        """
        
        assunto = f"TAX ALL - Solicitação de Documentos - {empresa} - {competencia}"
        # Alteração do encoding do assunto
        assunto = Header(assunto, "utf-8").encode()
        enviar_email(email_cliente, assunto, corpo_email)
    else:
        st.error("Por favor, preencha todos os campos!")
