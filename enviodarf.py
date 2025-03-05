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
CREDENTIALS_PATH = r"C:\Users\bruno.cardoso\Documents\projeto1\credentials.json"
TOKEN_PATH = r"C:\Users\bruno.cardoso\Documents\projeto1\token.pickle"

# Lista de empresas fornecidas
empresas = [
    "2MG", "ADTEL", "GSI SERVIÇOS ESPECIALIZADOS", "ALFA E OMEGA", "ALFASEG VIGILANCIA", "ANGELLO SOM",
    "DINÂMICA FACILITY ADMINISTRAÇÃO PREDIAL LTDA", "ATITUDE", "ESPARTA SEGURANÇA LTDA (ACIMA DE 3 MILHÕES)",
    # Restante da lista...
]

# Obter o mês e ano atual
competencia = datetime.datetime.now().strftime("%m/%Y")

def autenticar_gmail():
    creds = None

    if os.path.exists(TOKEN_PATH):
        with open(TOKEN_PATH, "rb") as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENTIALS_PATH):
                st.error(f"Arquivo de credenciais não encontrado: {CREDENTIALS_PATH}")
                return None
            
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_PATH, SCOPES)
            creds = flow.run_local_server(port=8080)
        
        with open(TOKEN_PATH, "wb") as token:
            pickle.dump(creds, token)

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
    with open(r"C:\Users\bruno.cardoso\Documents\projeto1\bruno_cardoso.jpg", "rb") as img_file:
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
    "DARF - INSS", "DARF - IRRF", "DARF - PIS", "DARF - COFINS", "Folha de Pagamento", 
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
        
        assunto = f"TAX ALL - Solicitação de Documentos - {empresa}"
        # Alteração do encoding do assunto
        assunto = Header(assunto, "utf-8").encode()
        enviar_email(email_cliente, assunto, corpo_email)
    else:
        st.error("Por favor, preencha todos os campos!")
