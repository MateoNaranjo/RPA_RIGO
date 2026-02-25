import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from Config.Senttings import CONFIG_EMAIL
from Repositorios.Correos import CorreosRepo


class EmailCorreos:

    def __init__(
            self, 
            smtp_server: str = None or CONFIG_EMAIL["smtp_server"], 
            smtp_port: int = None or CONFIG_EMAIL["smtp_port"], 
            email: str = None or CONFIG_EMAIL["email"], 
            password: str = None or CONFIG_EMAIL["password"]
        ):
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.email = email
        self.password = password


    def EnviarCorreo(
            self,
            to_emails: str,
            asunto: str,
            body: str,
            is_html: bool = False,
            cc_emails: str = None,
            bcc_emails: str = None
    ):

        msg = MIMEMultipart()
        msg["From"] = self.email
        msg["To"] = to_emails
        msg["Subject"] = asunto

        if cc_emails:
            msg["Cc"] = cc_emails

        tipo = "html" if is_html else "plain"
        msg.attach(MIMEText(body, tipo))

        destinatarios = []

        destinatarios += to_emails.split(",")

        if cc_emails:
            destinatarios += cc_emails.split(",")

        if bcc_emails:
            destinatarios += bcc_emails.split(",")

        with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
            server.starttls()
            server.login(self.email, self.password)
            server.sendmail(self.email, destinatarios, msg.as_string)

    
    def EnviarCorreoCod(self, cod_email):
        
        params = CorreosRepo.ObtenerParametrosCorreo(cod_email)
        
        self.EnviarCorreo(
            to_emails=params["to"],
            cc_emails=params["cc"],
            bcc_emails=params["bcc"],
            asunto=params["asunto"],
            body=params["body"],
            is_html=params["is_html"]
        )