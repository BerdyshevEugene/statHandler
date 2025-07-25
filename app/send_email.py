import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

from logger.logger import setup_logger

logger = setup_logger(module_name=__name__)

EMAIL_RECIPIENTS = os.getenv(
    "EMAIL_RECIPIENTS", "email1@example.com,email2@example.com,email3@example.com"
).split(",")
FROM_ADDRESS = os.getenv("FROM_ADDRESS", "your_email@example.com")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD", "your_password")
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))


def send_email_with_attachment(subject, body, file_path):
    """
    Отправляет email с вложением на несколько адресов
    """
    try:
        msg = MIMEMultipart()
        msg["From"] = FROM_ADDRESS
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain"))

        filename = os.path.basename(file_path)

        if filename.lower().endswith(".xlsx"):
            maintype = "application"
            subtype = "vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        elif filename.lower().endswith(".xls"):
            maintype = "application"
            subtype = "vnd.ms-excel"
        else:
            maintype = "application"
            subtype = "octet-stream"

        logger.info(f"attaching file: {filename} with MIME type: {maintype}/{subtype}")

        with open(file_path, "rb") as attachment:
            part = MIMEBase(maintype, subtype)
            part.set_payload(attachment.read())

        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="{filename}"')
        msg.attach(part)

        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(FROM_ADDRESS, EMAIL_PASSWORD)

        for to_address in EMAIL_RECIPIENTS:
            to_address = to_address.strip()
            msg["To"] = to_address
            text = msg.as_string()
            server.sendmail(FROM_ADDRESS, to_address, text)
            del msg["To"]

        server.quit()
        logger.success(f"excel file sent to {len(EMAIL_RECIPIENTS)} addresses")

    except Exception as e:
        logger.error(f"failed to send file via email: {e}")
