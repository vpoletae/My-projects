import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
from datetime import datetime
import os

      
def send_mail(files, send_to, subject, send_from=SEND_FROM,
                    text_1=TEXT_1, text_2=TEXT_2, path=PATH):
      if files:
            msg = MIMEMultipart()
            msg['From'] = send_from
            msg['To'] = COMMASPACE.join(send_to)
            msg['Date'] = formatdate(localtime=True)
            msg['Subject'] = subject
            msg.attach(MIMEText(text_1))

            for file in files:
                  attachment = MIMEBase('application', "octet-stream")
                  attachment.set_payload(open(os.path.join(path, file), "rb").read())
                  encoders.encode_base64(attachment)
                  attachment.add_header('Content-Disposition', 'attachment', filename=file)
                  msg.attach(attachment)

            smtp = smtplib.SMTP('outbound.cisco.com', 25)
            smtp.sendmail(send_from, send_to, msg.as_string())
            smtp.close()
      else:
            msg = MIMEMultipart()
            msg['From'] = send_from
            msg['To'] = COMMASPACE.join(send_to)
            msg['Date'] = formatdate(localtime=True)
            msg['Subject'] = subject
            msg.attach(MIMEText(text_2))

            smtp = smtplib.SMTP('change', 25)
            smtp.sendmail(send_from, send_to, msg.as_string())
            smtp.close()



































































      
      
