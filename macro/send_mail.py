from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib

def send(name, email, mail=None, excel=None, libreoffice=None):
    if mail is None:
        server = "smtp.gmail.com"
        port = "587"
        username = "your@mail.com"
        password = "yourpassword"
        sender = email
        receiver = "receiver@mail.com"
        subject = "Macro Contribute"
        msg = MIMEMultipart('alternative')
        html_message = """<body style="margin: 0; padding: 0;">
        <table border="1">
        <tbody>
        <tr>
        <td>Name</td>
        <td>""" + name + """</td>
        </tr>
        <tr>
        <td>E-mail</td>
        <td>""" + email + """</td>
        </tr>
        <tr>
        <td>Excel VBA</td>
        <td>""" + excel + """</td>
        </tr>
        <tr>
        <td>LibreOffice Basic</td>
        <td>""" + libreoffice + """</td>
        </tr>
        </tbody>
        </table>
        </body>"""
        msg['To'] = receiver
        msg['From'] = sender
        msg['Subject'] = subject
        message = MIMEText(html_message, 'html')
        msg.attach(message)
        message = msg.as_string()

        s = smtplib.SMTP(server, port)
        s.starttls()
        s.ehlo()
        try:
            s.login(username, password)
            s.sendmail(sender, receiver, message)
            s.quit()
            return 1
        except TimeoutError:
            return 0
    else:
        server = "smtp.gmail.com"
        port = "587"
        username = "your@mail.com"
        password = "yourpassword"
        sender = email
        receiver = "receiver@mail.com"
        subject = "Macro Contact Mail"
        msg = MIMEMultipart('alternative')
        html_message = """<body style="margin: 0; padding: 0;">
        <table border="1">
        <tbody>
        <tr>
        <td>Name</td>
        <td>""" + name + """</td>
        </tr>
        <tr>
        <td>E-mail</td>
        <td>""" + email + """</td>
        </tr>
        <tr>
        <td>Mail</td>
        <td>""" + mail + """</td>
        </tr>
        </tbody>
        </table>
        </body>"""
        msg['To'] = receiver
        msg['From'] = sender
        msg['Subject'] = subject
        message = MIMEText(html_message, 'html')
        msg.attach(message)
        message = msg.as_string()

        s = smtplib.SMTP(server, port)
        s.starttls()
        s.ehlo()
        try:
            s.login(username, password)
            s.sendmail(sender, receiver, message)
            s.quit()
            return 1
        except TimeoutError:
            return 0
