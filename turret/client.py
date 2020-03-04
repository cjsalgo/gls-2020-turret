import docx
import subprocess
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

class TurretClient:

    def __init__(self, sender_email, password):
        self.sender_email = sender_email
        self.password = password

    def normalize(self,name):
        return name.replace(" ", "_")

    def convert_docx_to_pdf(self, name):
        bashCommand = "doc2pdf files/" + name + ".docx"
        process = subprocess.Popen(bashCommand.split(), stdout=subprocess.PIPE)
        output, error = process.communicate()

    def create_certificate(self, name):
        doc = docx.Document('certificate.docx')
        doc.paragraphs[7].runs[4].text = self.normalize(name)
        doc.save('files/' + name + '.docx')

    def send_email(self, name, receiver):
        fromaddr = self.sender_email
        toaddr = receiver

        # instance of MIMEMultipart
        msg = MIMEMultipart()

        # storing the senders email address
        msg['From'] = fromaddr

        # storing the receivers email address
        msg['To'] = toaddr

        # storing the subject
        msg['Subject'] = 'Welcome to GLS Manila South 2020'

        # string to store the body of the mail
        body = '<html><body><img width="50%" height="50%" src="https://drive.google.com/uc?export=view&id=13YP-jf7zta9cjuSoA3TLrwCDqIen3wRn"/><br><br>Hi ' + name + ',<br><br>Welcome to the <b>Global Leadership Summit (GLS) Manila South.</b> We are glad that you are<br>here with us.<br><br>Thank you for providing your contact information to us. To show our appreciation, you are<br>receiving the link to Craig Groschel’s coaching event which will happen on April 2020.<br><br>https://intlsummitreg.regfox.com/online-coaching-with-craig-english<br><br>We hope that Craig Groschel’s coaching event will further help you in your leadership journey.<br><br>Enjoy the Summit!<br><br><b>GLS Manila South Team</b></body></html>'

        # attach the body with the msg instance
        msg.attach(MIMEText(body, 'html', 'utf-8'))

        # open the file to be sent
        # filename = "GLS_Certificate.pdf"
        # attachment = open("files/" + self.normalize(name) + ".pdf", "rb")

        # instance of MIMEBase and named as p
        # p = MIMEBase('application', 'octet-stream')

        # To change the payload into encoded form
        # p.set_payload((attachment).read())

        # encode into base64
        # encoders.encode_base64(p)

        # p.add_header('Content-Disposition', "attachment; filename= %s" % filename)

        # attach the instance 'p' to instance 'msg'
        # msg.attach(p)

        # creates SMTP session
        s = smtplib.SMTP('smtp.gmail.com', 587)

        # start TLS for security
        s.starttls()

        # Authentication
        s.login(fromaddr, self.password)

        # Converts the Multipart msg into a string
        text = msg.as_string()

        # sending the mail
        s.sendmail(fromaddr, toaddr, text)

        # terminating the session
        s.quit()
