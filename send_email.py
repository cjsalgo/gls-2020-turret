
# Python code to illustrate Sending mail with attachments 
# from your Gmail account
# libraries to be imported
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

fromaddr = "cjsalgo235@gmail.com"
toaddr = "cjsalgo235@gmail.com"

# instance of MIMEMultipart
msg = MIMEMultipart()

# storing the senders email address
msg['From'] = fromaddr

# storing the receivers email address
msg['To'] = toaddr

# storing the subject
msg['Subject'] = "Subject of the Mail"

# string to store the body of the mail
body = '<html><body><img src="https://drive.google.com/uc?export=view&id=13YP-jf7zta9cjuSoA3TLrwCDqIen3wRn"/><br><br>Hi Gerbert,<br><br>Thank you for attending the Global Leadership Summit (GLS) Manila South held on March 6-7, 2020 at GCF South Metro.<br><br>It was an honor having you considering your busy schedule. We hope that you found the event to be both interesting and informative.<br><br>You may refer to the attached file for your Certificate of Attendance.<br><br>We look forward to seeing you next year. Additionally, please do not hesitate to reach out to us if you are interested to hold GLS huddles in your organizations, companies or churches.<br><br>Thank you once again for being a part of this year\'s Global Leadership Summit.<br><br>Best Regards,<br><br><br><b>GLS Manila South Team<br></body></html>'

# attach the body with the msg instance
msg.attach(MIMEText(body, 'html', 'utf-8'))

# open the file to be sent
filename = "File_name_with_extension.pdf"
attachment = open("files/CJ Salgo.pdf", "rb")

# instance of MIMEBase and named as p
p = MIMEBase('application', 'octet-stream')

# To change the payload into encoded form
p.set_payload((attachment).read())

# encode into base64
encoders.encode_base64(p)

p.add_header('Content-Disposition', "attachment; filename= %s" % filename)

# attach the instance 'p' to instance 'msg'
msg.attach(p)

# creates SMTP session
s = smtplib.SMTP('smtp.gmail.com', 587)

# start TLS for security
s.starttls()

# Authentication
s.login(fromaddr, "Hearthstone2-0-5")

# Converts the Multipart msg into a string
text = msg.as_string()

# sending the mail
s.sendmail(fromaddr, toaddr, text)

# terminating the session
s.quit()
