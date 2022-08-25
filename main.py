import pandas as pd
import smtplib as smt
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from getpass import getpass

print("Reading mails from Excel")
data = pd.read_excel("info.xlsx")
emailData = data.get("Email")
emailList = list(emailData)

try:
    # SMTP Object
    print("In process...")
    mailServer = smt.SMTP("smtp.gmail.com", 587)
    mailServer.starttls()

    # Inputting email and password
    senderMail = input("Enter your Email address: ")
    senderPass = getpass("Input your password: ")
    # login via email
    mailServer.login(senderMail, senderPass)
    print("Login Sucessful")
    # Setting up the email subject, to, from, and message
    fromEmail = senderMail
    toEmail = emailList
    Emailmessage = MIMEMultipart("alternative")
    Emailmessage['Subject'] = "Testing for Python Project"
    Emailmessage['from'] = senderMail

    # Creating the text via HTML
    messageHTML = '''
    <html>
        <head>
            <style>
                .btn{
                  background:red;
                }
            </style>
        </head>
        <body>
            <h1>HEading 1</h1>
            <button id="btn">Test</button>
        </body>
    </html>
    '''

    # Attach the message and Multipart
    textMsg = MIMEText(messageHTML, "html")
    Emailmessage.attach(textMsg)

    # Sending the email
    mailServer.sendmail(fromEmail, toEmail, Emailmessage.as_string)
    for email in emailList:
        print("Message has been sent to ", emailList[email])

except Exception as error:
    print(error)
