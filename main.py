# Importing required modules

from asyncio import constants
import pandas as pd
import smtplib as smt
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from getpass import getpass


try:
    # Reading mails from excel file
    print("Reading mails from Excel")
    data = pd.read_excel("info.xlsx")
    emailData = data.get("Email")
    emailList = list(emailData)

    # Reading message from HTML file
    HTMLFile = open("index.html", "r")
    messageHTML = HTMLFile.read()

    # SMTP Object
    print("In process...")
    mailServer = smt.SMTP("smtp.gmail.com", 587)
    mailServer.starttls()  # Starting the server

    # Inputting Email and Password
    senderMail = input("Enter your Email address: ")
    senderPass = getpass("Input your password: ")

    # Setting up the email subject, to, from, and message
    fromEmail = senderMail
    toEmail = emailList
    Emailmessage = MIMEMultipart("alternative")
    Emailmessage['Subject'] = "Testing for Python Project"
    Emailmessage['from'] = senderMail

    # Login via email
    mailServer.login(senderMail, senderPass)
    print("Login Sucessful")

    # Attach the message and Multipart
    textMsg = MIMEText(messageHTML, "html")
    Emailmessage.attach(textMsg)

    # Sending the email
    mailServer.sendmail(fromEmail, toEmail, Emailmessage.as_string())
    mailServer.quit()

    # Displaying the successful message
    for email in range(len(emailList)):
        print("Message has been sent to", emailList[email])


except Exception as error:
    print(error)
