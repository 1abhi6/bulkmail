# Importing required modules

from asyncio import constants
import pandas as pd
import smtplib as smt
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from getpass import getpass
import re

# Regular expression for validating the mail
regex = re.compile(
    r"([-!#-'*+/-9=?A-Z^-~]+(\.[-!#-'*+/-9=?A-Z^-~]+)*|\"([]!#-[^-~ \t]|(\\[\t -~]))+\")@([-!#-'*+/-9=?A-Z^-~]+(\.[-!#-'*+/-9=?A-Z^-~]+)*|\[[\t -Z^-~]*])")

# Function to validate the mail using Regex module


def isValid(email):
    if re.fullmatch(regex, email):
        return email


try:
    # Reading mails from excel file
    print("Reading mails from Excel")
    data = pd.read_excel("info.xlsx")
    emailData = data.get("Email")
    emailList = list(emailData)

    # emailList contains all the mails read from the excel file
    # Now, we will check each mail using isValid function
    validEmail = map(isValid, emailList)
    validEmailListUnFiltered = list(validEmail)

    # validEmailListUnFiltered returns the valid mails as well as invalid mails as none in a list
    # To filter all the valid mails we will use filter() function
    validEmailListFiltered = filter(
        lambda emailValue: emailValue != None, validEmailListUnFiltered)
    listOfFinalEmails = list(validEmailListFiltered)

    print("Mail will be send of these emails:", listOfFinalEmails)

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
    toEmail = listOfFinalEmails
    Emailmessage = MIMEMultipart("alternative")
    Emailmessage['Subject'] = "Testing for Python Project"
    Emailmessage['from'] = senderMail

    # Login via email
    mailServer.login(senderMail, senderPass)
    print("Login Successful")

    # Attach the message and Multipart
    textMsg = MIMEText(messageHTML, "html")
    Emailmessage.attach(textMsg)

    # Sending the email
    mailServer.sendmail(fromEmail, toEmail, Emailmessage.as_string())
    mailServer.quit()

    # Displaying the successful message
    for email in range(len(listOfFinalEmails)):
        print("Message has been sent to", listOfFinalEmails[email])


except Exception as error:
    print(error)
