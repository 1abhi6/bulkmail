# Importing required modules

import pandas as pd
import smtplib as smt
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import re
from tkinter import *

# Regular expression for validating the mail
regex = re.compile(
    r"([-!#-'*+/-9=?A-Z^-~]+(\.[-!#-'*+/-9=?A-Z^-~]+)*|\"([]!#-[^-~ \t]|(\\[\t -~]))+\")@([-!#-'*+/-9=?A-Z^-~]+(\.[-!#-'*+/-9=?A-Z^-~]+)*|\[[\t -Z^-~]*])")

# Function to validate the mail using Regex module


def isValid(email):
    if re.fullmatch(regex, email):
        return email


def readMails():
    # Reading mails from excel file
    # print("Reading mails from Excel")
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
        lambda emailValue: emailValue != None, validEmailListUnFiltered
    )
    listOfFinalEmails = list(validEmailListFiltered)
    return listOfFinalEmails


def sendEmail():
    try:
        listOfFinalEmails = readMails()
        print("Mail will be send of these emails:", listOfFinalEmails)
        # Reading message from HTML file
        HTMLFile = open("index.html", "r")
        messageHTML = HTMLFile.read()

        # SMTP Object
        print("In process...")
        mailServer = smt.SMTP("smtp.gmail.com", 587)
        mailServer.starttls()  # Starting the server

        # Inputting Email and Password
        # senderMail = input("Enter your Email address: ")
        # senderPass = getpass("Input your password: ")

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
        print("sent")
        if(mailServer.send_message):
            return True
        else:
            return False

        # Displaying the successful message
        # for email in range(len(listOfFinalEmails)):
        #     print("Message has been sent to", listOfFinalEmails[email])
        # return True
    except Exception as error:
        print(error)

# Creating GUI
window = Tk()
window.geometry("500x350")
windowIcon = window.iconbitmap("icon.ico")
# Setting icon of master window
# window.iconphoto(False, windowIcon)

# This function will get the  sender's email credentials from GUI
def getEmailData():
    
    global senderMail
    senderMail = email.get()
    global senderPass
    senderPass = passw.get()
    # readMails() will return true if the mail is send successfully else false
    retsendEmail = sendEmail()

    # If mail is sent successfully if block will execute
    if(retsendEmail):
        dispMsg = readMails()
        Label(window, text="Email Sent Successfully!!!",
              font=('Arial 12')).place(x=60, y=190)
        Label(window, text="Message has been sent to: ",
              font=('Arial 12')).place(x=60, y=220)
        posy = 0
        for i in range(len(dispMsg)):
            Label(window, text=str(dispMsg[i]+"\n"),
                  font=('Arial 12')).place(x=60, y=245+posy)
            posy = posy+25
    # If mail is not sent
    else:
        Label(window, text="Not sent, Due to invalid login credentials.",
              font=('Arial 12')).place(x=60, y=190)
    
#  Create widgets
email = StringVar()
passw = StringVar()

# Label widget for email and password
emailLabel = Label(window, text="Email", font=('Arial 12')).place(x=30, y=50)
password = Label(window, text="Password", font=('Arial 12')).place(x=30, y=90)

# Entry widget field for email and password
entry_variable = Entry(window, textvariable=email, width=20,
                       font=('Arial 12')).place(x=120, y=50)
entry_variable = Entry(window, textvariable=passw, width=20,
                       font=('Arial 12'), show="*").place(x=120, y=90)

# Button widget to send the email
button_submit = Button(window, text="Send Email", font=('Arial 12'), padx=20,bg='#ffffff', activebackground='#00ff00', command=getEmailData).place(
    x=130, y=140
)
window.mainloop()
