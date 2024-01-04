from tkinter import *
from tkinter.font import BOLD
from PIL import Image, ImageFont, ImageDraw

import pandas
import smtplib

from email import encoders
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase

from pandas.io import excel

gui = Tk()
gui.title("CERTIFICATE")
gui.resizable(0, 0)

frame = Canvas(gui, width=360, height=380, bg='#b2d8d8', relief='raised')
frame.pack()

label1 = Label(gui, text='WELCOME', bg='#b2d8d8', fg='RoyalBlue4')
label1.config(font=('Times New Roman', 25, BOLD))
frame.create_window(180, 60, window=label1)

label2 = Label(gui, text="Let's get started!", bg='#b2d8d8')
label2.config(font=('Times New Roman', 14))
frame.create_window(180, 100, window=label2)

label3 = Label(gui, bg='#b2d8d8')
label3.config(font=('Arial', 12, BOLD))
frame.create_window(180, 250, window=label3)


def savePDF():
    global file, saved_name, imgName
    file = pandas.read_excel(
        r"details.xlsx")
    for i in file.index:
        now = pandas.Timestamp.today()
        joined = file["Date of Joining"][i]
        period = (now - joined).days
        if period >= 365:
            name = file["Name"][i]
            img = Image.open(r"Sample_Certificate.png").convert('RGB')
            certificate = ImageDraw.Draw(img)
            location = (260, 190)
            font = ImageFont.truetype("times.ttf", 20)
            certificate.text(location, name, (0, 0, 0), font=font)
            saved_name = 'Certificate_' + name
            imgName = "Certificate_" + name + ".pdf"
            img.save(imgName)

        # E-mail the pdf as an attachment
        sender = "mahindraaastha1716@gmail.com"
        receiver = file["E-mail"][i]

        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = receiver
        msg['Subject'] = "Certificate of Participation"

        # body of the mail
        body = '''Thank you for participating
                                
Thanks and Regards
Team _________
'''
        msg.attach(MIMEText(body, 'plain'))

        # Open the file to be sent
        filename = saved_name + ".pdf"
        attachment = open(imgName, "rb")
        p = MIMEBase('application', 'octet-stream')

        # To change the payload into encoded form
        p.set_payload((attachment).read())
        # encode into base64
        encoders.encode_base64(p)
        p.add_header('Content-Disposition',
                     "attachment; filename= %s" % filename)

        # attach the instance 'p' to instance 'msg'
        msg.attach(p)

        # creates SMTP session
        try:
            s = smtplib.SMTP('smtp.gmail.com', 587)
            s.ehlo()
        except:
            print('Something went wrong...')

        # start TLS for security
        s.starttls()

        # Authentication
        s.login(sender, 'Aastha16!')

        # Converts the Multipart msg into a string
        text = msg.as_string()

        # Sending the mail
        s.sendmail(sender, receiver, text)
        #print('sent to : ' + name)
        s.quit()
    label3.config(text="'Certificate Successfully Generated'\nE-mail Sent!")


btn = Button(text="     Generate Certificate     ", command=savePDF,
             cursor="hand2", bg='RoyalBlue3', fg='white', font=('Arial', 12, 'bold'))
frame.create_window(180, 180, window=btn)

gui.mainloop()
