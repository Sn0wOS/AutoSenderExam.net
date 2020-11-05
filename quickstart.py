import smtplib
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from os import walk
import csv


def send_mail(sent_from, to, subject, body, filename=None):
    gmail_user = 'shlomo.neuberger'
    gmail_password = 'ehkdetcekeshnjwq'
    try:
        msg = MIMEMultipart()
        msg['From'] = sent_from
        msg['To'] = to
        msg['Date'] = formatdate(localtime=True)
        msg['Subject'] = subject

        msg.attach(MIMEText(body))
        if type(filename) is str:
            if '&' in filename:
                listFiles = filename.split('&')
                for name in listFiles:
                    if name != "":
                        with open(name, "rb") as f:
                            part = MIMEApplication(f.read(), _subtype="pdf")
                        part.add_header('Content-Disposition', 'attachment',
                                        filename=str(name))
                        msg.attach(part)
            else:
                with open(filename, "rb") as f:
                    part = MIMEApplication(f.read(), _subtype="pdf")
                part.add_header('Content-Disposition', 'attachment',
                                filename=str(filename))
                msg.attach(part)

        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.ehlo()
        server.login(gmail_user, gmail_password)
        server.sendmail(sent_from, to, msg.as_string())
        server.close()

        print('Email sent!')
    except Exception as e:
        print('Something went wrong...\n', e)


csvList = []
childDict = {'FirstName': '', 'LastName': '', 'Email': '', 'file': ''}
with open("students-בחינה מיתוג-2.csv") as csvFile:
    reader = csv.reader(csvFile, delimiter=';')
    for row in reader:
        childDict['FirstName'] = row[0]
        childDict['LastName'] = row[1]
        childDict['Email'] = row[2]
        csvList.append(childDict.copy())

for root, di, fileNames in walk('.'):
    for name in fileNames:
        for student in csvList:
            if student['FirstName'] in name:
                if student['LastName'] in name:
                    #print('email:', student['Email'], 'file', name)
                    if student['file'] == "":
                        student['file'] = name
                    else:
                        student['file'] = student['file']+'&'+name
index = 9

"""for student in csvList:
    send_mail("Shlomo exam system",
              student["Email"], "exam of mitug", "your exam", student['file'])
"""
send_mail('grade system shlomo', "shlomo@neuberger.co.il", csvList[index]['FirstName'], csvList[index]
          ['FirstName']+' '+csvList[index]['LastName']+' '+csvList[index]['Email'], csvList[index]['file'])
