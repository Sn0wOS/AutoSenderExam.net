from sys import modules
from PyQt5.QtWidgets import QApplication, QMainWindow,\
    QFileDialog, QPushButton, QFileDialog, \
    QLabel, QTableView
from PyQt5.QtCore import Qt
from PyQt5.QtCore import QAbstractTableModel
import sys
import os
import csv
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate


class TableModel(QAbstractTableModel):
    def __init__(self, data):
        super(TableModel, self).__init__()
        self._data = data

    def data(self, index, role):
        if role == Qt.DisplayRole:
            # See below for the nested-list data structure.
            # .row() indexes into the outer list,
            # .column() indexes into the sub-list
            return self._data[index.row()][index.column()]

    def rowCount(self, index):
        # The length of the outer list.
        return len(self._data)

    def columnCount(self, index):
        # The following takes the first sub-list, and returns
        # the length (only works if all rows are an equal length)
        return len(self._data[0])


class App(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.title = "מערכת שליחת מיילים של קבצי בדיקה"
        self.left = 10
        self.top = 10
        self.width = 700
        self.height = 700
        self.csv_directory = ""
        self.answers_directory = ""
        self.dataTableSrc = []
        self.studentList = []
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        self.csv_load_btn = QPushButton("הוסף רשימת תלמדים", self)
        self.csv_file_name = QLabel(self)
        self.csv_file_name.setText("לא נבחר")

        self.folder_load_btn = QPushButton("הוסף תיקיית קבצים", self)
        self.folder_dir_name = QLabel(self)
        self.folder_dir_name.setText("לא נבחר")

        self.send_btn = QPushButton("שלח", self)

        self.table1 = QTableView(self)

        self.model = TableModel([])

        self.csv_load_btn.setGeometry(50, 50, 150, 25)
        self.csv_file_name.setGeometry(200, 50, 200, 25)

        self.folder_load_btn.setGeometry(50, 100, 150, 25)
        self.folder_dir_name.setGeometry(200, 100, 200, 25)

        self.send_btn.setGeometry(150, 150, 200, 25)

        self.table1.setGeometry(0, 200, self.width, 500)

        self.csv_load_btn.setToolTip("רשימת תלמידים שהופק על ידי exam.net")
        self.folder_load_btn.setToolTip(
            "תיקייה המכילה את הקבצים של התשובות")

        self.csv_load_btn.clicked.connect(self.open_dialog_get_students_csv)
        self.folder_load_btn.clicked.connect(self.open_dialog_get_answer_pdfs)

        self.send_btn.clicked.connect(
            lambda: self.send_mails("shlomo", "test", "some notes"))

        self.show()

    def open_dialog_get_students_csv(self):
        fileDialog = QFileDialog()
        # getOpen file return tuple (fileName,filter)
        fileName, _ = fileDialog. getOpenFileName(
            self, "בחר קובץ רשימת תלמדים", "", "Files (*.csv)")
        self.csv_directory = fileName
        self.csv_file_name.setText(os.path.basename(fileName))
        self.updateList()
        self.table1.setModel(TableModel(self.dataTableSrc))

    def open_dialog_get_answer_pdfs(self):
        fileDialog = QFileDialog()
        directory = fileDialog.getExistingDirectory(
            self, "הוסף תיקיית תשובות", "Desktop", QFileDialog.ShowDirsOnly)
        self.answers_directory = directory
        self.folder_dir_name.setText(os.path.basename(directory))
        self.updateList()
        self.table1.setModel(TableModel(self.dataTableSrc))

    def __getStudentListFromCSV(self, csv_directory: str) -> list:
        csvList = []
        childDict = {'FirstName': '', 'LastName': '',
                     'Email': '', 'file': 'Exam File', "grade": "grade"}
        try:
            with open(csv_directory) as csvFile:
                reader = csv.reader(csvFile, delimiter=';')
                for row in reader:
                    childDict['FirstName'] = row[0]
                    childDict['LastName'] = row[1]
                    childDict['Email'] = row[2]
                    csvList.append(childDict.copy())
        except Exception as e:
            print(e)
            return None

        return csvList

    def __findFilesInDirectory(self, answers_directory, studentList) -> list:
        for _, _, fileNames in os.walk(answers_directory):
            for name in fileNames:  # itrate on all files in given directory
                for student in studentList:  # iterate on all student name
                    # test that sure and last name apper in the file
                    if student['FirstName'] in name and student['LastName'] in name:
                        # print('email:', student['Email'], 'file', name)
                        if student['file'] == "Exam File":
                            student['file'] = os.path.join(
                                answers_directory, name)
                        else:
                            student['file'] = student['file']+'&'+os.path.join(
                                answers_directory, name)
        return studentList

    def updateList(self):
        self.studentList = self.__getStudentListFromCSV(self.csv_directory)
        if isinstance(self.studentList, list):
            self.studentList = self.__findFilesInDirectory(
                self.folder_dir_name, self.studentList)
            if isinstance(self.studentList, list):
                self.dataTableSrc = self.list_of_dict_to_list(self.studentList)

    def send_mails(self, From, examName, notes=""):
        bodyBaseString = "{}\n\
            שלום ל{} \n\
            להלן מצורפת הבחינה שלך ב{}\n\
            בהצלחה להבא מ{}"
        for student in self.studentList:
            bodyBaseString.format(
                notes, student['FirstName']+" "+student["LastName"], examName, f)
            self.__send_mail(From,
                             student["Email"], examName, bodyBaseString, student['file'])

    def __send_mail(sent_from, to, subject, body, filename=None):
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
                                part = MIMEApplication(
                                    f.read(), _subtype="pdf")
                            part.add_header('Content-Disposition', 'attachment',
                                            filename=os.path.basename(name))
                            msg.attach(part)
                else:
                    with open(filename, "rb") as f:
                        part = MIMEApplication(f.read(), _subtype="pdf")
                    part.add_header('Content-Disposition', 'attachment',
                                    filename=os.path.basename(filename))
                    msg.attach(part)
            else:
                return
            server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
            server.ehlo()
            server.login(gmail_user, gmail_password)
            server.sendmail(sent_from, to, msg.as_string())
            server.close()

            print('Email sent!')
        except Exception as e:
            print('Something went wrong...\n', e)

    def list_of_dict_to_list(self, src: list) -> list:
        temp = []
        ret = []
        for s in src:
            temp = []
            for _, value in s.items():
                temp.append(value)

            ret.append(temp)
        return ret


if __name__ == "__main__":
    app = QApplication([])
    ex = App()
    sys.exit(app.exec_())
