from PyQt5 import QtCore, QtGui, QtWidgets
import os
import sqlite3
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.setFixedSize(602, 240)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("convert.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        MainWindow.setWindowIcon(icon)
        MainWindow.setStyleSheet("background-color: rgb(37, 37, 38);")
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.btn1 = QtWidgets.QPushButton(self.centralwidget, clicked= lambda: self.browsefiles())
        self.btn1.setGeometry(QtCore.QRect(480, 30, 75, 31))
        self.btn1.setStyleSheet("background-color: rgb(51, 51, 51);\n"
        "color: rgb(250, 250, 250);")
        self.btn1.setObjectName("btn1")
        self.lbl1 = QtWidgets.QLineEdit(self.centralwidget)
        self.lbl1.setGeometry(QtCore.QRect(50, 30, 431, 31))
        self.lbl1.setStyleSheet("background-color: rgb(51, 51, 51);\n"
        "color: rgb(250, 250, 250);\n"
        "selection-color: rgb(0, 229, 255);")
        self.lbl1.setObjectName("lbl1")
        self.lbl3 = QtWidgets.QLineEdit(self.centralwidget)
        self.lbl3.setGeometry(QtCore.QRect(50, 160, 261, 31))
        self.lbl3.setStyleSheet("background-color: rgb(51, 51, 51);\n"
        "selection-color: rgb(0, 229, 255);\n"
        "color: rgb(250, 250, 250);")
        self.lbl3.setObjectName("lbl3")
        self.btn2 = QtWidgets.QPushButton(self.centralwidget, clicked= lambda: self.convert())
        self.btn2.setGeometry(QtCore.QRect(210, 60, 111, 21))
        self.btn2.setStyleSheet("background-color: rgb(51, 51, 51);\n"
        "color: rgb(250, 250, 250);")
        self.btn2.setObjectName("btn2")
        self.btn5 = QtWidgets.QPushButton(self.centralwidget, clicked= lambda: self.mail())
        self.btn5.setGeometry(QtCore.QRect(120, 190, 111, 21))
        self.btn5.setStyleSheet("background-color: rgb(51, 51, 51);\n"
        "color: rgb(250, 250, 250);")
        self.btn5.setObjectName("btn5")
        self.btn4 = QtWidgets.QPushButton(self.centralwidget, clicked= lambda: self.sendmail())
        self.btn4.setGeometry(QtCore.QRect(310, 160, 75, 31))
        self.btn4.setStyleSheet("background-color: rgb(51, 51, 51);\n"
        "color: rgb(250, 250, 250);")
        self.btn4.setObjectName("btn4")
        self.lbl2 = QtWidgets.QLineEdit(self.centralwidget)
        self.lbl2.setGeometry(QtCore.QRect(50, 120, 431, 31))
        self.lbl2.setStyleSheet("background-color: rgb(51, 51, 51);\n"
        "color: rgb(250, 250, 250);\n"
        "selection-color: rgb(0, 229, 255);")
        self.lbl2.setObjectName("lbl2")
        self.btn3 = QtWidgets.QPushButton(self.centralwidget, clicked= lambda: self.attachfile())
        self.btn3.setGeometry(QtCore.QRect(480, 120, 75, 31))
        self.btn3.setStyleSheet("background-color: rgb(51, 51, 51);\n"
        "color: rgb(250, 250, 250);")
        self.btn3.setObjectName("btn3")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 602, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
        self.adres()

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "CSVtoSQLITE"))
        self.btn1.setText(_translate("MainWindow", "Import .csv"))
        self.lbl1.setPlaceholderText(_translate("MainWindow", "  Select .csv"))
        self.lbl3.setPlaceholderText(_translate("MainWindow", "  e-mail"))
        self.btn2.setText(_translate("MainWindow", "Convert"))
        self.btn5.setText(_translate("MainWindow", "Change e-mail"))
        self.btn4.setText(_translate("MainWindow", "Send file"))
        self.lbl2.setPlaceholderText(_translate("MainWindow", "  Attach .db file to send by e-mail"))
        self.btn3.setText(_translate("MainWindow", "Attach .db"))

    def browsefiles(self):
        fpath = QtWidgets.QFileDialog.getOpenFileName(None, 'Open file', filter='*.csv')
        try:
                self.lbl1.setText(fpath[0])
        except: 
                self.lbl1.setText("  Must be .CSV file!")

    def convert(self):
        file = self.lbl1.text()
        name = (os.path.splitext(file)[0])
        
        try:
            csv_file = pd.read_csv(file)
            csv_file.columns = csv_file.columns.str.strip()
            
            saveopt, _ = QtWidgets.QFileDialog.getSaveFileName(None, caption='Save file', filter = '.db', directory= name + '.db')
        
            try:

                with sqlite3.connect(saveopt) as con:
                    dropTable1 = "DROP TABLE IF EXISTS Table1"
                    con.execute(dropTable1)
                    csv_file.to_sql("Table1", con)
                    self.lbl1.setText("  Converted and saved successfully!")
            except:
                self.lbl1.setText("  Save directory must be selected!")
        except:
            self.lbl1.setText("  Please import correct .CSV file!")


    def attachfile(self):
        apath = QtWidgets.QFileDialog.getOpenFileName(None, 'Open file', filter='*.db')
        try:
                self.lbl2.setText(apath[0])
        except: 
                self.lbl2.setText("  Must be .db file!")


    def mail(self):
        with open("mail.txt", "w") as file:
            file.write(self.lbl3.text())


    def adres(self): 
        with open("mail.txt", "r") as file:
            adres = file.read()

        self.lbl3.insert('  ' + adres)


    def sendmail(self):
        try:
            afile = self.lbl2.text()
            aname = os.path.basename(afile)
            mailaddr = self.lbl3.text()

            msg = MIMEMultipart()
            msg['Subject'] = "Your .db file"
            msg['Body'] = "Your converted file is in the attachements"
            msg['From'] = "exampel@gmail.com"
            msg['To'] = mailaddr

            part = MIMEBase('application', "octet-stream")
            part.set_payload(open(afile, "rb").read())
            encoders.encode_base64(part)

            part.add_header('Content-Disposition', 'attachment', filename=aname)

            msg.attach(part)

            if self.lbl3.text() == "":
                self.lbl3.setPlaceholderText("  Set e-mail address to send file")

            with smtplib.SMTP('smtp.gmail.com',587) as server:
                server.ehlo()
                server.starttls()
                server.ehlo()
                server.login('exampel@gmail.com','password')
                server.send_message(msg)
            
            self.lbl2.setPlaceholderText("  e-mail sent successfully!")


        except:
            self.lbl2.setPlaceholderText("  Choose your attachement")



if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
