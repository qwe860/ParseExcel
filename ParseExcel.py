import configparser
import sys
import os
from PyQt5.QtWidgets import QApplication, QWidget, QComboBox, QPushButton, QFileDialog, QVBoxLayout, QMessageBox
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import QMovie
import pandas as pd
import numpy as np
from datetime import datetime
from configparser import ConfigParser
from ParseExcel_gui import Ui_MainWindow
import ctypes

from Google import Create_Service
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


class MainWindow(QtWidgets.QMainWindow, QtWidgets.QWidget, Ui_MainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.setupUi(self)

        self.setWindowTitle("Kae's Excel Parser??!!")
        self.setWindowIcon(QtGui.QIcon('icon2.png'))    
        self.browseButton.clicked.connect(lambda index = 1: self.getFileName())
        self.readButton.setDisabled(True)
        self.emailButton.setDisabled(True)
        self.readButton.clicked.connect(lambda index = 1: self.readExcel(self.filePathLineEdit.text()))
        self.addButton.clicked.connect(lambda index = 1: self.add_email(self.emailLineEdit, self.emailListWidget)) #button pressed
        self.removeButton_2.clicked.connect(lambda index = 0: self.remove_email(self.emailListWidget))
        self.saveButton.clicked.connect(lambda index = 0: self.save_config())
        self.emailButton.clicked.connect(lambda index = 0: self.send_email())

        self.obtain_config()

        #self.filePathLineEdit.setText('C:/Users/kyeong1/Desktop/ParseExcel/Fab3busbar- Thermal Scan data-Rev1.xlsx')

        self.movie = QMovie('Loading.gif')
        self.loadingLabel.setMovie(self.movie)
        self.movie.start()
        self.loadingLabel.setHidden(True)
        self.loadingLabel_2.setHidden(True)

        self.statusLabel.setHidden(True)

    def add_email(self, textField, emailListWidget):
        textboxValue = textField.text().replace(" ","")

        if len(textboxValue) == 0:
            QMessageBox.question(self, 'Information', "Please do not leave blank!", QMessageBox.Ok)
        else: 
            matching_items = emailListWidget.findItems(textboxValue, Qt.MatchContains)
            if matching_items:
                QMessageBox.question(self, 'Information', "There is a duplicate entry! Please get rid of it and try again!", QMessageBox.Ok)
            else:
                emailListWidget.addItem(textboxValue)
                

    def remove_email(self, emailList):
        listItems = emailList.selectedItems()
        if not listItems: 
            return        
        for item in listItems:
            emailList.takeItem(emailList.row(item))

    def save_config(self):
        if os.path.exists("config.ini"):
                os.remove("config.ini")

        config = ConfigParser()

        config.read('config.ini')
        config.add_section('Email_List')

        for x in range (self.emailListWidget.count()):
                config.set('Email_List', 'email_'+str(x), self.emailListWidget.item(x).text())
                print("poop")

        config.add_section('Temperatures')
        config.set('Temperatures', 'Max_Temp', self.spinBox.text())
        config.set('Temperatures', 'Temp_Diff', self.doubleSpinBox.text())
        with open('config.ini', 'w') as f:
            config.write(f)

        self.emailListWidget.clear()
        QMessageBox.information(self, 'Information', "Configuration Saved!", QMessageBox.Ok)
        self.textBrowser.clear()
        self.obtain_config()

    def obtain_config(self):
        config = ConfigParser()
        config.read('config.ini')

        k = len(config['Email_List'])

        for i in range (k):
            print(config['Email_List']['email_' + str(i)])
            self.emailListWidget.addItem(config['Email_List']['email_' + str(i)])

        self.spinBox.setValue(int(config['Temperatures']['Max_Temp']))
        self.max_temp = int(config['Temperatures']['Max_Temp'])
        self.doubleSpinBox.setValue(float(config['Temperatures']['Temp_Diff']))
        self.temp_diff = float(config['Temperatures']['Temp_Diff'])


    def getFileName(self):
        file_filter = 'Data File (*.xlsx *.csv *.dat);; Excel File (*.xlsx *.xls)'
        response = QFileDialog.getOpenFileName(
            parent=self,
            caption='Select a data file',
            directory=os.getcwd(),
            filter=file_filter,
            initialFilter='Excel File (*.xlsx *.xls)'
        )

        if response[0] != '':
            self.sheetComboBox.clear()
            self.loadingLabel.setHidden(False)
            self.loadingLabel_2.setHidden(False)
            self.statusLabel.setHidden(False)
            
            self.browseButton.setDisabled(True)
            self.sheetComboBox.setDisabled(True)

            self.statusLabel.setText('Reading excel file. Please wait...')
            self.worker  = WorkerThread(response[0])
            self.worker.start()
            self.worker.update_sheets.connect(self.update_combobox)
            self.worker.insert_file_path.connect(self.InsertFilePath)
            self.worker.finished.connect(self.FinishedReadingFile)
            #load.stopAnimation()

        return response[0]

    def FinishedReadingFile(self):
        self.readButton.setDisabled(False)
        print('file reading complete!')
        self.loadingLabel.setHidden(True)
        self.loadingLabel_2.setHidden(True)

        self.browseButton.setDisabled(False)
        self.sheetComboBox.setDisabled(False)

        self.statusLabel.setHidden(True)

    def InsertFilePath(self, file_path):
        self.filePathLineEdit.setText(file_path)

        
    def update_combobox(self, sheet_names):
        self.sheetComboBox.addItems(sheet_names)

    def UpdateTextBrowser(self, content):
        self.textBrowser.append(content)

    def readExcel(self, file_path):

        sheet_to_read = self.sheetComboBox.currentText()
        print(file_path)
        self.browseButton.setDisabled(True)
        self.sheetComboBox.setDisabled(True)

        self.loadingLabel.setHidden(False)
        self.loadingLabel_2.setHidden(False)

        self.worker2 = Worker2(file_path, sheet_to_read, self.max_temp, self.temp_diff)
        self.worker2.start()
        self.worker2.update_textBrowser.connect(self.UpdateTextBrowser)
        self.worker2.keyword_exists.connect(self.keywordNotExist)
        self.worker2.finished.connect(self.excel_finished)


    def excel_finished(self):
        self.loadingLabel.setHidden(True)
        self.loadingLabel_2.setHidden(True)
        self.browseButton.setDisabled(False)
        self.sheetComboBox.setDisabled(False)
        self.emailButton.setDisabled(False)


        if self.didSheetLoadCorrectly:
            QMessageBox.information(self, 'Information', "Excel spreadsheet read successfully!", QMessageBox.Ok)

    def keywordNotExist(self, does_it_exist):
        if not does_it_exist:
            self.didSheetLoadCorrectly = False

            QMessageBox.critical(self, 'Error parsing sheet', 
            'The selected sheet is missing the "JOINTS" keyword. Please select another sheet or make sure "JOINTS" is in column A!',
            QMessageBox.Ok)

        else:
            self.didSheetLoadCorrectly = True

    def extract_time(Date_and_time, type):
        if type == 'hour':
            result = datetime.strptime(Date_and_time, '%d-%B-%Y %H:%M:%S').strftime('%H')
        if type == 'min':
            result = datetime.strptime(Date_and_time, '%d-%B-%Y %H:%M:%S').strftime('%M')
        if type == 'sec':
            result = datetime.strptime(Date_and_time, '%d-%B-%Y %H:%M:%S').strftime('%S')
        return result

    def convertDateTimeFormat(self, Date_and_time, type):
            if type == 'date':
                result = datetime.strptime(Date_and_time, '%Y-%m-%d %H:%M:%S').strftime('%d %b %Y')
            return result

    def send_email(self):

        content = self.textBrowser.toPlainText()
        email_addresses  = []
        for x in range(self.emailListWidget.count()):
            email_addresses.append(self.emailListWidget.item(x).text()) 
        
        #print(content)
        #print(email_addresses)
        
        CLIENT_SECRET_FILE = 'client_secret.json'
        API_NAME = 'gmail'
        API_VERSION = 'v1'
        SCOPES = ['https://mail.google.com/']

        service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)

        emailMsg = content
        mimeMessage = MIMEMultipart()
        mimeMessage['to'] = ';'.join(email_addresses)
        mimeMessage['subject'] = 'Busbar Temperature Report'
        mimeMessage.attach(MIMEText(emailMsg, 'plain'))
        raw_string = base64.urlsafe_b64encode(mimeMessage.as_bytes()).decode()

        try:
            message = service.users().messages().send(userId='me', body={'raw': raw_string}).execute()
            QMessageBox.information(self, 'Success!', 
            'Email sent successfully!',
            QMessageBox.Ok)
        except Exception as error:
            print("an error occured while trying to send mail! Attempt number: %d" % x)
            #subprocess.call([r'C:\Users\kyeong1\Desktop\NotificationApp\FlushDNS.bat'])
            QMessageBox.critical(self, 'Error', 
            'An unknown error occured. Please rectify it. I dunno what is going on. I just know that IT''S AN ERROR!: ' + error,
            QMessageBox.Ok)
            #time.sleep(5)


class WorkerThread(QThread):
    update_sheets = pyqtSignal(list)
    insert_file_path = pyqtSignal(str)

    def __init__(self, fileName):
        super(WorkerThread, self).__init__()
        self.file_name = fileName

    def run(self):

        xls = pd.ExcelFile(self.file_name)
        sheet_names = xls.sheet_names
        print(sheet_names)
        self.update_sheets.emit(sheet_names)
        self.insert_file_path.emit(self.file_name)

class Worker2(QThread):

    update_textBrowser = pyqtSignal(str)
    keyword_exists = pyqtSignal(bool)

    def __init__(self, filePath, sheetName, threshold, temp_diff):
        super(Worker2, self).__init__()
        self.file_path = filePath
        self.sheet_to_read = sheetName
        self.threshold = threshold
        self.temp_diff = temp_diff

    def run(self):
        try:
            df = pd.read_excel(self.file_path, sheet_name=self.sheet_to_read)
            mask = np.column_stack([df[col].str.contains(r"JOINTS", na=False) for col in df])
            header_row = df.loc[mask.any(axis=1)].index[0]
            self.keyword_exists.emit(True)

            print('joints exists ')
            df.columns = df.iloc[header_row]
            df = df.iloc[header_row+1:,:]

            starting_index = df.columns.get_loc('Ambient Temp Deg C')

            joint_list = []
            location_list = []
            temperature_list = []
            date_list = []

            total_rows_to_index = 2 * len(df.index)
            total_col_to_index = 2 *  (len(df.columns) - starting_index)
            total_length = total_rows_to_index + total_col_to_index
            #check if absolute temperature is more than 40 degrees
            for column in df:

                current_index = df.columns.get_loc(column)
                if current_index > starting_index:

                    #print(column)                       #print column name
                    #print(df.columns.get_loc(column))   #get index of column name

                    for i in range(len(df.index)): #i is the row
                        val = df[column].values[i]
                        #print (val)


                        try: 
                            difference = abs(val - df[column].values[current_index+1])
                            #print('temp difference is ', difference)
                            int(val)            #df[row, column]
                            if val > self.threshold:
                                date_list.append(column)
                                location = df.iloc[i , 1]
                                location_list.append(location)
                                joint = df.iloc[i , 0]
                                joint_list.append(joint)
                                print('DANGER: temperature reading of  ' ,val ,' is more than 40 degrees at joint ', joint ,' at location: ', location)
                                temperature_list.append(val)
                        except:
                            print('noep')
            
            col_length_to_index = len(df.columns) - starting_index
            difference_threshold = self.temp_diff

            diff_temp_list = []
            diff_date1_list = []
            diff_date2_list = []
            diff_joint_list = []
            diff_location_list = []

            for row in range(len(df.index)):
                for column in df:
                    currentIndex = df.columns.get_loc(column)
                    if currentIndex > starting_index:
                        #do the calculation here

                        if currentIndex < col_length_to_index:
                            value = df[column].values[row]
                            next_value =  df.iloc[row, currentIndex+1]

                            try:
                                difference = abs(float(next_value) - float(value))
                                if difference > difference_threshold:
                                    diff_temp_list.append("{:.1f}".format(difference))

                                   # format_float = "{:.2f}".format(float)


                                    diff_date1_list.append(column)
                                    diff_date2_list.append(df.columns[currentIndex+1])

                                    diff_joint_list.append(df.iloc[row, 0])
                                    diff_location_list.append(df.iloc[row,1])
                            except:
                                print('this entry is not a number')

                            
            print('------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------')
            print('Summary of temperature warnings for sheet ', self.sheet_to_read, ':\n')

            string0 = 'Reported generated on: ' + str(datetime.now().strftime('%d %b %Y %H:%M:%S'))  
            self.update_textBrowser.emit(string0)
            string1 = 'Summary of temperature warnings for sheet '+ self.sheet_to_read+ ':'
            self.update_textBrowser.emit(string1)
            self.update_textBrowser.emit('---------------------------------------------------------------------')

            for index, item in enumerate(joint_list):
                print('WARNING: temperature reading of  ' ,temperature_list[index] ,' taken on ', self.convertDateTimeFormat(str(date_list[index]), 'date') ,'is more than '+ str(self.threshold) +' degrees at joint ', item ,' at location: ', location_list[index])
                string2 = 'WARNING: temperature reading of  ' + str(temperature_list[index]) +' deg C taken on '+ self.convertDateTimeFormat(str(date_list[index]), 'date') +' is more than ' + str(self.threshold) +' deg at joint '+ str(item) +', at location: '+ str(location_list[index])
                self.update_textBrowser.emit(string2)
        

            print('------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------')
            print('Summary of temperature difference for sheet ', self.sheet_to_read, ':\n')

            string3 = '\nSummary of temperature difference for sheet ' + self.sheet_to_read + ':'
            self.update_textBrowser.emit(string3)
            self.update_textBrowser.emit('---------------------------------------------------------------------')

            for index, item in enumerate(diff_temp_list):
                print('DANGER: Temperature difference at joint ', diff_joint_list[index], 'at location ', diff_location_list[index], 'is ', item, 
                '. For readings taken between ', self.convertDateTimeFormat(str(diff_date1_list[index]), 'date') ,'and ',self.convertDateTimeFormat(str(diff_date2_list[index]), 'date'))

                string4 = 'DANGER: Temperature difference at joint ' + str(diff_joint_list[index])+ ' at location '+ str(diff_location_list[index])+ ' is '+ str(item) +' deg C. For readings taken between '+ self.convertDateTimeFormat(str(diff_date1_list[index]), 'date') +' & '+self.convertDateTimeFormat(str(diff_date2_list[index]), 'date')
                self.update_textBrowser.emit(string4)

        except Exception as e:
            self.keyword_exists.emit(False)
            print('joints dont exist')
            print(e)

    def convertDateTimeFormat(self, Date_and_time, type):
            if type == 'date':
                result = datetime.strptime(Date_and_time, '%Y-%m-%d %H:%M:%S').strftime('%d %b %Y')

            elif type == 'time':
                result = datetime.strptime(Date_and_time, '%d-%B-%Y %H:%M:%S').strftime('%H:%M:%S')
            return result


if __name__ == "__main__":
    import sys
    myappid = 'mycompany.myproduct.subproduct.version' # arbitrary string
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    app = QtWidgets.QApplication(sys.argv)
    myWidget = MainWindow()
    myWidget.show()
    sys.exit(app.exec_())
