import json
import shutil
import tempfile
from random import randint
import math
import requests
from PyQt6.QtWidgets import *
from PyQt6.QtCore import QThread,pyqtSignal
from PyQt6 import uic
import os
from time import sleep
import undetected_chromedriver.v2 as uc
from pandas import read_excel
from selenium import webdriver
_count = 0
_countSucess = 0
_countExcel = 0

class UI(QMainWindow):
    def __init__(self):
        super(UI, self).__init__()
        uic.loadUi("Data/FormGG.ui",self)
        self.pushButton = self.findChild(QPushButton,'pushButton')
        self.pushButton.setEnabled(False)
        self.pushButton.clicked.connect(self.start)
        self.pushButton_2 = self.findChild(QPushButton,'pushButton_2')
        self.pushButton_2.clicked.connect(self.diaLogExcelFile)
        self.pushButton_3 = self.findChild(QPushButton,'pushButton_3')
        self.pushButton_3.clicked.connect(self.stopThread)
        self.lineEdit = self.findChild(QLineEdit,'lineEdit')
        self.label_2 = self.findChild(QLabel,'label_2')
        self.spinBox = self.findChild(QSpinBox,'spinBox')
        self.spinBox_2 = self.findChild(QSpinBox,'spinBox_2')
        self.handleUrl(1)

        self.checkBox = self.findChild(QCheckBox,'checkBox')
        self.label_3 = self.findChild(QLabel,'label_3')
        self.threadHandel = {}
        self.show()

    def start(self):
        self.handleUrl(0)
        for i in range(0,int(self.spinBox.value())):
            self.threadHandel[i] = HandelThread(i)
            self.threadHandel[i].excelFiles = self.fileName
            self.threadHandel[i].labelSucess.connect(self.labelSucess)
            self.threadHandel[i].labelStatus.connect(self.labelStatus)
            self.threadHandel[i].checkBox = self.checkBox
            self.threadHandel[i].lineEdit = self.lineEdit
            self.threadHandel[i].spinBox_2 = self.spinBox_2
            self.threadHandel[i].start()
    def diaLogExcelFile(self):
        try:
            self.fileName, _ = QFileDialog.getOpenFileName(self, "QFileDialog.getOpenFileName()")
            self.pushButton.setEnabled(True)
        except:
            pass
    def labelSucess(self,index):
        self.label_2.setText(f'Thành Công: {index}')
    def labelStatus(self,text):
        self.label_3.setText(f'Trạng Thái: {text}')

    def handleUrl(self,index):
        # Index == 0 to save url
        if index == 0:
            with open('DATA/Reload/url.txt','w') as wURL:wURL.write(self.lineEdit.text())
        # Else show url to line Edit
        else:
            with open('DATA/Reload/url.txt', 'r') as rURL:
                rURL = rURL.read()
                self.lineEdit.setText(str(rURL))
    def stopThread(self):
        for i in range(0, int(self.spinBox.value())):
            self.threadHandel[i].terminate()
            self.threadHandel[i].closeBrowser()
class HandelThread(QThread):
    labelSucess = pyqtSignal(str)
    labelStatus = pyqtSignal(str)
    def __init__(self,index = 0):
        super(HandelThread,self).__init__()
        self.index = index
    def run(self):
        global _countExcel
        while True:
            try:
                check = self.handel()
                if check == 1:break
            except Exception as e:
                self.labelStatus.emit('Lỗi,Vui lòng liên hệ dev để kiểm tra lỗi !')
    def setBrowser(self):
        self.temp = os.path.normpath(tempfile.mkdtemp())
        opts = uc.ChromeOptions()
        # If spinBox use proxy checked,handel get proxy to api and use it !
        if self.checkBox.isChecked():
            proxy = self.getProxy()
            opts.add_argument('--proxy-server=%s' % proxy)
        opts.add_argument(f"--window-position={self.index * 200},0")
        opts.add_argument("--window-size=560,880")
        self.labelStatus.emit('Khởi tạo trình duyệt !')
        args = ["hide_console", ]
        opts.add_argument('--deny-permission-prompts')
        opts.add_argument(f'--user-data-dir={self.temp}')
        opts.add_argument("--disable-popup-blocking")
        opts.add_argument("--disable-gpu")
        self.browser = uc.Chrome(executable_path=os.getcwd()+'/chromedriver',options=opts,service_args=args)
        with self.browser:self.browser.get(self.lineEdit.text())
    def handel(self):
        # Create Browser
        self.setBrowser()
        check = self.waitBrowser(self.browser,'input[name="name"]')
        global _countExcel
        if check==0:
            dataExcel = read_excel(self.excelFiles)
            loop = dataExcel.shape[0]
            if _countExcel >= loop:return 1
            excel = dataExcel.iloc[_countExcel]
            _countExcel += 1
            try:
                firstname = excel.iloc[1]
                numberphone = excel.iloc[3]
                email = excel.iloc[2]
            except Exception as e:
                firstname = excel[0]
                numberphone = excel[2]
                email = excel[0]
            # send fistname
            self.browser.find_element_by_css_selector('input[name="name"]').send_keys(str(firstname))
            sleep(1)
            # send email
            self.browser.find_element_by_css_selector('input[name="email"]').send_keys(str(email if math.isnan(email) != True else ''))
            sleep(1)
            # Send phone
            numberphone = str(numberphone)
            self.browser.find_element_by_css_selector('input[name="phone"]').send_keys('0' + numberphone if len(numberphone) < 10 else numberphone)
            sleep(1)
            # Button Click
            self.browser.find_element_by_xpath('//*[@id="BUTTON473"]').click()
            # Save last email
            try:
                with open ('email.txt','w') as wEmail:
                    wEmail.write(email if math.isnan(email) != True else firstname)
            except:pass
            # Sleep time delay
            sleep(randint(5,8))
            # Check Sucessfull Or Failed
            check = self.waitBrowser(self.browser,'div[class="ladipage-message-text"]')
            if check == 0:
                statusText = self.browser.find_element_by_css_selector('div[class="ladipage-message-text"]')
                if 'Cảm ơn' in statusText.text:
                    global _countSucess
                    _countSucess += 1
                    self.labelSucess.emit(f'{_countSucess}')
        self.closeBrowser()
        sleep(int(self.spinBox_2.value()))

    def getProxy(self):
        with open('proxy.txt','r') as rProxy:rProxy = rProxy.readlines()
        while True:
            proxy = requests.get(
                f'http://proxy.tinsoftsv.com/api/changeProxy.php?key=[{rProxy[self.index]}]&location=[0]')
            proxy_data = json.loads(proxy.text)
            try:
                proxy = proxy_data['proxy']
                self.labelStatus.emit(f'Lấy Thành Công Proxy : {proxy}')
                return proxy
            except:
                timedelay = proxy_data['next_change']
                for i in range(int(timedelay), 0, -1):
                    self.labelStatus.emit(f'Hệ Thống Sẽ Khởi Động Sau {i}s')
                    sleep(1)
    def closeBrowser(self):
        if self.browser:
            self.browser.close()
            self.browser.quit()
            try:
                shutil.rmtree(r'{}'.format(self.temp))
            except:
                pass

    def waitBrowser(self,browser, options=None, options1=None, options2=None, options3=None, options4=None,
                    options5=None,
                    time_out=15, index=0, index1=0, index2=0, index3=0, index4=0, index5=0):
        check = None
        count = 0
        indexcheck = None
        while not check and count < time_out:
            try:
                check = browser.find_elements_by_css_selector(options)[index]
                indexcheck = 0
            except:
                pass
            try:
                check = browser.find_elements_by_css_selector(options1)[index1]
                indexcheck = 1
            except:
                pass
            try:
                check = browser.find_elements_by_css_selector(options2)[index2]
                indexcheck = 2
            except:
                pass
            try:
                check = browser.find_elements_by_css_selector(options3)[index3]
                indexcheck = 3
            except:
                pass
            try:
                check = browser.find_elements_by_css_selector(options4)[index4]
                indexcheck = 4
            except:
                pass
            try:
                check = browser.find_elements_by_css_selector(options5)[index5]
                indexcheck = 5
            except:
                pass
            sleep(1)
            count += 1

        return indexcheck
if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    UIwindow = UI()
    app.exec()


