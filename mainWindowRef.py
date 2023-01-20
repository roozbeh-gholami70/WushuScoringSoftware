from PyQt5 import QtWidgets, uic
from PyQt5.QtWidgets import *
from PyQt5 import QtCore, QtGui , QtSerialPort ,QtMultimedia
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtSerialPort import QSerialPortInfo
from PyQt5.QtMultimedia import QMediaPlayer, QMediaContent
from PyQt5.QtMultimedia import QSound
from playsound import playsound
from PyQt5.QtCore import Qt
from PyQt5.uic import loadUiType

import resources_rc


import string
from score import  Main

import openpyxl 
from openpyxl.styles import PatternFill
import pandas as pd

import sys
import os
import copy


def loadUiClass(path):
    stream = QFile(path)
    stream.open(QFile.ReadOnly)
    try:
        return loadUiType(stream)[0]
    finally:
        stream.close()

# pyrcc5 resources.qrc -o resources_rc.py
# qt keys https://web.mit.edu/~firebird/arch/sun4x_59/doc/html/qt.html

t = 0

redLcdList = ["lcdNumber", "lcdNumber_3", "lcdNumber_7","lcdNumber_9", "lcdNumber_13"]
blueLcdList = ["lcdNumber_2", "lcdNumber_4", "lcdNumber_8", "lcdNumber_10", "lcdNumber_14"]
colorDict = {"red":0, "blue":1, "white":2}

tableFirstColList = ["داور1","داور2","داور3","داور4","داور5","اخطار","ناکدان", "ناک اوت","تذکر"]
tableHeaderList = ["نام","راند","داور1","داور2","داور3","داور4","داور5","اخطار","ناکدان", "ناک اوت","تذکر"]
alphabetList = list(string.ascii_uppercase)


class MainWindow (QMainWindow, loadUiClass(':/ui_files/MainWindowReferee.ui')):
    def __init__(self):
        super( MainWindow, self).__init__()
        # uic.loadUi(Ui_Class, self)
        
        self.setupUi(self)
        self.show()
        porttnavn = pyqtSignal(str)
        self._getserial_ports()

        self.pushButton_5.clicked.connect(lambda:self.up1r())
        self.pushButton_9.clicked.connect(lambda:self.up2r())
        self.pushButton_22.clicked.connect(lambda:self.up3r())
        self.pushButton_28.clicked.connect(lambda:self.up4r())
        self.pushButton_36.clicked.connect(lambda:self.up5r())

        self.pushButton_7.clicked.connect(lambda:self.up1b())
        self.pushButton_14.clicked.connect(lambda:self.up2b())
        self.pushButton_23.clicked.connect(lambda:self.up3b())
        self.pushButton_26.clicked.connect(lambda:self.up4b())
        self.pushButton_34.clicked.connect(lambda:self.up5b())

        ##########################################################
        self.pushButton_6.clicked.connect(lambda:self.dn1r())
        self.pushButton_12.clicked.connect(lambda:self.dn2r())
        self.pushButton_24.clicked.connect(lambda:self.dn3r())
        self.pushButton_27.clicked.connect(lambda:self.dn4r())
        self.pushButton_35.clicked.connect(lambda:self.dn5r())

        self.pushButton_8.clicked.connect(lambda:self.dn1b())
        self.pushButton_13.clicked.connect(lambda:self.dn2b())
        self.pushButton_15.clicked.connect(lambda:self.dn3b())
        self.pushButton_25.clicked.connect(lambda:self.dn4b())
        self.pushButton_33.clicked.connect(lambda:self.dn5b())
        
        
        self.pushButton_3.setEnabled(False)
        self.endBtn.setEnabled(False)
        self.sendBtn.setEnabled(False)
        self.connectBtn.setEnabled(False)
        self.pushButton_2.clicked.connect(lambda:self.start1())
        self.pushButton_3.clicked.connect(lambda:self.puse())
        self.endBtn.clicked.connect(lambda:self.resetTime())
        self.endBtn.clicked.connect(lambda:self.receive())
        self.connectBtn.clicked.connect(lambda:self.on_toggled())
        # self.connectBtn.clicked.connect(lambda:self.on_toggled())
        self.sendBtn.clicked.connect(lambda:self.sendToArdo())
        # self.endBtn.clicked.connect(lambda:self.choose_music())

        self.uploadBtn.clicked.connect(lambda:self.uploadData())

        self.spinBox.valueChanged.connect(lambda:self.setTimer())
        self.spinBox_2.valueChanged.connect(lambda:self.setTimer())

        self.spinBox_5.valueChanged.connect(lambda:self.setRestTime())
        self.spinBox_6.valueChanged.connect(lambda:self.setRestTime())
        self.weightTxt.textChanged.connect(lambda:self.setWeight())
        
        self.gameNumCbx.activated.connect(lambda:self.setGameNum())
        self.portsCbx.activated.connect(lambda:self.setPortName())

        self.scoreBoardBtn.clicked.connect(lambda:self.openSecondWindow())

        self.playersNameBtn.clicked.connect(lambda:self.takePlayerNames())
        self.settingBtn.clicked.connect(lambda:self.getMatchSettings())
        self.saveCsvBtn.clicked.connect(lambda:self.saveToCsv())

        self.round1Rbtn.clicked.connect(lambda:self.radioBtnState(self.round1Rbtn))
        self.round2Rbtn.clicked.connect(lambda:self.radioBtnState(self.round2Rbtn))
        self.round3Rbtn.clicked.connect(lambda:self.radioBtnState(self.round3Rbtn))

        self.loadCsvBtn.clicked.connect(lambda: self.loadCsvFile())

        self.gameNumCbx.setStyleSheet("QComboBox QAbstractItemView { \
                                        selection-color: red; \
                                        selection-background-color: rgb(244, 244, 244);}")
        self.setSpinBoxReadOnly()
        
        self.round1Rbtn.setChecked(True)
        self.second = 0
        self.count = 0
        self.roundNum = 1
        self.numOfMatches = 0
        self.start = False
        self.mainTime = False
        self.restTime = False
        self.lockRef = True
        self.redPoints = 0
        self.bluePoints = 0
        self.redWinRand = [0, 0 , 0]
        self.blueWinRand = [0, 0 , 0]
        self.gameNum = ""
        self.redPlayer = ""
        self.bluePlayer = ""
        self.playerWeight = ""
        self.portName = "__"
        self.serial = QtSerialPort.QSerialPort(self.portName,
                                                baudRate=QtSerialPort.QSerialPort.Baud115200, 
                                                readyRead=self.receive)
        self.refColor = {"1":[0,0,0,0,0], "2":[0,0,0,0,0], "3":[0,0,0,0,0]}

        self.redPlayerRefsPoint = {"1":[0,0,0,0,0], "2":[0,0,0,0,0], "3":[0,0,0,0,0]}
        self.bluePlayerRefsPoint = {"1":[0,0,0,0,0], "2":[0,0,0,0,0], "3":[0,0,0,0,0]}

        self.redPlayerTablePoint = {"1":[0,0,0,0], "2":[0,0,0,0], "3":[0,0,0,0]}
        self.bluePlayerTablePoint = {"1":[0,0,0,0], "2":[0,0,0,0], "3":[0,0,0,0]}
        # my part start

        timer = QTimer(self)
        # adding action to timer
        timer.timeout.connect(self.showTime)

        # update the timer every tenth second
        timer.start(100)

        # my part end
    
        self.minet = 0
        self.t =0
        
        
    def setSpinBoxReadOnly(self):
        for roundNum in range(1,4):
            for idx in range(4):
                myObj = eval("self.red"+str(roundNum)+"_"+str(idx+1) + "spinBox")
                myObj.lineEdit().setReadOnly(True)

                myObj = eval("self.blue"+str(roundNum)+"_"+str(idx+1) + "spinBox")
                myObj.lineEdit().setReadOnly(True)
                
        
    def keyPressEvent(self, e):

        modifiers = QtWidgets.QApplication.keyboardModifiers()
      
        if e.modifiers() & Qt.ControlModifier:
            if ("Main" not in dir(self)):
                QMessageBox.about(self, "اخطار", "لطفا اسکوربورد را باز کنید.")
            else:
                if e.key() == Qt.Key_7:
                    myObj = eval("self.red"+str(self.roundNum)+"_1spinBox")
                    if (int(myObj.value())>0):
                        tmp = int(myObj.value()) - 1
                        myObj.setValue(tmp)         

                if e.key() == Qt.Key_4:
                    myObj = eval("self.red"+str(self.roundNum)+"_2spinBox")
                    if (int(myObj.text())>0):
                        tmp = int(myObj.value()) - 1
                        myObj.setValue(tmp)    
                
                if e.key() == Qt.Key_1:
                    myObj = eval("self.red"+str(self.roundNum)+"_3spinBox")
                    if (int(myObj.text())>0):
                        tmp = int(myObj.value()) - 1
                        myObj.setValue(tmp)    
                if e.key() == Qt.Key_0:
                    myObj = eval("self.red"+str(self.roundNum)+"_4spinBox")
                    if (int(myObj.text())>0):
                        tmp = int(myObj.value()) - 1
                        myObj.setValue(tmp)                           

                if e.key() == Qt.Key_9:
                    myObj = eval("self.blue"+str(self.roundNum)+"_1spinBox")
                    if (int(myObj.text())>0):
                        tmp = int(myObj.value()) - 1
                        myObj.setValue(tmp)    

                if e.key() == Qt.Key_6:
                    myObj = eval("self.blue"+str(self.roundNum)+"_2spinBox")
                    if (int(myObj.text())>0):
                        tmp = int(myObj.value()) - 1
                        myObj.setValue(tmp)
                
                if e.key() == Qt.Key_3:
                    myObj = eval("self.blue"+str(self.roundNum)+"_3spinBox")
                    if (int(myObj.text())>0):
                        tmp = int(myObj.value()) - 1
                        myObj.setValue(tmp) 
                if e.key() == Qt.Key_Period:
                    myObj = eval("self.blue"+str(self.roundNum)+"_4spinBox")
                    if (int(myObj.text())>0):
                        tmp = int(myObj.value()) - 1
                        myObj.setValue(tmp) 
        else:
            if ("Main" not in dir(self)):
                QMessageBox.about(self, "اخطار", "لطفا اسکوربورد را باز کنید.")
            else:
                if e.key() == Qt.Key_F1:
                    if (self.pushButton_2.isEnabled()):
                        self.start1()
                    elif ((self.pushButton_3.isEnabled()) & (self.count != 0)):
                        self.puse()

                if e.key() == Qt.Key_7:
                    myObj = eval("self.red"+str(self.roundNum)+"_1spinBox")
                    tmp = int(myObj.value()) + 1
                    myObj.setValue(tmp)        

                if e.key() == Qt.Key_4:
                    myObj = eval("self.red"+str(self.roundNum)+"_2spinBox")
                    tmp = int(myObj.value()) + 1
                    myObj.setValue(tmp) 
                
                if e.key() == Qt.Key_1:
                    myObj = eval("self.red"+str(self.roundNum)+"_3spinBox")
                    tmp = int(myObj.value()) + 1
                    myObj.setValue(tmp) 
                 
                if e.key() == Qt.Key_0:
                    myObj = eval("self.red"+str(self.roundNum)+"_4spinBox")
                    tmp = int(myObj.value()) + 1
                    myObj.setValue(tmp) 

                if e.key() == Qt.Key_9:
                    myObj = eval("self.blue"+str(self.roundNum)+"_1spinBox")
                    tmp = int(myObj.value()) + 1
                    myObj.setValue(tmp) 

                if e.key() == Qt.Key_6:
                    myObj = eval("self.blue"+str(self.roundNum)+"_2spinBox")
                    tmp = int(myObj.value()) + 1
                    myObj.setValue(tmp) 
                
                if e.key() == Qt.Key_3:
                    myObj = eval("self.blue"+str(self.roundNum)+"_3spinBox")
                    tmp = int(myObj.value()) + 1
                    myObj.setValue(tmp) 
                
                if e.key() == Qt.Key_Period:
                    myObj = eval("self.blue"+str(self.roundNum)+"_4spinBox")
                    tmp = int(myObj.value()) + 1
                    myObj.setValue(tmp)  

        
    def takePlayerNames(self):
        self.redPlayerData = {}
        self.bluePlayerData = {}

        if ("Main" not in dir(self)):
            QMessageBox.about(self, "اخطار", "لطفا اسکوربورد را باز کنید.")
        elif (self.numOfMatches == 0):
            QMessageBox.about(self, "اخطار", "لطفا تعداد مسابقات را مشخص کنید.")
        else:
            redPlayerName, done1 = QtWidgets.QInputDialog.getText(
                self, 'مشخصات مبارز قرمز', 'نام مبارز قرمز:')
    
            bluePlayerName, done2 = QtWidgets.QInputDialog.getText(
            self, 'مشخصات مبارز آبی', 'نام مبارز آبی:') 


            if done1 and done2 :
                msgBoxTopic = "مشخصات مبارزین"
                msgBoxQuestion = "مبارز قرمز: " + redPlayerName+ "\n" +"مبارز آبی: " + bluePlayerName + "\n"
                buttonReply = QMessageBox.question(self,msgBoxTopic ,msgBoxQuestion , QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
                if buttonReply == QMessageBox.Yes:
                    self.redPlayer = redPlayerName
                    self.bluePlayer = bluePlayerName
                    self.redPlayerData["name"] = redPlayerName
                    self.bluePlayerData["name"] = bluePlayerName
                    self.redPlayerData["weight"] = str(self.playerWeight)
                    self.bluePlayerData["weight"] = str(self.playerWeight)
                    self.redPlayerData["matchNum"] = str(self.gameNum)
                    self.bluePlayerData["matchNum"] = str(self.gameNum)
                    self.setPlayerName()
                else:
                    self.takePlayerNames()
    
    def getMatchSettings(self):
        if ("Main" not in dir(self)):
            QMessageBox.about(self, "اخطار", "لطفا اسکوربورد را باز کنید.")
        else:
            numOfMatches, done1 = QtWidgets.QInputDialog.getText(
                self, '  تعداد مسابقات', 'تعداد مسابقه: ')
            if done1:
                msgBoxTopic = "تعداد مسابقات"
                msgBoxQuestion = " تعداد مسابقه: "+ numOfMatches + "\n"
                buttonReply = QMessageBox.question(self,msgBoxTopic ,msgBoxQuestion , QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
                if buttonReply == QMessageBox.Yes:
                    self.numOfMatches = int(numOfMatches)
                    self.myDataTable = ["0"]*int(numOfMatches)
                    
                    self.gameNumCbx.clear()
                    items  = list(range(1, int(numOfMatches)+1))
                    self.gameNumCbx.addItems(map(str, items))
                    self.gameNum = self.gameNumCbx.currentText()
                    self.setGameNum()
                else:
                    self.getMatchSettings()


    def setTable(self):
        for i in range(self.gridTable.count()):
            item = self.gridTable.itemAt(i).widget()
            if isinstance(item,QSpinBox):
                item.setValue(0)

    def radioBtnState(self,myBtn):
        self.storeTableData()
        self.calcPoints()
        if myBtn.isChecked() == True:
            self.roundNum = int(str(myBtn.objectName())[5])
            if ("Main" in dir(self)):
                self.Main.roundNumLbl.setText(str(self.roundNum))

            for idx, item in enumerate(redLcdList):
                myObjRed = eval("self."+item)
                myObjBlue = eval("self."+blueLcdList[idx])
                myObjRed.display(self.redPlayerRefsPoint[str(self.roundNum)][idx])
                myObjBlue.display(self.bluePlayerRefsPoint[str(self.roundNum)][idx])

            for idx in range(4):
                myObj = eval("self.red"+str(self.roundNum)+"_"+str(idx+1) + "spinBox")
                myObj.setValue(int(self.redPlayerTablePoint[str(self.roundNum)][idx]))

                myObj = eval("self.blue"+str(self.roundNum)+"_"+str(idx+1) + "spinBox")
                myObj.setValue(int(self.bluePlayerTablePoint[str(self.roundNum)][idx]))               

    def setPlayerName(self):
        if ("Main" not in dir(self)):
            QMessageBox.about(self, "اخطار", "لطفا اسکوربورد را باز کنید.")
        else:
            self.Main.redPlayerNameTxt.setPlainText(self.matchData[str(self.gameNum)]["red"])
            self.Main.bluePlayerNameTxt.setPlainText(self.matchData[str(self.gameNum)]["blue"])
            self.weightTxt.setPlainText(self.matchData[str(self.gameNum)]["weight"])
            # self.Main.redPlayerNameTxt.insertPlainText(self.redPlayer)
            # self.Main.bluePlayerNameTxt.insertPlainText(self.bluePlayer)
    
    def setWeight(self):
        self.playerWeight = self.weightTxt.toPlainText()
        if ("Main" not in dir(self)):
            QMessageBox.about(self, "اخطار", "لطفا اسکوربورد را باز کنید.")
        else:
            self.Main.weightLbl.setText(str(self.playerWeight))
            try:
                self.redPlayerData["weight"] = str(self.playerWeight)
                self.bluePlayerData["weight"] = str(self.playerWeight)
            except:
                pass
    
    def setGameNum(self):
        
        if ("Main" not in dir(self)):
            QMessageBox.about(self, "اخطار", "لطفا اسکوربورد را باز کنید.")
        else:
            try:
                self.redPlayerData["matchNum"] = str(self.gameNum)
                self.bluePlayerData["matchNum"] = str(self.gameNum)
            except:
                pass
            self.gameNum = self.gameNumCbx.currentText()
            self.Main.gameNumTxt.setHtml("<p align='center'>" + str(self.gameNum) + "</p>")
            self.setPlayerName()
    
    def saveToCsv(self):
        if (self.gameNum!=""):
            try:
                excelFile = "match_"+str(self.gameNum)+".xlsx"
                dir_path = os.path.dirname(os.path.realpath(__file__))
                if os.path.exists(dir_path+"\\" + excelFile):
                    self.myExcel = openpyxl.load_workbook("match_"+str(self.gameNum)+".xlsx")
                    self.workSheet =  self.myExcel.active
                else:
                    self.myExcel = openpyxl.Workbook()
                    self.workSheet = self.myExcel.active
                    self.writeExcelFileHeader()
                
                self.writeExcelFile()
                self.myExcel.save(excelFile)


                self.setTable()
                self.redPlayerRefsPoint = {"1":[0,0,0,0,0], "2":[0,0,0,0,0], "3":[0,0,0,0,0]}
                self.bluePlayerRefsPoint = {"1":[0,0,0,0,0], "2":[0,0,0,0,0], "3":[0,0,0,0,0]}

                self.redPlayerTablePoint = {"1":[0,0,0,0], "2":[0,0,0,0], "3":[0,0,0,0]}
                self.bluePlayerTablePoint = {"1":[0,0,0,0], "2":[0,0,0,0], "3":[0,0,0,0]}
                index = self.gameNumCbx.currentIndex()
                self.gameNumCbx.model().item(index).setEnabled(False)
                self.gameNumCbx.setCurrentIndex(index+1)
                self.setGameNum()
                self.redWinRand = [0, 0 , 0]
                self.blueWinRand = [0, 0 , 0]
                print(self.redPlayerData)
                print(self.bluePlayerData)
            except Exception as e:
                if (("redPlayerData" in str(e)) | ("bluePlayerData" in str(e))):
                    QMessageBox.about(self, "خطا", "لطفا مشصخصات مبارزین را وارد کنید.")
        else:
            QMessageBox.about(self, "اخطار", "لطفا تعداد مسابقات را مشخص کنید.")
        
    def writeExcelFileHeader(self):
        red_cell = PatternFill(patternType='solid', fgColor='FB2D01')
        blue_cell = PatternFill(patternType='solid', fgColor='0163FB')
        self.workSheet['B1'] = str(self.redPlayerData["name"])
        self.workSheet['E1'] = str(self.bluePlayerData["name"])
        try:
            for idx,item in enumerate(alphabetList[1:7]):
                self.workSheet[item+"2"] = "راند" + str((idx%3)+1)
                if (idx < 3):
                    self.workSheet[item +'1'].fill = red_cell
                    self.workSheet[item+str(2)].fill = red_cell
                else:
                    self.workSheet[item +'1'].fill = blue_cell
                    self.workSheet[item+str(2)].fill = blue_cell
            for idx,item in enumerate(tableFirstColList):
                self.workSheet["A" + str(idx+3)] = item
            
            self.workSheet.merge_cells('B1:D1')
            self.workSheet.merge_cells('E1:G1')
            
        except Exception as e:
            print(str(e))
            pass
    
    def writeExcelFile(self):
        red_cell = PatternFill(patternType='solid', fgColor='FB2D01')
        blue_cell = PatternFill(patternType='solid', fgColor='0163FB')
        try:
            for randIdx in range(3):
                for idx in range(3, len(tableFirstColList)+3):
                    if (idx < 8):
                        self.workSheet[alphabetList[randIdx+1]+str(idx)] = self.redPlayerRefsPoint["1"][idx-3]
                        self.workSheet[alphabetList[randIdx+4]+str(idx)] = self.bluePlayerRefsPoint["1"][idx-3]
                    else:
                        self.workSheet[alphabetList[randIdx+1]+str(idx)] = self.redPlayerTablePoint["1"][idx-8]
                        self.workSheet[alphabetList[randIdx+4]+str(idx)] = self.bluePlayerTablePoint["1"][idx-8] 
                    self.workSheet[alphabetList[randIdx+1]+str(idx)].fill = red_cell
                    self.workSheet[alphabetList[randIdx+4]+str(idx)].fill = blue_cell
        except Exception as e:
            print(str(e))
            pass

    def update_gui1(self):
        
        self.label_14.setText("{:02d}".format(self.mints))
        self.label_16.setText("{:02d}".format( self.scnd))
        
        self.Main.label_3.setText("{:02d}".format(self.mints))
        self.Main.label_5.setText("{:02d}".format( self.scnd))
    
    def updateTimer(self):
        self.label_2.setText("{:02d}".format(self.mints))
        self.label_8.setText("{:02d}".format( self.scnd))
        
        self.Main.label_3.setText("{:02d}".format( self.mints))
        self.Main.label_5.setText("{:02d}".format(self.scnd))

    def setTimer(self):
        self.start = False
        second = self.spinBox_2.value()
        minute = self.spinBox.value()
        self.label_2.setText("{:02d}".format(minute))
        self.label_8.setText("{:02d}".format( second))
        self.pushButton_3.setEnabled(False)
        self.pushButton_3.setText("وقفه")
        self.pushButton_2.setEnabled(True)
        # changing the value of count
        self.secs = second * 10
        self.count = second * 10 + 9

        
        self.mainTime = True
        self.minutes = minute
        # setting text to the label
    
    def setRestTime(self):
        second = self.spinBox_6.value()
        minute = self.spinBox_5.value()
        if ((second != 0) | (minute != 0)):
            self.restTime = True
            self.label_16.setText("{:02d}".format(second))
            self.label_14.setText("{:02d}".format(minute ))
        else:
            self.restTime = False
    
    def showTime(self):
        self.setLcdColor()
        self.calcPoints()
        # checking if flag is true
        if self.start:
            self.spinBox.setEnabled(False)
            self.spinBox_6.setEnabled(False)
            self.spinBox_2.setEnabled(False)
            self.spinBox_5.setEnabled(False)
            # incrementing the counter
            self.count -= 1
            # timer is completed
            if self.count == 0:
                if (self.minutes > 0):
                    self.minutes -= 1
                    self.count = 599
                else:
                # making flag false
                    self.endtime()
                    self.start = False
                    self.spinBox_2.setValue(0)
                    self.spinBox.setValue(0)
                    # setting text to the label
                    self.mints = 0
                    self.scnd = 0
                    self.pushButton_3.setEnabled(False)
                    self.pushButton_2.setEnabled(True)
                    self.mainTime = False
                    if (self.lockRef):
                        self.lockRef = False
                        QTimer.singleShot(5000,lambda: self.sendToArdo(packet="999999"))
                    if (self.restTime):
                        self.restTimeFun()
                    else:
                        self.setZeroTime()
                        
            
            # self.count -= 1
            

        if self.start:
            # getting text from count
            self.scnd = int(self.count/10)
            self.mints = self.minutes
            if (self.mainTime==True):
                if (self.minutes == 0):
                    if ((self.count < 150) & (self.count > 0)):
                        if (self.count % 10 == 0):
                            self.secondsound()
                self.updateTimer()
            elif (self.resetTime):
                if (self.minutes == 0):
                    if ((self.count < 100) & (self.count > 0)):
                        if (self.count % 10 == 0):
                            self.secondsound()
                self.update_gui1()

    def restTimeFun(self):
        second = self.spinBox_6.value()
        minute = self.spinBox_5.value()
        if ((second != 0) | (minute != 0)):
            self.restTime = True
            self.secs = second * 10
            self.count = (second * 10) + 9
            self.minutes = minute

            self.start = True
            self.pushButton_2.setEnabled(False)
            self.pushButton_3.setEnabled(True)
            self.pushButton_3.setText("وقفه")
            self.spinBox_6.setValue(0)
            self.spinBox_5.setValue(0)
        else:
            self.setZeroTime()
    
    def resetTime(self):
        reply = QMessageBox.question(self, 'پایان راند', 'آیا برای اتمام راند اطمینان دارید؟',
				QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            
            if (self.start):
                self.start = False
                self.count = 0
                
                if (~self.restTime):
                    self.restTimeFun()
                else:
                    self.setZeroTime()
            else:
                self.count = 0
                if (~self.restTime):
                    self.restTimeFun()
                else:
                    self.setZeroTime()

    def setZeroTime(self):
        self.restTime = False
        self.lockRef = True
        self.mainTime = False
        self.pushButton_2.setEnabled(True)
        self.endBtn.setEnabled(False)
        self.pushButton_3.setEnabled(False)
        self.pushButton_3.setText("وقفه")
        self.spinBox_2.setValue(0)
        self.spinBox.setValue(0)
        self.count = 0
        self.minutes = 0
        self.mints = 0
        self.scnd = 0
        self.spinBox.setEnabled(True)
        self.spinBox_6.setEnabled(True)
        self.spinBox_2.setEnabled(True)
        self.spinBox_5.setEnabled(True)
        self.update_gui1()
        self.updateTimer()


    def start1(self):
        self.start = True
        self.redPoints = 0
        self.bluePoints = 0
        if self.count == 0:
            self.start = False
            QMessageBox.about(self, "اخطار", "لطفا زمان خود را تعیین کنید.")
        elif ("Main" not in dir(self)):
            self.start = False
            QMessageBox.about(self, "اخطار", "لطفا اسکوربورد را باز کنید.")
        else:
            self.pushButton_3.setEnabled(True)
            self.pushButton_2.setEnabled(False)
            self.endBtn.setEnabled(True)
        
        # self.timer_start()

    def puse(self):
        if self.start == True:
            self.start = False
            self.pushButton_3.setText("ادامه")
        else:
            self.start = True
            self.pushButton_3.setText("وقفه")




    def secondsound(self):
        pass
        # playsound(r"C:\Users\HP\Desktop\time3.mp3")
    def endtime(self):
        pass
        # playsound(r"C:\Users\HP\Desktop\end.mp3")    
        
    # def colorbux(self):
    #     if 
    #     self.stackedWidget.setCurrentIndex(1)
    def openSecondWindow(self):
        if ("Main" in dir(self)):
            if (self.Main.isVisible()):
                pass
            else:
                self.Main.setVisible(True)
        else:
            # opening the second window
            self.Main = Main() # same as the second window class name
        
        
        self.Main.roundNumLbl.setText(str(self.roundNum))
        
        
    def setPortName(self):
        self.portName = (self.portsCbx.currentText()).lower()
        if (self.portName != "__"):
            self.serial = QtSerialPort.QSerialPort(self.portName,
                                                    baudRate=QtSerialPort.QSerialPort.Baud115200, 
                                                    readyRead=self.receive)
            self.connectBtn.setEnabled(True)
        else: 
            self.connectBtn.setEnabled(False)
        print(self.portName)

    @QtCore.pyqtSlot()
    def receive(self):
        # "x12x2x01x45:32:22"
        while self.serial.canReadLine():            
            self.text = self.serial.readLine().data().decode()
            self.textSplit = self.text.split('x')[1:]
            self.consoleTxt.setPlainText("received message:\n"+self.text+"\n")
            self.ardo()

    @QtCore.pyqtSlot()
    def on_toggled(self):
        if not self.serial.isOpen():
            try:
                self.serial.open(QtCore.QIODevice.ReadWrite)
                if(self.serial.isOpen()):
                    print("connected")
                    self.sendBtn.setEnabled(True)
                    self.portsCbx.setEnabled(False)
                    self.connectBtn.setText("قطع اتصال")
                    self.connectionLbl.setProperty("checked","true")
                    self.connectionLbl.style().unpolish(self.connectionLbl)
                    self.connectionLbl.style().polish(self.connectionLbl)
                    self.connectionLbl.update()
                else:
                    print("not connected")
                    QMessageBox.about(self, "خطا", "امکان برقراری ارتباط وجود ندارد!")
                    self.connectionLbl.setProperty("checked","false")
                    self.connectionLbl.style().unpolish(self.connectionLbl)
                    self.connectionLbl.update()
            except:
                QMessageBox.about(self, "خطا", "امکان برقراری ارتباط وجود ندارد!")
                pass
                
        else:
            self.connectionLbl.setProperty("checked","false")
            self.connectionLbl.style().unpolish(self.connectionLbl)
            self.connectionLbl.update()
            self.serial.close()
            self.portsCbx.setEnabled(True)
            self.connectBtn.setText("اتصال")
            self.sendBtn.setEnabled(False)
                    
    @QtCore.pyqtSlot()
    def sendToArdo(self, packet=""):
        if (packet == ""):
            packet = self.pointPacket()
            if self.serial.isOpen():
                self.consoleTxt.clear()
                self.serial.write(packet.encode())
                self.consoleTxt.setPlainText("sent message:\n"+packet+"\n")
        else:
            if self.serial.isOpen():
                self.consoleTxt.clear()
                self.serial.write(packet.encode())
                self.consoleTxt.setPlainText("sent message:\n"+packet+"\n")
    
    @QtCore.pyqtSlot()
    def pointPacket(self):
        packet = "9"
        for idx in range(5):
            # tmp = "{}{}{:02d}{:02d}".format("9",str(self.refColor[str(self.roundNum)][idx]),
            #                                 int(self.redPlayerRefsPoint[str(self.roundNum)][idx]),
            #                                 int(self.bluePlayerRefsPoint[str(self.roundNum)][idx]) )
            tmp = "{}".format(str(self.refColor[str(self.roundNum)][idx]))
            packet += tmp
        packet += "\n"
        # if self.serial.isOpen():
        #     self.consoleTxt.clear()
        #     self.serial.write(packet.encode())
        #     self.consoleTxt.setPlainText("sent message:\n"+packet+"\n")
        return packet

    @QtCore.pyqtSlot()
    def ardo(self):
        refIdx = int(self.textSplit[2])-1
        refRedLcd = redLcdList[refIdx]
        refBlueLcd = blueLcdList[refIdx]
        redPoint = int(self.textSplit[0])
        bluePoint = int(self.textSplit[1])
        myObjRed = eval("self."+refRedLcd)
        myObjBlue = eval("self."+refBlueLcd)
        myObjRed.display(str(redPoint))
        myObjBlue.display(str(bluePoint))

        


        
    def setLcdColor(self):
        self.setStackedWidgetColor(self.lcdNumber, self.lcdNumber_2, self.stackedWidget, refIdx=0)
        self.setStackedWidgetColor(self.lcdNumber_3, self.lcdNumber_4, self.stackedWidget_3, refIdx=1)
        self.setStackedWidgetColor(self.lcdNumber_7, self.lcdNumber_8, self.stackedWidget_4, refIdx=2)
        self.setStackedWidgetColor(self.lcdNumber_9, self.lcdNumber_10, self.stackedWidget_5, refIdx=3)
        self.setStackedWidgetColor(self.lcdNumber_13, self.lcdNumber_14, self.stackedWidget_6, refIdx=4)



    def setStackedWidgetColor(self, lcdRed, lcdBlue, myStackedWidget, refIdx):
        if (int(lcdRed.value()) > int(lcdBlue.value())):
            myStackedWidget.setCurrentIndex(colorDict["red"])
            self.refColor[str(self.roundNum)][int(refIdx)] = colorDict["red"]
        elif (int(lcdRed.value()) < int(lcdBlue.value())):
            myStackedWidget.setCurrentIndex(colorDict["blue"])
            self.refColor[str(self.roundNum)][int(refIdx)] = colorDict["blue"]
        else:
            self.refColor[str(self.roundNum)][int(refIdx)] = colorDict["white"]
            myStackedWidget.setCurrentIndex(colorDict["white"])

    def uploadData(self):
        if ("Main" not in dir(self)):
            QMessageBox.about(self, "اخطار", "لطفا اسکوربورد را باز کنید.")
        else:
            self.calcPoints()
            self.storeTableData()
            self.sendToArdo()
            if (self.redPoints > self.bluePoints):
                myObj = eval("self.Main.stackedWidgetRed"+str(self.roundNum))
                myObj.setCurrentIndex(1)
                myObj = eval("self.Main.stackedWidgetBlue"+str(self.roundNum))
                myObj.setCurrentIndex(0)
                self.redWinRand[int(self.roundNum-1)] = 1
                self.blueWinRand[int(self.roundNum-1)] = 0

            elif (self.redPoints < self.bluePoints):
                myObj = eval("self.Main.stackedWidgetBlue"+str(self.roundNum))
                myObj.setCurrentIndex(1)
                myObj = eval("self.Main.stackedWidgetRed"+str(self.roundNum))
                myObj.setCurrentIndex(0)
                self.redWinRand[int(self.roundNum-1)] = 0
                self.blueWinRand[int(self.roundNum-1)] = 1
            else:
                myObj = eval("self.Main.stackedWidgetBlue"+str(self.roundNum))
                myObj.setCurrentIndex(0)
                myObj = eval("self.Main.stackedWidgetRed"+str(self.roundNum))
                myObj.setCurrentIndex(0)
            self.bluePoints = 0
            self.redPoints = 0
            
            if (sum(self.redWinRand)==2):
                self.Main.winnerStackedWidget.setCurrentIndex(0)
            elif (sum(self.blueWinRand)==2):
                self.Main.winnerStackedWidget.setCurrentIndex(2)
            else:
                self.Main.winnerStackedWidget.setCurrentIndex(1)

    def resetRefLcd(self):
        children = self.findChildren(QLCDNumber)
        for child in children:
            child.display("0")

    def storeTableData(self):
        for idx in range(4):
            myObj = eval("self.red"+str(self.roundNum)+"_"+str(idx+1) + "spinBox")
            self.redPlayerTablePoint[str(self.roundNum)][idx] =  int(myObj.value())

            myObj = eval("self.blue"+str(self.roundNum)+"_"+str(idx+1) + "spinBox")
            self.bluePlayerTablePoint[str(self.roundNum)][idx] =  int(myObj.value())




    def calcPoints(self):
        self.bluePoints = 0
        self.redPoints = 0
        for idx, item in enumerate(redLcdList):
            myObjRed = eval("self."+item)
            myObjBlue = eval("self."+blueLcdList[idx])
            self.redPlayerRefsPoint[str(self.roundNum)][idx] = int(myObjRed.value())
            self.bluePlayerRefsPoint[str(self.roundNum)][idx] = int(myObjBlue.value())
            if (int(myObjRed.value()) > int(myObjBlue.value())):
                self.redPoints += 1
            elif (int(myObjRed.value()) < int(myObjBlue.value())):
                self.bluePoints += 1
            else:
                self.redPoints += 1
                self.bluePoints += 1
        if ("Main" in dir(self)):
            myProgressVal = int((float(self.redPoints)/( float(self.redPoints) +  float(self.bluePoints)))*100)
            self.Main.progressBar.setValue(myProgressVal)



    def loadCsvFile(self):
        if ("Main" not in dir(self)):
            QMessageBox.about(self, "اخطار", "لطفا اسکوربورد را باز کنید.")
        else:
            path = self.getfile()
            if (path!=""):
                self.openExcelPandas(path)
                self.gameNumCbx.clear()            
                self.gameNumCbx.addItems(map(str, self.matchData.keys()))
                self.gameNum = self.gameNumCbx.currentText()
                self.setGameNum()
            else:
                pass

    def openExcelPandas(self, path):
        df = pd.read_excel(path)
        self.matchData = {}
        
        for idx, item in df.iterrows():
            tmp = {}
            tmp["red"] = copy.deepcopy(item["قرمز"])
            tmp["blue"] = copy.deepcopy(item["آبی"])
            tmp["weight"] = copy.deepcopy(str(item["وزن"]))
            self.matchData[str(item["شماره مسابقه"])] = copy.deepcopy(tmp)
        
        

    def open_workbook(self, path):
        workbook = openpyxl.load_workbook(filename=path)
        print(f"Worksheet names: {workbook.sheetnames}")
        sheet = workbook.active
        print(f"The title of the Worksheet is: {sheet.title}")
        print(f"Cells that contain data: {sheet.calculate_dimension()}")
        for row in sheet.rows:
            for cell in row:
                if isinstance(cell, openpyxl.cell.cell.MergedCell):
                    # Skip this cell
                    continue
                print(f"{cell.column_letter}{cell.row} = {cell.value}")

    def getfile(self):
        dir_path = os.path.dirname(os.path.realpath(__file__))
        fname = QFileDialog.getOpenFileName(self, 'Open file', 
         dir_path,"Excel files (*.csv *.XLSX)")
        if (fname[0]!= ""):
            fname = fname[0].replace("/","\\")
            return(fname)
        else:
            return("")

    def _getserial_ports(self):
        ports = QtSerialPort.QSerialPortInfo.availablePorts()
        for port in ports :
            self.portsCbx.addItem(str(port.portName()))
        self.portsCbx.addItem("__")
        self.portsCbx.setCurrentText("__")

    def up1r(self):
        v = self.lcdNumber.value()
        self.lcdNumber.display(v+1)  
    def up2r(self):
        v = self.lcdNumber_3.value()
        self.lcdNumber_3.display(v+1) 
    def up3r(self):
        v = self.lcdNumber_7.value()
        self.lcdNumber_7.display(v+1) 
    def up4r(self): 
        v = self.lcdNumber_9.value()
        self.lcdNumber_9.display(v+1)
    def up5r(self):       
        v = self.lcdNumber_13.value()
        self.lcdNumber_13.display(v+1)

    def up1b(self): 
        v = self.lcdNumber_2.value()
        self.lcdNumber_2.display(v+1)
    def up2b(self): 
        v = self.lcdNumber_4.value()
        self.lcdNumber_4.display(v+1)
    def up3b(self):
        v = self.lcdNumber_8.value()
        self.lcdNumber_8.display(v+1) 
    def up4b(self): 
        v = self.lcdNumber_10.value()
        self.lcdNumber_10.display(v+1)
    def up5b(self): 
        v = self.lcdNumber_14.value()
        self.lcdNumber_14.display(v+1)

    def dn1r(self):
        v = self.lcdNumber.value()
        if (v > 0):
            self.lcdNumber.display(v-1)
    def dn2r(self):
        v = self.lcdNumber_3.value()
        if (v > 0):
            self.lcdNumber_3.display(v-1)
    def dn3r(self):
        v = self.lcdNumber_7.value()
        self.lcdNumber_7.display(v-1)
    def dn4r(self):
        v = self.lcdNumber_9.value()
        if (v > 0):
            self.lcdNumber_9.display(v-1)
    def dn5r(self):
        v = self.lcdNumber_13.value()
        if (v > 0):
            self.lcdNumber_13.display(v-1)
    
    def dn1b(self):
        v = self.lcdNumber_2.value()
        if (v > 0):
            self.lcdNumber_2.display(v-1)
    def dn2b(self):
        v = self.lcdNumber_4.value()
        if (v > 0):
            self.lcdNumber_4.display(v-1)
    def dn3b(self):
        v = self.lcdNumber_8.value()
        if (v > 0):
            self.lcdNumber_8.display(v-1)
    def dn4b(self):
        v = self.lcdNumber_10.value()
        if (v > 0):
            self.lcdNumber_10.display(v-1)
    def dn5b(self):
        v = self.lcdNumber_14.value()
        if (v > 0):
            self.lcdNumber_14.display(v-1)


    def closeEvent(self, event):
        close = QtWidgets.QMessageBox.question(self,
                                         "خروج",
                                         "آیا برای خروج از برنامه اطمینان دارید؟",
                                         QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
        if close == QtWidgets.QMessageBox.Yes:
            if ("Main" in dir(self)):
                self.Main.close()
            if(self.serial.isOpen()):
                self.serial.close()
            event.accept()
        else:
            event.ignore()

if __name__ == "__main__":

    app = QtWidgets.QApplication(sys.argv)
    app.setStyle('Fusion')
    Ui =  MainWindow()
    Ui.show()
    sys.exit(app.exec_())