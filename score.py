from PyQt5 import QtWidgets, uic
from PyQt5.QtWidgets import *
from PyQt5 import QtCore, QtGui , QtSerialPort ,QtMultimedia
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtSerialPort import QSerialPortInfo
from PyQt5.QtMultimedia import QMediaPlayer, QMediaContent
from PyQt5.QtMultimedia import QSound
from playsound import playsound
import sys
import  resources_rc
from PyQt5.uic import loadUiType


def loadUiClass(path):
    stream = QFile(path)
    stream.open(QFile.ReadOnly)
    try:
        return loadUiType(stream)[0]
    finally:
        stream.close()

a = 1
class Main (QMainWindow, loadUiClass(':/ui_files/score.ui')):
    def __init__(self):
        super( Main, self).__init__()
        self.setupUi(self)
        self.setWindowFlags(Qt.FramelessWindowHint |  QtCore.Qt.MSWindowsFixedSizeDialogHint)
        # self.setWindowFlags(self.windowFlags() | QtCore.Qt.MSWindowsFixedSizeDialogHint)
        self.show()
        self.a = 1
        self.b = 0
        self.start = True


    def right_menu(self, pos):
        menu = QMenu()
        
        if self.isMaximized():
            maximize_option = menu.addAction('Minimize')
            exit_option = menu.addAction('Exit')
            # Menu option events
            maximize_option.triggered.connect(lambda:self.showNormal())
        else:
            maximize_option = menu.addAction('Maximize')
            exit_option = menu.addAction('Exit')
            # Menu option events
            maximize_option.triggered.connect(lambda:self.showMaximized())

        exit_option.triggered.connect(lambda: self.close())

        # Position
        menu.exec_(self.mapToGlobal(pos))
	

    def mousePressEvent(self, event):
        self.oldPosition = event.globalPos()
        
        if event.button() == Qt.RightButton:
            self.currentPosition = event.globalPos() - self.oldPosition
            self.right_menu(self.currentPosition)
            
    def keyPressEvent(self, e):  
        if e.key() == QtCore.Qt.Key_F11:
            if self.isMaximized():
                self.showNormal()
            else:
                self.showMaximized()
	
    def mouseMoveEvent(self, event):
        delta = QPoint(event.globalPos() - self.oldPosition)
        self.move(self.x() + delta.x(), self.y() + delta.y())
        
        self.oldPosition = event.globalPos()



    def r1(self):
        if self.a == 0:
            self.a = 1
        else:
            self.a = 0
        self.stackedWidgetRed1.setCurrentIndex(self.a)
    def r2(self):
        if self.a == 0:
            self.a = 1
        else:
            self.a = 0
        self.stackedWidgetRed2.setCurrentIndex(self.a)
    def r3(self):
        if self.a == 0:
            self.a = 1
        else:
            self.a = 0
        self.stackedWidgetRed3.setCurrentIndex(self.a)
   

    def b1(self):
        if self.a == 0:
            self.a = 1
        else:
            self.a = 0
        self.stackedWidgetBlue1.setCurrentIndex(self.a)
    def b2(self):
        if self.a == 0:
            self.a = 1
        else:
            self.a = 0
        self.stackedWidgetBlue2.setCurrentIndex(self.a)
    def b3(self):
        if self.a == 0:
            self.a = 1
        else:
            self.a = 0
        self.stackedWidgetBlue3.setCurrentIndex(self.a)
    def winer(self):
        if self.b == 0:
            self.b = +1
        
        else:
            self.b = 0
        self.stackedWidget_17.setCurrentIndex(self.b)



    def timer_start(self):
        self.a = 2
        self.time_left_int = 0
        

        self.my_qtimer = QTimer(self)
        self.my_qtimer.timeout.connect(self.timer_timeout)
        self.my_qtimer.start(1000) 
        self.update_gui()
        
    def start1(self):
        self.start = True
        self.timer_start()

    def puse(self):
        tm = 0
        if self.start == True:
            self.start = False
        else:
            self.start = True
    def end(self):
        self.start = False

    def timer_timeout(self):
       
        if self.start == True :
            self.time_left_int -= 1
          

            if self.time_left_int == -1:
                self.a -= 1                                                             
                self.time_left_int = 59
            if self.a == 0:
                if self.time_left_int < 10:
                    self.secondsound()
                if self.time_left_int == 0:
                    self.endtime()

            if self.a == 0:
                if self.time_left_int == 0:
                    self.start = False
                    # self.update_gui()
                    # time.sleep(5)
                    self.resttime()
                    
                    

        self.update_gui()

    def update_gui(self):
        
        self.label_3.setText(str(self.a))
        self.label_5.setText(str( self.time_left_int))

if __name__ == "__main__":

    app = QtWidgets.QApplication(sys.argv)
    Ui =  Main()
    Ui.show()
    sys.exit(app.exec_())