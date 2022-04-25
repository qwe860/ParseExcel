import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets  import QApplication, QWidget, QLabel
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QMovie

class LoadingScreen(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setFixedSize(200,200)
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.CustomizeWindowHint)

        self.label_animation = QLabel(self)

        self.movie = QMovie('Loading.gif')
        self.label_animation.setMovie(self.movie)
        self.startAnimation()
        print('address of self', id(self))
        self.show()


    def startAnimation(self):
        self.movie.start()
        print('animatin started')

    def stopAnimation(self):
        self.movie.stop()
        print('animationn stopped')
        self.close()