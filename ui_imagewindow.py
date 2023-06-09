# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'C:\Users\mmartinezh\Documents\Programacion\Licks Python OK\ui_imagewindow.ui'
#
# Created by: PyQt5 UI code generator 5.15.4
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_ImageWindow(object):
    def setupUi(self, ImageWindow):
        ImageWindow.setObjectName("ImageWindow")
        ImageWindow.resize(1201, 578)
        ImageWindow.setMinimumSize(QtCore.QSize(24, 24))
        ImageWindow.setMaximumSize(QtCore.QSize(999999, 999999))
        ImageWindow.setStyleSheet("QMainWindow {\n"
"    background-color: rgb(255, 255, 255);\n"
"}\n"
"\n"
"QPushButton {\n"
"    border: 1px solid;\n"
"    border-radius: 10px;\n"
"    border-color: rgb(173, 173, 173);\n"
"    background-color: rgb(225, 225, 225);\n"
"}\n"
"QPushButton:hover {\n"
"    background-color: rgb(218, 255, 215);\n"
"}\n"
"\n"
"QBoxLayout {\n"
"    border: 1px solid black;\n"
"    border-radius: 4px;\n"
"    background-color: rgb(255, 255, 255);\n"
"}\n"
"\n"
"QScrollArea {\n"
"    background-color: rgb(255, 255, 255);\n"
"}\n"
"\n"
"QSlider::groove { \n"
"\n"
"}\n"
"\n"
"QSlider::handle { \n"
"    background-color: rgb(0, 122, 217);\n"
"    border: 2px solid  rgb(0, 122, 217);\n"
"    border-radius: 4px;\n"
"}")
        self.centralwidget = QtWidgets.QWidget(ImageWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.centralwidget)
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout.setSpacing(0)
        self.verticalLayout.setObjectName("verticalLayout")
        self.cont_Main = QtWidgets.QWidget(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.cont_Main.sizePolicy().hasHeightForWidth())
        self.cont_Main.setSizePolicy(sizePolicy)
        self.cont_Main.setObjectName("cont_Main")
        self.verticalLayout_3 = QtWidgets.QVBoxLayout(self.cont_Main)
        self.verticalLayout_3.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout_3.setSpacing(0)
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.scrollArea = QtWidgets.QScrollArea(self.cont_Main)
        self.scrollArea.setStyleSheet("QScrollArea {\n"
"    background-color: rgb(255, 255, 255);\n"
"    border: none;\n"
"}")
        self.scrollArea.setWidgetResizable(True)
        self.scrollArea.setAlignment(QtCore.Qt.AlignCenter)
        self.scrollArea.setObjectName("scrollArea")
        self.scrollAreaWidgetContents = QtWidgets.QWidget()
        self.scrollAreaWidgetContents.setGeometry(QtCore.QRect(0, 0, 1201, 510))
        self.scrollAreaWidgetContents.setStyleSheet("#scrollAreaWidgetContents {\n"
"    background-color: rgb(255, 255, 255);\n"
"    border: 1px solid black;\n"
"    border-radius: 4px;\n"
"}")
        self.scrollAreaWidgetContents.setObjectName("scrollAreaWidgetContents")
        self.verticalLayout_4 = QtWidgets.QVBoxLayout(self.scrollAreaWidgetContents)
        self.verticalLayout_4.setObjectName("verticalLayout_4")
        self.vLayout = QtWidgets.QVBoxLayout()
        self.vLayout.setContentsMargins(0, 0, 0, 0)
        self.vLayout.setObjectName("vLayout")
        self.verticalLayout_4.addLayout(self.vLayout)
        self.scrollArea.setWidget(self.scrollAreaWidgetContents)
        self.verticalLayout_3.addWidget(self.scrollArea)
        self.verticalLayout.addWidget(self.cont_Main)
        self.widget = QtWidgets.QWidget(self.centralwidget)
        self.widget.setMinimumSize(QtCore.QSize(0, 68))
        self.widget.setMaximumSize(QtCore.QSize(16777215, 68))
        self.widget.setObjectName("widget")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(self.widget)
        self.verticalLayout_2.setContentsMargins(10, 0, 10, 10)
        self.verticalLayout_2.setSpacing(0)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_3.setContentsMargins(10, 10, -1, 10)
        self.horizontalLayout_3.setSpacing(10)
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.btn_play = QtWidgets.QPushButton(self.widget)
        self.btn_play.setMinimumSize(QtCore.QSize(148, 32))
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/images/iconos/play.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.btn_play.setIcon(icon)
        self.btn_play.setIconSize(QtCore.QSize(24, 24))
        self.btn_play.setObjectName("btn_play")
        self.horizontalLayout_3.addWidget(self.btn_play)
        self.btn_pause = QtWidgets.QPushButton(self.widget)
        self.btn_pause.setMinimumSize(QtCore.QSize(148, 32))
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap(":/images/iconos/pause.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.btn_pause.setIcon(icon1)
        self.btn_pause.setIconSize(QtCore.QSize(24, 24))
        self.btn_pause.setObjectName("btn_pause")
        self.horizontalLayout_3.addWidget(self.btn_pause)
        self.btn_stop = QtWidgets.QPushButton(self.widget)
        self.btn_stop.setMinimumSize(QtCore.QSize(148, 32))
        icon2 = QtGui.QIcon()
        icon2.addPixmap(QtGui.QPixmap(":/images/iconos/stop.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.btn_stop.setIcon(icon2)
        self.btn_stop.setIconSize(QtCore.QSize(20, 20))
        self.btn_stop.setObjectName("btn_stop")
        self.horizontalLayout_3.addWidget(self.btn_stop)
        self.sld_speed = QtWidgets.QSlider(self.widget)
        self.sld_speed.setMinimumSize(QtCore.QSize(210, 20))
        self.sld_speed.setMaximumSize(QtCore.QSize(210, 20))
        self.sld_speed.setMinimum(5)
        self.sld_speed.setMaximum(20)
        self.sld_speed.setProperty("value", 10)
        self.sld_speed.setOrientation(QtCore.Qt.Horizontal)
        self.sld_speed.setObjectName("sld_speed")
        self.horizontalLayout_3.addWidget(self.sld_speed)
        self.sld_vol = QtWidgets.QSlider(self.widget)
        self.sld_vol.setMinimumSize(QtCore.QSize(21, 52))
        self.sld_vol.setMaximumSize(QtCore.QSize(21, 16777215))
        self.sld_vol.setMaximum(100)
        self.sld_vol.setProperty("value", 80)
        self.sld_vol.setOrientation(QtCore.Qt.Vertical)
        self.sld_vol.setObjectName("sld_vol")
        self.horizontalLayout_3.addWidget(self.sld_vol)
        self.verticalLayout_2.addLayout(self.horizontalLayout_3)
        self.verticalLayout.addWidget(self.widget)
        ImageWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(ImageWindow)
        QtCore.QMetaObject.connectSlotsByName(ImageWindow)

    def retranslateUi(self, ImageWindow):
        _translate = QtCore.QCoreApplication.translate
        ImageWindow.setWindowTitle(_translate("ImageWindow", "MainWindow"))
        self.btn_play.setText(_translate("ImageWindow", "  Play"))
        self.btn_pause.setText(_translate("ImageWindow", "  Pause"))
        self.btn_stop.setText(_translate("ImageWindow", "  Stop"))
import images
