import os
import sys
from ui_main import Ui_MainWindow
from ui_subfolder import Ui_SubFolderWindow
from ui_imagewindow import Ui_ImageWindow

from PyQt5.QtCore import Qt, QSize, QUrl
from PyQt5.QtGui import QPixmap, QIcon
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QHBoxLayout, QLabel, QSizePolicy, QMessageBox, QSlider
from PyQt5.QtMultimedia import QMediaPlayer, QMediaContent

class ImageWindow(QMainWindow):
    def __init__(self, folder, image):
        super().__init__()
        self.ui = Ui_ImageWindow()
        self.ui.setupUi(self)
        self.setFixedWidth(1200)
        self.setWindowTitle(image.split('.')[0])
        pixmap = QPixmap(os.path.join(folder, image))
        label = QLabel(self)
        label.setMinimumWidth(self.width())
        label.setPixmap(pixmap.scaledToWidth(label.width(), Qt.SmoothTransformation))
        self.ui.vLayout.addWidget(label)

        self.player = QMediaPlayer()
        ruta = os.path.join(folder, image.split('.')[0] + '.mp3')
        url = QUrl.fromLocalFile(ruta)
        content = QMediaContent(url)
        self.player.setMedia(content)
        self.reproducir()


        self.player.mediaStatusChanged[QMediaPlayer.MediaStatus].connect(self.media_status_changed)

        self.playbackButtons()

    def media_status_changed(self, status):
        if status == QMediaPlayer.EndOfMedia:
            self.detener()

    def closeEvent(self, event):
        self.detener()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            self.close()
        if event.key() == Qt.Key_Space:
            if self.playing == False:
                self.reproducir()
            else:
                self.pausar()

    def reproducir(self):
        self.player.play()
        self.playing = True
    def detener(self):
        self.player.stop()
        self.playing = False
    def pausar(self):
        self.player.pause()
        self.playing = False
    def movSliderSpeed(self):
        self.player.setPlaybackRate(round(self.speed_slider.value()/10))
    def movSliderVol(self):
        self.player.setVolume(self.vol_slider.value())

    def playbackButtons(self):
        self.ui.btn_play.clicked.connect(self.reproducir)
        self.ui.btn_pause.clicked.connect(self.pausar)
        self.ui.btn_stop.clicked.connect(self.detener)

        self.speed_slider = self.ui.sld_speed
        self.speed_slider.sliderMoved.connect(self.movSliderSpeed)

        self.vol_slider = self.ui.sld_vol
        self.vol_slider.sliderMoved.connect(self.movSliderVol)

class SubFolderWindow(QMainWindow):
    def __init__(self, folder, images):
        super().__init__()
        self.ui = Ui_SubFolderWindow()
        self.ui.setupUi(self)
        self.setWindowTitle(os.path.split(folder)[1])
        #self.central_widget = QWidget()
        #self.setCentralWidget(self.central_widget)
        #layout = QVBoxLayout(self.central_widget)

        for image in images:
            extension = os.path.splitext(image)
            if str(extension[1]) == ".jpg":
                button = QPushButton(self)
                pixmap = QPixmap(folder + '/' + image)
                button.setIcon(QIcon(pixmap))
                button.setText("  " + os.path.splitext(image)[0])
                button.setIconSize(QSize(300, 80))
                button.clicked.connect(lambda _, image=image: self.showImage(folder, image))
                self.ui.vLayout.addWidget(button)

    def showImage(self, folder, image):
        self.image_window = ImageWindow(folder, image)
        self.image_window.show()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            self.close()

class MainWindow(QMainWindow):
    def __init__(self, folder):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        self.folder = folder
        self.setWindowTitle('Licks - Categorías')
        self.ui.btn_titulo.clicked.connect(self.mostrarMensaje)

        #self.central_widget = QWidget()
        #self.setCentralWidget(self.central_widget)

        #layout = QVBoxLayout(self.central_widget)
        subfolders = next(os.walk(self.folder))[1]
        for subfolder in subfolders:
            button = QPushButton(subfolder, self)
            button.clicked.connect(lambda _, subfolder=subfolder: self.showSubFolder(subfolder))
            button.setMinimumHeight(40)
            self.ui.vLayout.addWidget(button)

    def showSubFolder(self, subfolder):
        images = os.listdir(os.path.join(self.folder, subfolder))
        self.subfolder_window = SubFolderWindow(os.path.join(self.folder, subfolder), images)
        self.subfolder_window.show()

    def mostrarMensaje(self):
        self.mensaje = QMessageBox()
        self.mensaje.setIcon(QMessageBox.Information)
        self.mensaje.setText("Creado por M. Martinez 2023")
        #self.mensaje.setInformativeText("This is additional information")
        #self.mensaje.setWindowTitle("MessageBox demo")
        #self.mensaje.setDetailedText("The details are as follows:")
        self.mensaje.show()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            self.close()

    def closeEvent(self, event):
        for window in QApplication.topLevelWidgets():
            window.close()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow(os.path.dirname(os.path.abspath(__file__)))
    window.show()
    sys.exit(app.exec_())
