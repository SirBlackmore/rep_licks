import sys, threading
import funcionesConexiones
import funcionesExcel
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QLineEdit, QListWidget, QTextEdit, QDesktopWidget, QComboBox, QSlider, QTableWidget
from PyQt5 import QtWidgets, QtGui
from PyQt5.QtCore import QSize, QPoint, QPropertyAnimation, Qt

import funcionesGeneral
import funcionesSQL
from ui_main import Ui_MainWindow  # Porque el principal se ha creado de tipo main window (no widget)
from ui_tablaFilasResultados import Ui_WindowTablaResult
from ui_password import Ui_Form_password

class Expediente():
    def __init__(self):
        self.numExp = ""
        self.verificacion = ""
        self.tipo = ""
        self.marca = ""
        self.modelo = ""
        self.numSerie = ""
        self.modalidad = ""
        self.descripcion = ""
        self.norma = ""
        self.caracteristicas = ""
        self.emp = ""

class inicio(QMainWindow):  # Crea clase partiendo del tipo QMainWindow

    def __init__(self):  # Inicializa la clase
        super().__init__()  # Con ésto si hubiera clases dentro de ésta función también las inicializaría. Y también necesario al compilar .ui como .py

        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        self.anchoInicial = 924
        self.altoInicial = 720
        self.setGeometry(0, 0, self.anchoInicial, self.altoInicial)
        self.setMaximumSize(self.anchoInicial, self.altoInicial)
        self.qtRectangle = self.frameGeometry()
        self.centerPoint = QDesktopWidget().availableGeometry().center()
        self.qtRectangle.moveCenter(self.centerPoint)
        self.move(self.qtRectangle.topLeft())

        self.ui.lbl_mensajeEstado.setText("")
        self.ui.txt_expediente.setFocus()

        self.muestraConfig = False

        # Botones home
        self.ui.btn_buscarExp.clicked.connect(self.btn_buscarExpediente)
        self.ui.btn_guardarSesion.clicked.connect(self.btnGuardarExcel)
        self.ui.btn_cargarSesion.clicked.connect(self.btnCargarHojasExcel)
        self.ui.btn_abrirCarpeta.clicked.connect(self.btnAbrirCarpeta)
        self.ui.btn_addPrecinto.clicked.connect(self.addPrecinto)
        self.ui.btn_delPrecinto.clicked.connect(self.delPrecinto)
        self.ui.btn_convertir.clicked.connect(self.btnConvertirEnsayos)

        # Botones Certificado
        self.ui.btn_cargarPreview.clicked.connect(self.btn_cargarPreview)
        self.ui.btn_tablaResultados.clicked.connect(self.btn_tablaResultados)
        self.ui.btn_tablaEMP.clicked.connect(self.btn_tablaEMP)
        self.sliderRepMod = self.findChild(QSlider, "slider_repMod")
        self.sliderRepMod.valueChanged.connect(self.slideRepMod)
        self.sliderInstalacion = self.findChild(QSlider, "sld_instalacion")
        self.sliderInstalacion.valueChanged.connect(self.slideInst)
        self.ui.btn_tablaPdf.clicked.connect(self.btn_generarPdf)

        # Botones-acciones stacked
        self.ui.stackedWidget.setCurrentWidget(self.ui.p_home)
        self.ui.btn_home.clicked.connect(self.mostrarHome)
        self.ui.btn_SdV.clicked.connect(self.mostrarSdV)
        self.ui.btn_DdR.clicked.connect(self.mostrarDdR)
        self.ui.btn_certif.clicked.connect(self.mostrarCertif)
        self.ui.btn_config.clicked.connect(self.btn_mostrarPassword)

        for w in self.ui.menuContainer.findChildren(QWidget):
                w.clicked.connect(self.aplicarEstilo)
    
    def btnConvertirEnsayos(self):
        if expediente.numExp != '':
            funcionesExcel.convertirEnsayos(expediente.numExp)
        else: self.mensajesEstado("naranja", "No se ha buscado ningún expediente")

    def addPrecinto(self):
        if self.ui.txt_precinto.text():
            self.ui.lst_precEquipo.addItem(self.ui.txt_precinto.text())
            self.ui.txt_precinto.clear()
            self.ui.txt_precinto.setFocus()

    def delPrecinto(self):
        self.ui.lst_precEquipo.takeItem(self.ui.lst_precEquipo.currentRow())

    def aplicarEstilo(self):
        for w in self.ui.menuContainer.findChildren(QWidget):
            if w.objectName() != self.sender().objectName():
                w.setStyleSheet("QPushButton {border: none;}"
                                "QPushButton:hover{"
                                "background-color: rgb(250, 250, 250);"
                                "border-left: 4px solid rgb(27, 156, 216);}")
            else: w.setStyleSheet("background-color: rgb(250, 250, 250);"
                                "border: 1px solid grey;border-left: 4px solid rgb(27, 156, 216);")

    def mostrarHome(self):
        self.ui.stackedWidget.setCurrentWidget(self.ui.p_home)
    def mostrarSdV(self):
        self.ui.stackedWidget.setCurrentWidget(self.ui.p_SdV)
    def mostrarDdR(self):
        self.ui.stackedWidget.setCurrentWidget(self.ui.p_DdR)
    def mostrarCertif(self):
        self.ui.stackedWidget.setCurrentWidget(self.ui.p_Certif)
        self.ui.lbl_motivoMod.setVisible(False)
        self.ui.txt_motivoMod.setVisible(False)
        if not str("reparación").upper() in expediente.verificacion.upper() and not str("modificación").upper() in expediente.verificacion.upper():
            self.sliderRepMod.setVisible(False)
            self.ui.lbl_selectRep.setVisible(False)
            self.ui.lbl_selectMod.setVisible(False)
        else:
            self.sliderRepMod.setVisible(True)
            self.ui.lbl_selectRep.setVisible(True)
            self.ui.lbl_selectMod.setVisible(True)
    
    def btn_mostrarPassword(self):
        if self.muestraConfig == False:
            self.password = Clase_password()
            self.password.show()
        else:
            self.mostrarConfig()

    def btnAbrirCarpeta(self):
        if self.comprobarExpIntroducido() == False:
            return
        else:
            if funcionesGeneral.abrirCarpeta(self.ui.txt_expediente.text().upper()) == "nodisp":
                self.mensajesEstado("naranja", "No se ha encontrado la carpeta correspondiente al expediente")

    def mostrarConfig(self):
        self.lista = funcionesGeneral.mostrarConfigCategoria(["celdas_sdv", "celdas_ddr", "celdas_trafreal", "celdas_certificado"])

        # Tabla configuración certificados (celdas)
        self.ui.tbl_ConfigCertif.setRowCount(len(self.lista[0]))
        self.ui.tbl_ConfigCertif.setColumnCount(2)
        self.ui.tbl_ConfigCertif.setHorizontalHeaderLabels(["Nombre", "Celda Excel"])
        for i in range(len(self.lista[0])):
            self.ui.tbl_ConfigCertif.setItem(i, 0, QtWidgets.QTableWidgetItem(self.lista[0][i]))
            self.ui.tbl_ConfigCertif.setItem(i, 1, QtWidgets.QTableWidgetItem(self.lista[1][i]))

        self.ampliacion = 300
        self.anim = QPropertyAnimation(self, b"size")
        self.anim2 = QPropertyAnimation(self, b"pos")
        self.anim.setDuration(500)
        self.anim2.setDuration(500)
        self.ancho = self.size().width()
        self.alto = self.size().height()
        if self.muestraConfig == False:
            self.setMaximumWidth(self.anchoInicial + self.ampliacion)
            self.anim.setStartValue(QSize(self.anchoInicial, self.altoInicial))
            self.anim.setEndValue(QSize(self.anchoInicial + self.ampliacion, self.altoInicial))
            self.anim2.setStartValue(QPoint(self.pos().x(), self.pos().y()))
            self.anim2.setEndValue(QPoint(self.pos().x() - round(self.ampliacion / 2), self.pos().y()))
            self.muestraConfig = True
            self.anim.start()
            self.anim2.start()
            self.activarBotones(True)
        else:
            self.anim.setStartValue(QSize(self.ancho, self.alto))
            self.anim.setEndValue(QSize(self.anchoInicial, self.altoInicial))
            self.anim2.setStartValue(QPoint(self.pos().x(), self.pos().y()))
            self.anim2.setEndValue(QPoint(self.pos().x() + round(self.ampliacion / 2), self.pos().y()))
            self.muestraConfig = False
            self.anim.start()
            self.anim2.start()
            self.anim.finished.connect(self.restaurarAncho)

    def restaurarAncho(self):
        self.setMaximumWidth(self.anchoInicial)

    def comprobarExpIntroducido(self):
        if self.ui.txt_expediente.text() == "":
            self.ui.txt_expediente.setStyleSheet("border: 1px solid red; background-color: rgb(255, 255, 225);")
            self.mensajesEstado("naranja", "No se ha introducido ningún expediente.")
            return False
        else:
            self.ui.txt_expediente.setText(self.ui.txt_expediente.text().upper())
            self.ui.txt_expediente.setStyleSheet("background-color: rgb(255, 255, 225);")
            self.ui.lbl_mensajeEstado.setText("")
            self.ui.lbl_mensajeEstado.setStyleSheet("")
            return True

    def btn_buscarExpediente(self):
        if self.comprobarExpIntroducido() == False:
            return
        else:
            expediente.numExp = self.ui.txt_expediente.text()
            # Borrado campos
            for w in self.findChildren(QLineEdit):
                if w.objectName() != "txt_expediente":
                    w.clear()
            for w in self.findChildren(QTextEdit):
                w.clear()
            for w in self.findChildren(QListWidget):
                w.clear()
            for w in self.findChildren(QTableWidget):
                w.clear()
                w.setRowCount(0)
                w.horizontalHeader().hide()

            conexion = funcionesConexiones.conectarGESLAB(expediente)

            if conexion == "error":
                self.mensajesEstado("rojo","Error de conexión")
            elif conexion == "expFalse":
                self.mensajesEstado("naranja", "Error: No se ha encontrado el expediente.")
            elif conexion == "eqFalse":
                self.mensajesEstado("naranja", "Error: No se ha encontrado el equipo.")
            else: # Si ha conectado correctamente
                self.mensajesEstado("verde", "Conectado")

                self.ui.txt_verificacion.setText(expediente.verificacion)
                self.ui.txt_marca.setText(expediente.marca)
                self.ui.txt_modelo.setText(expediente.modelo)
                self.ui.txt_numSerie.setText(expediente.numSerie)
                self.ui.txt_tipo.setText(expediente.tipo)
                self.ui.txt_modalidad.setText(expediente.modalidad)
                self.ui.txt_descripcion.setText(expediente.descripcion)
                self.ui.txt_norma.setText(expediente.norma)

                expediente.caracteristicas = funcionesSQL.caractEquipo(expediente.verificacion, "EquivVerif", 2) + ";" + funcionesSQL.caractEquipo(expediente.norma, "EquivNormas", 1) + ";" \
                                             + funcionesSQL.caractEquipo(expediente.tipo, "EquivTipos", 2) + ";" + funcionesSQL.caractEquipo(expediente.modalidad, "EquivModalidad", 1)

                expediente.emp = funcionesSQL.caractEquipo(expediente.verificacion, "EquivVerif", 2) + ";" + funcionesSQL.caractEquipo(expediente.norma, "EquivNormas", 2) + ";" + funcionesSQL.caractEquipo(expediente.tipo, "EquivTipos", 3)

                self.ui.txt_seleccionado.setText(expediente.caracteristicas)
                self.ui.txt_emp.setText(expediente.emp)

                self.activarBotones(True)

    def activarBotones(self, activar):
        self.activar = activar
        if self.activar == True:
            self.ui.btn_SdV.setEnabled(True)
            self.ui.btn_DdR.setEnabled(True)
            self.ui.btn_certif.setEnabled(True)
        else:
            self.ui.btn_SdV.setEnabled(False)
            self.ui.btn_DdR.setEnabled(False)
            self.ui.btn_certif.setEnabled(False)

    def btnGuardarExcel(self):
        resultado = funcionesExcel.guardarExcel(self.ui.txt_expediente.text())
        if resultado == True:
            self.mensajesEstado("verde", "El archivo se ha guardado correctamente.")

    def btnCargarHojasExcel(self):
        self.ui.lst_ensayos.clear()

        ensayos = funcionesExcel.cargarHojasExcel(self.ui.txt_expediente.text())

        if ensayos == "error":
            self.mensajesEstado("rojo", "Error al cargar los archivos.")
        elif ensayos == "nodisp":
            self.ui.lst_ensayos.addItem("No disponible")
            self.mensajesEstado("naranja", "No se ha encontrado el archivo de ensayos")
        else:
            for ensayo in ensayos:
                self.ui.lst_ensayos.addItem(ensayo)
            self.mensajesEstado("verde","Se han cargado correctamente los archivos.")

    def btn_tablaResultados(self):
        self.tablaRes = Clase_tablaFilasResultados("filas_resultados")
        # self.hide()
        self.tablaRes.showMaximized()

    def btn_tablaEMP(self):
        self.tablaEMP = Clase_tablaFilasResultados("emp")
        self.tablaEMP.showMaximized()

    def limpiarMensajeEstado(self):
        self.ui.lbl_mensajeEstado.setStyleSheet("background-color: rgb(235, 245, 255);")
        self.ui.lbl_mensajeEstado.setText("")

    def mensajesEstado(self, tipo, mensaje):
        t = threading.Timer(5, self.limpiarMensajeEstado)
        if tipo == "verde":
            self.ui.lbl_mensajeEstado.setStyleSheet("background-color: rgb(208, 255, 219);")
        elif tipo == "naranja":
            self.ui.lbl_mensajeEstado.setStyleSheet("background-color: rgb(255, 170, 0, 140);")
        elif tipo == "rojo":
            self.ui.lbl_mensajeEstado.setStyleSheet("background-color: rgba(255, 0, 0, 140);")
        self.ui.lbl_mensajeEstado.setText("  " + mensaje)
        t.start()

    def btn_cargarPreview(self):
        if expediente.caracteristicas == "": self.mensajesEstado("naranja", "Se debe cargar un expediente")
        else:
            if self.sliderRepMod.isVisible() == True and self.sliderRepMod.value() == 1 and self.ui.txt_motivoMod.text() == "":
                self.mensajesEstado("naranja", "No se ha introducido el motivo de la modificación")
                return

            # Diccionario filas-plantillas resultados
            self.df = funcionesSQL.buscarFilasResultadosPreview(expediente, self.ui.sld_instalacion.value())
            # Diccionario EMP
            self.dict_emp = funcionesSQL.buscarEMP(expediente.emp)
            # Diccionario resultados
            self.dict_resultados = funcionesExcel.cargarDatosEnsayos(expediente.numExp)
            if self.dict_resultados == "error":
                self.mensajesEstado("naranja", "No se han podido cargar los datos de los ensayos")
                return
            elif self.dict_resultados == "nodisp":
                self.mensajesEstado("naranja", "No se ha encontrado el archivo de resultados de los ensayos")
                return
            else:
                self.mensajesEstado("verde", "Se han cargado los datos correctamente")
            # Se unen los diccionarios EMP y resultados y se suma el texto ET y num ET
            self.dict_resultados.update(self.dict_emp)
            self.dict_resultados["texto_ET"] = funcionesSQL.buscarValor("EquivNormas", "textoEMP", expediente.norma, "nombre")
            self.dict_resultados["num_ET"] = funcionesSQL.buscarValor("extipo", "Modelo", expediente.modelo, "numET")

            self.lst_Aptos = []
            # Creación cmb_aptos y núm columnas aptos
            for i in range(len(self.df)):
                for j in range(3, len(self.df.columns)):
                    if "Apto" in str(self.df.iat[i,j]):
                        self.numAptoLinea = str(self.df.iat[i,j]).count("Apto") # Lee número de "Aptos" en col 1 (Campo)
                        for k in range(self.numAptoLinea):
                            self.cmb_Apto = QComboBox()
                            self.cmb_Apto.setEditable(True)
                            self.cmb_Apto.lineEdit().setAlignment(Qt.AlignCenter)
                            self.cmb_Apto.lineEdit().setAlignment(Qt.AlignVCenter)
                            self.cmb_Apto.setEditable(False)
                            self.cmb_Apto.setFont(QtGui.QFont('Segoe UI', 10))
                            self.cmb_Apto.setStyleSheet("border: 1px solid  rgb(216, 216, 216); background-color: rgb(255, 255, 235);")
                            self.cmb_Apto.addItem("Apto")
                            self.cmb_Apto.addItem("No Apto")
                            self.lst_Aptos.append(self.cmb_Apto)

            # Configuración tabla
            self.ui.tablePreview.clear()
            self.ui.tablePreview.setRowCount(0)
            self.numColOriginal = len(self.df.columns)
            self.numColAdicional = 0
            for nombreCol in self.df.columns:
                if "Campo" in str(nombreCol):
                    self.numColAdicional += 1
            self.ui.tablePreview.setColumnCount(self.numColOriginal + self.numColAdicional)
            self.ui.tablePreview.horizontalHeader().show()
            self.ui.tablePreview.setHorizontalHeaderLabels(self.df.columns)
            for i in range(self.numColAdicional):
                self.ui.tablePreview.setColumnHidden(self.numColOriginal-i-1, True)
                self.ui.tablePreview.setHorizontalHeaderItem(self.numColOriginal + i,QtWidgets.QTableWidgetItem("Valor " + str(i+1)))
                self.ui.tablePreview.setColumnWidth(self.numColOriginal + i, 136)
            self.ui.tablePreview.setColumnHidden(0, True)
            self.ui.tablePreview.setColumnHidden(1, True)
            self.ui.tablePreview.setColumnWidth(2, 360)
            self.numAptoCont = 0

            # Recorrido de DataFrame para rellenar tabla
            for i in range(len(self.df)):
                self.texto = ""
                #if str(self.df.iat[i, 2]) in self.dict_resultados: self.texto = self.dict_resultados[str(self.df.iat[i, 2])]
                #else: self.texto = str(self.df.iat[i, 2])
                self.texto = str(self.df.iat[i, 2])

                if str(self.df.iat[i,1]) != 'N':    # Si no está vacío
                    self.ui.tablePreview.setRowCount(self.ui.tablePreview.rowCount()+1)

                    if "espacio" in self.texto:
                        self.ui.tablePreview.hideRow(self.ui.tablePreview.rowCount()-1)

                    for j in range(self.numColAdicional):  # Mira valores en las posibles columnas
                        self.emp = ""
                        self.valor = ""
                        if self.df.iat[i, 3+j]: # Si además hay algún valor, apto, emp
                            # Asignación desplegables combo
                            if "Apto" in str(self.df.iat[i, 3+j]):
                                self.ui.tablePreview.setCellWidget(self.ui.tablePreview.rowCount() - 1, self.numColOriginal+j, self.lst_Aptos[self.numAptoCont])
                                self.numAptoCont = self.numAptoCont + 1
                            elif "resultado_cmb" in str(self.df.iat[i, 3+j]):
                                self.cmb_superado = QComboBox(self.cmb_Apto)
                                self.cmb_superado.setEditable(True)
                                self.cmb_superado.lineEdit().setAlignment(Qt.AlignCenter)
                                self.cmb_superado.lineEdit().setAlignment(Qt.AlignVCenter)
                                self.cmb_superado.setEditable(False)
                                self.cmb_superado.setFont(QtGui.QFont('Segoe UI', 10))
                                self.cmb_superado.setStyleSheet(
                                    "border: 1px solid  rgb(216, 216, 216); background-color: rgb(255, 255, 235);")
                                self.cmb_superado.clear()
                                self.cmb_superado.addItem("SUPERADO")
                                self.cmb_superado.addItem("NO SUPERADO")
                                self.ui.tablePreview.setCellWidget(self.ui.tablePreview.rowCount()-1, self.numColOriginal+j, self.cmb_superado)
                            else:
                                self.listaValores = str(self.df.iat[i, 3+j]).split(" ")
                                if "valor" in self.df.iat[i, 3+j]:
                                    for k in range(len(self.listaValores)):
                                        if self.listaValores[k] in self.dict_resultados:
                                            if self.dict_resultados[self.listaValores[k]]: # Si el elemento buscado no está vacío en el diccionario
                                                if "valor" in str(self.listaValores[k]):
                                                    self.valor = self.valor + str(max(self.dict_resultados[self.listaValores[k]], key=abs)).replace(".",",")
                                                else: self.valor = self.valor + " " + str(self.dict_resultados[self.listaValores[k]])
                                            else:
                                                self.valor = "No encontrado"
                                                break
                                        else: self.valor = self.valor + " " + self.listaValores[k]
                                    self.ui.tablePreview.setItem(self.ui.tablePreview.rowCount() - 1, self.numColOriginal+j, QtWidgets.QTableWidgetItem(self.valor))
                                    self.ui.tablePreview.item(self.ui.tablePreview.rowCount() - 1, self.numColOriginal+j).setTextAlignment(Qt.AlignCenter)
                                    self.ui.tablePreview.item(self.ui.tablePreview.rowCount() - 1, self.numColOriginal+j).setBackground(QtGui.QColor(255, 230, 172))
                                else:
                                    if "objeto_texto" in self.df.iat[i, 3+j]:
                                        if self.sliderRepMod.isVisible() == True and self.sliderRepMod.value() == 1:
                                            self.valor = funcionesSQL.caractEquipo("modificación", "EquivVerif", 3) + " por " + self.ui.txt_motivoMod.text() + ", " \
                                            "" + funcionesSQL.caractEquipo(expediente.caracteristicas.split(";")[2] + ";" + expediente.caracteristicas.split(";")[0], "textosProced", 1)
                                        elif self.sliderRepMod.isVisible() == True and self.sliderRepMod.value() == 0:
                                            self.valor = funcionesSQL.caractEquipo("reparación", "EquivVerif",3) + " " \
                                            "" + funcionesSQL.caractEquipo(expediente.caracteristicas.split(";")[2] + ";" + expediente.caracteristicas.split(";")[0], "textosProced", 1)
                                        else:
                                            self.valor = funcionesSQL.caractEquipo(expediente.verificacion, "EquivVerif",3) + " " \
                                            "" + funcionesSQL.caractEquipo(expediente.caracteristicas.split(";")[2] + ";" + expediente.caracteristicas.split(";")[0], "textosProced", 1)
                                        self.ui.txt_objeto.setText(self.valor)
                                    else:
                                        for k in range(len(self.listaValores)):
                                            if self.listaValores[k] in self.dict_resultados:
                                                if self.dict_resultados[self.listaValores[k]]: # Si el elemento buscado no está vacío en el diccionario
                                                    self.valor = self.valor + " " + str(self.dict_resultados[self.listaValores[k]])
                                            else:
                                                self.valor = self.valor + " " + str(self.listaValores[k])

                                    self.ui.tablePreview.setItem(self.ui.tablePreview.rowCount() - 1, self.numColOriginal+j, QtWidgets.QTableWidgetItem(self.valor))
                                    self.ui.tablePreview.item(self.ui.tablePreview.rowCount() - 1, self.numColOriginal+j).setTextAlignment(Qt.AlignCenter)
                                    self.ui.tablePreview.item(self.ui.tablePreview.rowCount() - 1, self.numColOriginal+j).setBackground(QtGui.QColor(235, 245, 255))
                            # Ésta además pone el valor original de "Campo" para mantenerlo en el df al exportar a PDF
                            self.ui.tablePreview.setItem(self.ui.tablePreview.rowCount() - 1, 3+j, QtWidgets.QTableWidgetItem(self.df.iat[i, 3+j]))
                    self.ui.tablePreview.setItem(self.ui.tablePreview.rowCount()-1, 0, QtWidgets.QTableWidgetItem(str(self.df.iat[i, 0]))) # Nº Fila Excel
                    self.ui.tablePreview.setItem(self.ui.tablePreview.rowCount()-1, 1, QtWidgets.QTableWidgetItem(str(self.df.iat[i, 1]))) # SI / NO
                    self.ui.tablePreview.setItem(self.ui.tablePreview.rowCount()-1, 2, QtWidgets.QTableWidgetItem(self.texto)) # Título (texto)
            self.ui.btn_tablaPdf.setEnabled(True)

        def closeEvent(self, event):
            for window in QApplication.topLevelWidgets():
                window.close()

    def btn_generarPdf(self):
        
        self.columnas = [i for i in range(self.ui.tablePreview.columnCount())]
        df = pd.DataFrame(columns=[self.ui.tablePreview.horizontalHeaderItem(col).text() for col in self.columnas],
                          index=range(self.ui.tablePreview.rowCount()))

        for i in range(self.ui.tablePreview.rowCount()):
            for j in range(self.ui.tablePreview.columnCount()):
                if self.ui.tablePreview.item(i,j):  # Si no está vacío
                    df.iat[i,j] = self.ui.tablePreview.item(i, j).text()
                else:
                    self.widget_cmb = self.ui.tablePreview.cellWidget(i, j)
                    if isinstance(self.widget_cmb, QComboBox):
                        df.iat[i, j] = self.widget_cmb.currentText()
                    else: df.iat[i,j] = ""

        self.lst_precintos = []
        for i in range(self.ui.lst_precEquipo.count()):
            self.lst_precintos.append(self.ui.lst_precEquipo.item(i).text())

        self.mensaje = funcionesExcel.generarPdf(df, expediente.numExp, self.numColAdicional, self.lst_precintos)
        if self.mensaje == "ok":
            self.mensajesEstado("verde", "Se ha generado el PDF correctamente.")
        elif self.mensaje == "nodisp":
            self.mensajesEstado("naranja", "No se ha encontrado la plantilla de resultados.")
        elif self.mensaje == "error":
            self.mensajesEstado("naranja", "Ha ocurrido un error al generar la tabla de resultados.")
        elif self.mensaje == "errorPDFabierto":
            self.mensajesEstado("naranja", "Cerrar archivo PDF antes de generar de nuevo")

    def slideRepMod(self, valor):
        if valor == 0:
            self.ui.lbl_selectMod.setStyleSheet("color: rgb(221, 221, 221);"), self.ui.lbl_selectRep.setStyleSheet("color: black;")
            self.ui.lbl_motivoMod.setVisible(False)
            self.ui.txt_motivoMod.setVisible(False)
        else:
            self.ui.lbl_selectRep.setStyleSheet("color: rgb(221, 221, 221);"), self.ui.lbl_selectMod.setStyleSheet("color: black;")
            self.ui.lbl_motivoMod.setVisible(True)
            self.ui.txt_motivoMod.setVisible(True)
    
    def slideInst(self, valor):
        if valor == 0:
            self.ui.lbl_selectInstPortico.setStyleSheet("color: rgb(221, 221, 221);"), self.ui.lbl_selectInstOtros.setStyleSheet("color: black;")
        else:
            self.ui.lbl_selectInstOtros.setStyleSheet("color: rgb(221, 221, 221);"), self.ui.lbl_selectInstPortico.setStyleSheet("color: black;")

class Clase_tablaFilasResultados(QMainWindow):  # Crea clase partiendo del tipo QMainWindow

    def __init__(self, tabla):  # Inicializa la clase
        super().__init__()  # Con ésto si hubiera clases dentro de ésta función también las inicializaría. Y también necesario al compilar .ui como .py

        self.tabla = tabla
        self.ui = Ui_WindowTablaResult()
        self.ui.setupUi(self)

        self.df = funcionesSQL.datosFilasResultados(self.tabla)
        self.ui.tbl_filasResultados.setColumnCount(len(self.df.columns))
        self.ui.tbl_filasResultados.setRowCount(len(self.df))
        self.ui.tbl_filasResultados.setHorizontalHeaderLabels(self.df.columns)

        for i in range(len(self.df)):
            for j in range(len(self.df.columns)):
                self.ui.tbl_filasResultados.setItem(i,j, QtWidgets.QTableWidgetItem(str(self.df.iat[i,j])))

        self.ui.btn_insertarCol.clicked.connect(self.btnInsertarCol)
        self.ui.btn_insertarFila.clicked.connect(self.btnInsertarFila)
        self.ui.btn_eliminarFila.clicked.connect(self.btnEliminarFila)
        self.ui.btn_eliminarCol.clicked.connect(self.btnEliminarCol)
        self.ui.btn_confirmar.clicked.connect(self.btnConfirmar)
        self.ui.btn_crearDesdeExcel.clicked.connect(funcionesExcel.tablaDesdeExcel)

    def btnInsertarCol(self):
        self.columnaSel = self.ui.tbl_filasResultados.currentColumn()
        self.ui.tbl_filasResultados.insertColumn(self.columnaSel)

    def btnInsertarFila(self):
        self.filaSel = self.ui.tbl_filasResultados.currentRow()
        self.ui.tbl_filasResultados.insertRow(self.filaSel)

    def btnEliminarCol(self):
        self.columnaSel = self.ui.tbl_filasResultados.currentColumn()
        self.ui.tbl_filasResultados.removeColumn(self.columnaSel)

    def btnEliminarFila(self):
        self.filaSel = self.ui.tbl_filasResultados.currentRow()
        self.ui.tbl_filasResultados.removeRow(self.filaSel)

    def btnConfirmar(self):
        # for i in range(self.ui.tbl_filasResultados.)
        columnas = []
        for i in range(self.ui.tbl_filasResultados.columnCount()):
            columnas.append(self.ui.tbl_filasResultados.horizontalHeaderItem(i).text())
        df = pd.DataFrame(columns=columnas)

        nuevaFila = []
        for i in range(self.ui.tbl_filasResultados.rowCount()):
            for j in range(self.ui.tbl_filasResultados.columnCount()):
                        nuevaFila.append(self.ui.tbl_filasResultados.item(i, j).text())
            df.loc[len(df.index)] = nuevaFila
            nuevaFila = []

        funcionesSQL.guardarFilasResultados(df, self.tabla)

class Clase_password(QWidget):
    def __init__(self):
        super().__init__()

        self.password = "labcinem"
        self.m_ui = Ui_Form_password()
        self.m_ui.setupUi(self)
        self.m_ui.btn_aceptar.clicked.connect(self.btn_aceptar)

    def btn_aceptar(self):
        if self.m_ui.txt_password.text() == self.password:
            ventanaPrincipal.mostrarConfig()
            self.close()
        else:
            self.m_ui.lbl_mensajeEstado.setStyleSheet("background-color: rgb(255, 170, 0, 140);")
            self.m_ui.lbl_mensajeEstado.setText("Contraseña incorrecta")


if __name__ == '__main__':
    expediente = Expediente()
    app = QApplication(sys.argv)
    ventanaPrincipal = inicio()
    ventanaPrincipal.show()

    sys.exit(app.exec_())

