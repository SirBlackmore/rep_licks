# -------------------------------------------------------------------------------------------------------------
# |                         TAREAS PENDIENTES
# |
# |     1: Hacer que en config cinem sea independiente simul manual/auto, retorno manual/auto
# |
# -------------------------------------------------------------------------------------------------------------




import sys
from threading import Timer
import funcionesConexiones
import funcionesExcel
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QLineEdit, QListWidget, QTextEdit, QDesktopWidget, QComboBox, QSlider, QTableWidget, QPlainTextEdit, QLabel, QStyledItemDelegate, QPushButton, QFileDialog
from PyQt5 import QtWidgets, QtGui
from PyQt5.QtCore import QSize, QPoint, QPropertyAnimation, Qt

import funcionesGeneral
import funcionesSQL
from ui_main import Ui_MainWindow  # Porque el principal se ha creado de tipo main window (no widget)
from ui_tablaFilasResultados import Ui_WindowTablaResult
from ui_trafReal import Ui_WindowTrafReal
from ui_password import Ui_Form_password
# from ventana_ddr import Clase_ddr
import funciones_SdV

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
        self.precintos = []

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
        self.provincias = funcionesSQL.buscarColumna("id_provincia", "Provincia")

        for provincia in self.provincias:
            self.ui.cmb_idProvincia.addItem(provincia)

        self.ui.cmb_idProvincia.activated.connect(self.cambiarIdProvincia)

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
        self.ui.btn_trafReal.clicked.connect(self.btn_trafReal)

        # Botones SdV
        self.ui.txt_velocidad_sdv.textChanged.connect(self.calcFrecuencia)
        self.ui.btn_enviar_velocidad.clicked.connect(self.enviarVelocidad)
        self.ui.btn_add_sdv.clicked.connect(self.addConfig)
        self.ui.sld_sdv.valueChanged.connect(self.sld_sdv)
        self.ui.sld_retorno.valueChanged.connect(self.sld_retorno)
        self.ui.lst_modelos_sdv.itemClicked.connect(self.lst_configClickedEvent)
        self.ui.btn_del_sdv.clicked.connect(self.eliminarModelo)
        self.ui.btn_update_sdv.clicked.connect(self.actualizarConfigModelo)
        self.ui.cmb_puertos.activated.connect(self.cambiarPuerto)

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

    def cambiarIdProvincia(self):
        self.ui.lbl_idProvincia.setText(funcionesSQL.buscarValor("id_provincia", "Provincia", self.ui.cmb_idProvincia.currentText(), "Id"))

    def mostrarHome(self):
        self.ui.stackedWidget.setCurrentWidget(self.ui.p_home)
    
    def mostrarSdV(self):
        self.ui.stackedWidget.setCurrentWidget(self.ui.p_SdV)
        self.actualizarListaModelos()
        try:
            self.cargarConfigCinem(expediente.modelo)
        except: # Si no encuentra la configuación (da error), sólo copia marca y modelo
            self.ui.txt_marca_sdv.setText(expediente.marca)
            self.ui.txt_modelo_sdv.setText(expediente.modelo)
        
        self.puertos = funciones_SdV.listaPuertos()
        self.ui.cmb_puertos.clear()
        for puerto in self.puertos:
            self.ui.cmb_puertos.addItem(puerto)

        self.puertoSel = funcionesGeneral.buscarConfig("config_sdv", "puerto")
        if self.puertoSel in self.puertos:
            self.ui.cmb_puertos.setCurrentText(self.puertoSel)

        self.ui.lst_sentido.setCurrentRow(0)

    def mostrarDdR(self):
        self.ventana_DdR = Clase_ddr()
        self.ventana_DdR.showMaximized()

    def btn_trafReal(self):
        self.tablaRes = Clase_tablaTrafReal()
        # self.hide()
        self.tablaRes.showMaximized()

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
                if w.objectName() != "lst_sentido":
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

                for precinto in expediente.precintos:
                    self.ui.lst_precEquipo.addItem(precinto)

                self.activarBotones(True)

    def activarBotones(self, activar):
        self.activar = activar
        if self.activar == True:
            self.ui.btn_SdV.setEnabled(True)
            self.ui.btn_DdR.setEnabled(True)
            self.ui.btn_trafReal.setEnabled(True)
            self.ui.btn_certif.setEnabled(True)
        else:
            self.ui.btn_SdV.setEnabled(False)
            self.ui.btn_DdR.setEnabled(False)
            self.ui.btn_trafReal.setEnabled(False)
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
        t = Timer(5, self.limpiarMensajeEstado)
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
            # Código provincia
            if self.ui.txt_numIdProvincia.text() != "":
                self.dict_resultados["id_provincia"] = self.ui.lbl_idProvincia.text() + self.ui.txt_numIdProvincia.text()
            else:
                self.dict_resultados["id_provincia"] = "No aplicable"

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
            self.ui.tablePreview.horizontalHeader().setFont(QtGui.QFont('', weight = QtGui.QFont.Bold)) # En '' iría la fuente
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
                self.itemTexto = QtWidgets.QTableWidgetItem("")
                self.itemTexto.setFlags(self.itemTexto.flags() & ~(Qt.ItemIsEditable))
                self.itemTexto.setText(str(self.df.iat[i, 2]))

                if str(self.df.iat[i,1]) != 'N':    # Si no está vacío
                    self.ui.tablePreview.setRowCount(self.ui.tablePreview.rowCount()+1)

                    if "espacio" in self.itemTexto.text():
                        self.ui.tablePreview.hideRow(self.ui.tablePreview.rowCount()-1)

                    for j in range(self.numColAdicional):  # Mira valores en las posibles columnas
                        self.valor = ""
                        if self.df.iat[i, 3+j]: # Si además hay algún valor, apto, emp
                            # Asignación desplegables combo
                            if "Apto" in str(self.df.iat[i, 3+j]):
                                self.ui.tablePreview.setCellWidget(self.ui.tablePreview.rowCount() - 1, self.numColOriginal+j, self.lst_Aptos[self.numAptoCont])
                                self.numAptoCont = self.numAptoCont + 1
                            elif "Precintos" in str(self.df.iat[i, 3+j]):
                                self.cmb_precintos = QComboBox(self.cmb_Apto)
                                self.cmb_precintos.setEditable(True)
                                self.cmb_precintos.lineEdit().setAlignment(Qt.AlignCenter)
                                self.cmb_precintos.lineEdit().setAlignment(Qt.AlignVCenter)
                                self.cmb_precintos.setEditable(False)
                                self.cmb_precintos.setFont(QtGui.QFont('Segoe UI', 10))
                                self.cmb_precintos.setStyleSheet(
                                    "border: 1px solid  rgb(216, 216, 216); background-color: rgb(255, 255, 235);")
                                self.cmb_precintos.clear()
                                self.cmb_precintos.addItem("Precintado")
                                self.cmb_precintos.addItem("Nuevos")
                                self.cmb_precintos.addItem("No precintado")
                                if str("periódica").upper() not in expediente.verificacion.upper(): self.cmb_precintos.setCurrentIndex(1)
                                self.ui.tablePreview.setCellWidget(self.ui.tablePreview.rowCount()-1, self.numColOriginal+j, self.cmb_precintos)
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
                            elif "texto_columna" in str(self.df.iat[i, 3+j]):
                                self.plainText = QPlainTextEdit()
                                self.ui.tablePreview.setCellWidget(self.ui.tablePreview.rowCount()-1, self.numColOriginal+j, self.plainText)
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
                                    self.item = QtWidgets.QTableWidgetItem(self.valor)
                                    self.ui.tablePreview.setItem(self.ui.tablePreview.rowCount() - 1, self.numColOriginal+j, self.item)
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
                                    #elif "texto_columna" in self.df.iat[i, 3+j]:
                                    #    self.valor = ""
                                    else:
                                        for k in range(len(self.listaValores)):
                                            if self.listaValores[k] in self.dict_resultados:
                                                if self.dict_resultados[self.listaValores[k]]: # Si el elemento buscado no está vacío en el diccionario
                                                    self.valor = self.valor + " " + str(self.dict_resultados[self.listaValores[k]])
                                            else:
                                                self.valor = self.valor + " " + str(self.listaValores[k])
                                    self.item = QtWidgets.QTableWidgetItem(self.valor)
                                    self.ui.tablePreview.setItem(self.ui.tablePreview.rowCount() - 1, self.numColOriginal+j, self.item)
                                    #self.ui.tablePreview.setItem(self.ui.tablePreview.rowCount() - 1, self.numColOriginal+j, QtWidgets.QTableWidgetItem(self.valor))
                                    self.ui.tablePreview.item(self.ui.tablePreview.rowCount() - 1, self.numColOriginal+j).setBackground(QtGui.QColor(235, 245, 255))
                                self.ui.tablePreview.item(self.ui.tablePreview.rowCount() - 1, self.numColOriginal+j).setTextAlignment(Qt.AlignCenter)
                                # if "ddr_frec_anterior" in self.df.iat[i, 3+j]: # or "texto_columna" in self.df.iat[i, 3+j] or self.df.iat[i, 3+j] == "":
                                #     self.item.setFlags(self.item.flags() | Qt.ItemIsEditable)
                                # else:
                                #     self.item.setFlags(self.item.flags() & ~(Qt.ItemIsEditable))
                                # self.item.setFlags(self.item.flags() & ~(Qt.ItemIsEditable))
                                    
                            # Ésta además pone el valor original de "Campo" para mantenerlo en el df al exportar a PDF
                            self.ui.tablePreview.setItem(self.ui.tablePreview.rowCount() - 1, 3+j, QtWidgets.QTableWidgetItem(self.df.iat[i, 3+j]))
                    self.ui.tablePreview.setItem(self.ui.tablePreview.rowCount()-1, 0, QtWidgets.QTableWidgetItem(str(self.df.iat[i, 0]))) # Nº Fila Excel
                    self.ui.tablePreview.setItem(self.ui.tablePreview.rowCount()-1, 1, QtWidgets.QTableWidgetItem(str(self.df.iat[i, 1]))) # SI / NO
                    self.ui.tablePreview.setItem(self.ui.tablePreview.rowCount()-1, 2, self.itemTexto) # Título (texto)
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
                    self.widget = self.ui.tablePreview.cellWidget(i, j)
                    if isinstance(self.widget, QComboBox):
                        df.iat[i, j] = self.widget.currentText()
                    elif isinstance(self.widget, QPlainTextEdit):
                        df.iat[i, j] = self.widget.toPlainText()
                    else: df.iat[i,j] = ""

        self.lst_precintos = []
        for i in range(self.ui.lst_precEquipo.count()):
            self.lst_precintos.append(self.ui.lst_precEquipo.item(i).text())

        self.mensaje = funcionesExcel.generarPdf(df, expediente.numExp, self.numColAdicional, self.lst_precintos)
        if self.mensaje == "ok":
            funcionesSQL.guardarLogCertif(expediente.numExp, expediente.numSerie, expediente.tipo)
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

# Tab SDV --------------------------------------------------------------------------------------------------------
    def cambiarPuerto(self):
        funcionesGeneral.modificarConfig('config_sdv', 'puerto', self.ui.cmb_puertos.currentText())
    
    def calcFrecuencia(self):
        if self.ui.txt_velocidad_sdv.text() != "" and self.ui.txt_frec_sdv_config.text() != "" and self.ui.txt_angulo_sdv.text() != "":
             frecuencia = funciones_SdV.calcularFrecuencia(self.ui.txt_velocidad_sdv.text(), self.ui.txt_frec_sdv_config.text(), self.ui.txt_angulo_sdv.text())
             self.ui.txt_frec_sdv.setText(str(round(frecuencia,2)).replace('.', ','))
        else:
            self.ui.txt_frec_sdv.setText('')
    
    def enviarVelocidad(self):
        if self.ui.txt_velocidad_sdv.text() != '' and self.ui.txt_frec_sdv.text() != '':
            self.frecuencia = funciones_SdV.calcularFrecuencia(self.ui.txt_velocidad_sdv.text(), self.ui.txt_frec_sdv_config.text(), self.ui.txt_angulo_sdv.text())
            if self.ui.txt_flujo.text == 'S': self.flujo = True
            else: self.flujo = False
            self.datoVelocidad = funciones_SdV.enviarVelocidad(self.frecuencia, self.ui.txt_amplitud.text() , self.ui.txt_duracion.text(), "Alejamiento",
                                                               self.ui.cmb_puertos.currentText(), int(self.ui.txt_velocidad_rs232.text()), int(self.ui.txt_bits_datos.text()),
                                                               self.ui.txt_paridad.text(), int(self.ui.txt_bits_parada.text()), self.flujo, int(self.ui.txt_timeout.text()))

            if self.datoVelocidad == "error_generador":
                self.mensajesEstado("rojo", "No se puede conectar con el generador de onda")
            else:
                self.ui.txt_datosRecibidos.setText(self.datoVelocidad)
                self.ui.txt_velocidadRecib.setText(self.datoVelocidad[int(self.ui.txt_inicio_vel.text()):int(self.ui.txt_inicio_vel.text())+int(self.ui.txt_longVel.text())])
        else:
            self.mensajesEstado("naranja", "No se ha introducido ninguna velocidad o frecuencia")

    def addConfig(self):
        if self.ui.txt_modelo_sdv.text() != "":
            # Si el modelo ya existe
            if funcionesSQL.buscarValor("configCinem", "Modelo", self.ui.txt_modelo_sdv.text(), "Modelo") != '':
                self.mensajesEstado("naranja", "El modelo introducido ya existe en la configuración")
            else:
                datosCinem = [self.ui.txt_marca_sdv.text(), self.ui.txt_modelo_sdv.text(), self.ui.txt_frec_sdv_config.text(), self.ui.txt_angulo_sdv.text(), self.ui.txt_amplitud.text(),
                        self.ui.txt_duracion.text(), self.ui.txt_tiempo_pulsos.text(), self.ui.sld_sdv.value(), self.ui.txt_timeout.text(), self.ui.txt_velocidad_rs232.text(),
                        self.ui.txt_paridad.text(), self.ui.txt_bits_datos.text(), self.ui.txt_bits_parada.text(), self.ui.txt_flujo.text(), self.ui.txt_inicio_vel.text(), 
                        self.ui.txt_longVel.text(), self.ui.sld_retorno.value()]
                funcionesSQL.insertarConfigCinem(datosCinem)
                self.actualizarListaModelos()

    def sld_sdv(self, valor):
        if valor == 0:
            self.ui.lbl_sdv_man.setStyleSheet("color: rgb(221, 221, 221);"), self.ui.lbl_sdv_auto.setStyleSheet("color: black;")
            for w in self.ui.grp_configCinem.findChildren(QLabel):
                    w.setEnabled(True)
            for w in self.ui.grp_configCinem.findChildren(QLineEdit):
                    w.setEnabled(True)
        else:
            self.ui.lbl_sdv_auto.setStyleSheet("color: rgb(221, 221, 221);"), self.ui.lbl_sdv_man.setStyleSheet("color: black;")
            for w in self.ui.grp_configCinem.findChildren(QLabel):
                    if not 'marca_sdv' in w.objectName() and not 'modelo_sdv' in w.objectName():
                        w.setEnabled(False)
            for w in self.ui.grp_configCinem.findChildren(QLineEdit):
                    if not 'marca_sdv' in w.objectName() and not 'modelo_sdv' in w.objectName():
                        w.setEnabled(False)

    def sld_retorno(self, valor):
        if valor == 0:
            self.ui.lbl_retorno_man.setStyleSheet("color: rgb(221, 221, 221);"), self.ui.lbl_retorno_auto.setStyleSheet("color: black;")
            for w in self.ui.grp_rs232.findChildren(QLabel):
                    w.setEnabled(True)
            for w in self.ui.grp_rs232.findChildren(QLineEdit):
                    w.setEnabled(True)
        else:
            self.ui.lbl_retorno_auto.setStyleSheet("color: rgb(221, 221, 221);"), self.ui.lbl_retorno_man.setStyleSheet("color: black;")
            for w in self.ui.grp_rs232.findChildren(QLabel):
                    w.setEnabled(False)
            for w in self.ui.grp_rs232.findChildren(QLineEdit):
                    w.setEnabled(False)

    def lst_configClickedEvent(self, item):
        self.cargarConfigCinem(item.text())

    def cargarConfigCinem(self, modelo):
        datos = funcionesSQL.leerConfigCinem(modelo)

        for w in self.ui.grp_rs232.findChildren(QLineEdit):
            w.clear()
        for w in self.ui.grp_configCinem.findChildren(QLineEdit):
            w.clear()
        self.ui.txt_marca_sdv.setText(datos[0])
        self.ui.txt_modelo_sdv.setText(datos[1])
        if datos[7] == '1':
            for w in self.ui.grp_configCinem.findChildren(QLineEdit):
                if not 'marca_sdv' in w.objectName() and not 'modelo_sdv' in w.objectName():
                    w.setEnabled(False)
        else:
            for w in self.ui.grp_configCinem.findChildren(QLineEdit):
                    w.setEnabled(True)
            self.ui.txt_frec_sdv_config.setText(datos[2])
            self.ui.txt_angulo_sdv.setText(datos[3])
            self.ui.txt_amplitud.setText(datos[4])
            self.ui.txt_duracion.setText(datos[5])
            self.ui.txt_tiempo_pulsos.setText(datos[6])
        if datos[16] == '1':
            for w in self.ui.grp_rs232.findChildren(QLineEdit):
                w.setEnabled(False)
        else:
            for w in self.ui.grp_rs232.findChildren(QLineEdit):
                w.setEnabled(True)
            self.ui.txt_timeout.setText(datos[8])
            self.ui.txt_velocidad_rs232.setText(datos[9])
            self.ui.txt_paridad.setText(datos[10])
            self.ui.txt_bits_datos.setText(datos[11])
            self.ui.txt_bits_parada.setText(datos[12])
            self.ui.txt_flujo.setText(datos[13])
            self.ui.txt_inicio_vel.setText(datos[14])
            self.ui.txt_longVel.setText(datos[15])
        self.ui.sld_sdv.setValue(int(datos[7]))
        self.ui.sld_retorno.setValue(int(datos[16]))

    def eliminarModelo(self):
        if self.ui.lst_modelos_sdv.currentRow() < 0:
            self.mensajesEstado("naranja", "No se ha seleccionado ningún elemento")
        else:
            funcionesSQL.eliminarConfigCinem(self.ui.lst_modelos_sdv.currentItem().text())
            self.actualizarListaModelos()

    def actualizarListaModelos(self):
        self.modelos = funcionesSQL.leerConfigDisponibles()
        self.ui.lst_modelos_sdv.clear()
        for modelo in self.modelos:
            self.ui.lst_modelos_sdv.addItem(modelo[0])

    def actualizarConfigModelo(self):
        if self.ui.lst_modelos_sdv.currentRow() < 0:
            self.mensajesEstado("naranja", "No se ha seleccionado ningún elemento")
        else:
            datosCinem = [self.ui.txt_marca_sdv.text(), self.ui.txt_modelo_sdv.text(), self.ui.txt_frec_sdv_config.text(), self.ui.txt_angulo_sdv.text(), self.ui.txt_amplitud.text(),
                     self.ui.txt_duracion.text(), self.ui.txt_tiempo_pulsos.text(), self.ui.sld_sdv.value(), self.ui.txt_timeout.text(), self.ui.txt_velocidad_rs232.text(),
                     self.ui.txt_paridad.text(), self.ui.txt_bits_datos.text(), self.ui.txt_bits_parada.text(), self.ui.txt_flujo.text(), self.ui.txt_inicio_vel.text(),
                     self.ui.txt_longVel.text(), self.ui.sld_retorno.value()]
            funcionesSQL.actualizarConfigModelo(datosCinem, self.ui.lst_modelos_sdv.currentItem().text())
            self.actualizarListaModelos()
# ----------------------------------------------------------------------------------------------------------------

class Clase_tablaTrafReal(QMainWindow):  # Crea clase partiendo del tipo QMainWindow

    def __init__(self):  # Inicializa la clase
        super().__init__()  # Con ésto si hubiera clases dentro de ésta función también las inicializaría. Y también necesario al compilar .ui como .py

        # self.tabla = tabla
        self.ui = Ui_WindowTrafReal()
        self.ui.setupUi(self)
        self.permitirCalculoError = False

        self.sliderFuente = self.findChild(QSlider, "sld_fuente")
        self.sliderFuente.valueChanged.connect(self.slideFuente)
        self.valorAnterior = self.sliderFuente.value()
        self.ui.btn_guardarInforme.clicked.connect(self.guardarInforme)
        self.ui.btn_cargarCSV.clicked.connect(self.cargarCSV)

        self.fuente = QtGui.QFont()
        self.fuente.setFamily("Segoe UI")
        self.fuente.setPointSize(9)
        self.ui.tbl_trafReal.setFont(self.fuente)
        self.ui.tbl_trafReal.setItemDelegate(FontDelegate(self.ui.tbl_trafReal))
        # Pasa item delegate para aplicar la fuente al texto también mientras se edita

        header = self.ui.tbl_trafReal.verticalHeader()
        header.setDefaultAlignment(Qt.AlignHCenter | Qt.AlignVCenter)

        self.ui.btn_tr_cinemCab.clicked.connect(self.select_cinemCab)
        self.ui.btn_tr_cinemEst.clicked.connect(self.select_cinemEst)
        self.ui.btn_tr_cinemMov.clicked.connect(self.select_cinemMov)
        self.ui.btn_tr_cabina.clicked.connect(self.select_cabina)

        self.ui.tbl_trafReal.currentItemChanged.connect(self.calcularErrores)

        self.tipoSeleccionado = ''

        for i in range(self.ui.tbl_trafReal.columnCount()):
            self.ui.tbl_trafReal.setColumnWidth(i, 112)
            self.ui.tbl_trafReal.EditingState
            for j in range(self.ui.tbl_trafReal.rowCount()):
                self.item = QtWidgets.QTableWidgetItem("")
                if i == 5 or i == 6:
                    self.item.setFlags(self.item.flags() & ~(Qt.ItemIsEditable))
                    self.item.setBackground(QtGui.QColor(235, 245, 255))
                    #self.item.setBackground(QtGui.QColor(255, 230, 172))
                self.ui.tbl_trafReal.setItem(j, i, self.item)
                self.ui.tbl_trafReal.item(j, i).setTextAlignment(Qt.AlignCenter)

        for i in range(1,5):
            self.ui.tbl_trafReal.setColumnWidth(i, 160)     

        self.permitirCalculoError = True

    def calcularErrores(self):
        if self.permitirCalculoError == True:
            for i in range(self.ui.tbl_trafReal.currentRow(), self.ui.tbl_trafReal.rowCount()):
                if self.ui.tbl_trafReal.item(i, 1).text() != '' and self.ui.tbl_trafReal.item(i, 3).text() != '':
                    self.datoEUT = float(self.ui.tbl_trafReal.item(i, 1).text().replace(",", "."))
                    self.datoREF = float(self.ui.tbl_trafReal.item(i, 3).text().replace(",", "."))
                    if self.datoREF <= 100:
                        self.item = QtWidgets.QTableWidgetItem(str(round(self.datoEUT-self.datoREF, 2)).replace(".", ","))
                        self.columna = 5
                    else:
                        self.item = QtWidgets.QTableWidgetItem(str(round(((self.datoEUT-self.datoREF)/self.datoREF)*100, 2)).replace(".", ","))
                        self.columna = 6
                    self.item.setFlags(self.item.flags() & ~(Qt.ItemIsEditable))
                    self.item.setBackground(QtGui.QColor(255, 230, 172))
                    self.ui.tbl_trafReal.setItem(i, self.columna, self.item)
                    self.ui.tbl_trafReal.item(i, self.columna).setTextAlignment(Qt.AlignCenter)
    def cargarCSV(self):
        self.permitirCalculoError = False
        self.ui.tbl_trafReal.clearContents()
        archivo = QFileDialog.getOpenFileName(self, 'Open file', funcionesGeneral.buscarConfig("rutas", "ruta_base_exp") + str(funcionesSQL.datetime.today().year) + "/", "Archivos CSV (*.csv)")
        if archivo[0]:
            self.df = pd.read_csv(archivo[0], sep=";", skiprows=6, header=None, keep_default_na=False) #, header=None, keep_default_na=False)
            for i in range (50):
                self.ui.tbl_trafReal.setItem(i, 1, QtWidgets.QTableWidgetItem(str(self.df.iat[i, 1]))) # EUT
                self.ui.tbl_trafReal.item(i, 1).setTextAlignment(Qt.AlignCenter)
                self.ui.tbl_trafReal.setItem(i, 3, QtWidgets.QTableWidgetItem(str(self.df.iat[i, 0]))) # REF
                self.ui.tbl_trafReal.item(i, 3).setTextAlignment(Qt.AlignCenter)
                self.ui.tbl_trafReal.setItem(i, 0, QtWidgets.QTableWidgetItem(str(self.df.iat[i, 4]))) # Hora (Id)
                self.ui.tbl_trafReal.item(i, 0).setTextAlignment(Qt.AlignCenter)
        self.ui.tbl_trafReal.setCurrentCell(0,0)
        self.permitirCalculoError = True
        self.calcularErrores()

    def cambiarEstilo(self, botonSel):
        for w in self.ui.wid_modalidad.findChildren(QPushButton):
            w.setStyleSheet("QPushButton {border: 1px solid; border-radius: 10px; border-color: rgb(173, 173, 173); background-color: rgb(225, 225, 225);} QPushButton:hover {background-color: rgb(218, 255, 215);}")
        botonSel.setStyleSheet("border: 1px solid; border-radius: 10px; border-color: rgb(173, 173, 173); background-color: rgb(0, 255, 127);")
        self.ui.btn_guardarInforme.setEnabled(True)

    def select_cinemCab(self):
        self.cambiarEstilo(self.ui.btn_tr_cinemCab)
        self.tipoSeleccionado = 'cinemCab'
    
    def select_cinemEst(self):
        self.cambiarEstilo(self.ui.btn_tr_cinemEst)
        self.tipoSeleccionado = 'cinemEst'

    def select_cinemMov(self):
        self.cambiarEstilo(self.ui.btn_tr_cinemMov)
        self.tipoSeleccionado = 'cinemMov'
    
    def select_cabina(self):
        self.cambiarEstilo(self.ui.btn_tr_cabina)
        self.tipoSeleccionado = 'cabina'

    def slideFuente(self, valor):
        self.fuente.setPointSize(valor)
        self.ui.tbl_trafReal.setFont(self.fuente)
        for i in range(self.ui.tbl_trafReal.columnCount()):
            self.ui.tbl_trafReal.horizontalHeaderItem(i).setFont(self.fuente)
            if valor <= self.valorAnterior:
                self.ui.tbl_trafReal.setColumnWidth(i, self.ui.tbl_trafReal.columnWidth(i) - valor*2)
            else:
                self.ui.tbl_trafReal.setColumnWidth(i, self.ui.tbl_trafReal.columnWidth(i) + valor*2)
        # for i in range(self.ui.tbl_trafReal.rowCount()):
            # self.ui.tbl_trafReal.verticalHeaderItem(i).setFont(self.fuente)
        self.valorAnterior = valor
    
    def guardarInforme(self):
        self.columnas = [i for i in range(self.ui.tbl_trafReal.columnCount())]
        df = pd.DataFrame(columns=[self.ui.tbl_trafReal.horizontalHeaderItem(col).text() for col in self.columnas],
                          index=range(self.ui.tbl_trafReal.rowCount()))

        for i in range(self.ui.tbl_trafReal.rowCount()):
            for j in range(self.ui.tbl_trafReal.columnCount()):
                    if self.ui.tbl_trafReal.item(i, j):
                        if j > 0 and j < 7 and self.ui.tbl_trafReal.item(i, j).text() != '':
                            df.iat[i,j] = float(self.ui.tbl_trafReal.item(i, j).text().replace(",", "."))
                        else:
                            df.iat[i,j] = self.ui.tbl_trafReal.item(i, j).text()
                    else:
                        df.iat[i,j] = None                  
                        
                    #df.iat[i,j] = self.ui.tablePreview.item(i, j).text()
        funcionesExcel.guardarTrafReal(df, expediente.numExp, self.tipoSeleccionado)

class FontDelegate(QStyledItemDelegate):
    def createEditor(self, parent, opt, index):
        editor = super().createEditor(parent, opt, index)
        font = index.data(Qt.FontRole)
        if font is not None:
            editor.setFont(font)
        return editor

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

