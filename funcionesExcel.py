#import win32com.client as win32 # Solo para convertir xls -> xlsx
import xlrd
import pandas as pd
from openpyxl import Workbook, load_workbook
import pandas
import win32com.client

import funcionesGeneral
import funcionesSQL
import os
from funcionesGeneral import buscarConfig

def cargarHojasExcel(expediente):
    year = "20" + str(expediente[3:5])
    ruta = buscarConfig("rutas", "ruta_base_exp") + year + '/' + expediente + '/' + expediente + '.xlsx'
    if os.path.isfile(ruta) == False:
        return "nodisp"
    else:
        try:
            libro = load_workbook(filename=ruta)
            return libro.sheetnames
        except:
            return "error"

def guardarExcel(expediente):

    libro = Workbook()
    sheet = libro.active

    sheet["B1"] = "Número de expediente"
    sheet["B2"] = expediente

    libro.save(filename=expediente + ".xlsx")

    return True

def cargarDatosEnsayos(expediente):
    year = "20" + str(expediente[3:5])
    ruta = buscarConfig("rutas", "ruta_base_exp") + year + '/' + expediente + '/' + expediente + '.xlsx'
    if os.path.isfile(ruta) == False:
        return "nodisp"
    else:
        try:
            libro = load_workbook(filename=ruta, data_only=True)
            datos = {"sdv_fecha": [], "sdv_bajo_valor": [], "sdv_alto_valor": [], "sdv_alarma_tension_valor": [],
                     "ddr_h_fecha": [], "ddr_h_ancho_lob_valor": [], "ddr_h_aten_sec_valor": [], "ddr_h_desv_eje_valor": [], "ddr_h_frec_media_valor": [],
                     "ddr_v_fecha": [], "ddr_v_ancho_lob_valor": [], "ddr_v_aten_sec_valor": [], "ddr_v_desv_eje_valor": [], "ddr_v_frec_media_valor": [],
                     "ddr_frec_anterior": "",
                     "tr_e_fecha": [], "tr_e_bajo_valor": [], "tr_e_alto_valor": [],
                     "tr_m_fecha": [], "tr_m_bajo_valor": [], "tr_m_alto_valor": []}

            for hoja in libro.sheetnames:
                if str("alej").upper() in hoja.upper() or str("aprox").upper() in hoja.upper():
                    datos["sdv_bajo_valor"].append(round(libro[hoja][funcionesGeneral.buscarConfig("celdas_sdv", "sdv_bajo")].value, 2))
                    datos["sdv_alto_valor"].append(round(libro[hoja][funcionesGeneral.buscarConfig("celdas_sdv", "sdv_alto")].value, 2))
                    datos["sdv_fecha"].append(libro[hoja][funcionesGeneral.buscarConfig("celdas_sdv", "sdv_fecha")].value)
                    if libro[hoja][funcionesGeneral.buscarConfig("celdas_sdv", "sdv_alarma_tension")].value is not None:
                        datos["sdv_alarma_tension_valor"].append(libro[hoja][funcionesGeneral.buscarConfig("celdas_sdv", "sdv_alarma_tension")].value)
                elif str("traf").upper() in hoja.upper() and str("_e").upper() in hoja.upper():
                    datos["tr_e_bajo_valor"].append(round(libro[hoja][funcionesGeneral.buscarConfig("celdas_trafreal", "tr_bajo")].value, 2))
                    datos["tr_e_alto_valor"].append(round(libro[hoja][funcionesGeneral.buscarConfig("celdas_trafreal", "tr_alto")].value, 2))
                    datos["tr_e_fecha"].append(libro[hoja][funcionesGeneral.buscarConfig("celdas_trafreal", "tr_fecha")].value)
                elif str("traf").upper() in hoja.upper() and str("_m").upper() in hoja.upper():
                    datos["tr_m_bajo_valor"].append(round(libro[hoja][funcionesGeneral.buscarConfig("celdas_trafreal", "tr_bajo")].value, 2))
                    datos["tr_m_alto_valor"].append(round(libro[hoja][funcionesGeneral.buscarConfig("celdas_trafreal", "tr_alto")].value, 2))
                    datos["tr_m_fecha"].append(libro[hoja][funcionesGeneral.buscarConfig("celdas_trafreal", "tr_fecha")].value)
                elif str("ddr_h").upper() in hoja.upper() and len(hoja) <= 5:
                    datos["ddr_h_fecha"].append(libro[hoja][funcionesGeneral.buscarConfig("celdas_ddr", "ddr_fecha")].value)
                    datos["ddr_h_ancho_lob_valor"].append(round(libro[hoja][funcionesGeneral.buscarConfig("celdas_ddr", "ddr_ancho_lob")].value, 2))
                    datos["ddr_h_aten_sec_valor"].append(round(libro[hoja][funcionesGeneral.buscarConfig("celdas_ddr", "ddr_aten_sec")].value, 2))
                    datos["ddr_h_desv_eje_valor"].append(round(libro[hoja][funcionesGeneral.buscarConfig("celdas_ddr", "ddr_desv_eje")].value, 2))
                    datos["ddr_h_frec_media_valor"].append(round(libro[hoja][funcionesGeneral.buscarConfig("celdas_ddr", "ddr_frec_media")].value, 3))
                    if libro[hoja][funcionesGeneral.buscarConfig("celdas_ddr", "ddr_frec_anterior")].value is not None:
                        datos["ddr_frec_anterior"] = str(round(libro[hoja][funcionesGeneral.buscarConfig("celdas_ddr", "ddr_frec_anterior")].value, 3)).replace(".",",")
                elif str("ddr_v").upper() in hoja.upper() and len(hoja) <= 5:
                    datos["ddr_v_fecha"].append(libro[hoja][funcionesGeneral.buscarConfig("celdas_ddr", "ddr_fecha")].value)
                    datos["ddr_v_ancho_lob_valor"].append(round(libro[hoja][funcionesGeneral.buscarConfig("celdas_ddr", "ddr_ancho_lob")].value, 2))
                    datos["ddr_v_aten_sec_valor"].append(round(libro[hoja][funcionesGeneral.buscarConfig("celdas_ddr", "ddr_aten_sec")].value, 2))
                    datos["ddr_v_desv_eje_valor"].append(round(libro[hoja][funcionesGeneral.buscarConfig("celdas_ddr", "ddr_desv_eje")].value, 2))
                    datos["ddr_v_frec_media_valor"].append(round(libro[hoja][funcionesGeneral.buscarConfig("celdas_ddr", "ddr_frec_media")].value, 3))

            return datos
        except:
            print(hoja)
            return "error"

def generarPdf(df, numExp, numColAdicional, precintos):
    ruta = 'plantillas/plantilla_resultados.xlsx'

    if os.path.isfile(ruta) == False:
        return "nodisp"
    else:
        try:
            libro = load_workbook(filename=ruta)
            hoja = libro.active

            for i in range(len(df)):
                for j in range(numColAdicional):
                    if df.iat[i,3+j]: # Da igual buscar así (en columnas "Campo"), que sumar numColAdicional y buscar en columnas "Valor"
                        for columna in range(1,19):
                            if str(hoja.cell(int(df.iat[i, 0]), columna).value) == str(j+1): # df.iat[i, 0] es donde está el número de fila en el df
                                hoja.cell(int(df.iat[i, 0]), columna).value = str(df.iat[i, 3+j+numColAdicional])

            # Se recorren todas las filas de la plantilla y se comparan con el df para ver si el número de fila está contenido en él (si no, se oculta)
            for i in range(int(funcionesGeneral.buscarConfig("celdas_certificado", "inicio_certif")), int(funcionesGeneral.buscarConfig("celdas_certificado", "final_certif"))+1):
                if str(i) not in df["Fila"].values:
                    hoja.row_dimensions[i].hidden = True
            
            if precintos:
                hoja[funcionesGeneral.buscarConfig("celdas_certificado", "celda_precintos")].value =  'Precintos:\n' + '\n'.join(str(p) for p in precintos)

            year = "20" + str(numExp[3:5])
            ruta = buscarConfig("rutas", "ruta_base_exp") + year + '/' + numExp + '/'

            libro.save(ruta + 'resultados_temp.xlsx')

            excel = win32com.client.Dispatch('Excel.Application')
            libro_excel = excel.Workbooks.Open(os.path.abspath(ruta + 'resultados_temp.xlsx'))

            libro_excel.ExportAsFixedFormat(0, os.path.abspath(ruta + numExp + '.pdf'))
            libro_excel.Close(SaveChanges=False)
            os.remove(os.path.abspath(ruta + 'resultados_temp.xlsx'))
            # excel.Quit()

            return "ok"
        except:
            return "error"

def convertirEnsayos(expediente):
    year = "20" + str(expediente[3:5])
    ruta = buscarConfig("rutas", "ruta_base_exp") + year + '/' + expediente

    excel = win32com.client.Dispatch('Excel.Application')
    libroNuevo = excel.Workbooks.Add()
    libroNuevo.SaveAs(expediente + ".xlsx")

    for archivo in os.listdir(os.path.abspath(ruta)):
        if os.path.splitext(archivo)[1] == ".XLS" or os.path.splitext(archivo)[1] == ".xls":
            if os.path.splitext(archivo)[0] == expediente + "a":
                rutaCompleta = ruta + '/' + archivo
                libro = excel.Workbooks.Open(rutaCompleta, False, False, None, password:="cinem123")
                libro.Sheets(1).Copy(Before:=libroNuevo.Sheets(1))
                libroNuevo.Sheets(1).Name="Aproximacion" 
                libro.Close(savechanges:=False)
            if os.path.splitext(archivo)[0] == expediente + "d":
                rutaCompleta = ruta + '/' + archivo
                libro = excel.Workbooks.Open(rutaCompleta, False, False, None, password:="cinem123")
                libro.Sheets(1).Copy(Before:=libroNuevo.Sheets(1))
                libroNuevo.Sheets(1).Name="Alejamiento" 
                libro.Close(savechanges:=False)
            if os.path.splitext(archivo)[0] == expediente + "h":
                rutaCompleta = ruta + '/' + archivo
                libro = excel.Workbooks.Open(rutaCompleta, False, False, None, password:="cinem123", WriteResPassword:="cinem123")
                libro.Sheets(2).Copy(Before:=libroNuevo.Sheets(1))
                libroNuevo.Sheets(1).Name="DdR_H"
                libro.Sheets(1).Copy(Before:=libroNuevo.Sheets(1))
                libroNuevo.Sheets(1).Name="Datos_DdR_H"  
                libro.Close(savechanges:=False)
            if os.path.splitext(archivo)[0] == expediente + "v":
                rutaCompleta = ruta + '/' + archivo
                libro = excel.Workbooks.Open(rutaCompleta, False, False, None, password:="cinem123", WriteResPassword:="cinem123")
                libro.Sheets(2).Copy(Before:=libroNuevo.Sheets(1))
                libroNuevo.Sheets(1).Name="DdR_V"
                libro.Sheets(1).Copy(Before:=libroNuevo.Sheets(1))
                libroNuevo.Sheets(1).Name="Datos_DdR_V"  
                libro.Close(savechanges:=False)
            
    libroNuevo.Application.DisplayAlerts = False
    libroNuevo.Sheets("Hoja1").Delete
    libroNuevo.Close(savechanges:=True)


#Éste no hará falta
def tablaDesdeExcel():
    pd_Excel = pandas.read_excel('plantillas/matrices.xlsx')
    funcionesSQL.crearSQLExcel(pd_Excel)


