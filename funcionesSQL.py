import sqlite3
import pandas as pd
import numpy as np
import funcionesGeneral
import os
from datetime import datetime

def actualizarConfigModelo(datos, modelo):
    conn = sqlite3.connect(funcionesGeneral.buscarConfig("rutas", "ruta_db_matrices"))
    cur = conn.cursor()

    if datos[7] == 1: datos[2:7] = ["", "", "", "", ""]
    if datos[16] == 1: datos[8:16] = ["", "", "", "", "", "", "", ""]

    cur.execute("UPDATE configCinem SET Marca = '" + datos[0] + "', Modelo = '" + datos[1] + "', Frecuencia = '" + datos[2] + "', Angulo = '" + datos[3]
                + "', Amplitud = '" + datos[4] + "', Duracion = '" + datos[5] + "', T_Pulsos = '" + datos[6] + "', Simulacion = '" + str(datos[7]) + "', Timeout = '" + datos[8]
                + "', Velocidad = '" + datos[9] + "', Paridad = '" + datos[10] + "', Bits_datos = '" + datos[11] + "', Bits_parada = '" + datos[12] + "', Control_flujo = '" + datos[13]
                + "', Caracter_inicio = '" + datos[14] + "', Longitud_vel = '" + datos[15] + "', Retorno = '" + str(datos[16]) + "' WHERE Modelo = '" + modelo + "'")
    conn.commit()
    conn.close()

def eliminarConfigCinem(modelo):
    conn = sqlite3.connect(funcionesGeneral.buscarConfig("rutas", "ruta_db_matrices"))
    cur = conn.cursor()
    cur.execute("DELETE FROM configCinem WHERE Modelo = '" + modelo + "'")
    conn.commit()
    conn.close()

def insertarConfigCinem(datos):
    conn = sqlite3.connect(funcionesGeneral.buscarConfig("rutas", "ruta_db_matrices"))
    cur = conn.cursor()

    if datos[7] == 1: datos[2:7] = ["", "", "", "", ""]
    if datos[16] == 1: datos[8:16] = ["", "", "", "", "", "", "", ""]

    cur.execute('INSERT INTO configCinem VALUES ("' + datos[0] + '", "' + datos[1] + '", "' + datos[2] + '", "' + datos[3] + '", "' + datos[4]
            + '", "' + datos[5] + '", "' + datos[6] + '", "' + str(datos[7]) + '", "' + datos[8] + '", "' + datos[9] + '", "' + datos[10]
            + '", "' + datos[11] + '", "' + datos[12] + '", "' + datos[13] + '", "' + datos[14] + '", "' + datos[15] + '", "' + str(datos[16])+ '")')

    conn.commit()
    conn.close()

def leerConfigDisponibles():
    conn = sqlite3.connect(funcionesGeneral.buscarConfig("rutas", "ruta_db_matrices"))
    cur = conn.cursor()
    cur.execute('SELECT Modelo FROM configCinem')
    resultado = cur.fetchall()
    conn.close()
    return resultado

def leerConfigCinem(modelo):
    try:
        conn = sqlite3.connect(funcionesGeneral.buscarConfig("rutas", "ruta_db_matrices"))
        cur = conn.cursor()
        cur.execute('SELECT * FROM configCinem WHERE Modelo = "' + modelo + '"')
        resultado = cur.fetchone()
        conn.close()
        return resultado
    except:
        return ""

# Éste no hará falta
def crearSQLExcel(pandasExcel):
    cabecera = ""
    pd_Excel = pandasExcel
    pd_Excel.replace(np.nan, 0)
    conn = sqlite3.connect(funcionesGeneral.buscarConfig("rutas", "ruta_db_matrices"))
    cur = conn.cursor()

    # Cabeceras a partir de Excel
    for nombre in pd_Excel.columns:
        cabecera = cabecera + "'" + nombre + "'" + ' TEXT, '
    cabecera = cabecera[:-2]
    cabecera = cabecera.replace(';', '-')
    print(cabecera)

    #Crear tabla
    cur.execute("DROP TABLE filas_resultados")
    cur.execute("CREATE TABLE IF NOT EXISTS filas_resultados (" + cabecera + ")")
    conn.commit()

    # Datos Excel a tabla
    #pd_Excel = pd_Excel[:2]
    print(pd_Excel)

    pd_Excel.to_sql('filas_resultados', conn, if_exists='replace',index=False)
    conn.commit()

    # Si no existe la tabla, siempre debería crearla
    # cur.execute("CREATE TABLE IF NOT EXISTS libro (id INTEGER PRIMARY KEY, titulo TEXT,"" autor TEXT, ano INTEGER, isbn INTEGER)")

    conn.close()

def guardarFilasResultados(df, tabla):
    cabecera = ""
    for nombre in df.columns:
        cabecera = cabecera + "'" + nombre + "'" + ' TEXT, '
    cabecera = cabecera[:-2]

    conn = sqlite3.connect(funcionesGeneral.buscarConfig("rutas", "ruta_db_matrices"))
    cur = conn.cursor()
    cur.execute("DROP TABLE " + tabla)
    cur.execute("CREATE TABLE IF NOT EXISTS " + tabla + " (" + cabecera + ")")
    df.to_sql(tabla, conn, if_exists='replace', index=False)
    conn.commit()

    conn.close()

def datosFilasResultados(tabla):
    conn = sqlite3.connect(funcionesGeneral.buscarConfig("rutas", "ruta_db_matrices"))
    df = pd.read_sql_query("SELECT * from " + tabla, conn)
    conn.close()
    return df

def buscarValor(tabla, nombreCol, texto, colRet):
    try:
        conn = sqlite3.connect(funcionesGeneral.buscarConfig("rutas", "ruta_db_matrices"))
        cur = conn.cursor()
        cur.execute('SELECT ' + colRet + ' FROM ' + tabla + ' WHERE ' + nombreCol + ' = "' + texto + '"')
        resultado = cur.fetchone()[0]
        conn.close()
        return resultado
    except:
        return ""

def buscarColumna(tabla, columna):
    try:
        conn = sqlite3.connect(funcionesGeneral.buscarConfig("rutas", "ruta_db_matrices"))
        cur = conn.cursor()
        encontrado = False
        respuesta = ""
        cur.execute("SELECT " + columna + " FROM " + tabla)
        respuesta = [""]
        for fila in cur.fetchall():
            respuesta.append(fila[0])
        conn.close()
        return respuesta
    except:
        return ""

def caractEquipo(texto, tabla, columna):
    conn = sqlite3.connect(funcionesGeneral.buscarConfig("rutas", "ruta_db_matrices"))
    cur = conn.cursor()
    encontrado = False
    respuesta = ""
    cur.execute("SELECT * FROM " + tabla)
    for fila in cur.fetchall():
        if fila[0].upper() in texto.upper():
            respuesta = fila[columna]
            encontrado = True
    conn.close()
    if encontrado == False:
            respuesta = "error_" + tabla
    return respuesta

def buscarFilasResultadosPreview(objExpediente, instalacion):
    conn = sqlite3.connect(funcionesGeneral.buscarConfig("rutas", "ruta_db_matrices"))
    consulta = 'SELECT Fila, "' + objExpediente.caracteristicas + '", Titulo, Campo1, Campo2, Campo3 from filas_resultados ORDER BY Fila ASC'
    # consulta = 'SELECT Fila, "' + seleccionado + '", Titulo, Campo from filas_resultados WHERE NOT "' + seleccionado + '" = "N"'
    df = pd.read_sql_query(consulta, conn)
    conn.close()
    valor = "S"
    lista = []
    for i in range(len(df)):
        if df.iat[i, 1] != "S" and df.iat[i, 1] != "N":
            lista = str(df.iat[i, 1]).split(",")
            for j in range(len(lista)):
                if "Planos" in lista[j]:
                    if buscarValor("exTipo", "Modelo", objExpediente.modelo, "dosPlanos") == "N": valor = "N"
                if "Norma" in lista[j]:
                    if objExpediente.norma == "OM 1994":
                        valor = "N"
                if "Instalacion" in lista[j]:
                    if instalacion == 0:
                        valor = "N"
            df.iat[i, 1] = valor
    del objExpediente
    return df

def buscarEMP(txt_emp):
    # Devuelve un diccionario con los valores para la combinación seleccionada
    conn = sqlite3.connect(funcionesGeneral.buscarConfig("rutas", "ruta_db_matrices"))
    consulta = 'SELECT Campo, "' + txt_emp + '" from emp'
    df = pd.read_sql_query(consulta, conn)
    conn.close()
    dict_emp = df.set_index('Campo').transpose().to_dict('index')
    dict_emp = dict_emp[txt_emp]
    return dict_emp

def guardarLogCertif(exp, numSerie, tipo):
    try:
        conn = sqlite3.connect(funcionesGeneral.buscarConfig("rutas", "ruta_db_matrices"))
        cur = conn.cursor()
        cur.execute("INSERT INTO logCertif VALUES ('" + exp + "', '" + numSerie + "', '" + tipo + "', '" + os.getlogin() + "', '" + datetime.now().strftime("%d/%m/%Y %H:%M:%S") + "')")

        conn.commit()
        conn.close()
    except:
        pass