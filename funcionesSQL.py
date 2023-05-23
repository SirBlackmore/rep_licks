import sqlite3
import pandas as pd
import numpy as np
import main
import os

# Éste no hará falta
def crearSQLExcel(pandasExcel):
    cabecera = ""
    pd_Excel = pandasExcel
    pd_Excel.replace(np.nan, 0)
    conn = sqlite3.connect("database/matrices.db")
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

    conn = sqlite3.connect("database/matrices.db")
    cur = conn.cursor()
    cur.execute("DROP TABLE " + tabla)
    cur.execute("CREATE TABLE IF NOT EXISTS " + tabla + " (" + cabecera + ")")
    df.to_sql(tabla, conn, if_exists='replace', index=False)
    conn.commit()

    conn.close()

def datosFilasResultados(tabla):
    conn = sqlite3.connect("database/matrices.db")
    df = pd.read_sql_query("SELECT * from " + tabla, conn)
    conn.close()
    return df

def buscarValor(tabla, nombreCol, texto, colRet):
    try:
        conn = sqlite3.connect("database/matrices.db")
        cur = conn.cursor()
        cur.execute('SELECT ' + colRet + ' FROM ' + tabla + ' WHERE ' + nombreCol + ' = "' + texto + '"')
        resultado = cur.fetchone()[0]
        conn.close()
        return resultado
    except:
        return ""
def caractEquipo(texto, tabla, columna):
    conn = sqlite3.connect("database/matrices.db")
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
    conn = sqlite3.connect("database/matrices.db")
    consulta = 'SELECT Fila, "' + objExpediente.caracteristicas + '", Titulo, Campo1, Campo2, Campo3 from filas_resultados'
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
    conn = sqlite3.connect("database/matrices.db")
    consulta = 'SELECT Campo, "' + txt_emp + '" from emp'
    df = pd.read_sql_query(consulta, conn)
    conn.close()
    dict_emp = df.set_index('Campo').transpose().to_dict('index')
    dict_emp = dict_emp[txt_emp]
    return dict_emp






