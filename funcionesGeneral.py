from configparser import ConfigParser
import os
import subprocess

def buscarConfig(categoria, nombre):

    try:
        config = ConfigParser()
        config.read("config.ini", encoding='utf-8')

        return config[categoria][nombre]
    except:
        return "error"

def mostrarConfigCategoria(categorias):
    try:
        config = ConfigParser()
        config.read("config.ini", encoding='utf-8')

        claves = []
        valores = []

        for categoria in categorias:
            for linea in config.items(categoria):
                claves.append(linea[0])
                valores.append(linea[1])

        lista = [claves,valores]

        return lista
    except:
        return "error"

def abrirCarpeta(expediente):
    year = "20" + str(expediente[3:5])
    ruta = str(buscarConfig("rutas", "ruta_base_exp") + year + '/' + expediente + '/')
    if os.path.exists(ruta) == False:
        return "nodisp"
    else:
        subprocess.Popen('explorer "' + os.path.abspath(ruta) + '"')

def modificarConfig(categoria, nombre, valor):

    config = ConfigParser()
    config.read("config.ini", encoding='utf-8')
    config[categoria][nombre] = valor

    with open("config.ini", 'w') as configfile:
        config.write(configfile)
