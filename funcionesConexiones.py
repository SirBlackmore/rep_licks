import requests
from datetime import datetime
import json
# response = requests.get('https://website.example/id', headers={'Authorization': 'access_token myToken'})

urlExp = "https://wswecogeslab.cem.es/api/Expedientes/GetExpedientesByLaboratorio?nombre=Cinem%C3%B3metros"
urlEquipos = "https://wswecogeslab.cem.es/api/Instrumentos/GetIdLabInstrumento?idLab=7"
urlES = "https://wswecogeslab.cem.es/api/EntradasSalidas/GetEntradaSalidasFiltrados?FechaDesde=2020-01-01&" + \
        "FechaHasta=" + str(datetime.now().year) + "-" + str(datetime.now().month) + "-" + str(datetime.now().day) + "&IdLaboratorio=7"
token = "pgH7QzFHJx4w46fI~5Uzi4RvtTwlEXp"

def conectarGESLAB(objExpediente):
    conexion = "OK"
    expEncontrado = False
    eqEncontrado = False
    try:
        #Conexión a tabla expedientes
        respuesta = requests.get(urlExp,headers={'XApiKey': token})

        if respuesta.status_code != 200:
            conexion = "error"
            # TERMINAR FUNCIÓN SI HAY ERROR!!!!! Y PONER MENSAJE ERROR

        else:
            for registro in respuesta.json():
                if registro["numeroExpediente"] == objExpediente.numExp:
                    expEncontrado = True
                    #print("Expediente: ")
                    #print(registro)
                    objExpediente.verificacion = registro["servicioActividad"]
                    idCem = registro["lstEquipos"][0]["idCem"]

            if expEncontrado == False:
                conexion = "expFalse"
                return conexion

    except:
        conexion = "error"
        return conexion

    try:
        # Conexión a tabla equipos
        respuestaEquipos = requests.get(urlEquipos,headers={'XApiKey': token})

        if respuestaEquipos.status_code != 200:
            conexion = "error"
        else:
            for registro in respuestaEquipos.json():
                if registro["idCem"] == idCem:
                    eqEncontrado = True
                    #print("Equipo: ")
                    #print(registro)
                    objExpediente.tipo = registro["nombre"]
                    objExpediente.marca = registro["lstMarca"][0]["nombreMarca"]
                    objExpediente.modelo = registro["lstModelo"][0]["nombreModelo"]
                    objExpediente.numSerie = registro["nroSerie"]
                    objExpediente.modalidad = registro["lstTiposInstrumento"][0]["nombreTipoInstrumento"]
                    objExpediente.descripcion = registro["especificaciones"]
                    objExpediente.norma = registro["lstNorma"][0]["nombreNorma"]

            if eqEncontrado == False:
                conexion = "eqFalse"
                return conexion

    except:
        conexion = "error"
        return conexion

    try:
        # Conexión a tabla E/S
        respuestaEquipos = requests.get(urlES,headers={'XApiKey': token})

        if respuestaEquipos.status_code != 200:
            conexion = "error"
        else:
            for registro in respuestaEquipos.json():
                if registro["expediente"] == objExpediente.numExp:
                    # print(registro)
                    eqEncontrado = True
                    print("Equipo!: ")
                    print(registro)
                    # objExpediente.tipo = registro["nombre"]
                    # objExpediente.marca = registro["lstMarca"][0]["nombreMarca"]
                    # objExpediente.modelo = registro["lstModelo"][0]["nombreModelo"]
                    # objExpediente.numSerie = registro["nroSerie"]
                    # objExpediente.modalidad = registro["lstTiposInstrumento"][0]["nombreTipoInstrumento"]
                    # objExpediente.descripcion = registro["especificaciones"]
                    # objExpediente.norma = registro["lstNorma"][0]["nombreNorma"]

            if eqEncontrado == False:
                conexion = "esFalse"
                return conexion

    except:
        conexion = "error"

    return conexion