import pyvisa
import math
import serial.tools.list_ports
import serial
from time import sleep

V_LUZ = 299792458

def enviarVelocidad(frecuencia, amplitud, duracion, sentido, puerto, baudrate, bytesize, parity, stopbits, xonxoff, timeout):   

# HACER FUNCIÓN "Enviar velocidad", y separadas en "Generar velocidad (si hay generador) y recibir rs232 (si hay retorno)"

    try:
        rm = pyvisa.ResourceManager()
        #print(rm.list_resources())

        generador = rm.open_resource('GPIB0::10::INSTR')

        generador.clear()

        # print(generador.query('*IDN?'))
        # print(generador.read())

        # Reseteo valores Keysight
        generador.write("*RST")
        generador.write("*WAI") # Esperar a que se completen las operaciones pendientes
        # generador.write("VOLT:LIM:HIGH 4") # Limites voltaje
        # generador.write("VOLT:LIM:LOW -4")
        # generador.write("VOLT:LIM:STAT ON")

        ciclos = str(frecuencia * int(duracion)).replace(",", ".")

        if sentido == "Alejamiento":
            fase = -90
        else:
            fase = 90

        # comando = "FUNC SIN;FREQ " + str(frecuencia) + ";VOLT " + amplitud + ";VOLT:OFFS 0"
        # print(comando)
        # generador.write(comando)
        # generador.write("SOUR1:PHAS 30")
        

        # Canal 1 (frecuencia y amplitud)
        comando = "FUNC SIN;FREQ " + str(frecuencia) + ";VOLT " + amplitud + ";VOLT:OFFS 0"
        generador.write(comando)
        # Canal 2 (frecuencia y amplitud)
        comando = "SOUR2:FUNC SIN;FREQ " + str(frecuencia) + ";VOLT " + amplitud + ";VOLT:OFFS 0"
        generador.write(comando)

        # Ciclos canal 1
        comando = "BURS:MODE TRIG;NCYC " + ciclos + ";PHAS 0"
        generador.write(comando)
        generador.write("TRIG:SOUR BUS")
        generador.write("BURS:STAT ON")
        generador.write("OUTP ON")
        # Ciclos canal 2
        comando = "SOUR2:BURS:MODE TRIG;NCYC " + ciclos + ";PHAS " + str(fase)
        generador.write(comando)
        generador.write("TRIG2:SOUR BUS")
        generador.write("SOUR2:BURS:STAT ON")
        generador.write("OUTP2 ON")

        generador.write("*TRG")
        generador.write("*WAI")
        #generador.write("*RST")
        
        sleep(2) # Parece que hay que esperar a que haga la foto

        ser = serial.Serial(port=puerto, baudrate=baudrate, bytesize=bytesize, parity=parity, stopbits=stopbits, timeout=timeout, xonxoff=xonxoff, rtscts=False)

        generador.clear()
        try:
            datos = ser.readline()
        except serial.SerialTimeoutException: # REVISAR, NO FUNCIONA CON TIMEOUT
            print('Data could not be read')

        # s = ser.read(10)

        generador.close()

        return datos.decode()

    except:
        return "error_generador"
    
def calcularFrecuencia(velocidad, frecCinem, angulo):
    frecuencia = float(velocidad) * 2 / 3.6 * float(str(frecCinem).replace(',', '.')) * 1000000000 / int(V_LUZ) * math.cos(int(angulo) * math.pi / 180)
    return frecuencia

def listaPuertos():
    listaPuertos = []
    for port, desc, hwid in sorted(serial.tools.list_ports.comports()):
        listaPuertos.append(port)
    return listaPuertos
        
# def leerSerie():
    

