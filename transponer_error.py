print("              POR FAVOR, ESPERE...                     ")
import os
import time

def testSystem():
    if os.name != 'nt':
        print("                                                       ")
        print("                                                       ")
        print("                     ERROR:                            ")
        print("ESTA HERRAMIENTA SOLO FUNCIONA PARA SISTEMAS OPERATIVOS")
        print("                   WINDOWS NT                          ")
        print("                                                       ")
        print("                                                       ")
        print("       TERMINANDO EJECUCIÓN DEL PROGRAMA...             ")
        time.sleep(5)
        return False
    else:
        return True
