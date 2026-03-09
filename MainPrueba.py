import traceback

from Funciones.ConexionSAP import ConexionSAP

from Config.Settings import SAP_CONFIG
import pyperclip
from HU.HU01_EgresosCuentasPorPagar import Facturas


import os
from Funciones.FuncionesExcel import ExcelService as ServicioExcel




if __name__ =="__main__":
        pruebaH01=Facturas()
        
        pruebaH01.descargar_XML()

    
