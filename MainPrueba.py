import traceback

from Funciones.ConexionSAP import ConexionSAP

from Config.Settings import SAP_CONFIG
from Config.init_config import in_config
from HU.HU01_LoginSAP import ObtenerSesionActiva
import pyperclip

from HU.HU01_EgresosCuentasPorPagar import Facturas

from Funciones.GuiShellFunciones import AbrirTransaccion

import os
from Funciones.FuncionesExcel import ExcelService as ServicioExcel
from Config.init_config import in_config as inConfig




if __name__ =="__main__":
        pruebaH01=Facturas()
        
        pruebaH01.descargar_XML()

    
