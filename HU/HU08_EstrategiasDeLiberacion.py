import re
import logging
import pandas as pd
from datetime import datetime
import time
import threading
from config.settings import SAP_CONFIG
from config.init_config import in_config
from funciones.ConexionSAP import ConexionSAP
from funciones.consultarOC import consultarOC
from funciones.CargarAnexo import cargar_archivo_gos # Asegúrate de que este archivo exista
from repositorios.Excel import Excel as ExcelDB
from funciones.GuiShellFunciones import AbrirTransaccion,ObtenerSesionActiva

class  HU08_EstrategiasDeLiberacion:
    def __init__(self):
        """
        Inicializa los componentes de conexión y logging.
        """
        self.logger = logging.getLogger("HU07_ClasificarOC")
        self.sap = ConexionSAP(
            SAP_CONFIG.get('SAP_USUARIO'),
            SAP_CONFIG.get('SAP_PASSWORD'),
            in_config('SAP_CLIENTE'),
            in_config('SAP_IDIOMA'),
            in_config('SAP_PATH'),
            in_config('SAP_SISTEMA')
        )
    


    def ejecutar(self):
        """
        Docstring for ejecutar
        
        :param self: Description
        """
        session = ObtenerSesionActiva()
        AbrirTransaccion(session, "ZMM_68")

        session.findById("wnd[0]/usr/ctxtR_BEDAT-LOW").text = "01.01.2026"
        session.findById("wnd[0]/usr/ctxtR_BEDAT-HIGH").text = "31.01.2026"

 