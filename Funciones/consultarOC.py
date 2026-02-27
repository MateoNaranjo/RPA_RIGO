import win32com.client  # Para SAP
import pandas as pd     # Para la lógica de negocio y Excel
import os    
import time           # Para rutas del File Server
from datetime import datetime, timedelta
from Funciones.ConexionSAP import ConexionSAP
from Config.settings import SAP_CONFIG
from Config.init_config import in_config

def consultarOC(sesion, numeroOC):
    """
    Consulta una OC en SAP, valida su existencia y extrae Estado y Monto.
    """
    try:
        # 1. Navegar a la transacción de visualización
        sesion.findById("wnd[0]/tbar[0]/okcd").text = "/nME23N"
        sesion.findById("wnd[0]").sendVKey(0)
        
        # 2. Cargar el número de Orden
        sesion.findById("wnd[0]/tbar[1]/btn[17]").press()
        sesion.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN").text = str(numeroOC)
        sesion.findById("wnd[1]").sendVKey(0)
        
        # 3. VALIDACIÓN DE EXISTENCIA
        barra_estado = sesion.findById("wnd[0]/sbar")
        if barra_estado.messagetype == "E": # Si hay error en la barra de SAP
            return {"status": "Error", "detalle": barra_estado.text, "monto": 0}

        # 4. EXTRACCIÓN DE DATOS (Solo si la OC existe)
        try:
            # El campo de estatus suele estar en la cabecera
            # Nota: Los IDs pueden variar levemente según la versión de SAP
            estado_texto = sesion.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT10/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1232/txtMEPO1232-STATUS01").text
            
            # El campo de valor neto (Monto)
            monto_raw = sesion.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT10/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1232/ssubHEADER_CUM_2:SAPLMEGUI:1235/txtMEPO1235-VALUE01").text
            
            # Limpieza del monto (Ej: "1.500.250,00" -> 1500250.0)
            monto_limpio = float(monto_raw.replace(".", "").replace(",", ".").strip())
            
        except Exception as e:
            estado_texto = "No detectado"
            monto_limpio = 0.0
            print(f"Advertencia: No se pudieron leer detalles de la OC {numeroOC}: {e}")

        return {
            "status": "OK",
            "detalle": estado_texto,
            "monto": monto_limpio
        }

    except Exception as e:
        return {"status": "Error", "detalle": str(e), "monto": 0}