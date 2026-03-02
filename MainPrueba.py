import traceback

from Funciones.ConexionSAP import ConexionSAP

from Config.Settings import SAP_CONFIG
from Config.init_config import in_config
from HU.HU01_LoginSAP import ObtenerSesionActiva
import pyperclip

from Funciones.GuiShellFunciones import AbrirTransaccion

import os
from Funciones.FuncionesExcel import ExcelService as ServicioExcel
from Config.init_config import in_config as inConfig


def main():
    # 1. Conexión
    sap = ConexionSAP(
                SAP_CONFIG.get('user'),
                SAP_CONFIG.get('password'),
                in_config('SapMandante'),
                in_config('SapIdioma'),
                in_config('SapRutaLogon'),
                in_config('SapSistema')
            )
    sap.iniciar_sesion_sap()
    #session = conectar_sap(in_config("SapSistema"),in_config("SAP_MANDANTE") ,SAP_CONFIG["user"],SAP_CONFIG["password"],)
    #print(session)
    session = ObtenerSesionActiva()

   
    try : 
            #TODO: hacer el bulk por hoja a tabla en base de datos     
            #TicketInsumoRepo.crearPCTicketInsumo( estado=0, observaciones= "Cargue de insumo")
            rutaParametros = os.path.join(inConfig("PathTemp"),"EstrategiasDeLibreacion.xlsx")
            ServicioExcel.ejecutarBulkDesdeExcel(rutaParametros, sheet="ALL")
            #TicketInsumoRepo.crearPCTicketInsumo( estado=100, observaciones= "Cargue de insumo")
    except: 
            #TicketInsumoRepo.crearPCTicketInsumo( error= 99, observaciones="Carge de insumo " )
            print("Error al cargar insumo MainPrueba HU00 Estrategias de liberacion")
            traceback.print_exc()
    


if __name__ =="__main__":
        main()

    
