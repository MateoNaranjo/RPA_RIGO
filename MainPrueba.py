import traceback

from funciones.ConexionSAP import ConexionSAP

from config.settings import SAP_CONFIG
from config.init_config import in_config
from HU.HU01_LoginSAP import ObtenerSesionActiva
import pyperclip

from funciones.GuiShellFunciones import AbrirTransaccion

import os
from funciones.FuncionesExcel import ExcelService as ServicioExcel
from config.init_config import in_config as inConfig


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

    AbrirTransaccion(session, "ZMM_68")
    """
    pues digamos, la mayoría de órdenes se hace en enero, sí, pero pues acá sería bueno colocar,
    digamos en el mes en que está, porque pronto generemos alguna orden durante el mes o algo. 
    Entonces para que la consulte si alguna cosa o k listo, pues coloquemosle acá hasta que que me que estrés.
    """

    import datetime
    # Obtenemos la fecha y hora actual
    ahora = datetime.datetime.now()
    fecha_formateada = ahora.strftime("%d.%m.%Y") # Ejemplo de salida: 01.01.2026
    # Crear una fecha usando el año actual, mes 1, día 1
    primer_dia_anio = datetime.date(ahora.year, 1, 1)
    primer_dia_anio = primer_dia_anio.strftime("%d.%m.%Y")  # Ejemplo de salida: 01.01.2026


    session.findById("wnd[0]/usr/ctxtR_BEDAT-LOW").text = primer_dia_anio #Primer dia del año actual 
    session.findById("wnd[0]/usr/ctxtR_BEDAT-HIGH").text = fecha_formateada #Fecha actual
    session.findById("wnd[0]/usr/btn%_R_EKORG_%_APP_%-VALU_PUSH").press() # Abre VEntana org de Compras 
    session.findById("wnd[1]/tbar[0]/btn[16]").press() #Boton basura, borrar datos 

    grupoOrgCompras = ["OC03","OC30","OC02"]
    texto_sap = "\r\n".join(grupoOrgCompras)
    pyperclip.copy(texto_sap)

    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    # Estado de la OC
    session.findById("wnd[0]/usr/ctxtR_FRGKE-LOW").text = "B" # se Filtra por estado de bloqueo, B 
    """
    4001109218
    4001109602
    4001109605
    4001109690
    4001109698
    4001109712
    4001109718
    4001109720
    4001110010

    4001155953
    4001155956
    4001155955
    4001155957

    """
    # Número de Pedido  
    session.findById("wnd[0]/usr/ctxtR_EBELN-LOW").text = "4001155953"
    session.findById("wnd[0]/usr/txtR_ERNAM-LOW").text = "FERNCAMS" #Responsable
    session.findById("wnd[0]/tbar[1]/btn[8]").press() #Ejecutar búsqueda


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
    


if __name__ == "__main__":
    main()