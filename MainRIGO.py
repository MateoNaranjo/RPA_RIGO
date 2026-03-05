from Repositorios.Excel import Excel
from HU.HU07_ClasificarOrdenesOC import HU07_ClasificarOC
#from HU.HU08_EstrategiasDeLiberacion import HU08_EstrategiasDeLiberacion
from HU.HU00_Despliegue import Reutilizables

from HU.HU04_NotificarOCSinFacturar import HU04_Auditoria
from HU.HU03_OCSinFactura import HU03_DiagnosticoCierre
from HU.HU02_ValidacionFAC import HU02_VerificacionDiaria
from HU.HU05_GestionAnexos import HU05_CargueSQL
from HU.HU01_EgresosCuentasPorPagar import Facturas
from datetime import datetime

import os

def cerrar_sap():
    try:
        # cubrir variantes de SAP GUI
        for proceso in ["saplogon.exe", "saplgpad.exe", "sapgui.exe"]:
            os.system(f"taskkill /f /im {proceso}")
        print("[+] Intento de cierre de SAP completado.")
    except Exception as e:
        print(f"[-] Error al cerrar SAP: {e}")


def cerrar_chrome():
    try:
        os.system("taskkill /f /im chrome.exe")
        print("[+] Chrome cerrado correctamente.")
    except Exception as e:
        print(f"[-] Error al cerrar Chrome: {e}")


cerrarSAP=cerrar_sap()
cerrarGoogle=cerrar_chrome()

if __name__ == "__main__":

   

    '''pruebaExcel=HU08_EstrategiasDeLiberacion()
    pruebaExcel.ejecutar()'''
   

    """
    HU07: Valida la creación y liberación de ordenes de compra
    HU01: Ingresa cadena descarga xml y hace la comparación de valores en sap
    hu04: valida tema de las facturas
    Hu03: se encargan de validar las facturas, valida que el monto este cargado y la factura este creada, 
    si no esta creada es culpa del proovedor y si no esta paga es culpa del cliente
    hu02: Validcion de facturas, por medio del hu03 en casos de error:
    hu06: Validacion de presupuesto
    hu05: Cargar reportes a la bse de datos
    """
    pruebaExcel=HU07_ClasificarOC()
    pruebaExcel.ejecutar()
    cerrar_sap()


    ejecucionHu01 = Facturas()
    ejecucionHu01.ejecutar()
    cerrar_sap()
    cerrar_chrome()

    ejecucionHu04=HU04_Auditoria()
    ejecucionHu04.ejecutar()
    cerrar_sap()

    fecha_hora = datetime.now().strftime('%Y%m%d')
    nombre_archivo=rf"Informe_Auditoria_Facturacion_{fecha_hora}.xlsx"

    ejecucionHu03=HU03_DiagnosticoCierre()
    ejecucionHu03.procesar_desde_excel(nombre_archivo)

    ejecucionHu02=HU02_VerificacionDiaria()
    ejecucionHu02.ejecutar()
    cerrar_sap()
    
    # ruta = r"\\192.168.50.169\RPA_RIGO_GestionPagodeArrendamientos\Resultados\Reporte_HU03_Cierre_20260116_1306.xlsx"
    # HU05_CargueSQL.ejecutar_cargue_desde_excel(ruta)

    