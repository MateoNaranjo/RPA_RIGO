from Repositorios.Excel import Excel
from HU.HU07_ClasificarOrdenesOC import HU07_ClasificarOC
#from HU.HU08_EstrategiasDeLiberacion import HU08_EstrategiasDeLiberacion
from HU.HU00_Despliegue import Reutilizables

from HU.HU04_NotificarOCSinFacturar import HU04_Auditoria
from HU.HU03_OCSinFactura import HU03_DiagnosticoCierre
from HU.HU02_ValidacionFAC import HU02_VerificacionDiaria
from HU.HU05_GestionAnexos import HU05_CargueSQL
from HU.HU01_EgresosCuentasPorPagar import Facturas


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

    #ejecucionHu01 = Facturas()
    #ejecucionHu01.ejecutar()

    #ejecucionHu04=HU04_Auditoria()
    #ejecucionHu04.ejecutar()


    # nombre_archivo=r"Informe_Auditoria_Facturacion_20260116_0942.xlsx"

    # ejecucionHu03=HU03_DiagnosticoCierre()
    # ejecucionHu03.procesar_desde_excel(nombre_archivo)

    # ejecucionHu02=HU02_VerificacionDiaria()
    # ejecucionHu02.ejecutar()
    
    # ruta = r"\\192.168.50.169\RPA_RIGO_GestionPagodeArrendamientos\Resultados\Reporte_HU03_Cierre_20260116_1306.xlsx"
    # HU05_CargueSQL.ejecutar_cargue_desde_excel(ruta)

    