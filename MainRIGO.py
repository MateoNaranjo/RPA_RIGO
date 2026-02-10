from Repositorios.Excel import Excel
from HU.HU07_ClasificarOrdenesOC import HU07_ClasificarOC
from HU.HU04_NotificarOCSinFacturar import HU04_Auditoria
from HU.HU02_ValidacionFAC import HU02_VerificacionDiaria
from HU.HU05_GestionAnexos import HU05_CargueSQL
from HU.HU03_OCSinFactura import HU03_DiagnosticoCierre


if __name__ == "__main__":
    # pruebaExcel=HU07_ClasificarOC()
    # pruebaExcel.ejecutar()

     ejecucionHu04=HU04_Auditoria()
     ejecucionHu04.ejecutar()

    # ejecucionHu02=HU02_VerificacionDiaria()
    # ejecucionHu02.ejecutar()

    # nombre_archivo=r"Informe_Auditoria_Facturacion_20260116_0942.xlsx"

    # ejecucionHu03=HU03_DiagnosticoCierre()
    # ejecucionHu03.procesar_desde_excel(nombre_archivo)

    
    # ruta = r"\\192.168.50.169\RPA_RIGO_GestionPagodeArrendamientos\Resultados\Reporte_HU03_Cierre_20260116_1306.xlsx"
    # HU05_CargueSQL.ejecutar_cargue_desde_excel(ruta)

    