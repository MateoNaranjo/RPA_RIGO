import time
import re  
import win32com.client
import threading

from Funciones.CargarAnexo import _interaccion_ventana_windows
from pywinauto import Desktop
class TransaccionME2L:

    """Clase para gestionar la consulta de pedidos por proveedor en SAP"""
    
    def __init__(self, sap_conexion):
        self.sap = sap_conexion
        self.sesion = sap_conexion.sesion
        self.logger = sap_conexion.logger

    def buscar_oc_activa(self, nit_proveedor):
        """Busca el primer número de Orden de Compra pendiente en ME2L"""
        try:
            self.sap.abrir_transaccion("ME2L")
            
            # Filtros de búsqueda
            self.sesion.findById("wnd[0]/usr/ctxtEL_LIFNR-LOW").text = nit_proveedor
            self.sesion.findById("wnd[0]/usr/ctxtLISTU").text = "ALV" # Formato tabla
            
            # Ejecutar (F8)
            self.sesion.findById("wnd[0]/tbar[1]/btn[8]").press()
            time.sleep(2)
            
            # Leer el grid de resultados
            grid = self.sesion.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
            
            # Recorremos las filas (hasta un máximo de 10 para no demorar)
            filas_a_revisar = min(grid.rowCount, 10)
            columnas = grid.ColumnOrder
            
            self.logger.info(f"Escaneando {filas_a_revisar} filas en busca de la OC...")

            for fila in range(filas_a_revisar):
                for col_idx in range(columnas.Count):
                    valor = str(grid.getCellValue(fila, columnas(col_idx))).strip()
                    
                    # Buscamos el patrón: empieza por 4, tiene 10 dígitos
                    if re.match(r"^4[0-9]{9}$", valor):
                        self.logger.info(f"¡OC encontrada en fila {fila}!: {valor}")
                        return valor
            
            self.logger.error("No se encontró ningún número de OC válido en las primeras filas.")
            return None
                
        except Exception as e:
            self.logger.error(f"Error escaneando tabla ME2L: {str(e)}")
            return None
        
    def exportar_tabla(self, ruta_archivo):

        desktop = Desktop(backend="win32")
        ventana = desktop.window(title_re="(?i)Save.*As.*")
        hilo_externo = threading.Thread(target=_interaccion_ventana_windows, args=(ruta_archivo, ventana,))
        hilo_externo.daemon = True
        hilo_externo.start()

        self.sesion.findById("wnd[0]/tbar[1]/btn[43]").press()
        self.sesion.findById("wnd[1]/tbar[0]/btn[0]").press()

        # 2. LANZAR HILO PARA VENTANA DE WINDOWS