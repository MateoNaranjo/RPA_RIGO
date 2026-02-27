from Funciones.CargarAnexo import _interaccion_ventana_windows
import threading
import pyautogui

class ME80FN:

    def __init__(self, sap_conexion):
        self.sap_conexion= sap_conexion
        self.sesion = sap_conexion.sesion
        self.logger = sap_conexion.logger

    def ingresar_oc(self, oc):
        try:
            self.sesion.findById("wnd[0]/usr/ctxtSP$00003-LOW").text = oc
            self.sesion.findById("wnd[0]/tbar[1]/btn[8]").press()

        except Exception as e:
            print("Error al ingresar la orden de compra en la ME80FN", e)


    def entrar_repartos(self):
        try:
            self.sesion.findById("wnd[0]/usr/cntlMEALV_GRID_CONTROL_80FN/shellcont/shell").pressToolbarContextButton("DETAIL_MENU")
            self.sesion.findById("wnd[0]/usr/cntlMEALV_GRID_CONTROL_80FN/shellcont/shell").selectContextMenuItem("TO_EINT")
        except Exception as e:
            self.logger.info(f"Error al entrar a repartos: {e}")   
    
    def exportar_tabla(self, ruta_archivo, nombre: str):
        titulo = "Save As"

        try:
            print("Exportando tabla en ME2L")
            hilo_externo = threading.Thread(target=_interaccion_ventana_windows, args=(ruta_archivo, titulo,))
            hilo_externo.daemon = True
            hilo_externo.start()

            if nombre == "cabecera":
                self.sesion.findById("wnd[0]/usr/cntlMEALV_GRID_CONTROL_80FN/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
                self.sesion.findById("wnd[0]/usr/cntlMEALV_GRID_CONTROL_80FN/shellcont/shell").selectContextMenuItem("&XXL")
                print("Exportando la cabecera")

            elif nombre == "repartos": 
                self.sesion.findById("wnd[0]/usr/cntlMEALV_GRID_CONTROL_80FN_EINT/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")    
                self.sesion.findById("wnd[0]/usr/cntlMEALV_GRID_CONTROL_80FN_EINT/shellcont/shell").selectContextMenuItem("&XXL")
                print("Exportando Repartos")
            
        except Exception as e:
            print("Error al exportar la tabla en la ME80FN", e)
            raise
