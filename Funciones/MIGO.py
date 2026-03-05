import time

class TransaccionMIGO:
    """Clase para realizar entradas de mercancía en SAP"""
    
    def __init__(self, sap_conexion):
        self.sap = sap_conexion
        self.sesion = sap_conexion.sesion
        self.logger = sap_conexion.logger

    def contabilizar_entrada(self, num_oc, referencia_factura):
        """Ejecuta el proceso MIGO para legalizar la factura"""
        try:
            self.sap.abrir_transaccion("/n")
            self.sap.abrir_transaccion("MIGO")
            
            # 1. Ingresar Pedido y Cargar datos
            id_oc = "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2000/ctxtGODYNPRO-PO_NUMBER"
            self.sesion.findById(id_oc).text = num_oc
            self.sesion.findById("wnd[0]").sendVKey(0) 
            self.logger.info(f"OC {num_oc} cargada. Esperando tabla...")
            time.sleep(5) 

            # 3. Ingresar la Referencia (Delivery Note)
            id_ref = "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0110/txtGOHEAD-LFSNR"
            self.sesion.findById(id_ref).text = referencia_factura
            self.logger.info(f"Referencia '{referencia_factura}' ingresada.")

            # 4. Marcar 'Item OK' en la fila 2 (Arrendamiento)
            # Usamos el ID exacto que grabaste: chkGOITEM-TAKE_IT[3,1]
            try:
                id_check_arriendo = "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/chkGOITEM-TAKE_IT[3,1]"
                self.sesion.findById(id_check_arriendo).selected = True
                self.logger.info("Item OK de Arrendamiento (Fila 2) seleccionado.")
            except Exception as e:
                self.logger.error(f"No se pudo marcar el checkbox de arriendo: {e}")

            # 5. CONTABILIZAR (F11)
           
            self.sesion.findById("wnd[0]/tbar[0]/btn[11]").press()
            time.sleep(4)

            # 6. Resultado final
            resultado = self.sesion.findById("wnd[0]/sbar").text
            self.logger.info(f"Resultado final SAP: {resultado}")
            return resultado

        except Exception as e:
            self.logger.error(f"Error crítico en MIGO: {str(e)}")
            return None
          