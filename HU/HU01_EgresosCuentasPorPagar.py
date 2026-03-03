from Funciones.ConexionSAP import ConexionSAP
from datetime import datetime
from pathlib import Path
from Funciones.LeerXML import LectorFacturaXML
from Funciones.ME2L import TransaccionME2L
from Funciones.MIGO import TransaccionMIGO
import pandas as pd
from Config.Settings import SAP_CONFIG, CADENA_CONFIG
from Config.init_config import in_config
from Config.Database import Database
from Funciones.DescargarXML import login_colsubsidio, realizar_consulta, descargar_xml_final, mover_archivos, renombrar_archivo
import logging
looger = logging.getLogger(__name__)



class Facturas:
    def __init__(self):
        self.sap=ConexionSAP(
            SAP_CONFIG.get('SAP_USUARIO'),
            SAP_CONFIG.get('SAP_PASSWORD'),
            in_config('SAP_CLIENTE'),
            in_config('SAP_IDIOMA'),
            in_config('SAP_PATH'),
            in_config('SAP_SISTEMA')
        )
        self.sesion = None

        self.pathXML=Path(in_config('PathXML'))

        self.rutaTemporal=in_config('PathTemp')
        self.carpeta=self.rutaTemporal+"\\HU01"
        self.rutaHU01=Path(self.carpeta)

        self.cadenaUsuario=CADENA_CONFIG.get('usuario')
        self.cadenaContraseña=CADENA_CONFIG.get('contrasena')
        self.cadenaRuta=CADENA_CONFIG.get('ruta')

        

    def obtener_documentos(self, columna):
        query=f"SELECT {columna} FROM PagoArriendos.ReporteHU07 WHERE EstadoSAP != 'No existe / Error'"
        db= Database()
        with db.get_connection() as conn:
            cursor =conn.cursor()
            cursor.execute(query)
            resultados =cursor.fetchall()
            return [row[0] for row in resultados]
    
    def obtener_documentos_oc(self, columna, nit):
        query=f"SELECT {columna} FROM PagoArriendos.ReporteHU07 WHERE NIT = '{nit}'"
        db = Database()
        with db.get_connection() as conn:
            cursor =conn.cursor()
            cursor.execute(query)
            resultados =cursor.fetchall()
            return [row[0] for row in resultados]
        
            

    def descargar_XML(self):
        documentos_no_encontrados = []
        try:
            documentos = self.obtener_documentos('NIT')  # consulta a la DB
            
            sesion = login_colsubsidio(self.cadenaUsuario, self.cadenaContraseña, self.cadenaRuta)
            contadorOc=0
            contador=0
            
            for nro_documento in documentos:
                contador+=1
                try:
                    looger.info("AAAAAAAAA")
                    oc=self.obtener_documentos_oc('Oc', nro_documento)
                    looger.info("EEEEEEEE")
                    realizar_consulta(contador, sesion, oc=nro_documento )
                    looger.info("IIIIIIIII")
                    descargar_xml_final(sesion)
                    looger.info("OOOOOOOOOOOO")
                    looger.info(f"[+] XML descargado para documento {nro_documento}")

                    #mover_archivos(r"C:\ProgramData\RPA_RIGO", self.pathXML, 'xml')
                    looger.info("uuuuuuuuuuuuuu")
                    renombrar_archivo(self.pathXML, oc[contadorOc], 'xml')
                    looger.info(oc[contadorOc])
                    contadorOc+=1

                    
                except Exception as e:
                    looger.info(f"[-] No se encontró XML para documento {nro_documento} {e}")
                    documentos_no_encontrados.append({
                        "NumeroDocumento": nro_documento,
                        "Estado": "No encontrado"
                    })

            #   Generar reporte en Excel si hay documentos no encontrados
            if documentos_no_encontrados:
                df = pd.DataFrame(documentos_no_encontrados)
                timestamp = datetime.now().strftime("%Y%m%d")
                ruta_reporte = f"{self.rutaHU01}"+f"\\HU01\\XML_No_Encontrados_{timestamp}.xlsx"
                df.to_excel(ruta_reporte, index=False)
                print(f"[!] Reporte generado: {ruta_reporte}")

        except Exception as e:
            print(f"Error al ingresar al aplicativo cadena: {e}")

    def comparar_XML_SAP(self,oc):
        self.sap.iniciar_sesion_sap()
        xml_path = rf"{self.pathXML}\{oc}.xml"
        datos = LectorFacturaXML(xml_path).obtener_datos()
        
        me2l = TransaccionME2L(self.sap)
        oc = me2l.buscar_oc_activa(datos['nit'])

        if oc:
            migo = TransaccionMIGO(self.sap)
            migo.contabilizar_entrada(oc, datos['factura'])



    def ejecutar(self):
       print("RUTAAAAAAAAAAAAAA", self.cadenaRuta)
       self.descarga= Facturas.descargar_XML(self)
       self.consultaSAP=Facturas.comparar_XML_SAP(self)
       self.descarga
       self.consultaSAP


        


        
        




