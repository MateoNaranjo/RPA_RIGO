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
from Funciones.DescargarXML import login_colsubsidio, realizar_consulta, descargar_xml_final, renombrar_archivo


class Facturas:
    def __init__(self):
        self.sap=ConexionSAP(
            SAP_CONFIG.get('user'),
            SAP_CONFIG.get('password'),
            in_config('SapMandante'),
            in_config('SapIdioma'),
            in_config('SapRutaLogon'),
            in_config('SapSistema')
        )
        self.sesion = None
        self.consultaSAP=None
        self.descarga=None
        self.pathXML=Path(in_config('PathXML'))

        self.rutaTemporal=in_config('PathTemp')
        self.carpeta=self.rutaTemporal+"\\HU01"
        self.rutaHU01=Path(self.carpeta)

        self.cadenaUsuario=CADENA_CONFIG.get('usuario')
        self.cadenaContraseña=CADENA_CONFIG.get('contrasena')
        self.cadenaRuta=CADENA_CONFIG.get('ruta')

        

    def obtener_documentos(self, columna):
        query=f"SELECT {columna} FROM PagoArriendos.ReporteHU07 WHERE EstadoSAP != 'No existe / Error'"
        with Database.get_connection() as conn:
            cursor =conn.cursor()
            cursor.execute(query)
            resultados =cursor.fetchall()
            return [row[0] for row in resultados]
    
    def obtener_documentos_oc(self, columna, nit):
        query=f"SELECT {columna} FROM PagoArriendos.ReporteHU07 WHERE NIT = '{nit}'"
        with Database.get_connection() as conn:
            cursor =conn.cursor()
            cursor.execute(query)
            resultados =cursor.fetchall()
            return [row[0] for row in resultados]
        
    def obtener_documentos_oc_comparar(self, columna):
        query=f"SELECT {columna} FROM PagoArriendos.ReporteHU07 WHERE NIT != 'No existe / Error'"
        with Database.get_connection() as conn:
            cursor =conn.cursor()
            cursor.execute(query)
            resultados =cursor.fetchall()
            return [row[0] for row in resultados]
        
    def obtener_monto(self, oc):
        query =f"SELECT Monto FROM PagoArriendos.ReporteHU07 WHERE OC = {oc}"
        with Database.get_connection() as conn:
            cursor=conn.cursor()
            cursor.execute(query)
            resultados=cursor.fetchall()
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
                    
                    oc=self.obtener_documentos_oc('Oc', nro_documento)
                    
                    realizar_consulta(contador, sesion, oc=nro_documento )

                    monto=self.obtener_monto(oc[contadorOc])
                    
                    descargar_xml_final(sesion, monto)
                    
                    print(f"[+] XML descargado para documento {nro_documento}")

                    
                    
                    renombrar_archivo(self.pathXML, oc[contadorOc], 'xml')
                    print(oc[contadorOc])
                    contadorOc+=1

                    
                except Exception as e:
                    print(f"[-] No se encontró XML para documento {nro_documento} {e}")
                    documentos_no_encontrados.append({
                        "NumeroDocumento": nro_documento,
                        "Estado": "No encontrado"
                    })

            #   Generar reporte en Excel si hay documentos no encontrados
            if documentos_no_encontrados:
                df = pd.DataFrame(documentos_no_encontrados)
                timestamp = datetime.now().strftime("%Y%m%d")
                ruta_reporte = f"{self.rutaHU01}"+f"\\HU01_{timestamp}.xlsx"
                df.to_excel(ruta_reporte, index=False)
                print(f"[!] Reporte generado: {ruta_reporte}")

        except Exception as e:
            print(f"Error al ingresar al aplicativo cadena: {e}")

    def comparar_XML_SAP(self, oc, contador):
        self.sap.iniciar_sesion_sap()
        nit=self.documentos = self.obtener_documentos('NIT')
        try:
            xml_path = rf"{self.pathXML}\{oc}.xml"
            datos = LectorFacturaXML(xml_path).obtener_datos()
            
            me2l = TransaccionME2L(self.sap)
            oc = me2l.buscar_oc_activa(nit[contador])

            
            migo = TransaccionMIGO(self.sap)
            migo.contabilizar_entrada(oc, datos['factura'])
        except Exception as e:
            print(e)


    def ejecutar(self):
        self.descarga = Facturas.descargar_XML(self)
        oc = self.obtener_documentos_oc_comparar('OC')
        
        print(oc)
        contador=0
        if oc:  # solo si hay elementos
            for i in oc:
                print(i)
                self.consultaSAP = Facturas.comparar_XML_SAP(self, i, contador)
                contador+=1

        # Evita acceder si no existe
        if self.consultaSAP is not None:
            print(self.consultaSAP)


        


        
        




