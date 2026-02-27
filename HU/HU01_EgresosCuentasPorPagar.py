from Funciones.ConexionSAP import ConexionSAP
from datetime import datetime
from Funciones.LeerXML import LectorFacturaXML
from Funciones.ME2L import TransaccionME2L
from Funciones.MIGO import TransaccionMIGO
from Config.Settings import SAP_CONFIG, CADENA_CONFIG
from Config.init_config import in_config
from Funciones.DescargarXML import login_colsubsidio, realizar_consulta, descargar_xml_final


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

        self.cadenaUsuario=CADENA_CONFIG.get('CADENA_USUARIO')
        self.cadenaContraseña=CADENA_CONFIG.get('CADENA_CONTRASEÑA')
        self.cadenaRuta=CADENA_CONFIG.get('CADENA_RUTA')

    def descargar_XML(self):
        nro_documento="4001249504"
        try:
            sesion = login_colsubsidio(self.cadenaUsuario, self.cadenaContraseña, self.cadenaRuta)
            realizar_consulta(sesion, oc=nro_documento)
            descargar_xml_final(sesion)
        except Exception:
            print("error al ingresar al aplicativo cadena")

    def comparar_XML_SAP(self):
        self.sap.iniciar_sesion_sap()
        xml_path = r"C:\ProgramData\RIGO\Insumo\ad090063145002525021701C7.xml"
        datos = LectorFacturaXML(xml_path).obtener_datos()
        
        me2l = TransaccionME2L(self.sap)
        oc = me2l.buscar_oc_activa(datos['nit'])

        if oc:
            migo = TransaccionMIGO(self.sap)
            migo.contabilizar_entrada(oc, datos['factura'])



    def ejecutar(self):
       self.descarga= Facturas.descargar_XML(self)
       self.consultaSAP=Facturas.comparar_XML_SAP(self)
       self.descarga
       self.consultaSAP


        


        
        




