
from Funciones.LeerXML import LectorFacturaXML
from Funciones.ME2L import TransaccionME2L
from Funciones.MIGO import TransaccionMIGO
from Config.Senttings import SAP_CONFIG
from Config.init_config import in_config
from Funciones.ConexionSAP import ConexionSAP
from HU.HU01_EgresosCuentasPorPagar import Facturas


if __name__ =="__main__":
    pruebaH01=Facturas()
    pruebaH01.ejecutar()

    
