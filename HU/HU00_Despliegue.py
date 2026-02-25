import logging
from pathlib import Path
from datetime import datetime
import traceback
from config.init_config import init_config, in_config

class Reutilizables:
    """Clase para manejo de ambiente y logging del proyecto"""
    
    def __init__(self, path_proyecto, path_audit, path_logs, path_temp, path_insumo, path_resultado):
        self.path_proyecto = Path(path_proyecto)
        self.path_audit = Path(path_audit)
        self.path_logs = Path(path_logs)
        self.path_temp = Path(path_temp)
        self.path_insumo = Path(path_insumo)
        self.path_resultado = Path(path_resultado)
        
        # Configurar logger
        self._configurar_logger()
    
    def _configurar_logger(self):
        """Configura el sistema de logging"""
        # Crear carpeta de logs si no existe
        self.path_logs.mkdir(parents=True, exist_ok=True)
        
        # Nombre de archivo con timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = self.path_logs / f"RPA_Arriendos_{timestamp}.log"
        
        # Configuración del logger
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s | %(levelname)-8s | %(funcName)-20s | %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S',
            handlers=[
                logging.FileHandler(log_file, encoding='utf-8'),
                logging.StreamHandler()  # También mostrar en consola
            ]
        )
        
        self.logger = logging.getLogger(__name__)
        self.logger.info("=" * 80)
        self.logger.info("Sistema de logging inicializado")
        self.logger.info("=" * 80)
    
    def crear_carpetas(self):
        """Crea todas las carpetas necesarias para el proyecto"""
        try:
            carpetas = {
                'Proyecto': self.path_proyecto,
                'Auditoría': self.path_audit,
                'Logs': self.path_logs,
                'Temporal': self.path_temp,
                'Insumos': self.path_insumo,
                'Resultados': self.path_resultado
            }
            
            for nombre, carpeta in carpetas.items():
                if not carpeta.exists():
                    carpeta.mkdir(parents=True, exist_ok=True)
                    self.logger.info(f"✓ Carpeta creada: {nombre} -> {carpeta}")
                else:
                    self.logger.debug(f"Carpeta ya existe: {nombre} -> {carpeta}")
            
            self.logger.info("Despliegue de ambiente completado exitosamente")
            return True
            
        except Exception as e:
            self.logger.error(f"Error al crear carpetas: {str(e)}", exc_info=True)
            return False
    
    def audit_log(self, mensaje, tipo='INFO'):
        """Log de auditoría"""
        if tipo == 'INFO':
            self.logger.info(mensaje)
        elif tipo == 'WARNING':
            self.logger.warning(mensaje)
        elif tipo == 'ERROR':
            self.logger.error(mensaje)
        elif tipo == 'DEBUG':
            self.logger.debug(mensaje)
    
    def limpiar_carpeta_temp(self):
        """Limpia archivos temporales"""
        try:
            archivos_eliminados = 0
            for archivo in self.path_temp.glob('*'):
                if archivo.is_file():
                    archivo.unlink()
                    archivos_eliminados += 1
            
            self.logger.info(f"Carpeta temporal limpiada. {archivos_eliminados} archivos eliminados")
            return True
            
        except Exception as e:
            self.logger.error(f"Error al limpiar carpeta temporal: {str(e)}")
            return False
    
    def validar_archivo_existe(self, ruta_archivo):
        """Valida si un archivo existe"""
        archivo = Path(ruta_archivo)
        if archivo.exists():
            self.logger.debug(f"Archivo encontrado: {archivo.name}")
            return True
        else:
            self.logger.warning(f"Archivo NO encontrado: {archivo}")
            return False
    
    def get_ruta_insumo(self, nombre_archivo):
        """Obtiene ruta completa de archivo en carpeta insumo"""
        return self.path_insumo / nombre_archivo
    
    def get_ruta_resultado(self, nombre_archivo):
        """Obtiene ruta completa de archivo en carpeta resultado"""
        return self.path_resultado / nombre_archivo
    
    def get_ruta_temp(self, nombre_archivo):
        """Obtiene ruta completa de archivo en carpeta temp"""
        return self.path_temp / nombre_archivo
    
    def cargar_configuracion():
        init_config()
        print("In_config cargado:", in_config("PathProyecto"))
        print("Configuracion global iniciada")

Reutilizables.cargar_configuracion()


# Inicializar ambiente al importar
ambiente = Reutilizables(
    in_config("PathProyecto"),
    in_config("PathAudit"),
    in_config("PathLogs"),
    in_config("PathTemp"),
    in_config("PathInsumos"),
    in_config("PathResultados")
)

ambiente.crear_carpetas()




# ================================
# GestionSOLPED – HU00: DespliegueAmbiente
# Autor: Paula Sierra - NetApplications
# Descripcion: Carga parámetros, valida carpetas y prepara entorno
# Ultima modificacion: 30/11/2025
# Propiedad de Colsubsidio
# Cambios: Ajuste ruta base dinámica + estándar Colsubsidio
# ================================

import os
import json

from config.init_config import init_config, in_config
from funciones.FuncionesExcel import ExcelService

import os
#from funciones.FuncionesExcel import ExcelService
#from repositorios.Excel import Excel as ServicioExcel
from config.init_config import in_config as inConfig



#from config.initconfig import init_config

# Inicializar ambiente al importar
ambiente = Reutilizables(
    in_config("PathProyecto"),
    in_config("PathAudit"),
    in_config("PathLogs"),
    in_config("PathTemp"),
    in_config("PathInsumos"),
    in_config("PathResultados")
)


def EjecutarHU00():
    """
    Prepara el entorno: valida carpetas, carga parámetros y estructura inicial.
    """

    # ==========================================================
    # 1. Ruta base del proyecto (importante)
    # ==========================================================
    ruta_base = os.path.dirname(os.path.abspath(__file__))  # ruta de HU00
    ruta_base = os.path.abspath(os.path.join(ruta_base, ".."))
    # Sube un nivel para quedar en /AutomatizacionGestionSolped

    # ==========================================================
    # 2. Definir las carpetas obligatorias según estándar
    # ==========================================================
    carpetas = [
        "Audit/Logs",
        "Audit/Screenshots",
        "Temp",
        "Insumo",
        "Resultado",
        "Funciones",
        "HU",
    ]

    for carpeta in carpetas:
        ruta_completa = os.path.join(ruta_base, carpeta)

        if not os.path.exists(ruta_completa):
            os.makedirs(ruta_completa)

    # ==========================================================
    # 3. Cargar parámetros desde o BD
    # ==========================================================
    init_config()
 
    # ==========================================================
    # 4. Cargar Ecxel con hojas que van a ser las tablas de parametros en la BD
    # ==========================================================

    try : 
            #TODO: hacer el bulk por hoja a tabla en base de datos     
            #TicketInsumoRepo.crearPCTicketInsumo( estado=0, observaciones= "Cargue de insumo")
            rutaParametros = os.path.join(inConfig("PathTemp"),"EnvioCorreos.xlsx")
            ExcelService.ejecutarBulkDesdeExcel(rutaParametros, sheet="ALL")
            #TicketInsumoRepo.crearPCTicketInsumo( estado=100, observaciones= "Cargue de insumo")
    except Exception as e: 
            #TicketInsumoRepo.crearPCTicketInsumo( error= 99, observaciones="Carge de insumo " )
            print("Error al cargar insumo Despliegue HU00 envio correos")
            traceback.print_exc()

   

    

    ruta_config = os.path.join(ruta_base, "config.json")

    if os.path.exists(ruta_config):
        with open(ruta_config, "r", encoding="utf-8") as f:
            config = json.load(f)
    else:
        config = {}

    return config
