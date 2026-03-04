import getpass 
import logging
import os
from pathlib import Path
from datetime import datetime
import socket
import traceback
from Config.init_config import init_config, in_config

import pandas as pd
from sqlalchemy import text
from Config.Database import Database


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
        maquina = socket.gethostname()
        usuario = getpass.getuser()
        # Nombre de archivo con timestamp
        timestamp = datetime.now().strftime("%d%m%Y")
        log_file = self.path_logs / f"Log_{maquina}_{usuario}_{timestamp}.txt"
        robbot = in_config("CodigoRobot")
        
        # Configuración del logger
        logging.basicConfig(
            level=logging.DEBUG, # Stev:  Cambiar a INFO, para borrar todos los mensajes modo DEBUG
            # FECHA HORA | ESTADO | MENSAJE | CODIGOROBOT | TASKNAME   
            format=rf'%(asctime)s | %(levelname)-2s | %(message)-10s | {robbot} | %(funcName)-20s ',
            #format='%(asctime)s | %(levelname)-8s | %(message)s | RIGO | %(funcName)-20s ',
            datefmt='%Y-%m-%d %H:%M:%S',
            handlers=[
                logging.FileHandler(log_file, encoding='utf-8'),
                logging.StreamHandler()  # También mostrar en consola
            ]
        )
        
        self.logger = logging.getLogger(__name__)
        self.logger.debug("=" * 80)
        self.logger.debug("Sistema de logging inicializado")
        self.logger.debug("=" * 80)
    
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
                    self.logger.debug(f" Carpeta creada: {nombre} -> {carpeta}")
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
            self.logger.info(mensaje)
    
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
            self.logger.info(f"Archivo encontrado: {archivo.name}")
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
        """Carga parámetros desde la base de datos en un diccionario """
        import logging
        logger = logging.getLogger(__name__)
        
        init_config()
        logger.info("In_config cargado:", in_config("PathProyecto"))
        #self.logger.info("Configuracion global iniciada")
    def cargarParametros():
        """Carga parámetros desde el archivo de configuración"""
        import logging
        logger = logging.getLogger(__name__)
        db = Database()
        engine = db.get_engine()
    
        # 1. Ruta del archivo (Asegúrate de tenerla en tu config)
        ruta_excel = rf"{in_config("PathParametrosRIGO")}"
        esquema_destino = "PagoArriendos"

        try:
            # 2. Leer TODAS las hojas a la vez (sheet_name=None devuelve un dict)
            dict_hojas = pd.read_excel(ruta_excel, sheet_name=None)
            
            logger.info(f"--- Iniciando despliegue de parámetros desde {ruta_excel} ---")

            for nombre_hoja, df in dict_hojas.items():
                # Limpiar nombres de columnas (quitar espacios, puntos, etc.)
                df.columns = [col.strip().replace(' ', '_') for col in df.columns]
                
                # 3. Cargar a la base de datos, El nombre de la tabla será el mismo que el de la hoja
                df.to_sql(
                    name=nombre_hoja, 
                    con=engine, 
                    schema=esquema_destino, 
                    if_exists='replace', # Reemplaza la tabla en cada despliegue # Stev: IMPORTANTE! remplaza toda la tabla con registros nuevos
                    index=False
                )
                
                logger.info(f" Tabla [{esquema_destino}].[{nombre_hoja}] cargada exitosamente ({len(df)} registros).")

            logger.info("--- Proceso de despliegue finalizado con éxito ---")

        except Exception as e:
            logger.exception(f" Error critico durante el despliegue: {str(e)}")



Reutilizables.cargar_configuracion()
Reutilizables.cargarParametros()

# Inicializar ambiente al importar
ambiente = Reutilizables(
    in_config("PathProyecto"),
    in_config("PathAudit"),
    in_config("PathLog"),
    in_config("PathTemp"),
    in_config("PathInsumo"),
    in_config("PathResultado")
)

ambiente.crear_carpetas()






    

   

 