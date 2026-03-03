# config/Settings.py

import os
from dotenv import load_dotenv
from pathlib import Path
#from Config.init_config import in_config

# Cargar .env
load_dotenv()

# Ruta base del proyecto
BASE_DIR = Path(__file__).resolve().parent.parent

def get_env_variable(key: str, required: bool = True):
    value = os.getenv(key)

    if required and not value:
        raise EnvironmentError(f"La variable '{key}' no está definida en .env")

    return value



# Configuración del proceso
PROCESO_CONFIG = {
    'DIAS_ESPERA_LIBERACION': 2,
    'HORA_LIMITE_ENVIO': '11:45',
    'ReIntentos': 3,
    'TIMEOUT_SAP': 30
}

# ========= CONFIG SAP ==========
SAP_CONFIG = {
    "user": get_env_variable("SAP_USUARIO"),
    "password": get_env_variable("SAP_PASSWORD"),
}

CADENA_CONFIG={
    "usuario":get_env_variable("CADENA_USUARIO"),
    "contrasena":get_env_variable("CADENA_CONTRASENA"),
    "ruta":get_env_variable("CADENA_RUTA")
}

# ========= CONEXION BASE DE DATOS ==========
DB_CONFIG = {
    "host": get_env_variable("SERVERDB"),
    "Database": get_env_variable("NAMEDB"),
    "user": get_env_variable("USERDB"),
    "password": get_env_variable("PASSWORDDB"),    
}


# ========= CONFIG EMAIL ==========
'''CONFIG_EMAIL = {
    "smtp_server": get_env_variable("EMAIL_SMTP_SERVER"),
    "smtp_port": get_env_variable("EMAIL_SMTP_PORT"),
    "email": get_env_variable("EMAIL_USER"),
    "password": get_env_variable("EMAIL_PASSWORD"),  # IMPORTANTE: Cambiar por variable de entorno en producción
}'''

# ========= RUTAS =========
# RUTAS = {
#     "PathLog": get_env_variable("PATHLOG"),
#     "PathLogError": get_env_variable("PATHLOGERROR"),
#     "PathResultados": get_env_variable("PATHRESULTADOS"),
#     "PathReportes": get_env_variable("PATHREPORTES"),
#     "PathInsumo": get_env_variable("PATHINSUMO"),
#     "PathTexto": get_env_variable("PATHTEXTO_SAP"),
#     "PathRuta": get_env_variable("PATHRUTA_SAP"),
#     "PathTempFileServer": get_env_variable("SAP_TEMP_PATH"),
#     # Archivo de configuración de correos
#     "ArchivoCorreos": os.path.join(BASE_DIR, "Insumo", "EnvioCorreos.xlsx"),
#     # Rutas de archivos
#     "PathInsumos": os.path.join(BASE_DIR, "Insumo"),
#     "PathSalida": os.path.join(BASE_DIR, "Salida"),
#     "PathTemp": os.path.join(BASE_DIR, "Temp"),
#     "PathResultado": os.path.join(BASE_DIR, "Resultado"),
# }

# Crear carpetas si no existen
# for key, path in RUTAS.items():
#     if key.startswith("Path") and key not in [
#         "PathLog",
#         "PathLogError",
#         "ArchivoCorreos",
#     ]:
#         os.makedirs(path, exist_ok=True)

# Crear carpeta de logs
#os.makedirs(os.path.dirname(in_config("PathLogs")), exist_ok=True)

