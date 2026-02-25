# ================================
# Funcion: WriteLog
# Autor: Paula Sierra - Steven Navarro - NetApplications
# Descripcion: Registrar eventos en archivo log con estructura estándar
# Ultima modificacion: 11/02/2026
# Propiedad de Colsubsidio
# Cambios:
# - Se agregan validaciones
# - Se corrige creación de carpetas
# - Se normaliza estructura línea
# - Se agrega argumento fijo para ruta de logs desde inconfig
# ================================

import datetime
import os
import getpass
import socket
import traceback
from config.init_config import inConfig

 
def WriteLog(mensaje: str = "Mensaje vacío", estado: str = "INFO", nombreTarea: str = "TaskNoDefinida", rutaRegistro = inConfig("PathLog"),):
    """
    mensaje  : Texto del log
    estado   : INFO, DEBUG, WARN, ERROR
    nombreTarea: Nombre de HU o Main
    rutaRegistro : Ruta de carpeta Logs o archivo .log
    """

    try:
        # ==========================================================
        # 1. Validaciones básicas
        # ==========================================================
        if not mensaje:
            mensaje = "Mensaje vacío"

        if not nombreTarea:
            nombreTarea = "TaskNoDefinida"

        estado = estado.upper()
        if estado not in ["INFO", "DEBUG", "WARN", "ERROR"]:
            estado = "INFO"

        # ==========================================================
        # 2. Fecha
        # ==========================================================
        ahora = datetime.datetime.now()
        fecha_linea = ahora.strftime("%d/%m/%Y %H:%M:%S")
        fechaArchivo = ahora.strftime("%Y%m%d")

        # ==========================================================
        # 3. Datos del sistema
        # ==========================================================
        nombre_maquina = socket.gethostname()
        usuario = getpass.getuser()

        # ==========================================================
        # 4. Determinar ruta final
        # ==========================================================
        base, extension = os.path.splitext(rutaRegistro)

        if extension:
            carpeta_logs = os.path.dirname(rutaRegistro)
            rutaArchivo = rutaRegistro
        else:
            carpeta_logs = rutaRegistro
            nombreArchivo = f"Log_{nombre_maquina}_{usuario}_{fechaArchivo}.txt"
            rutaArchivo = os.path.join(carpeta_logs, nombreArchivo)

        os.makedirs(carpeta_logs, exist_ok=True)

        # ==========================================================
        # 5. Construcción línea estándar
        # ==========================================================
        linea = (
            f"{fecha_linea} | "
            f"{estado} | "
            f"{mensaje} | "
            f"{inConfig('CodigoRobot')} | "
            f"{nombreTarea} | "
            "\n"
        )

        # ==========================================================
        # 6. Escritura
        # ==========================================================
        with open(rutaArchivo, "a", encoding="utf-8") as f:
            f.write(linea)

    except Exception as e:
        # Nunca debe romper el bot por logging
        print(f"ERROR | Fallo WriteLog: {e} - {traceback.format_exc()}")
        # TODO : pasar error a la funcion writelog 
