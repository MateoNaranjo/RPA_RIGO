# ================================
# Funcion: WriteLog
# Autor: Paula Sierra - NetApplications
# Descripcion: Registrar eventos en archivo log con estructura estándar
# Ultima modificacion: 24/11/2025
# Propiedad de Colsubsidio
# Cambios: Versión inicial
# ================================

import datetime
import os
import getpass
import socket


def WriteLog(mensaje: str, estado: str, task_name: str, path_log: str):
    """
    mensaje  : Texto del log
    estado   : INFO, DEBUG, WARN, ERROR
    task_name: Nombre de HU o Main (ej. 'HU00_DespliegueAmbiente')
    path_log : str → ruta de la carpeta de logs
    """

    # === Fecha para línea y archivo ===
    ahora = datetime.datetime.now()
    fecha_linea = ahora.strftime("%d/%m/%Y %H:%M:%S")
    fecha_archivo = ahora.strftime("%Y%m%d")

    # === Datos del sistema ===
    nombre_maquina = socket.gethostname()
    usuario = getpass.getuser()

    # ------------------------------------------------------------------
    # 1. Determinar si path_log es carpeta o archivo
    # ------------------------------------------------------------------
    base, extension = os.path.splitext(path_log)

    if extension:
        # → es un archivo: "C:/.../error.log"
        carpeta_logs = os.path.dirname(path_log)
        ruta_archivo = path_log
    else:
        # → es carpeta: "C:/.../Logs"
        carpeta_logs = path_log
        nombre_archivo = f"Log_{nombre_maquina}_{usuario}_{fecha_archivo}.log"
        ruta_archivo = os.path.join(carpeta_logs, nombre_archivo)

    # === Asegurar que la carpeta existe ===
    os.makedirs(path_log, exist_ok=True)

    # === Construcción de línea con estructura estándar ===
    """
    FECHA HORA | ESTADO | TASKNAME | MENSAJE | NOMBRE_MAQUINA | USUARIO
    """
    linea = (
        f"{fecha_linea} | "
        f"{estado:<5} | "
        f"{task_name} | "
        f"{mensaje} | "
        f"{nombre_maquina} | "
        f"{usuario}\n"
    )

    # === Guardar log ===
    with open(ruta_archivo, "a", encoding="utf-8") as f:
        f.write(linea)