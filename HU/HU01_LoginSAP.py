import win32com.client  # pyright: ignore[reportMissingModuleSource]
import time
import getpass
import subprocess
import os
from Config.init_config import in_config

import logging
looger = logging.getLogger(__name__)

from Funciones.ValidacionM21N import ventana_abierta

import pyautogui



def abrir_sap_logon():
    """Abre SAP Logon si no está ya abierto."""
    #SAP_CONFIG = get_sap_config()
    try:
        # Verificar si SAP ya está abierto
        sapgui = win32com.client.GetObject("SAPGUI")
        return True
    except:
        # Si no está abierto, se lanza el ejecutable
        #"logon_path": get_env_variable("SAP_LOGON_PATH"),
        subprocess.Popen(in_config("SAP_LOGON_PATH"))
        time.sleep(5)  # Esperar a que abra SAP Logon
        return False


def conectar_sap(conexion, mandante, usuario, password, idioma="ES"):

    abrir_sap = abrir_sap_logon()
    time.sleep(3)
    if abrir_sap:
        print(" SAP Logon 750 ya se encuentra abierto")
    else:
        print(" SAP Logon 750 abierto ")

    try:
        print("Iniciando conexion con SAP...")

        # 1️⃣ Obtener objeto SAPGUI
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        if not sap_gui_auto:
            raise Exception(
                "No se pudo obtener el objeto SAPGUI. Asegúrate de que SAP Logon esté instalado y el scripting habilitado."
            )

        application = sap_gui_auto.GetScriptingEngine  # motor de Scripting

        # 2️⃣ Buscar conexión activa
        # application.Connections → lista de conexiones (entradas en SAP Logon).
        connection = None
        for item in application.Connections:
            if item.Description.strip().upper() == conexion.strip().upper():
                connection = item
                break

        # 3️⃣ Si no existe conexión abierta, abrir una nueva
        if not connection:
            print(f"Abriendo nueva conexion a {conexion}...")
            connection = application.OpenConnection(conexion, True)
            time.sleep(3)  # Esperar que abra
        else:
            looger.info(f"✅ Conexion existente encontrada con {conexion}.")

        # 4️⃣ Verificar sesión
        if connection.Children.Count > 0:
            session = connection.Children(0)
            looger.info("Sesion existente reutilizada.")
        else:
            session = connection.Children(0).CreateSession()
            looger.info(" Nueva sesion creada.")

        # 5️⃣ Si la pantalla está en login, ingresar credenciales
        # if "RSYST-BNAME" in session.findById("wnd[0]/usr").Text:
        #     print("🧩 Ingresando credenciales...")
        # if password is None:
        #     password = getpass.getpass("Contraseña SAP: ")
        # Ingresar datos de login
        time.sleep(3)
        session.findById("wnd[0]/usr/txtRSYST-MANDT").text = mandante
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = usuario
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password
        session.findById("wnd[0]/usr/txtRSYST-LANGU").text = idioma
        session.findById("wnd[0]").sendVKey(0)
        looger.info(" Conectado correctamente a SAP.")

        if ventana_abierta(session, "Copyrigth"):
            pyautogui.press("enter")

        try:
            if validarLoginDiag(
                ruta_imagen=rf".\img\logindiag.png",
                confidence=0.5,
                intentos=20,
                espera=0.5
            ):
                looger.info("Ventana loginDiag Copyrigth inesperada superada correctamente")
        except Exception as e:
            looger.info(f"no se encontro ventana Copyrigth en login {e}")

        if ventana_abierta(session, "Info de licencia en entrada al sistema múltiple"):
            
            #print("entro a la funcion click")
            time.sleep(20)  
            pyautogui.click()
            pyautogui.press("enter")
               
            try:
                if validarLoginDiag(
                    ruta_imagen=rf".\img\Infodelicenciaenentradaalsistemamultiple.png",
                    confidence=0.8,
                    intentos=20,
                    espera=0.5
                ):  
                    pyautogui.click()
                    #print("encontro la imagen ")
                    #looger.info("Ventana info de licencia inesperada superada correctamente")
            except Exception as e:
                looger.info(f"no se encontro ventana Copyrigth en login {e}")
        return session

    except Exception as e:
        looger.error(f" Error al conectar a SAP: {e}")
        return None


def ObtenerSesionActiva():
    """Obtiene una sesión SAP ya iniciada (con usuario logueado)."""
    try:
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        application = sap_gui_auto.GetScriptingEngine

        # Buscar una conexión activa con sesión
        for conn in application.Connections:
            if conn.Children.Count > 0:
                session = conn.Children(0)
                looger.info(f" Sesion encontrada en conexión: {conn.Description}")
                return session

        looger.warning(" No se encontró ninguna sesion activa.")
        return None

    except Exception as e:
        looger.error(f" Error al obtener la sesion activa: {e}")
        return None



def validarLoginDiag(ruta_imagen, confidence=0.5, intentos=3, espera=0.5):
    """
    Busca una imagen en pantalla y hace Enter cuando la encuentra.

    Args:
    
        ruta_imagen (str): Ruta de la imagen a buscar.
        confidence (float): Confianza para el match (requiere OpenCV).
        intentos (int): Número de intentos antes de fallar.
        espera (float): Tiempo entre intentos en segundos.

    Returns:
        bool: True si hizo click, False si no encontró la imagen.
    """

    for _ in range(intentos):
        pos = pyautogui.locateCenterOnScreen(ruta_imagen, confidence=confidence)
        if pos:
            pyautogui.press("enter")
            return True
        time.sleep(espera)

    looger.warning(f" No se encontró la ventana login diag: {ruta_imagen}")
    return False

