

import win32com.client  # pyright: ignore[reportMissingModuleSource]
import logging
looger = logging.getLogger(__name__)


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