import win32com.client  # pyright: ignore[reportMissingModuleSource]


def ObtenerSesionActiva():
    """Obtiene una sesión SAP ya iniciada (con usuario logueado)."""
    try:
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        application = sap_gui_auto.GetScriptingEngine

        # Buscar una conexión activa con sesión
        for conn in application.Connections:
            if conn.Children.Count > 0:
                session = conn.Children(0)
                print(f" Sesion encontrada en conexión: {conn.Description}")
                return session

        print(" No se encontró ninguna sesion activa.")
        return None

    except Exception as e:
        print(f" Error al obtener la sesion activa: {e}")
        return None