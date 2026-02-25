import win32com.client
import time
import subprocess
from pathlib import Path
from Config.settings import SAP_CONFIG, PROCESO_CONFIG
from HU.HU00_Despliegue import ambiente

class ConexionSAP:

    def __init__(self, usuario, contrasena, cliente, idioma, aplicativo, sistema):
        self.usuario = usuario
        self.contrasena = contrasena
        self.cliente = cliente
        self.idioma = idioma
        self.aplicativo = aplicativo
        self.sistema = sistema
        self.sesion = None
        self.connection = None
        self.logger = ambiente.logger
    
    def abrir_SAP(self):
        """Abre SAP GUI si no está abierto"""
        try:
            win32com.client.GetObject("SAPGUI")
            self.logger.info("SAP GUI ya está en ejecución")
            return True
        except:
            try:
                self.logger.info(f"Iniciando SAP GUI desde: {self.aplicativo}")
                subprocess.Popen(self.aplicativo)
                time.sleep(5)
                self.logger.info("SAP GUI iniciado exitosamente")
                return True
            except FileNotFoundError:
                self.logger.error(f"Ejecutable no encontrado: {self.aplicativo}")
                return False
            except Exception as e:
                self.logger.error(f"Error al iniciar SAP GUI: {str(e)}")
                return False
    
    def conectar_SAP(self):
        """Establece conexión con SAP"""
        if not self.abrir_SAP():
            return None
        
        max_intentos = PROCESO_CONFIG['ReIntentos']
        
        for intento in range(1, max_intentos + 1):
            try:
                self.logger.info(f"Intento {intento} de {max_intentos} - Conectando a SAP")
                
                # Obtener objeto de automatización
                sap_gui_auto = win32com.client.GetObject("SAPGUI")
                if not sap_gui_auto:
                    raise Exception("No se pudo obtener objeto SAPGUI")
                
                application = sap_gui_auto.GetScriptingEngine
                
                # Verificar si ya hay conexión abierta
                if application.Children.Count > 0:
                    self.connection = application.Children(0)
                    self.logger.info("Usando conexión SAP existente")
                else:
                    # Abrir nueva conexión
                    self.logger.info(f"Abriendo conexión con: {self.sistema}")
                    self.connection = application.OpenConnection(self.sistema, True)
                    time.sleep(2)
                
                # Obtener sesión
                if self.connection.Children.Count > 0:
                    self.sesion = self.connection.Children(0)
                    self.logger.info("Sesión SAP obtenida exitosamente")
                    return self.sesion
                else:
                    raise Exception("No hay sesiones activas")
                
            except Exception as e:
                self.logger.error(f"Error en intento {intento}: {str(e)}")
                if intento < max_intentos:
                    self.logger.info(f"Reintentando en 3 segundos...")
                    time.sleep(3)
                else:
                    self.logger.error("Máximo de intentos alcanzado")
                    return None
        
        return None
    
    def ingresar_SAP(self, sesion):
        """Realiza login en SAP"""
        try:
            
            self.logger.info("Iniciando proceso de login")
            sesion.findById("wnd[0]").maximize()
            sesion.findById("wnd[0]/usr/txtRSYST-BNAME").text = self.usuario
            sesion.findById("wnd[0]/usr/pwdRSYST-BCODE").text = self.contrasena
            sesion.findById("wnd[0]/usr/txtRSYST-LANGU").text = self.idioma
            sesion.findById("wnd[0]/usr/txtRSYST-MANDT").text = self.cliente
            sesion.findById("wnd[0]").sendVKey(0)
            time.sleep(2)
            
            # Verificar si el login fue exitoso
            try:
                # Si hay error de login, aparecerá una ventana de mensaje
                mensaje_error = sesion.findById("wnd[1]/usr/txtMESSTXT1").text
                self.logger.error(f"Error de login: {mensaje_error}")
                return False
            except:
                # Si no hay ventana de error, el login fue exitoso
                self.logger.info(f"Login exitoso - Usuario: {self.usuario}")
                return True
                
        except Exception as e:
            self.logger.error(f"Error en ingresar_SAP: {str(e)}")
            return False
    
    def iniciar_sesion_sap(self):
        """Proceso completo de iniciar sesión"""
        try:
            
            sesion = self.conectar_SAP()
            if not sesion:
                raise Exception("No se pudo crear la sesión SAP")
            
            if not self.ingresar_SAP(sesion):
                raise Exception("Error en el login")
            
            self.sesion = sesion
            time.sleep(2)
            return sesion
            
        except Exception as e:
            self.logger.error(f"Error en iniciar_sesion_sap: {str(e)}")
            return None
    
    def verificar_sesion_activa(self):
        """Verifica si la sesión sigue activa"""
        try:
            if self.sesion:
                _ = self.sesion.Info.SystemName
                return True
        except:
            self.logger.warning("Sesión SAP no está activa")
            return False
        return False
    
    def abrir_transaccion(self, transaccion):
        """Abre una transacción en SAP"""
        if not self.verificar_sesion_activa():
            self.logger.error("Sesión no activa. No se puede abrir transacción")
            return False
        
        try:
            self.logger.info(f"Abriendo transacción: {transaccion}")
            self.sesion.findById("wnd[0]/tbar[0]/okcd").text = transaccion
            self.sesion.findById("wnd[0]").sendVKey(0)
            time.sleep(1)
            self.logger.info(f"Transacción {transaccion} abierta")
            return True
        except Exception as e:
            self.logger.error(f"Error al abrir transacción {transaccion}: {str(e)}")
            return False





'''ejecutarMain=PagoArriendos(SAP_CONFIG.get('SAP_USUARIO'),
                            SAP_CONFIG.get('SAP_PASSWORD'),
                            SAP_CONFIG.get('SapMandante'),
                            SAP_CONFIG.get('SapIdioma'),
                            SAP_CONFIG.get('SapRutaLogon'),
                            SAP_CONFIG.get('SapSistema')
                            )
ejecutarMain.abrir_SAP()
ejecutarMain.ingresar_SAP(ejecutarMain.conectar_SAP())'''

