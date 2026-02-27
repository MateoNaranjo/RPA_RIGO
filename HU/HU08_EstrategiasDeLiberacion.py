import os
import re
import logging
import pandas as pd
from datetime import datetime
import time
import threading
import pyperclip
import traceback


from Config.settings import SAP_CONFIG
from Config.init_config import in_config
from Funciones.ConexionSAP import ConexionSAP
from Funciones.consultarOC import consultarOC
from Funciones.CargarAnexo import cargar_archivo_gos # Asegúrate de que este archivo exista
from repositorios.Excel import Excel as ExcelDB
from Funciones.GuiShellFunciones import AbrirTransaccion,ObtenerSesionActiva,LeerTXT_SAP_Universal,validar_estrategias_sap
from Funciones.EmailSender import EmailSender, EnviarNotificacionCorreo

class  HU08_EstrategiasDeLiberacion:
    def __init__(self):
        """
        Inicializa los componentes de conexión y logging.
        """
        self.logger = logging.getLogger("HU07_ClasificarOC")
        self.sap = ConexionSAP(
            SAP_CONFIG.get('user'),
            SAP_CONFIG.get('password'),
            in_config('SapMandante'),
            in_config('SapIdioma'),
            in_config('SapRutaLogon'),
            in_config('SapSistema')
        )
    


    def ejecutar(self):
        """
        Docstring for ejecutar
        
        :param self: Description
        """
        session = self.sap.iniciar_sesion_sap()
        if not session: return
      
        session = ObtenerSesionActiva()

        AbrirTransaccion(session, "ZMM_68")
        """
        pues digamos, la mayoría de órdenes se hace en enero, sí, pero pues acá sería bueno colocar,
        digamos en el mes en que está, porque pronto generemos alguna orden durante el mes o algo. 
        Entonces para que la consulte si alguna cosa o k listo, pues coloquemosle acá hasta que que me que estrés.
        """

        import datetime
        # Obtenemos la fecha y hora actual
        ahora = datetime.datetime.now()
        fecha_formateada = ahora.strftime("%d.%m.%Y") # Ejemplo de salida: 01.01.2026
        # Crear una fecha usando el año actual, mes 1, día 1
        primer_dia_anio = datetime.date(ahora.year, 1, 1)
        primer_dia_anio = primer_dia_anio.strftime("%d.%m.%Y")  # Ejemplo de salida: 01.01.2026


        session.findById("wnd[0]/usr/ctxtR_BEDAT-LOW").text = primer_dia_anio #Primer dia del año actual 
        session.findById("wnd[0]/usr/ctxtR_BEDAT-HIGH").text = fecha_formateada #Fecha actual
        
        # Grupo de Organización de Compras
        grupoOrgCompras = ["OC03","OC30","OC02"]
        texto_sap = "\r\n".join(grupoOrgCompras)
        pyperclip.copy(texto_sap)
        session.findById("wnd[0]/usr/btn%_R_EKORG_%_APP_%-VALU_PUSH").press() # Abre Ventana org de Compras 
        session.findById("wnd[1]/tbar[0]/btn[16]").press() #Boton basura, borrar datos 
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()

        # Estado de la OC
        session.findById("wnd[0]/usr/ctxtR_FRGKE-LOW").text = "B" # se Filtra por estado de bloqueo, B 

        # Número de Pedido  
        #session.findById("wnd[0]/usr/ctxtR_EBELN-LOW").text = "4001155953"
        listaOC = ["4001109218","4001109602","4001109605","4001109690","4001109698","4001109712","4001109718","4001109720","4001110010",
                   "4001155953","4001155956","4001155955","4001155957"]
        texto_sap = "\r\n".join(listaOC)
        pyperclip.copy(texto_sap)
        session.findById("wnd[0]/usr/btn%_R_EBELN_%_APP_%-VALU_PUSH").press() # Abre ventana numero de pedido
        session.findById("wnd[1]/tbar[0]/btn[16]").press() #Boton basura, borrar datos 
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
              
        # Responsable
        responsable = ["FERNCAMS","ERIIGUZV"]
        texto_sap = "\r\n".join(responsable)
        pyperclip.copy(texto_sap)
        session.findById("wnd[0]/usr/btn%_R_ERNAM_%_APP_%-VALU_PUSH").press() # Abre ventana responsable de la OC
        session.findById("wnd[1]/tbar[0]/btn[16]").press() # Boton basura, borrar datos
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        #session.findById("wnd[0]/usr/txtR_ERNAM-LOW").text = "FERNCAMS" #Responsable ERIIGUZV
        
        
        # Ejecutar búsqueda
        session.findById("wnd[0]/tbar[1]/btn[8]").press() #Ejecutar búsqueda


        # Guardar resultados en Excel
        rutaGuardar = in_config("PathTemp") 
        session.findById("wnd[0]/tbar[1]/btn[45]").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = rutaGuardar
        #time.sleep(10)  # Esperar a que se abra la ventana de guardado
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = f"EstrategiasDeLiberacion{fecha_formateada}.txt"
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        df = LeerTXT_SAP_Universal(os.path.join(rutaGuardar, f"EstrategiasDeLiberacion{fecha_formateada}.txt"))
        # Limpiar espacios en los nombres de las columnas
        df.columns = df.columns.str.strip()
        # Eliminar filas duplicadas basándonos en la columna 'Doc.compr.'
        #df.drop_duplicates(subset=['Doc.compr.'], inplace=True)
        
        
        # Filtramos solo las columnas que existan en el DataFrame original #2
        columnas_interes = ['Fecha doc.', 'Estr.', 'Doc.compr.', 'Status Lib', 'Precio neto', 'Fecha Lib', 'Usuario Li', 'Fecha Lib.']
        columnas_validas = [col for col in columnas_interes if col in df.columns]
        df_final = df[columnas_validas].copy() # Aseguramos que solo trabajamos con las columnas que realmente existen en el DataFrame original
        # Convertir 'Precio neto' a numérico, manejando comas y puntos
        df_final['Precio neto'] = pd.to_numeric(df_final['Precio neto'].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False),errors='coerce').fillna(0)

       
        # Agrupar por 'Doc.compr.' y sumar 'Precio neto'
        df_sum = df_final.groupby("Doc.compr.") .agg({  
                "Fecha doc.": "first",
                "Estr.": "first", 
                "Status Lib": "first",               
                "Precio neto": "sum"
                # "Fecha Lib": "first",
                # "Usuario Li": "first",
                # "Fecha Lib.": "first"
            }).reset_index()
        # Guardar el DataFrame resultante en un nuevo archivo CSV para revisar resultados intermedios
        df_sum.to_csv(os.path.join(rutaGuardar, f"EstrategiasDeLiberacion{fecha_formateada}.csv"), index=False)

        # Cargar la hoja específica
        ruta_excel = rf"{in_config('PathInsumo')}\EnvioCorreos.xlsx"
        nombre_pestana = "EstrategiasDeLiberacionOC" # Cambia esto por el nombre real        
        df_excel = pd.read_excel(ruta_excel, sheet_name=nombre_pestana)
        # Limpiar espacios en los nombres de las columnas
        df_excel.columns = df_excel.columns.str.strip()

        # Ejecutar validación
        df_sap_validado = validar_estrategias_sap(df_sum, df_excel)
        # Ver resultados
        print("Columnas obtenidas del df_sap_validado :")
        print(df_sap_validado[['Doc.compr.', 'Precio neto', 'Estr.', 'Resultado_Validacion']])



        EnviarNotificacionCorreo( codigoCorreo=1,nombreTarea="Prueba RIGO - Notificacion",)

           

   


        try : 
                #TODO: hacer el bulk por hoja a tabla en base de datos     
                #TicketInsumoRepo.crearPCTicketInsumo( estado=0, observaciones= "Cargue de insumo")
                rutaParametros = os.path.join(in_config("PathTemp"),"EstrategiasDeLibreacion.xlsx")
                #ServicioExcel.ejecutarBulkDesdeExcel(rutaParametros, sheet="ALL")
                #TicketInsumoRepo.crearPCTicketInsumo( estado=100, observaciones= "Cargue de insumo")
        except: 
                #TicketInsumoRepo.crearPCTicketInsumo( error= 99, observaciones="Carge de insumo " )
                print("Error al cargar insumo MainPrueba HU00 Estrategias de liberacion")
                traceback.print_exc()

 