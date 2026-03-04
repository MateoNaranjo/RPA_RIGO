import os
import re
import logging
import pandas as pd
import datetime
import pyperclip

from Config.Settings import SAP_CONFIG
from Config.init_config import in_config
from Funciones.ConexionSAP import ConexionSAP
from Repositorios.Excel import Excel as ExcelDB
from Config.Database import Database
from Funciones.GuiShellFunciones import AbrirTransaccion, NotificarErroresEstrategia,ObtenerSesionActiva,LeerTXT_SAP_Universal,validar_estrategias_sap
from Funciones.EmailSender import EmailSender, EnviarCorreoPersonalizado, EnviarNotificacionCorreo

from sqlalchemy import text

import logging
logger = logging.getLogger(__name__)

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
          
        ahora = datetime.datetime.now() # Obtenemos la fecha y hora actual
        fecha_formateada = ahora.strftime("%d.%m.%Y") # Ejemplo de salida: 01.01.2026
        primer_dia_anio = datetime.date(ahora.year, 1, 1)    # Crear una fecha usando el año actual, mes 1, día 1
        primer_dia_anio = primer_dia_anio.strftime("%d.%m.%Y")  # Ejemplo de salida: 01.01.2026

        session.findById("wnd[0]/usr/ctxtR_BEDAT-LOW").text = primer_dia_anio #Primer dia del año actual 
        session.findById("wnd[0]/usr/ctxtR_BEDAT-HIGH").text = fecha_formateada #Fecha actual
        
        # Grupo de Organización de Compras
        grupoOrgCompras = ["OC03","OC30","OC02"]# Esto lo puedes traer de la tabla de la base de datos db, parametros 
        texto_sap = "\r\n".join(grupoOrgCompras)
        pyperclip.copy(texto_sap) # copia al portapapeles la informacion 
        session.findById("wnd[0]/usr/btn%_R_EKORG_%_APP_%-VALU_PUSH").press() # Abre Ventana org de Compras 
        session.findById("wnd[1]/tbar[0]/btn[16]").press() #Boton basura, borrar datos 
        session.findById("wnd[1]/tbar[0]/btn[24]").press() #Boton pegar datos 
        session.findById("wnd[1]/tbar[0]/btn[8]").press() # Ejecutar Filtro 

        # Estado de la OC // Actualizacion : se traen todos los estados, no es necesario filtrar.
        #session.findById("wnd[0]/usr/ctxtR_FRGKE-LOW").text = "B" # se Filtra por estado de bloqueo, B 

        # Número de Pedido  
        #session.findById("wnd[0]/usr/ctxtR_EBELN-LOW").text = "4001155953" # solo para validar una OC. para pruebas 
        listaOC = ["4001155953","4001155956","4001155955","4001155957"] # Esto lo puedes traer de tu tabla de la base de datos db, base medicamentos 
        texto_sap = "\r\n".join(listaOC)
        pyperclip.copy(texto_sap)
        session.findById("wnd[0]/usr/btn%_R_EBELN_%_APP_%-VALU_PUSH").press() # Abre ventana numero de pedido
        session.findById("wnd[1]/tbar[0]/btn[16]").press() #Boton basura, borrar datos 
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
              
        # Responsable
        #session.findById("wnd[0]/usr/txtR_ERNAM-LOW").text = "FERNCAMS" #Responsable ERIIGUZV
        responsable = ["FERNCAMS","ERIIGUZV"] # Esto lo puedes traer de la tabla de la base de datos db, parametros 
        texto_sap = "\r\n".join(responsable)
        pyperclip.copy(texto_sap)
        session.findById("wnd[0]/usr/btn%_R_ERNAM_%_APP_%-VALU_PUSH").press() # Abre ventana responsable de la OC
        session.findById("wnd[1]/tbar[0]/btn[16]").press() # Boton basura, borrar datos
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
               
        
        # Ejecutar búsqueda
        session.findById("wnd[0]/tbar[1]/btn[8]").press() #Ejecutar búsqueda

        # Guardar resultados en Excel
        rutaGuardar = f"{in_config('PathTemp')}\HU08"
        session.findById("wnd[0]/tbar[1]/btn[45]").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = rutaGuardar
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = f"EstrategiasDeLiberacion{fecha_formateada}.txt"
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        #df = LeerTXT_SAP_Universal(os.path.join(rutaGuardar, f"EstrategiasDeLiberacion{fecha_formateada}.txt"))
        df = LeerTXT_SAP_Universal(os.path.join(rutaGuardar, f"EstrategiasDeLiberacion2.txt"))
        #df = pd.read_excel(os.path.join(rutaGuardar, f"EstdeliberacionEjemplos.xlsx"))
        
        # Limpiar espacios en los nombres de las columnas
        df.columns = [re.sub(r'\s+', ' ', str(col)).strip() for col in df.columns]
 
        # Identificar y renombrar duplicados
        cols = pd.Series(df.columns)
        for i in cols[cols.duplicated()].unique():
            cols[cols == i] = [f"{i}_{j}" if j != 0 else i for j in range(sum(cols == i))]
        df.columns = cols # Ahora tus columnas se llamarán "Nombre 1" (la primera) y "Nombre 1_1" (la segunda)

        #df.columns = pd.io.common.dedup_names(df.columns, is_unique=False)  # Stev: Se rompe dependiendo la versión de pandas, se deja esta línea comentada por ahora.
         
           
        # Filtramos solo las columnas que existan en el DataFrame original #2
        columnas_interes = ['Fecha doc.','Acreedor','Nombre 1','Estr.', 'Doc.compr.', 'Status Lib', 'Precio neto', ]
        columnas_validas = [col for col in columnas_interes if col in df.columns]
        df = df[columnas_validas].copy() # Aseguramos que solo trabajamos con las columnas que realmente existen en el DataFrame original
        # Convertir 'Precio neto' a numérico, manejando comas y puntos
        df['Precio neto'] = pd.to_numeric(df['Precio neto'].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False),errors='coerce').fillna(0)

        # Agregar la fecha y hora actual, Usamos format para que SQL lo reconozca como DATETIME fácilmente
        df['FechaActualizacion'] = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        # Agregar el estado de notificación inicial, Lo marcamos como 'Pendiente' para que el módulo de correo sepa qué procesar
        df['EstadoNotificacion'] = 'Pendiente'
               
        # Agrupar por 'Doc.compr.' y sumar 'Precio neto'
        df = df.groupby("Doc.compr.") .agg({  
                "Fecha doc.": "first",
                "Acreedor": "first",
                "Nombre 1": "first",
                "Estr.": "first", 
                "Status Lib": "first",
                "Precio neto": "sum", # Sumamos el precio neto para cada documento de compra  // STEV : se deja fuera del alcance por ahora.
                "FechaActualizacion": "first",
                "EstadoNotificacion": "first",
                # "Fecha Lib": "first",
                # "Usuario Li": "first",
                # "Fecha Lib.": "first",
                # "Usuario Li": "first"
            }).reset_index()



        db= Database()
        engine = db.get_engine()

        df.to_sql("Temp_Estrategias", con=engine, if_exists='replace', index=False)

        # 2. Ejecutar el MERGE en SQL
        with engine.connect() as conn:
            conn.execute(text("""
                MERGE [PagoArriendos].[EstrategiasDeLiberacion] AS Target
                USING [Temp_Estrategias] AS Source
                ON (Target.[Doc.compr.] = Source.[Doc.compr.])
                -- SI YA EXISTE: Actualizamos solo si cambió el Status o la Estrategia
                WHEN MATCHED AND (Target.[Status Lib] <> Source.[Status Lib] OR Target.[Estr.] <> Source.[Estr.]) THEN
                    UPDATE SET 
                        Target.[Status Lib] = Source.[Status Lib],
                        Target.[Estr.] = Source.[Estr.],
                        Target.[EstadoNotificacion] = 'Pendiente', 
                        Target.[FechaActualizacion] = GETDATE()
                -- SI NO EXISTE: Lo insertamos como nuevo
                WHEN NOT MATCHED THEN
                    INSERT ([Doc.compr.], [Status Lib], [Estr.], [EstadoNotificacion], [FechaActualizacion])
                    VALUES (Source.[Doc.compr.], Source.[Status Lib], Source.[Estr.], 'Pendiente', GETDATE());
            """))
            conn.commit()

       
       
        df = pd.read_sql_table("EstrategiasDeLiberacion", engine, schema="PagoArriendos")

        df = df[df['EstadoNotificacion'] == 'Pendiente'].copy()

        print(type(df))
        print("Columnas obtenidas del df de la base de datos:")
        print(df.columns.tolist())
        print("Columnas obtenidas del list(df):")
        print(list(df))
        print("Columnas obtenidas del df.head():")
        print(df.head())
        print("Columnas obtenidas del  df.info()")
        print(df.info())

        

        # === 1. PREPARACIÓN DE GRUPOS ===
        # Grupo de Bloqueadas (B) - Se envían todas juntas a un correo fijo
        df_bloqueadas = df[(df['Status Lib'] == 'B') & (df['EstadoNotificacion'] == 'Pendiente')].copy()
        # Filtro para Pendientes (P) que están Pendientes de notificación
        df_pendientes = df[(df['Status Lib'] == 'P') & (df['EstadoNotificacion'] == 'Pendiente')].copy()
        # Filtro para Liberadas (L) que están Pendientes de notificación
        df_liberadas = df[(df['Status Lib'] == 'L') & (df['EstadoNotificacion'] == 'Pendiente')].copy()


        # === 2. CRUCE PARA CASO 3 (LIBERADAS) ===
        # Leemos el excel y aseguramos que NIT sea string para un cruce limpio
        df_excel_proveedores = pd.read_excel(in_config("ArchivoCorreos"), sheet_name="Proveedores")
        df_excel_proveedores['NIT'] = df_excel_proveedores['NIT'].astype(str).str.strip()
        df_liberadas['Acreedor'] = df_liberadas['Acreedor'].astype(str).str.strip()

       
        #  Si tienes una tabla de proveedores (df_proveedores) con columnas ['NIT', 'Correo', 'Inmueble', 'Contrato']
        # Hacemos un merge para traer la información necesaria
        df_final_L = pd.merge(
            df_liberadas, 
            df_excel_proveedores, 
            left_on='Acreedor', 
            right_on='NIT', 
            how='left'
        )

        


        # --- CASO 1: STATUS 'B' (Correo Único) ---
        if not df_bloqueadas.empty:
            # 1. Preparación de datos para el correo
            destinatario_b = "chuto00@gmail.com" 
            fecha_hoy_str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            asunto_b = f"Reporte OC Bloqueadas (Status B) - {fecha_hoy_str}"
            cuerpo_b = f"<h3>Listado de OCs Bloqueadas:</h3> {df_bloqueadas.to_html(index=False)}"

            try:
                # 2. Enviamos el correo
                EnviarCorreoPersonalizado(
                    destinatario=destinatario_b,
                    asunto=asunto_b,
                    cuerpo=cuerpo_b,
                    nombreTarea="Notificacion_Bloqueadas"
                )
                logger.info("Correo de Bloqueadas enviado con éxito.")

                # 3. ACTUALIZACIÓN EN BASE DE DATOS
                # Extraemos los Documentos de Compras (IDs únicos) para marcarlos como 'Enviado'
                # Ajusta 'Documento compras' al nombre real de tu columna de ID
                lista_ids = df_bloqueadas['Doc.compr.'].astype(str).tolist()
                ids_para_sql = ", ".join([f"'{id}'" for id in lista_ids])

                with engine.connect() as conn:
                    query = text(f"""
                        UPDATE [PagoArriendos].[EstrategiasDeLiberacion]
                        SET EstadoNotificacion = 'EnviadoB', 
                            FechaActualizacion = '{fecha_hoy_str}'
                        WHERE [Doc.compr.] IN ({ids_para_sql})
                    """)
                    conn.execute(query)
                    conn.commit()
                    
                logger.info(f" Se marcaron {len(lista_ids)} - {lista_ids} registros como 'Enviado' en SQL.")

            except Exception as e:
                logger.error(f"Error al procesar notificaciones de Bloqueadas: {str(e)}")

        # --- CASO 2: STATUS 'P' (Correo según Estrategia) ---
        if not df_pendientes.empty:
            # Generamos la fecha una sola vez para este bloque
            fecha_hoy_str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            # Agrupamos por Estrategia para enviar un correo por cada una
            for estrategia, grupo in df_pendientes.groupby("Estr."):
                try:
                    # 1. Configuración del destinatario y contenido
                    # Reemplaza esta línea con tu lógica de búsqueda (ej. desde un diccionario de parámetros)
                    correo_estrategia = "Steven.navarro@netapplications.com.co" 
                    
                    asunto_p = f"OC Pendientes por Liberar - Estrategia: {estrategia} - {fecha_hoy_str}"
                    cuerpo_p = f"<h3>OCs asignadas a su estrategia {estrategia}:</h3> {grupo.to_html(index=False)}"
                    
                    # 2. Enviamos el correo
                    EnviarCorreoPersonalizado(
                        destinatario=correo_estrategia,
                        asunto=asunto_p,
                        cuerpo=cuerpo_p,
                        nombreTarea=f"Notificacion_Pendientes_{estrategia}"
                    )
                    logger.info(f"Correo enviado con éxito para la estrategia: {estrategia}")

                    # 3. ACTUALIZACIÓN EN BASE DE DATOS (Solo para los registros de este grupo/estrategia)
                    lista_ids_p = grupo['Doc.compr.'].astype(str).tolist()
                    ids_para_sql_p = ", ".join([f"'{id}'" for id in lista_ids_p])

                    with engine.connect() as conn:
                        query = text(f"""
                            UPDATE [PagoArriendos].[EstrategiasDeLiberacion]
                            SET EstadoNotificacion = 'EnviadoP', 
                                FechaActualizacion = '{fecha_hoy_str}'
                            WHERE [Doc.compr.] IN ({ids_para_sql_p})
                        """)
                        conn.execute(query)
                        conn.commit()
                    
                    logger.info(f" SQL: Se marcaron {len(lista_ids_p)} - {lista_ids_p} registros de la estrategia {estrategia} como 'EnviadoP'.")

                except Exception as e:
                    logger.error(f" Error al procesar la estrategia {estrategia}: {str(e)}")

        # # --- CASO 2: STATUS 'P' (Correo según Estrategia) ---
        # if not df_pendientes.empty:
        #     # Agrupamos por Estrategia para enviar un correo por cada una
        #     for estrategia, grupo in df_pendientes.groupby("Estr."):
        #         # AQUÍ BUSCAMOS EL CORREO SEGÚN LA ESTRATEGIA
        #         # Supongamos que tienes un diccionario o lo buscas en tu df_excel
        #         # correo_estrategia = buscar_correo_por_estrategia(estrategia)
        #         correo_estrategia = "Steven.navarro@netapplications.com.co" # Esto es solo un ejemplo, debes reemplazarlo con tu lógica real para obtener el correo
                
        #         asunto_p = f"OC Pendientes por Liberar - Estrategia: {estrategia}"
        #         cuerpo_p = f"<h3>OCs asignadas a su estrategia {estrategia}:</h3> {grupo.to_html(index=False)}"
                
        #         EnviarCorreoPersonalizado(
        #             destinatario=correo_estrategia,
        #             asunto=asunto_p,
        #             cuerpo=cuerpo_p,
        #             nombreTarea=f"Notificacion_Pendientes_{estrategia}"


        # --- CASO 3: STATUS 'L' (Liberadas / Notificación a Arrendadores) ---
        if not df_final_L.empty:
            # Generamos la fecha con el formato estándar del Bot
            fecha_hoy_str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

            # Agrupamos por Acreedor para no saturar al arrendador con múltiples correos
            for acreedor, grupo in df_final_L.groupby("Acreedor"):
                try:
                    # 1. Extracción de datos del Arrendador
                    correo_proveedor = grupo['Correo Proveedor'].iloc[0]
                    nombre_arrendador = grupo['Nombre 1'].iloc[0]
                    asunto = f"COLSUBSIDIO - Orden de compra 2026 Arrendamiento y/o Administración"
                    
                    # 2. Construcción dinámica de la tabla HTML
                    filas_tabla = ""
                    for _, row in grupo.iterrows():
                        filas_tabla += f"""
                        <tr>
                            <td style='border: 1px solid #0056b3; padding: 8px;'>{row.get('Inmueble', 'N/A')}</td>
                            <td style='border: 1px solid #0056b3; padding: 8px;'>{row.get('No de contrato', 'N/A')}</td>
                            <td style='border: 1px solid #0056b3; padding: 8px;'>{row['Acreedor']}</td>
                            <td style='border: 1px solid #0056b3; padding: 8px;'>{nombre_arrendador}</td>
                            <td style='border: 1px solid #0056b3; padding: 8px;'>ARRIENDO</td>
                            <td style='border: 1px solid #0056b3; padding: 8px;'><b>{row['Doc.compr.']}</b></td>
                        </tr>
                        """

                    cuerpo_html = f"""
                    <html>
                    <body style='font-family: Arial, sans-serif; color: #333;'>
                        <p>Buen día, espero se encuentren muy bien.</p>
                        <p>A continuación comparto el (los) número(s) de la(s) órden(es) de compra, correspondiente al canon y/o administración para el periodo comprendido de Enero a Diciembre 2026.</p>
                        <p>Adjunto lineamientos de facturacion electronica... <b>Lo anterior ayudará a que podamos identificar con mayor facilidad su factura...</b></p>
                        
                        <table style='border-collapse: collapse; width: 100%; text-align: center;'>
                            <tr style='background-color: #0056b3; color: white;'>
                                <th>Inmueble</th>
                                <th>No de contrato</th>
                                <th>NIT</th>
                                <th>Arrendador</th>
                                <th>TIPO</th>
                                <th>ORDEN 2026</th>
                            </tr>
                            {filas_tabla}
                        </table>
                        
                        <p>**Recuerde que las facturas electrónicas deben ser enviadas al correo <a href='mailto:recepcion.facturaelectronica@colsubsidio.com'>recepcion.facturaelectronica@colsubsidio.com</a></p>
                        <p>Cordial saludo,</p>
                        <div style='color: #0056b3;'>
                            <b>Atentamente, Robot RIGO | Administración Inmobiliaria</b><br>
                            Gerencia de servicios administrativos
                        </div>
                    </body>
                    </html>
                    """
                    
                    # 3. Envío del Correo
                    logger.info(f" Enviando notificación de OC Liberada al arrendador: {nombre_arrendador} ({correo_proveedor})")
                    EnviarCorreoPersonalizado(
                        destinatario=correo_proveedor,
                        asunto=asunto,
                        cuerpo=cuerpo_html,
                        nombreTarea=f"Notificacion_Liberadas_{acreedor}"
                    )

                    # 4. ACTUALIZACIÓN EN BASE DE DATOS (Estado EnviadoL)
                    lista_ids_l = grupo['Doc.compr.'].astype(str).tolist()
                    ids_para_sql_l = ", ".join([f"'{id}'" for id in lista_ids_l])

                    with engine.connect() as conn:
                        query = text(f"""
                            UPDATE [PagoArriendos].[EstrategiasDeLiberacion]
                            SET EstadoNotificacion = 'EnviadoL', 
                                FechaActualizacion = '{fecha_hoy_str}'
                            WHERE [Doc.compr.] IN ({ids_para_sql_l})
                        """)
                        conn.execute(query)
                        conn.commit()
                    
                    logger.info(f" SQL: Se marcaron {len(lista_ids_l)} - {lista_ids_l}registros del arrendador {acreedor} como 'EnviadoL'.")

                except Exception as e:
                    logger.error(f" Error al notificar al arrendador {acreedor}: {str(e)}")
                

        # # --- CASO 3: STATUS 'L' (Solo guardamos el Excel con la información completa) ---
        # if not df_final_L.empty:
        #     for acreedor, grupo in df_final_L.groupby("Acreedor"):
        #         # Extraer datos básicos del proveedor (asumiendo que vienen en el merge)
        #         correo_proveedor = grupo['Correo Proveedor'].iloc[0]
        #         nombre_arrendador = grupo['Nombre 1'].iloc[0]
                
        #         asunto = f"COLSUBSIDIO-Orden de compra 2026 Arrendamiento y/o Administración"
                
        #         # Construcción de las filas de la tabla (Dinámico por si hay varias OCs para un mismo NIT)
        #         filas_tabla = ""
        #         for _, row in grupo.iterrows():
        #             filas_tabla += f"""
        #             <tr>
        #                 <td style='border: 1px solid #0056b3; padding: 8px;'>{row.get('Inmueble', 'N/A')}</td>
        #                 <td style='border: 1px solid #0056b3; padding: 8px;'>{row.get('No de contrato', 'N/A')}</td>
        #                 <td style='border: 1px solid #0056b3; padding: 8px;'>{row['Acreedor']}</td>
        #                 <td style='border: 1px solid #0056b3; padding: 8px;'>{nombre_arrendador}</td>
        #                 <td style='border: 1px solid #0056b3; padding: 8px;'>ARRIENDO</td>
        #                 <td style='border: 1px solid #0056b3; padding: 8px;'><b>{row['Doc.compr.']}</b></td>
        #             </tr>
        #             """

        #         cuerpo_html = f"""
        #         <html>
        #         <body style='font-family: Arial, sans-serif; color: #333;'>
        #             <p>Buen día, espero se encuentren muy bien.</p>
        #             <p>A continuación comparto el (los) número(s) de la(s) órden(es) de compra, correspondiente al canon y/o administración para el periodo comprendido de Enero a Diciembre 2026.</p>
        #             <p>Adjunto lineamientos de facturacion electronica... <b>Lo anterior ayudará a que podamos identificar con mayor facilidad su factura...</b></p>
                    
        #             <table style='border-collapse: collapse; width: 100%; text-align: center;'>
        #                 <tr style='background-color: #0056b3; color: white;'>
        #                     <th>Inmueble</th>
        #                     <th>No de contrato</th>
        #                     <th>NIT</th>
        #                     <th>Arrendador</th>
        #                     <th>TIPO</th>
        #                     <th>ORDEN 2026</th>
        #                 </tr>
        #                 {filas_tabla}
        #             </table>
                    
        #             <p>**Recuerde que las facturas electrónicas deben ser enviadas al correo <a href='mailto:recepcion.facturaelectronica@colsubsidio.com'>recepcion.facturaelectronica@colsubsidio.com</a></p>
        #             <p>Cordial saludo,</p>
        #             <div style='color: #0056b3;'>
        #                 <b>Yohan Guzmán | Analista Administración Inmobiliaria</b><br>
        #                 Gerencia de servicios administrativos
        #             </div>
        #         </body>
        #         </html>
        #         """
                
        #         # Enviar
        #         logger.info(f"📧 Enviando Caso 3 a: {correo_proveedor}")
        #         EnviarCorreoPersonalizado(
        #             destinatario=correo_proveedor,
        #             asunto=asunto,
        #             cuerpo=cuerpo_html,
        #             nombreTarea=f"Notificacion_Liberadas_Arrendatarios{estrategia}"
                    
        #         )
          
   

