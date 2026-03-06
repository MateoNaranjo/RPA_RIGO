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
from Funciones.GuiShellFunciones import AbrirTransaccion,ObtenerSesionActiva,LeerTXT_SAP_Universal,impimmirdf,fomatodf,descargadataestliberacion
from Funciones.EmailSender import EmailSender, EnviarCorreoPersonalizado, EnviarNotificacionCorreo

from sqlalchemy import text

import logging
logger = logging.getLogger(__name__)

class  HU08_EstrategiasDeLiberacion:
    def __init__(self):
        """
        Inicializa los componentes de conexión y logging.
        """
        self.logger = logging.getLogger("HU8")
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
        db= Database()
        engine = db.get_engine()
        #descargadataestliberacion (session = self.sap.iniciar_sesion_sap())
        EnviarNotificacionCorreo(codigoCorreo=1,nombreTarea="Probando db")  #Probando metodos de envio de correo
        rutaGuardar = fr"{in_config('PathTemp')}\HU08"
        ahora = datetime.datetime.now() # Obtenemos la fecha y hora actual
        fecha_formateada = ahora.strftime("%d.%m.%Y") # Ejemplo de salida: 01.01.2026

        #df = LeerTXT_SAP_Universal(os.path.join(rutaGuardar, f"EstrategiasDeLiberacion{fecha_formateada}.txt"))
        #df = LeerTXT_SAP_Universal(os.path.join(rutaGuardar, f"EstrategiasDeLiberacion5.txt"))
        df = pd.read_excel(os.path.join(rutaGuardar, f"EstdeliberacionEjemplosRIGO.xlsx"))
        
        df = fomatodf(df)
        logger.debug("imprimir df para pruebas: ")
        #impimmirdf(df)
         
        #df.to_sql("EstrategiasDeLiberacion", engine, schema="PagoArriendos", if_exists='replace', index=False)
    
        df.to_sql("Temp_Estrategias", con=engine, if_exists='replace', index=False)

        #  Ejecutar el MERGE en SQL
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
                        Target.[FechaActualizacion] = GETDATE(),
                        Target.[ContadorEnvio] = 0
                -- SI NO EXISTE: Lo insertamos como nuevo
                WHEN NOT MATCHED THEN
                    INSERT (
                        [Doc.compr.], [Fecha doc.], [Acreedor], [Nombre 1], [Creado], 
                        [Estr.], [Status Lib], [Precio neto],[FechaActualizacion],[EstadoNotificacion],[ContadorEnvio]
                    )
                    VALUES (
                        Source.[Doc.compr.], Source.[Fecha doc.], Source.[Acreedor], Source.[Nombre 1], Source.[Creado],
                        Source.[Estr.], Source.[Status Lib], Source.[Precio neto], GETDATE(),'Pendiente', 0
                    );
            """))
            conn.commit()
        #*************************
        # revisar fechas y si pasan 2 dias, cambiar estado a pendiente 
        #*********************

        # --- CONFIGURACIÓN PARA PRUEBAS ---
        MODO_PRUEBA = True  # Cambia a False para producción

        if MODO_PRUEBA:
            intervalo_sql = "second"
            valor_tiempo = -5  # 10 segundos para probar rápido
        else:
            intervalo_sql = "day"
            valor_tiempo = -2   # 2 días para producción


        with engine.connect() as conn:
        # Usamos f-string para {intervalo_sql} porque DATEADD no acepta parámetros en esa posición
            query_update = text(f"""
                UPDATE [PagoArriendos].[EstrategiasDeLiberacion]
                SET EstadoNotificacion = 'Pendiente'
                WHERE [Status Lib] <> 'L' 
                AND [FechaActualizacion] <= DATEADD({intervalo_sql}, :valor, GETDATE())
                AND EstadoNotificacion LIKE 'Enviado%'
            """)
            
            conn.execute(query_update, {"valor": valor_tiempo})
            conn.commit()
            logger.info(f"Reset de estados ejecutado usando intervalo: {valor_tiempo} {intervalo_sql}")

        
        # with engine.connect() as conn:
        #     conn.execute(text("""
        #         UPDATE [PagoArriendos].[EstrategiasDeLiberacion]
        #         SET EstadoNotificacion = 'Pendiente'
        #         WHERE [Status Lib] <> 'L' 
        #         AND [FechaActualizacion] <= DATEADD(day, -2, GETDATE())
        #         AND EstadoNotificacion LIKE 'Enviado%'
        #     """))
        #     conn.commit()
              
       
        df = pd.read_sql_table("EstrategiasDeLiberacion", engine, schema="PagoArriendos")

        #df = df[df['ContadorEnvio'] == 0].copy()
        
        df = df[df['EstadoNotificacion'] == 'Pendiente'].copy()

        # === 1. PREPARACIÓN DE GRUPOS ===
        # Grupo de Bloqueadas (B) - Se envían todas juntas a un correo fijo
        df_bloqueadas = df[(df['Status Lib'] == 'B') & (df['EstadoNotificacion'] == 'Pendiente')].copy()
        # Filtro para Pendientes (P) que están Pendientes de notificación
        df_pendientes = df[(df['Status Lib'] == 'P') & (df['EstadoNotificacion'] == 'Pendiente')].copy()
        # Filtro para Liberadas (L) que están Pendientes de notificación
        df_liberadas = df[(df['Status Lib'] == 'L') & (df['EstadoNotificacion'] == 'Pendiente')].copy()
        #df_liberadas = df_liberadas[df_liberadas['ContadorEnvio'] == 0].copy()
        # Filtro para Oc que llevan mas de 3 notificaciones 
        df_atrasadas = df[df['ContadorEnvio'] > 3].copy()

        
        impimmirdf(df_bloqueadas)
        impimmirdf(df_pendientes)
        impimmirdf(df_liberadas)
        impimmirdf(df_atrasadas)

             


        # === 2. CRUCE PARA CASO 3 (LIBERADAS) ===
        # Leemos el excel y aseguramos que NIT sea string para un cruce limpio
        df_proveedores = pd.read_sql_table("Proveedores", engine, schema="PagoArriendos")
        df_proveedores['NIT'] = df_proveedores['NIT'].astype(str).str.strip()
        df_liberadas['Acreedor'] = df_liberadas['Acreedor'].astype(str).str.strip()
       
        #  Si tienes una tabla de proveedores (df_proveedores) con columnas ['NIT', 'Correo', 'Inmueble', 'Contrato']
        # Hacemos un merge para traer la información necesaria
        df_final_L = pd.merge(
            df_liberadas, 
            df_proveedores, 
            left_on='Acreedor', 
            right_on='NIT', 
            how='left'
        )
        impimmirdf(df_final_L)

        


        # --- CASO 1: STATUS 'B' (Correo Único) ---
        if not df_bloqueadas.empty:
            # 1. Preparación de datos para el correo
            destinatario_b = "Steven.navarro@netapplications.com.co" 
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
                            FechaActualizacion = '{fecha_hoy_str}',
                            ContadorEnvio = ContadorEnvio + 1
                        WHERE [Doc.compr.] IN ({ids_para_sql})
                    """)
                    conn.execute(query)
                    conn.commit()
                    
                logger.info(f" Se marcaron {len(lista_ids)} - {lista_ids} registros como 'EnviadoB' en SQL.")

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
                    correos = {
                        'SX': "Steven.navarro@netapplications.com.co",
                        'JG': "Steven.navarro@netapplications.com.co",
                        'RH': "Steven.navarro@netapplications.com.co",
                        'OP': "Steven.navarro@netapplications.com.co"
                    }

                    correo_estrategia = correos.get(estrategia)
                    
                    
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
                                FechaActualizacion = '{fecha_hoy_str}',
                                ContadorEnvio = ContadorEnvio + 1
                            WHERE [Doc.compr.] IN ({ids_para_sql_p})
                        """)
                        conn.execute(query)
                        conn.commit()
                    
                    logger.info(f" SQL: Se marcaron {len(lista_ids_p)} - {lista_ids_p} registros de la estrategia {estrategia} como 'EnviadoP'.")

                except Exception as e:
                    logger.error(f" Error al procesar la estrategia {estrategia}: {str(e)}")


        # --- CASO 3: STATUS 'L' (Liberadas / Notificación a Arrendadores) ---
        if not df_final_L.empty:
            # Generamos la fecha con el formato estándar del Bot
            fecha_hoy_str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            ano_actual = datetime.datetime.now().strftime('%Y')
            logger.debug(f"fecha_hoy_str{fecha_hoy_str}")
            logger.debug(f"ano_actual{ano_actual}")

            # Agrupamos por Acreedor para no saturar al arrendador con múltiples correos
            for acreedor, grupo in df_final_L.groupby("Acreedor"):
                try:
                    # 1. Extracción de datos del Arrendador
                    correo_proveedor = grupo['Correo Proveedor'].iloc[0]
                    nombre_arrendador = grupo['Nombre 1 (Arrendador)'].iloc[0]
                    asunto = f"COLSUBSIDIO - Orden de compra {ano_actual} Arrendamiento y/o Administración"
                    
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
                        <p>A continuación comparto el (los) número(s) de la(s) órden(es) de compra, correspondiente al canon y/o administración para el periodo comprendido de Enero a Diciembre {ano_actual}.</p>
                        <p>Adjunto lineamientos de facturacion electronica... <b>Lo anterior ayudará a que podamos identificar con mayor facilidad su factura...</b></p>
                        
                        <table style='border-collapse: collapse; width: 100%; text-align: center;'>
                            <tr style='background-color: #0056b3; color: white;'>
                                <th>Inmueble</th>
                                <th>No de contrato</th>
                                <th>NIT</th>
                                <th>Arrendador</th>
                                <th>TIPO</th>
                                <th>ORDEN {ano_actual}</th>
                            </tr>
                            {filas_tabla}
                        </table>
                        
                        <p>**Recuerde que las facturas electrónicas deben ser enviadas al correo <a href='mailto:recepcion.facturaelectronica@colsubsidio.com'>recepcion.facturaelectronica@colsubsidio.com</a></p>
                        <p>Cordial saludo,</p>
                        <div style='color: #0056b3;'>
                            <b>Atentamente, Robot RIGO | Administración Inmobiliaria</b><br>
                        </div>
                    </body>
                    </html>
                    """
                    
                    # 3. Envío del Correo
                    logger.info(f" Enviando notificacion de OC Liberada al arrendador: {nombre_arrendador} ({correo_proveedor})")
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
                                FechaActualizacion = '{fecha_hoy_str}',
                                ContadorEnvio = ContadorEnvio + 1
                            WHERE [Doc.compr.] IN ({ids_para_sql_l})
                        """)
                        conn.execute(query)
                        conn.commit()
                    
                    logger.info(f" SQL: Se marcaron {len(lista_ids_l)} - {lista_ids_l}registros del arrendador {acreedor} como 'EnviadoL'.")

                except Exception as e:
                    logger.exception(f" Error al notificar al arrendador {acreedor}: {str(e)}")

        # --- CASO 4: Atrasadas por liberacion STATUS =!'L' (no Liberadas / Notificación a Administración Inmobiliaria ) ---  eri.guzman@colsubsidio.com, FernandoEjemplo@colsubsidio.com
        # --- CASO 4: OC ATRASADAS (> 3 intentos de notificación) ---
        if not df_atrasadas.empty:
            logger.info("Procesando OCs atrasadas para escalamiento...")
    
            # Definimos el mapeo de Creador a Correo (Sugerencia: mover esto a Config_Compras en DB)
            mapeo_administradores = {
                'ERIIGUZV': 'Steven.navarro@netapplications.com.co',
                'FERNCAMS': 'Steven.navarro@netapplications.com.co'
            }

            # Iteramos por cada administrador que tenga OCs atrasadas
            for creador, grupo in df_atrasadas.groupby("Creado"):
                destinatario_admin = mapeo_administradores.get(creador) 
                
                if destinatario_admin:
                    fecha_hoy_str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    asunto_admin = f"ALERTA: OCs con Liberación Crítica - Creador: {creador}"
                    
                    # Construcción del cuerpo con una tabla HTML de las atrasadas
                    cuerpo_admin = f"""
                    <html>
                    <body>
                        <p>Estimado(a), las siguientes Órdenes de Compra han superado los 3 recordatorios 
                        sin ser liberadas. Por favor, verificar su gestión de forma prioritaria:</p>
                        {grupo[['Doc.compr.', 'Fecha doc.', 'Acreedor', 'Nombre 1', 'Estr.', 'Status Lib','ContadorEnvio']].to_html(index=False, border=1)}
                        <p>Este es un correo automático de escalamiento del Bot RIGO.</p>
                    </body>
                    </html>
                    """
                    
                    try:
                        EnviarCorreoPersonalizado(
                            destinatario=destinatario_admin,
                            asunto=asunto_admin,
                            cuerpo=cuerpo_admin,
                            nombreTarea=f"Escalamiento_Atrasadas_{creador}"
                        )
                        
                        # Marcar como 'Escalado' para que no siga sumando intentos infinitamente
                        lista_ids_atrasados = grupo['Doc.compr.'].astype(str).tolist()
                        ids_sql = ", ".join([f"'{id}'" for id in lista_ids_atrasados])
                        
                        with engine.connect() as conn:
                            conn.execute(text(f"""
                                UPDATE [PagoArriendos].[EstrategiasDeLiberacion]
                                SET EstadoNotificacion = 'Escalado',
                                    FechaActualizacion = '{fecha_hoy_str}'
                                WHERE [Doc.compr.] IN ({ids_sql})
                            """))
                            conn.commit()
                        logger.info(f"Escalamiento enviado a {creador} ({destinatario_admin})")
                        
                    except Exception as e:
                        logger.error(f"Error en escalamiento para {creador}: {e}")
                

    
   

