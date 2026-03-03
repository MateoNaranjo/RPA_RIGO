import re
import logging
import pandas as pd
from datetime import datetime
from pathlib import Path
import time
from Config.Settings import SAP_CONFIG
from Config.init_config import in_config
from Config.Database import Database
from Funciones.ConexionSAP import ConexionSAP
from Funciones.consultarOC import consultarOC
from Funciones.CargarAnexo import cargar_archivo_gos
from Repositorios.Excel import Excel as ExcelDB
from sqlalchemy import types

class HU07_ClasificarOC:
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
        
        self.rutaTemporal=in_config('PathTemp')
        self.carpeta=self.rutaTemporal+"\\HU07"
        self.rutaHU07=Path(self.carpeta)
        self.sesion = None
        self.nombreTabla = "BaseMedicamentoslimpio"

    def ejecutar(self):
        base_datos_reporte = []

        try:
            # 1. Obtener datos de la base (Excel de entrada)
            self.logger.info(f"Leyendo registros de {self.nombreTabla}...")
            registros = ExcelDB.obtener_datos_por_posicion(self.nombreTabla)
            if not registros:
                self.logger.warning("No se encontraron registros para procesar.")
                return

            # 2. Iniciar Sesión en SAP
            self.sesion = self.sap.iniciar_sesion_sap()
            self.sap.abrir_transaccion("ME23N")

            print("\n>>> INICIANDO PROCESAMIENTO DE ÓRDENES...")
            contador = 0

            for registro in registros:
                try:
                    oc_raw = str(registro.get('orden_2025', ''))
                    proveedor = registro.get('nombre_facturador', 'Sin Proveedor')
                    cod_fin = registro.get('cod_fin', 'N/A')
                    contador += 1

                    # Limpieza de OC con Regex
                    match = re.search(r'400\d{7}', oc_raw)
                    if not match:
                        base_datos_reporte.append({
                            "OC": oc_raw, "Proveedor": proveedor, "Monto": 0,
                            "Estado SAP": "Formato Incorrecto", "Anexo GOS": "N/A"
                        })
                        continue

                    oc_numero = match.group(0)

                    # 3. Consultar OC y Monto en SAP
                    resultado = consultarOC(self.sesion, oc_numero)

                    if resultado["status"] == "OK":
                        monto = resultado["monto"]
                        detalle_sap = resultado["detalle"].lower()

                        es_liberada = any(palabra in detalle_sap for palabra in ["liberada", "active", "concluida"])
                        estado_final = "Liberada" if es_liberada else "Pendiente Liberación"

                        anexo_status = "No corresponde"

                        # 4. Cargar Anexo si está liberada
                        if es_liberada:
                            ruta_pdf = r"\\192.168.50.169\RPA_RIGO_GestionPagodeArrendamientos\Insumo\Anexos\Prueba.txt"
                            exito_carga = cargar_archivo_gos(self.sesion, oc_numero, ruta_pdf, self.logger)
                            anexo_status = "Cargado Exitosamente" if exito_carga else "Error en Carga"

                            time.sleep(1)
                            self.sesion.findById("wnd[0]/tbar[0]/okcd").text = "/n"
                            self.sesion.findById("wnd[0]").sendVKey(0)
                        nit = registro.get('nit', 'N/A')
                        base_datos_reporte.append({
                            "OC": oc_numero,
                            "Proveedor": proveedor,
                            "Monto": monto,
                            "EstadoSAP": estado_final,
                            "Anexo": anexo_status,
                            "NIT": nit
                        })
                        print(f"[*] Procesada OC {oc_numero} - {estado_final}")

                    else:
                        base_datos_reporte.append({
                            "OC": oc_numero, "Proveedor": proveedor, "Monto": 0,
                            "EstadoSAP": "No existe / Error", "Anexo": "N/A", "NIT": nit
                        })
                        print(f"[-] OC {oc_numero} no encontrada.")

                except Exception as e:
                    self.logger.error(f"Error procesando registro {registro}: {e}")
                    print(f"Error procesando registro {registro}: {e}")

            # 5. Generar Reporte Final
            self.generar_reporte_excel(base_datos_reporte)
            
            timestamp = datetime.now().strftime("%Y%d%m")
            db= Database()
            engine = db.get_engine()
            df = pd.read_excel(f"{self.rutaTemporal}"+f"\HU07\Reporte_HU07{timestamp}.xlsx")

            dydtype = {
                'Oc': types.VARCHAR(20),
                'Proveedor': types.VARCHAR(100),
                'Monto': types.VARCHAR(100),
                'EstadoSAP': types.VARCHAR(20),
                'Anexo': types.VARCHAR(20),
                'ClasificacionMonto': types.VARCHAR(20),
                'NIT': types.VARCHAR(20)
              
            }

            df.to_sql(f"ReporteHU07", con=engine, if_exists='replace', index=False, schema='PagoArriendos', dtype=dydtype)

            #self.ejecutar_cargue_desde_excel(f"{self.rutaTemporal}"+f"\HU07\Reporte_HU07{timestamp}.xlsx")

        except Exception as e:
            self.logger.error(f"Falla crítica en HU07: {e}")
            print(f"Falla crítica: {e}")

    def generar_reporte_excel(self, lista_datos):
        """
        Crea un archivo Excel consolidado y lo guarda en la carpeta de Reportes.
        """
        if not lista_datos:
            print("No hay datos para generar el reporte.")
            return

        df = pd.DataFrame(lista_datos)

        def clasificar_monto(m):
            if m > 10000000: return "Monto Alto (>10M)"
            if m > 1000000: return "Monto Medio (1M-10M)"
            return "Monto Bajo"
        
        df['Clasificación Monto'] = df['Monto'].apply(clasificar_monto)
        df = df.astype(str)
        
        if not self.rutaHU07.exists():
            self.rutaHU07.mkdir(parents=True, exist_ok=True)
            print('carpeta {self.carpeta} creada con exito')
        else:
            print('la carpeta {self.carpeta} ya existe')
        timestamp = datetime.now().strftime("%Y%d%m")
        ruta_reporte = f"{self.rutaTemporal}"+f"\HU07\Reporte_HU07{timestamp}.xlsx"

        try:
            df.to_excel(ruta_reporte, index=False)
            print("\n" + "="*50)
            print(f"REPORTE GENERADO: {ruta_reporte}")
            
            resumen = df.groupby(['Proveedor', 'Estado SAP']).size()
            print("\nResumen por Proveedor y Estado:")
            print(resumen)
            print("="*50)
            
        except Exception as e:
            print(f"Error al guardar el Excel: {e}")
    

    def crear_tabla_HU07(self):
        db= Database()
        tabla="PagoArriendos.ReporteHU07"

        query=f"""
        IF OBJECT_ID('{tabla}', 'U') IS NOT NULL
            DROP TABLE {tabla};

        CREATE TABLE {tabla} (
            Oc VARCHAR(20) PRIMARY KEY,
            Proveedor VARCHAR(100),
            Monto VARCHAR(100),
            EstadoSAP VARCHAR(20),
            Anexo VARCHAR(20),
            ClasificacionMonto VARCHAR(20),
            NIT VARCHAR(20)

            )
        """

        try:
            with db.get_connection(self) as conn:
                cursor = conn.cursor()
                cursor.execute(query)
                conn.commit()
                print(f"[+] Tabla {tabla} configuraicon con los titulos de la HU07")
                return True
        except Exception as e:
            print(f"[-] Error creando tabla: {e}")
            return False



    def ejecutar_cargue_desde_excel(self,ruta_excel):
        db= Database()
        tabla="PagoArriendos.ReporteHU07"
        if not HU07_ClasificarOC.crear_tabla_HU07():
            return
        
        try:
            df=pd.read_excel(ruta_excel)

            query_insert = f"""
                INSERT INTO {tabla} (
                Oc, Proveedor, Monto, EstadoSAP, Anexo, ClasificacionMonto, NIT) VALUES (?,?,?,?,?,?,?)
            """

            with db.get_connection(self) as conn:
                cursor =conn.cursor()

                for _, fila in df.iterrows():

                    valores=(
                        str(fila.get('OC')),
                        str(fila.get('Proveedor')),
                        str(fila.get('Monto')),
                        str(fila.get('Estado SAP')),
                        str(fila.get('Anexo GOS')),
                        str(fila.get('Clasificación Monto')),
                        str(fila.get('NIT'))
                    )

                    cursor.execute(query_insert, valores)

                conn.commit()
                print(f"[+] Éxito: Se cargaron {len(df)} registros a SQL Server.")

        except Exception as e:
            print(f"[-] Error en el cargue: {e}")

    
