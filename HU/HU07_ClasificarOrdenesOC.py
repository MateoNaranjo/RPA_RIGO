import re
import logging
import pandas as pd
from datetime import datetime
import time
import threading
from Config.settings import SAP_CONFIG
from Config.init_config import in_config
from Funciones.ConexionSAP import ConexionSAP
from Funciones.consultarOC import consultarOC
from Funciones.CargarAnexo import cargar_archivo_gos # Asegúrate de que este archivo exista
from repositorios.Excel import Excel as ExcelDB

class HU07_ClasificarOC:
    def __init__(self):
        """
        Inicializa los componentes de conexión y logging.
        """
        self.logger = logging.getLogger("HU07_ClasificarOC")
        self.sap = ConexionSAP(
            SAP_CONFIG.get('SAP_USUARIO'),
            SAP_CONFIG.get('SAP_PASSWORD'),
            in_config('SapMandante'),
            in_config('SapIdioma'),
            in_config('SapRutaLogon'),
            in_config('SapSistema')
        )
        self.sesion = None
        self.nombreTabla = "BaseMedicamentoslimpio"

    def ejecutar(self):
        # Esta lista almacenará los diccionarios de cada fila para el reporte final
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

            print("\n>>> INICIANDO PROCESAMIENTO DE ÓRDENES...")
            contador = 0
            for registro in registros:
                oc_raw = str(registro.get('Orden2025', ''))
                proveedor = registro.get('Proveedor', 'Sin Proveedor')
                cod_fin = registro.get('CodFin', 'N/A')
                contador =+1
                # Limpieza de OC con Regex
                match = re.search(r'400\d{7}', oc_raw)
                if not match:
                    base_datos_reporte.append({
                        "OC": oc_raw, "Proveedor": proveedor, "Monto": 0,
                        "Estado SAP": "Formato Incorrecto", "Anexo GOS": "N/A"
                    })
                    continue

                oc_numero = match.group(0)
                    
                # 3. Consultar OC y Monto en SAP (FASE 1)
                resultado = consultarOC(self.sesion, oc_numero)
                
                if resultado["status"] == "OK":
                    monto = resultado["monto"]
                    detalle_sap = resultado["detalle"].lower()
                    
                    # Verificar si está liberada
                    es_liberada = any(palabra in detalle_sap for palabra in ["liberada", "active", "concluida"])
                    estado_final = "Liberada" if es_liberada else "Pendiente Liberación"
                    
                    anexo_status = "No corresponde"
                    
                    # 4. Cargar Anexo si está liberada (FASE 2)
                    if es_liberada:
                        ruta_pdf = f"\\\\192.168.50.169\\RPA_RIGO_GestionPagodeArrendamientos\\Insumos\\Anexos\\Prueba.txt"
                        # Llamada a la función que maneja el hilo de Windows
                        exito_carga = cargar_archivo_gos(self.sesion, oc_numero, ruta_pdf, self.logger)
                        anexo_status = "Cargado Exitosamente" if exito_carga else "Error en Carga"
                        
                        time.sleep(1)

                        self.sesion.findById("wnd[0]/tbar[0]/okcd").text = "/n"
                        self.sesion.findById("wnd[0]").sendVKey(0)
                    
                    # Guardar información para el reporte
                    base_datos_reporte.append({
                        "OC": oc_numero,
                        "Proveedor": proveedor,
                        "Monto": monto,
                        "Estado SAP": estado_final,
                        "Anexo GOS": anexo_status
                    })
                    print(f"[*] Procesada OC {oc_numero} - {estado_final}")

                else:
                    # Si la OC no existe o hubo error de red
                    base_datos_reporte.append({
                        "OC": oc_numero, "Proveedor": proveedor, "Monto": 0,
                        "Estado SAP": "No existe / Error", "Anexo GOS": "N/A"
                    })
                    print(f"[-] OC {oc_numero} no encontrada.")

            # 5. Generar Reporte Final (FASE 3)
            self.generar_reporte_excel(base_datos_reporte)

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

        # Crear DataFrame
        df = pd.DataFrame(lista_datos)

        # Clasificación por montos aprobados (Lógica de negocio)
        def clasificar_monto(m):
            if m > 10000000: return "Monto Alto (>10M)"
            if m > 1000000: return "Monto Medio (1M-10M)"
            return "Monto Bajo"
        
        df['Clasificación Monto'] = df['Monto'].apply(clasificar_monto)

        # Configurar nombre del archivo
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        ruta_reporte = f"\\\\192.168.50.169\\RPA_RIGO_GestionPagodeArrendamientos\\Resultados\\Reporte_Gestion_HU07_{timestamp}.xlsx"

        try:
            # Guardar Excel
            df.to_excel(ruta_reporte, index=False)
            print(f"\n" + "="*50)
            print(f"REPORTE GENERADO: {ruta_reporte}")
            
            # Resumen por proveedor en consola
            resumen = df.groupby(['Proveedor', 'Estado SAP']).size()
            print("\nResumen por Proveedor y Estado:")
            print(resumen)
            print("="*50)
            
        except Exception as e:
            print(f"Error al guardar el Excel: {e}")