import pandas as pd
import glob
import os
import time
from datetime import datetime
from funciones.ConexionSAP import ConexionSAP
from funciones.DatosHU04 import consultar_datos_hu04 # Reutilizamos la conexión
from config.settings import SAP_CONFIG
from config.init_config import in_config

class HU02_VerificacionDiaria:
    def __init__(self):
        self.sap = ConexionSAP(
            SAP_CONFIG.get('SAP_USUARIO'),
            SAP_CONFIG.get('SAP_PASSWORD'),
            in_config('SapMandante'),
            in_config('SapIdioma'),
            in_config('SapRutaLogon'),
            in_config('SapSistema')
        )
        self.ruta_input = r"\\192.168.50.169\RPA_RIGO_GestionPagodeArrendamientos\Resultados"
        self.ruta_output = r"\\192.168.50.169\RPA_RIGO_GestionPagodeArrendamientos\Resultados\Reportes_HU02"

    def ejecutar(self):
        print(">>> Iniciando HU02: Verificación Diaria de Facturación...")
        sesion = self.sap.iniciar_sesion_sap()
        if not sesion: return

        # 1. Buscar último reporte HU07
        archivos = glob.glob(os.path.join(self.ruta_input, "Reporte_Gestion_HU07_*.xlsx"))
        if not archivos:
            print("[-] No se encontró insumo HU07.")
            return
        
        archivo_reciente = max(archivos, key=os.path.getctime)
        df = pd.read_excel(archivo_reciente)

        # 2. Filtrar solo las que NO tienen factura (Estado FAC vacío o pendiente)
        # Ajusta este filtro según como venga tu columna de factura en el Excel
        ocs_pendientes = df[df['Estado SAP'].isin(['Liberada', 'Pendiente Liberación'])]

        resultados = []

        for _, fila in ocs_pendientes.iterrows():
            oc = str(fila['OC'])
            print(f"[*] Analizando historial de OC {oc}...")
            
            # Llamamos a la lógica de extracción
            res = consultar_datos_hu04(sesion, oc)
            
            if res["status"] == "OK":
                # Lógica de diagnóstico HU02
                diagnostico = ""
                responsable = ""
                
                if not res["facturada"]:
                    if res.get("tiene_hes") == "SÍ":
                        diagnostico = "Servicio recibido, proveedor no ha cobrado."
                        responsable = "PROVEEDOR"
                    else:
                        diagnostico = "Servicio no recibido en sistema (Falta HES)."
                        responsable = "GESTOR INTERNO"
                else:
                    diagnostico = "Factura ya registrada."
                    responsable = "N/A"

                resultados.append({
                    "OC": oc,
                    "Proveedor": fila['Proveedor'],
                    "Dias desde Creación": res["dias"],
                    "¿Tiene HES/Entrada?": res.get("tiene_hes", "NO"),
                    "¿Tiene Factura?": "SÍ" if res["facturada"] else "NO",
                    "Diagnóstico": diagnostico,
                    "Responsable Acción": responsable
                })

        self.guardar_reporte(resultados)

    def guardar_reporte(self, datos):
        if not datos: return
        df_final = pd.DataFrame(datos)
        os.makedirs(self.ruta_output, exist_ok=True)
        nombre = f"HU02_Control_Facturacion_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        df_final.to_excel(os.path.join(self.ruta_output, nombre), index=False)
        print(f">>> HU02 Finalizada. Reporte generado: {nombre}")

