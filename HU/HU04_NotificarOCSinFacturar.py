import pandas as pd
import glob
import time
import os
from datetime import datetime
from funciones.DatosHU04 import consultar_datos_hu04
from funciones.ConexionSAP import ConexionSAP
from config.settings import SAP_CONFIG
from config.init_config import in_config

class HU04_Auditoria:
    def __init__(self):
        self.sap = ConexionSAP(
            SAP_CONFIG.get('SAP_USUARIO'),
            SAP_CONFIG.get('SAP_PASSWORD'),
            in_config('SapMandante'),
            in_config('SapIdioma'),
            in_config('SapRutaLogon'),
            in_config('SapSistema')
        )
        self.sesion = None
        self.ruta_input = r"\\192.168.50.169\RPA_RIGO_GestionPagodeArrendamientos\Resultados"
        self.ruta_output = r"\\192.168.50.169\RPA_RIGO_GestionPagodeArrendamientos\Resultados\Reportes_HU04"

    def buscar_ultimo_reporte_hu07(self):
        """Busca el archivo más reciente de la HU07 usando comodines."""
        patron = os.path.join(self.ruta_input, "Reporte_Gestion_HU07_*.xlsx")
        archivos = glob.glob(patron)
        return max(archivos, key=os.path.getctime) if archivos else None

    def ejecutar(self):
        print(">>> Conectando a SAP para Auditoría HU04...")
        self.sesion = self.sap.iniciar_sesion_sap()
        
        if not self.sesion:
            print("[-] Error crítico: No se pudo iniciar sesión.")
            return

        time.sleep(2)

        # Buscar el reporte de insumo
        archivo_hu07 = self.buscar_ultimo_reporte_hu07()
        if not archivo_hu07:
            print("[-] No se encontró reporte HU07 en la ruta de red.")
            return

        print(f">>> Leyendo reporte: {os.path.basename(archivo_hu07)}")
        df_hu07 = pd.read_excel(archivo_hu07)

        # Filtrar registros que necesitan auditoría
        ocs_para_auditar = df_hu07[df_hu07['Estado SAP'].isin(['Liberada', 'Pendiente Liberación'])]
        
        if ocs_para_auditar.empty:
            print("[!] No hay órdenes para procesar.")
            return

        resultados_auditoria = []
        print(f"\n>>> Procesando {len(ocs_para_auditar)} órdenes...")
        
        for _, fila in ocs_para_auditar.iterrows():
            oc = str(fila['OC'])
            print(f"[*] Consultando OC {oc}...")

            res = consultar_datos_hu04(self.sesion, oc)
            
            if res["status"] == "OK":
                # Lógica: 2+ días sin factura es crítico
                es_critico = res["dias"] >= 2 and not res["facturada"]
                
                resultados_auditoria.append({
                    "OC": oc,
                    "Proveedor": fila['Proveedor'],
                    "Monto": fila['Monto'],
                    "Fecha Creación SAP": res["fecha_sap"],
                    "Antigüedad (Días)": res["dias"],
                    "Facturada": "SÍ" if res["facturada"] else "NO",
                    "Requiere Acción": "SÍ" if es_critico else "NO"
                })
            else:
                print(f"    [!] Error en OC {oc}: {res['detalle']}")

        self.guardar_informe(resultados_auditoria)

    def guardar_informe(self, datos):
        if not datos: return
        df_final = pd.DataFrame(datos)
        
        if not os.path.exists(self.ruta_output):
            os.makedirs(self.ruta_output)
            
        fecha_hora = datetime.now().strftime('%Y%m%d_%H%M')
        nombre = f"Informe_Auditoria_Facturacion_{fecha_hora}.xlsx"
        ruta_completa = os.path.join(self.ruta_output, nombre)
        
        df_final.to_excel(ruta_completa, index=False)
        print(f"\n[!] PROCESO FINALIZADO. Reporte en: {nombre}")

# --- PUNTO DE ENTRADA ---
if __name__ == "__main__":
    bot = HU04_Auditoria()
    bot.ejecutar()