import pandas as pd
import os
from datetime import datetime
from config.init_config import in_config

class HU03_DiagnosticoCierre:
    def __init__(self):
        """
        Clase para identificar si el bloqueo del pago es interno (Falta HES)
        o externo (Falta Factura del proveedor).
        """
        self.ruta_reportes = in_config("PathResultado")

    def ejecutar_diagnostico(self, resultado_sap):
        """
        Analiza un diccionario de datos (procedente de SAP o de una fila de Excel).
        """
        estado_fac = "REGISTRADA" if resultado_sap.get("facturada") else "PENDIENTE"
        tiene_hes = resultado_sap.get("tiene_hes", "NO")
        
        # Lógica de diagnóstico
        if estado_fac == "REGISTRADA":
            diagnostico = "PROCESO COMPLETO"
            responsable = "N/A"
            accion = "Ninguna, la factura ya existe en SAP."
        elif tiene_hes == "SÍ":
            diagnostico = "PENDIENTE PROVEEDOR"
            responsable = "PROVEEDOR / CUENTAS POR PAGAR"
            accion = "El servicio ya se recibió (HES), pero el proveedor no ha cobrado o no se ha registrado la FAC."
        else:
            diagnostico = "PENDIENTE GESTOR"
            responsable = "GESTOR INTERNO"
            accion = "No se ha realizado la entrada de servicio (MIGO/HES). El gestor debe cargar el cumplido."

        return {
            "OC": resultado_sap.get("oc_numero"),
            "Estado_FAC": estado_fac,
            "¿Tiene HES?": tiene_hes,
            "Diagnóstico de Cierre": diagnostico,
            "Responsable Acción": responsable,
            "Acción Sugerida": accion,
            "Fecha_Analisis": datetime.now().strftime("%d/%m/%Y")
        }

    def procesar_desde_excel(self, nombre_archivo_hu04):
        """
        Lee el Excel generado por la HU04 y genera el reporte de la HU03.
        """
        ruta_completa = os.path.join(r"\\192.168.50.169\RPA_RIGO_GestionPagodeArrendamientos\Resultados\Reportes_HU04", nombre_archivo_hu04)
        
        if not os.path.exists(ruta_completa):
            print(f"[-] Error: No se encontró el archivo {nombre_archivo_hu04}")
            return None

        # 1. Leer archivo de la HU04
        df_hu04 = pd.read_excel(ruta_completa)
        resultados_hu03 = []

        print(f"[*] Procesando {len(df_hu04)} registros para diagnóstico de cierre...")

        # 2. Iterar y convertir cada fila al formato que entiende el diagnóstico
        for _, fila in df_hu04.iterrows():
            # Mapeo de columnas (asegúrate que coincidan con los nombres de tu HU04)
            datos_preparados = {
                "oc_numero": fila.get('OC'),
                "facturada": True if str(fila.get('Facturada')).upper() == "SÍ" else False,
                "tiene_hes": fila.get('¿Tiene HES/Entrada?', 'NO')
            }
            
            # 3. Llamar al método que causaba el TypeError (ahora con argumento)
            diag = self.ejecutar_diagnostico(datos_preparados)
            resultados_hu03.append(diag)

        # 4. Crear DataFrame final
        df_final = pd.DataFrame(resultados_hu03)
        
        # Guardar resultado de la HU03
        nombre_salida = f"Reporte_HU03_Cierre_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        ruta_salida = os.path.join(self.ruta_reportes, nombre_salida)
        df_final.to_excel(ruta_salida, index=False)
        
        print(f"[+] HU03 Finalizada. Reporte guardado como: {nombre_salida}")
        return df_final

# --- EJEMPLO DE USO EN TU SCRIPT PRINCIPAL ---
# hu03 = HU03_DiagnosticoCierre()
# hu03.procesar_desde_excel("Reporte_Gestion_HU04_20240320.xlsx")