# ===============================
# HU06: Validacion de presupuesto
# Autor: Santiago Pinzon - Desarrollador RPA
# Descripcion: Descripcion de la HU 
# Ultima modificacion: 2/1/2026
# Propiedad de Colsubsidio
# Cambios: Si aplica
# ===============================
import os
import time
from Funciones.ControlHU import control_hu
from Funciones.EscribirLog import WriteLog
from Funciones.ConexionSAP import ConexionSAP
from Config.Settings import SAP_CONFIG
from Config.init_config import in_config
from Funciones.Excel import ExcelService
from Repositorios.Excel import ExcelRepo
from Funciones.ME80FN import ME80FN
from datetime import datetime, date
import pandas as pd
import warnings


def HU01_Prueba():
    """
    Docstring for HU01_Prueba
    """
    # =========================
    # CONFIGURACION DEL PROCESO
    # =========================
    task_name = "HU01_Prueba"

    try:
        # === Inicio HU01 ===
        control_hu(task_name, 0)
        warnings.filterwarnings("ignore",category=UserWarning, module="openpyxl")
        # GestionTicketInsumo(estado, id, maquina, observaciones)
        # WriteLog(mensaje="Inicio HU01", estado="INFO", task_name=task_name)
        
        # ============================= Inicio acciones =============================
        sap = ConexionSAP(SAP_CONFIG.get('SAP_USUARIO'),
                        SAP_CONFIG.get('SAP_PASSWORD'),
                        in_config('SAP_CLIENTE'),
                        in_config('SAP_IDIOMA'),
                        in_config('SAP_PATH'),
                        in_config('SAP_SISTEMA')
                    )
        sap.iniciar_sesion_sap()
        ruta_insumo = in_config("PathInsumos")+"\BaseMedicamentos.xlsx"

        try:
            TablaBase = ExcelRepo.obtener_valores("BaseMedicamentosLimpio")
        except:
            
            columnas_medicamentos = {
                "cod_fin": "cod_fin",
                "nit": "nit",
                "orden_2025": "orden_2025",
                "mts2_segun_contrato": "mts2",
                "iva": "iva",
                "tipo": "tipo",
                "enero": "enero",
                "febrero": "febrero",
                "marzo": "marzo",
                "abril": "abril",
                "mayo": "mayo",
                "junio": "junio",
                "julio": "julio",
                "agosto": "agosto",
                "septiembre": "septiembre",
                "actubre": "octubre",
                "noviembre": "noviembre",
                "diciembre": "diciembre",
                "observacion_de_pagos": "observaciones",
                "no_de_contrato": "numero_contrato",
                "nombre_facturador": "nombre_facturador"
            }
            
            ruta_excel = ExcelService.limpiar_excel(ruta_insumo, columnas_medicamentos, header=3)
            ExcelService.ejecutar_bulk_desde_excel(ruta_excel)
            os.remove(ruta_excel)

            TablaBase = ExcelRepo.obtener_valores("BaseMedicamentosLimpio")


        reporte_validacion = []
        for registro in TablaBase[28:]:

            ruta_cabecera= in_config("PathTemp")+"\cabecera.xlsx"
            ruta_repartos= in_config("PathTemp")+"\Reparto.xlsx"
            ruta_tabla_final=in_config("PathTemp")+"\ValoresAComparar.xlsx"

            sap.MenuPrincipal()
            sap.abrir_transaccion("ME80FN")
            me80fn = ME80FN(sap)
            me80fn.ingresar_oc(registro["orden_2025"])
            time.sleep(2)
            me80fn.exportar_tabla(ruta_cabecera, "cabecera")
            time.sleep(3)
            os.system("taskkill /f /im excel.exe")
            time.sleep(5)
            me80fn.entrar_repartos()
            me80fn.exportar_tabla(ruta_repartos, "repartos")
            time.sleep(3)
            os.system("taskkill /f /im excel.exe")
            time.sleep(5)
            # Operaciones con los excel
            try:
                cabecera = pd.read_excel(ruta_cabecera, header=None)
                reparto = pd.read_excel(ruta_repartos, header=None)

                cabecera= cabecera.dropna(how="all").reset_index(drop=True)
                reparto= reparto.dropna(how="all").reset_index(drop=True)

                cabecera = cabecera.rename(columns={
                    0: "OC",
                    1: "posicion",
                    2: "material",
                    3: "descripcion",
                    10: "proveedor",
                    12: "grupo_de_Compras",
                    14: "valor_neto"
                })

                reparto = reparto.rename(columns={
                    4: "posicion",
                    9: "fecha_entrega"
                })

                cabecera["posicion"] = cabecera["posicion"].astype(str)
                reparto["posicion"] = reparto["posicion"].astype(str)

                cabecera["fecha_entrega"] = cabecera["posicion"].map(
                    reparto.set_index("posicion")["fecha_entrega"]
                )

                cabecera = cabecera[
                        [
                            "OC",
                            "posicion",
                            "material",
                            "descripcion",
                            "proveedor",
                            "grupo_de_Compras",
                            "valor_neto",
                            "fecha_entrega"
                        ]
                    ]

                cabecera.to_excel(ruta_tabla_final,index=False)
                print("Cruce completado")

                # Se ejecuta el bulk en la base de datos
                ExcelService.ejecutar_bulk_desde_excel(ruta_tabla_final)
                
            except Exception as e:
                print("Error en combinar columnas", e)

            

            DatosME80FN = ExcelRepo.obtener_valores("valoresacomparar")

            MAPA_MESES = {
                1: "enero",
                2: "febrero",
                3: "marzo",
                4: "abril",
                5: "mayo",
                6: "junio",
                7: "julio",
                8: "agosto",
                9: "septiembre",
                10: "octubre",
                11: "noviembre",
                12: "diciembre",
            }

            for d in DatosME80FN:
                fecha = datetime.strptime(d["fecha_entrega"], "%Y-%m-%d %H:%M:%S")
                valor_sap = float(d["valor_neto"])
                fecha_actual = date.today()
                mes_actual = fecha_actual.month
                

                if not fecha:
                    continue  

                mes = fecha.month
                columna_mes = MAPA_MESES.get(mes)

                if not columna_mes:
                    continue

                valor_excel = registro.get(columna_mes)

                if valor_excel is None:
                    print(
                        f"OC {registro['orden_2025']} | "
                        f"Mes {columna_mes.upper()} no existe en Excel"
                    )
                    continue

                valor_excel = float(valor_excel)

                if valor_sap == valor_excel:
                    estado = "OK"
                else:
                    estado = "DIFERENCIA"

                if mes == mes_actual:
                    vmt2 = float(registro['mts2'])   
                    vumt2 = valor_excel / vmt2   
                    print(
                        f"OC {registro['orden_2025']} | "
                        f"Fecha: {fecha.date()} | "
                        f"Mes: {columna_mes} | "
                        f"SAP: {valor_sap:,.0f} | "
                        f"Excel: {valor_excel:,.0f} | "
                        f"Resultado: {estado} | "
                        f"Valor Unitario MT2: {vumt2:,.0f} | " 
                    )
                    reporte_validacion.append({
                        "OC": registro["orden_2025"],
                        "Fecha_entrega": fecha.date(),
                        "Mes": columna_mes,
                        "Valor_SAP": valor_sap,
                        "Valor_Excel": valor_excel,
                        "Resultado": estado
                    })

                    ruta_reporte = in_config("PathTemp") + rf"\Reporte_Validacion_Presupuesto_{columna_mes}.xlsx"

            # Finaliza proceso de operaciones
            if os.path.exists(ruta_cabecera) and os.path.exists(ruta_repartos) and os.path.exists(ruta_tabla_final) :              
                os.remove(ruta_cabecera)
                os.remove(ruta_repartos)
                os.remove(ruta_tabla_final)
                print("Archivos temporales eliminados")

            print("Finalizacion del proceso el registro con oc:", registro["orden_2025"])

            for i in range(2):
                sap.MenuPrincipal()

        df_reporte = pd.DataFrame(reporte_validacion)

        df_reporte.to_excel(
                ruta_reporte,
                index=False,
                sheet_name="Validacion"
            )
        # ============================= Finalizacion HU =============================

        control_hu(task_name, 100)

        Estado = 100
        return Estado
    
    except Exception as e:
        print(f"Error en ejecucion: ({e}) ")
        # WriteLog()
        # GestionTicketInsumo(id, observaciones, estado, maquina)
        control_hu(task_name, 99)
        # Estado = 99
        # return Estado
        raise

    finally:
        # WriteLog()
        log = "Finalizacion HU"
        print(log)