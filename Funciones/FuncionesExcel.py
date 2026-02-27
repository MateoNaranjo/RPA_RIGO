import os
import csv
import re
import unicodedata
import warnings
import pandas as pd

from Config.init_config import in_config
from repositorios.Excel import Excel



class ExcelService:

    @staticmethod
    def limpiar_excel(
        ruta_entrada: str,
        columnas_mapeo: dict,
        hoja: str | int = 0,
        header: int = 0
    ) -> str:
        """
        Limpia un archivo Excel dejando solo las columnas requeridas,
        las renombra y guarda el resultado en un archivo nuevo.

        :param ruta_entrada: Ruta del archivo Excel original
        :param ruta_salida: Ruta donde se guardará el Excel limpio
        :param columnas_mapeo: Diccionario {col_original: col_final}
        :param hoja: Nombre o índice de la hoja (default 0)
        """
        nombre_archivo = os.path.splitext(os.path.basename(ruta_entrada))[0]

        # Leer archivo
        df = pd.read_excel(ruta_entrada, sheet_name=hoja, header=header)

        # Normalizar nombres de columnas
        df.columns = [ExcelService.normalize_column(c) for c in df.columns]

        # Columnas que realmente existen en el Excel
        columnas_existentes = [
            col for col in columnas_mapeo.keys() if col in df.columns
        ]

        # Advertencia si faltan columnas
        columnas_faltantes = set(columnas_mapeo.keys()) - set(columnas_existentes)
        if columnas_faltantes:
            print(f"Advertencia: columnas faltantes -> {columnas_faltantes}")

        # Filtrar columnas necesarias
        df_limpio = df[columnas_existentes]

        # Renombrar columnas
        df_limpio = df_limpio.rename(columns=columnas_mapeo)

        # Guardar archivo limpio
        
        carpeta_temp = in_config("PathTemp")
        ruta_salida = os.path.join(carpeta_temp, f"{nombre_archivo}limpio.xlsx")
        df_limpio.to_excel(ruta_salida, index=False)

        print(f"Archivo limpio generado correctamente en: {ruta_salida}")

        return ruta_salida

    # -----------------------------
    # NORMALIZACIÓN DE NOMBRES
    # -----------------------------
    @staticmethod
    def normalize_column(nombre: str) -> str:
        nombre = nombre.strip().lower()
        nombre = unicodedata.normalize('NFKD', nombre).encode('ascii', 'ignore').decode()
        nombre = re.sub(r'[^\w]', '_', nombre)
        nombre = re.sub(r'_+', '_', nombre)
        return nombre

    # -----------------------------
    # LIMPIEZA DE DATOS
    # -----------------------------
    @staticmethod
    def limpiar_texto(valor):
        if pd.isna(valor):
            return ""
        valor = str(valor)
        valor = unicodedata.normalize("NFKC", valor)
        valor = re.sub(r"[\x00-\x1F\x7F]", " ", valor)
        return valor.replace("\n", " ").replace("\r", " ").strip()

    @staticmethod
    def sanitize_text(value: str) -> str:
        if value is None:
            return "NULL"
        value = unicodedata.normalize("NFKC", str(value))
        value = re.sub(r"[\x00-\x1F\x7F]", " ", value)
        value = value.replace('"', "").strip()
        return value if value else "NULL"

    # -----------------------------
    # OBTENER COLUMNAS DEL EXCEL
    # -----------------------------
    @staticmethod
    def obtener_columnas_excel(ruta_excel: str, header: int = 0) -> list[str]:
        df = pd.read_excel(
            ruta_excel,
            header=header,
            nrows=0,
            engine="openpyxl"
        )
        return [ExcelService.normalize_column(c) for c in df.columns]
    
    
    # -----------------------------
    # EXCEL → CSV
    # -----------------------------
    @staticmethod
    def excel_a_csv(ruta_excel: str, header: int = 0) -> tuple[str, list]:

        warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
        try:
            df = pd.read_excel(
                ruta_excel,
                header=header,
                dtype=str,
                engine="openpyxl"
            )

            df.columns = [ExcelService.normalize_column(c) for c in df.columns]
            orden_columnas = list(df.columns)

            df = df.map(ExcelService.limpiar_texto)

            nombre_base = os.path.splitext(os.path.basename(ruta_excel))[0]
            carpeta_temp = in_config("PathTemp")
            ruta_csv = os.path.join(carpeta_temp, f"{nombre_base}.csv")

            df.to_csv(
                ruta_csv,
                sep=";",
                index=False,
                encoding="utf-8-sig"
            )
            return ruta_csv, orden_columnas
        except Exception as e:
            print(f"Error convirtiendo Excel a CSV: {e}")
            raise

    # -----------------------------
    # CSV → TXT
    # -----------------------------
    @staticmethod
    def convertir_txt(csv_path: str) -> str:

        txt_path = os.path.splitext(csv_path)[0] + ".txt"

        with open(csv_path, "r", encoding="latin1", newline="") as csv_file, \
             open(txt_path, "w", encoding="utf-8", newline="\n") as txt_file:

            reader = csv.reader(csv_file)

            for row in reader:
                cleaned = [ExcelService.sanitize_text(v) for v in row]

                if all(v == "NULL" for v in cleaned):
                    continue

                txt_file.write(";".join(cleaned) + "\n")

        return txt_path

    # -----------------------------
    # ORQUESTADOR FINAL
    # -----------------------------
    @staticmethod
    def ejecutar_bulk_desde_excel(ruta_excel: str, header: int = 0):
        """
        Punto único de entrada:
        - El Excel define columnas
        - El nombre del archivo define la tabla
        """

        nombre_tabla = ExcelService.normalize_column(
            os.path.splitext(os.path.basename(ruta_excel))[0]
        )

        ruta_csv = None
        ruta_txt = None

        try:
            # 1. Columnas (orden exacto del Excel)
            orden_columnas = ExcelService.obtener_columnas_excel(ruta_excel, header)

            # 2. Excel → CSV
            ruta_csv, orden_columnas = ExcelService.excel_a_csv(ruta_excel, header)

            # 3. CSV → TXT
            ruta_txt = ExcelService.convertir_txt(ruta_csv)

            # 4. Bulk con tabla temporal + final
            VariableExcelRepo = Excel(schema="PagoArriendos")
            print(orden_columnas)
            VariableExcelRepo.ejecutar_bulk_dinamico(
                ruta_txt=ruta_txt,
                tabla=nombre_tabla,
                columnas=orden_columnas
            )

        except Exception as e:
            print(f"Error ejecutando bulk desde Excel: {e}")
            raise

        finally:
            for f in (ruta_csv, ruta_txt):
                if f and os.path.exists(f):
                    os.remove(f)
                    
    @staticmethod
    def ejecutarBulkDesdeExcel(rutaExcel: str, header: int = 0, sheet: str | int | None= None):
        """
        Punto único de entrada:
        - El Excel define columnas
        - El nombre del archivo define la tabla
        """
 
        rutaCSV = None
        rutaTXT = None
 
        try:
            resultados = ExcelService.excelACSV(rutaExcel, header, sheet)
 
            for hoja, (rutaCSV, ordenColumnas) in resultados.items():
 
                nombreTabla = ExcelService.normalizacionColumna(str(hoja))
 
                rutaTXT = ExcelService.convertirTXT(rutaCSV)
 
                VariableExcelRepo = Excel(schema="PagoArriendos")
                VariableExcelRepo.ejecutarBulkDinamico(
                    rutaTXT=rutaTXT,
                    tabla=nombreTabla,
                    columnas=ordenColumnas
                )
 
                if os.path.exists(rutaCSV):
                    os.remove(rutaCSV)
 
                if os.path.exists(rutaTXT):
                    os.remove(rutaTXT)
 
        except Exception as e:
            print(f"Error ejecutando bulk desde Excel: {e}")
            raise[]
