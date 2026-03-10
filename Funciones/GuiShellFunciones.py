# ============================================
# Función Local: validacionME53N
# Autor: Paula Sierra Steven Navarro- NetApplications
# Descripcion: Ejecuta ME5A y exporta archivo TXT según estado.
# Ultima modificacion: 24/11/2025
# Propiedad de Colsubsidio
# Cambios: Ajustado Funciones para Arriendos 
# ============================================
from requests import session
import win32com.client
import traceback
import pandas as pd
import re
import subprocess
import time
import os
import pyperclip
from Funciones.EmailSender import EnviarCorreoPersonalizado
from Funciones.EscribirLog import WriteLog
import pyautogui
from pyautogui import ImageNotFoundException
from Funciones.Login import ObtenerSesionActiva
from typing import List, Literal, Optional
from Config.init_config import in_config
from Config.Database import Database
import logging
logger = logging.getLogger(__name__)
import datetime

import calendar



def obtener_correos(texto: str, dominio: Optional[str] = None) -> List[str]:
    """
    Obtiene correos electrónicos desde un texto.
    - Si se especifica dominio, filtra solo los correos que pertenezcan a ese dominio.
    - Si no se especifica dominio, retorna todos los correos encontrados.

    Args:
        texto (str): Texto multilínea donde buscar.
        dominio (Optional[str]): Dominio a filtrar (ej: '@gmail', '@gmail.com').

    Returns:
        List[str]: Lista de correos encontrados.
    """

    # Patrón general para correos
    patron_general = re.compile(
        r"\b[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}\b",
        re.IGNORECASE
    )

    correos = patron_general.findall(texto)

    if dominio:
        dominio = dominio.lower()

        # Normaliza dominio: asegura que empiece con '@'
        if not dominio.startswith("@"):
            dominio = "@" + dominio

        correos = [
            correo for correo in correos
            if correo.lower().endswith(dominio)
        ]

    return correos

def obtener_valor(texto: str, contiene: List[str]) -> Optional[str]:
    """
    Busca un valor numérico en una línea que contenga
    alguna de las palabras clave especificadas, con o sin símbolo $.

    Args:
        texto (str): Texto multilínea donde buscar.
        contiene (List[str]): Palabras clave a buscar en la línea.

    Returns:
        Optional[str]: Valor numérico encontrado (como string) o None.
    """

    # Patrón: opcional $, números con separadores de miles
    patron = re.compile(r"(?:\$?\s*)(\d{1,3}(?:[.,]\d{3})*|\d+)")

    contiene_upper = [c.upper() for c in contiene]

    for linea in texto.splitlines():
        linea_upper = linea.upper()

        if any(c in linea_upper for c in contiene_upper):
            match = patron.search(linea)
            if match:
                # Normalizar valor (quita separadores)
                valor = match.group(1).replace(".", "").replace(",", "")
                return valor

    return None

def leer_solpeds_desde_archivo(ruta_archivo):
    """
    Lee un archivo de texto plano con formato de tabla (| separado) y extrae
    información de Solicitudes de Pedido (SOLPEDs), agrupando por número de SOLPED.

    Args:
        ruta_archivo (str): La ruta completa al archivo de texto a leer.

    Returns:
        dict: Un diccionario donde cada clave es un número de SOLPED y el valor
              es otro diccionario con el conteo de 'items' y un 'set' de 'estados'.
              Ej: {'10023456': {'items': 3, 'estados': {'Estado A', 'Estado B'}}}
    """
    resultados = {}

    with open(ruta_archivo, "r", encoding="utf-8", errors="ignore") as f:
        for linea in f:
            # Todas las líneas útiles empiezan con '|'
            if not linea.strip().startswith("|"):
                continue

            partes = [p.strip() for p in linea.split("|")]

            # Esperamos al menos 16 columnas por la estructura del archivo
            if len(partes) < 16:
                continue

            try:
                purch_req = partes[1]       # PurchReq
                estado    = partes[15]      # Estado
            except IndexError:
                continue  # Si alguna fila viene corrupta

            # Evitar encabezados
            if purch_req.lower().startswith("purch"):
                continue

            # Inicializar
            if purch_req not in resultados:
                resultados[purch_req] = {
                    "items": 0,
                    "estados": set()
                }

            # Sumar item
            resultados[purch_req]["items"] += 1
            resultados[purch_req]["estados"].add(estado)

    return resultados

def obtener_numero_oc(session):
    """
    Obtiene el número de la Orden de Compra creada desde la barra de estado.
    """
    try:
        # El mensaje de exito con el número de OC suele aparecer en la barra de estado.
        status_text = session.findById("wnd[0]/sbar").text
        # Usamos una expresión regular para buscar un número que sigue a un texto estándar.
        # "Standard PO created under the number 4500021244" -> Ejemplo
        match = re.search(r'(\d{10,})', status_text) # Busca 10 o más dígitos
        if match:
            numero_oc = match.group(1)
            print(f"Número de OC extraído: {numero_oc}")
            return numero_oc
        else:
            print("No se pudo encontrar el numero de OC en la barra de estado.")
            return None
    except Exception as e:
        print(f"Error al obtener el número de OC: {e}")
        return None

def esperar_sap_listo(session, timeout=10):
    """
    Espera hasta que la sesión de SAP GUI no esté ocupada (session.Busy es False).

    Args:
        session: La sesión activa de SAP GUI.
        timeout (int): Tiempo máximo de espera en segundos.

    Raises:
        TimeoutError: Si SAP sigue ocupado después del tiempo de espera.
    """
    inicio = time.time()

    while time.time() - inicio < timeout:
        try:
            if not session.Busy:
                return True
        except:
            pass
        time.sleep(0.2)

    raise TimeoutError("SAP GUI no terminó de cargar (session.Busy)")


    """
    Cambia el Grupo de Compras ('EKGRP') basado en la Organización de Compras ('EKORG') actual.

    Args:
        session: La sesión activa de SAP GUI.

    Returns:
        list: Una lista de strings con las acciones realizadas.

    Raises:
        ValueError: Si la Organización de Compras actual no está en el mapa de condiciones.
    """
    # Obtener el valor actual de la organización de compra
    obj_orgCompra = get_GuiCabeceraTextField_text(session, "EKORG")
    if not obj_orgCompra:
        obj_orgCompra = obj_orgCompra.upper()

    #print(f"Valor de OrgCompra: {obj_orgCompra}")
    condiciones = {
        "s":"RCC",
        "S":"RCC",
        "":"RCC",
        "OC15": "RCC",
        "OC26": "HAB",
        "OC25": "HAB",
        "OC28": "AC2",
        "OC27": "AC2"
    }

    if obj_orgCompra not in condiciones:
        raise ValueError(f"Organización de compra '{obj_orgCompra}' no reconocida.")

    obj_grupoCompra = condiciones[obj_orgCompra]

    set_GuiCabeceraTextField_text(session, "EKGRP", obj_grupoCompra)
    #print(f"Grupo de compra actualizado a: {obj_grupoCompra}")
    acciones = []
    acciones.append(f"Valor de OrgCompra: {obj_orgCompra}")
    acciones.append(f"Grupo de compra actualizado a: {obj_grupoCompra}")
    return acciones

def normalizar_precio_sap(precio: str) -> int:
    """
    Convierte un precio SAP tipo '2.750.000,00' en entero 2750000
    para comparaciones confiables.
    """
    if not precio:
        return 0

    # Quitar separador de miles y decimales
    limpio = precio.replace(".", "").replace(",00", "")

    return int(limpio)

def MostrarCabecera():
    """
    Asegura que las secciones principales de la interfaz (Cabecera, Resumen, Detalle)
    estén visibles en la transacción ME21N para prevenir errores de "objeto no encontrado".
    """
    session = ObtenerSesionActiva()
    #time.sleep(0.2)
    esperar_sap_listo(session)
    pyautogui.hotkey("ctrl","F2")
    esperar_sap_listo(session)
    #time.sleep(0.2)
    pyautogui.hotkey("ctrl","F3")
    esperar_sap_listo(session)
    #time.sleep(0.5)
    pyautogui.hotkey("ctrl","F4")
    esperar_sap_listo(session)
    #time.sleep(0.5)
    pyautogui.hotkey("ctrl","F8")
    esperar_sap_listo(session)

def ProcesarTabla(name, dias=None):
    """name: nombre del txt a utilizar
    return data frame
    Procesa txt estructura ME5A y devuelve un df con manejo de columnas dinamico.
    dias: int|None -> número de días a mantener (si None, no aplica filtro por fecha)"""

    try:
        WriteLog(
            mensaje=f"Procesar archivo nombre {name}",
            estado="INFO",
            nombreTarea="procesarTablaME5A",)
  
        # path = f".\\AutomatizacionGestionSolped\\Insumo\\{name}"
        path = rf"{in_config('PathInsumos')}\{name}"

        # INTENTAR LEER CON DIFERENTES CODIFICACIONES
        lineas = []
        codificaciones = ["latin-1", "cp1252", "iso-8859-1", "utf-8"]

        for codificacion in codificaciones:
            try:
                with open(path, "r", encoding=codificacion) as f:
                    lineas = f.readlines()
                #print(f"EXITO: Archivo leido con codificacion {codificacion}")
                break
            except UnicodeDecodeError as e:
                print(f"ERROR con {codificacion}: {e}")
                continue
            except Exception as e:
                print(f"ERROR con {codificacion}: {e}")
                continue

        if not lineas:
            print("ERROR: No se pudo leer el archivo con ninguna codificacion")
            return pd.DataFrame()

        # Filtrar solo lineas de datos
        filas = [l for l in lineas if l.startswith("|") and not l.startswith("|---")]

        # DETECTAR ESTRUCTURA DE COLUMNAS DINAMICAMENTE
        if not filas:
            print("No se encontraron filas de datos en el archivo")
            return pd.DataFrame()

        # Analizar la primera fila para determinar estructura
        primera_fila = filas[0].strip().split("|")[1:-1]  # Quitar | inicial y final
        primera_fila = [p.strip() for p in primera_fila]

        num_columnas = len(primera_fila)
        #print(f"Estructura detectada: {num_columnas} columnas")
        #print(f"   Encabezados: {primera_fila}")

        # DEFINIR COLUMNAS BASE SEGUN ESTRUCTURA
        if num_columnas == 14:
            # Estructura original (sin Estado ni Observaciones)
            columnas_base = [
                "PurchReq",
                "Item",
                "ReqDate",
                "Material",
                "Created",
                "ShortText",
                "PO",
                "Quantity",
                "Plnt",
                "PGr",
                "Blank1",
                "D",
                "Requisnr",
                "ProcState",
            ]
            columnas_extra = ["Estado", "Observaciones"]

        elif num_columnas == 15:
            # Verificar si la columna 15 es "Estado" o "Observaciones"
            ultima_columna = primera_fila[-1].lower()
            if "estado" in ultima_columna:
                # Estructura con Estado pero sin Observaciones
                columnas_base = [
                    "PurchReq",
                    "Item",
                    "ReqDate",
                    "Material",
                    "Created",
                    "ShortText",
                    "PO",
                    "Quantity",
                    "Plnt",
                    "PGr",
                    "Blank1",
                    "D",
                    "Requisnr",
                    "ProcState",
                    "Estado",
                ]
                columnas_extra = ["Observaciones"]
            else:
                # Estructura con Observaciones pero sin Estado
                columnas_base = [
                    "PurchReq",
                    "Item",
                    "ReqDate",
                    "Material",
                    "Created",
                    "ShortText",
                    "PO",
                    "Quantity",
                    "Plnt",
                    "PGr",
                    "Blank1",
                    "D",
                    "Requisnr",
                    "ProcState",
                    "Observaciones",
                ]
                columnas_extra = ["Estado"]

        elif num_columnas == 16:
            # Estructura completa con Estado y Observaciones
            columnas_base = [
                "PurchReq",
                "Item",
                "ReqDate",
                "Material",
                "Created",
                "ShortText",
                "PO",
                "Quantity",
                "Plnt",
                "PGr",
                "Blank1",
                "D",
                "Requisnr",
                "ProcState",
                "Estado",
                "Observaciones",
            ]
            columnas_extra = []
        else:
            print(f"ERROR: Estructura no soportada: {num_columnas} columnas")
            return pd.DataFrame()

        # PROCESAR TODAS LAS FILAS
        filas_proc = []
        for i, fila in enumerate(filas):
            partes = fila.strip().split("|")[1:-1]
            partes = [p.strip() for p in partes]

            # Validar que tenga el numero correcto de columnas
            if len(partes) == num_columnas:
                filas_proc.append(partes)
            elif len(partes) == num_columnas + 1 and partes[-1] == "":
                # Caso: columna extra vacia al final
                filas_proc.append(partes[:num_columnas])
                if i < 3:  # Solo log primeras filas
                    print(f"   ADVERTENCIA Fila {i+1}: Columna extra vacia removida")
            else:
                print(
                    f"   ERROR Fila {i+1} ignorada: {len(partes)} columnas vs {num_columnas} esperadas"
                )
                if i == 0:  # Solo mostrar detalle para primera fila
                    print(f"      Contenido: {partes}")
                continue

        # CREAR DATAFRAME
        df = pd.DataFrame(filas_proc, columns=columnas_base)

        # AGREGAR COLUMNAS FALTANTES
        for col_extra in columnas_extra:
            if col_extra not in df.columns:
                df[col_extra] = ""
                print(f"EXITO: Columna '{col_extra}' agregada al DataFrame")

        # FILTRAR: Si la primera fila es encabezado, eliminarla
        primera_fila_es_encabezado = any(
            col in df.iloc[0].values if not df.empty else False
            for col in [
                "Purch.Req.",
                "Item",
                "Req.Date",
                "Short Text",
                "PurchReq",
                "Estado",
                "Observaciones",
            ]
        )

        if not df.empty and primera_fila_es_encabezado:
            df = df.iloc[1:].reset_index(drop=True)
            #print("EXITO: Fila de encabezado removida")

        #print(f"EXITO: Archivo procesado: {len(df)} filas de datos")
        #print(f"   - Columnas: {list(df.columns)}")

        if not df.empty:
            print(f"   - SOLPEDs: {df['PurchReq'].nunique()}")
            if "Estado" in df.columns:
                print(f"   - Estados unicos: {df['Estado'].value_counts().to_dict()}")

        # Normalizar formato fecha
        df["ReqDate_fmt"] = pd.to_datetime(
            df["ReqDate"], errors="coerce", dayfirst=True
        )

        df["ReqDate_fmt"] = pd.to_datetime(
            df["ReqDate"], errors="coerce", dayfirst=True
        )

        if dias is not None:
            hoy = pd.Timestamp.today().normalize()
            limite = hoy - pd.Timedelta(days=int(dias))
            filas_antes = len(df)
            df = df[df["ReqDate_fmt"] >= limite].reset_index(drop=True)
            filas_despues = len(df)
            print(
                f"EXITO: Filtrado por ReqDate últimos {dias} días -> {filas_despues}/{filas_antes}"
            )
        else:
            print("INFO: No se aplicó filtro por ReqDate (dias=None)")

        # opcional: eliminar columna auxiliar
        df.drop(columns=["ReqDate_fmt"], inplace=True)

        return df

    except Exception as e:
        WriteLog(
            mensaje=f"Error en procesarTablaME5A: {e}",
            estado="ERROR",
            nombreTarea="procesarTablaME5A",)
        print(f"ERROR en procesarTablaME5A: {e}")
        traceback.print_exc()
        return pd.DataFrame()

def ProcesarTablaMejorada(name, dias=None):
    try:
        # 1. Carga de archivo con manejo de rutas
        path = rf"{in_config('PathTemp')}\{name}"
        lineas_puras = []
        for cod in ["latin-1", "utf-8", "cp1252"]:
            try:
                with open(path, "r", encoding=cod) as f:
                    lineas_puras = [l.strip() for l in f.readlines()]
                break
            except: continue

        if not lineas_puras: return pd.DataFrame()

        # 2. Unificación de filas (Manejo de multilinealidad de SAP)
        filas_unificadas = []
        buffer_fila = ""
        for linea in lineas_puras:
            # Ignorar separadores visuales de SAP
            if not linea.startswith("|") or linea.strip().startswith("|---"):
                continue
            
            # Si la línea tiene muchos campos (pipes), es una nueva entrada [cite: 1, 4]
            if linea.count("|") > 10: 
                if buffer_fila: filas_unificadas.append(buffer_fila)
                buffer_fila = linea
            else:
                # Es continuación de la línea anterior (ej. Valor Neto o Moneda) [cite: 3, 6]
                buffer_fila += linea[1:]

        if buffer_fila: filas_unificadas.append(buffer_fila)

        # 3. Limpieza de datos y normalización de columnas
        data_final = []
        for f in filas_unificadas:
            # Dividir y limpiar espacios, ignorando elementos vacíos resultantes del split lateral
            partes = [p.strip() for p in f.split("|")]
            # Eliminar el primer y último elemento si son vacíos (por los pipes laterales)
            if partes[0] == "": partes.pop(0)
            if partes and partes[-1] == "": partes.pop(-1)
            
            if partes and not all(x == "*" for x in partes):
                data_final.append(partes)

        if not data_final: return pd.DataFrame()

        # 4. Construcción del DataFrame con validación de longitud
        encabezados = data_final[0]
        cuerpo = data_final[1:]
        
        # Validar si el primer elemento del cuerpo es en realidad el resto del encabezado
        # (A veces SAP usa 2 filas para el encabezado) 
        if cuerpo and "Material" not in encabezados and "Material" in cuerpo[0]:
            encabezados = [f"{e} {c}".strip() for e, c in zip(encabezados, cuerpo[0])]
            cuerpo = cuerpo[1:]

        # Forzar a que cada fila tenga exactamente la longitud de 'encabezados'
        cuerpo_ajustado = []
        for fila in cuerpo:
            if len(fila) > len(encabezados):
                cuerpo_ajustado.append(fila[:len(encabezados)]) # Recortar excedente
            elif len(fila) < len(encabezados):
                cuerpo_ajustado.append(fila + [""] * (len(encabezados) - len(fila))) # Rellenar faltante
            else:
                cuerpo_ajustado.append(fila)

        df = pd.DataFrame(cuerpo_ajustado, columns=encabezados)

        # 5. Limpieza de columnas "fantasma" y duplicados de encabezado
        df = df[df.iloc[:, 0] != encabezados[0]] # Eliminar si el encabezado se repite en medio
        
        # 6. Filtro por fecha (ReqDate o Fecha doc.) [cite: 4, 11, 48]
        col_fecha = next((c for c in df.columns if any(x in c for x in ["Date", "Fecha", "ReqDate"])), None)
        
        if col_fecha and not df.empty:
            df[col_fecha] = pd.to_datetime(df[col_fecha], errors="coerce", dayfirst=True)
            if dias is not None:
                limite = pd.Timestamp.today().normalize() - pd.Timedelta(days=int(dias))
                df = df[df[col_fecha] >= limite]

        return df.reset_index(drop=True)

    except Exception as e:
        print(f"Error crítico en ProcesarTablaMejorada: {e}")
        traceback.print_exc()
        return pd.DataFrame()

def buscar_objeto_por_id_parcial(session, id_parcial):
    """
    Busca de forma recursiva un objeto en la sesión de SAP cuyo ID 
    contenga la cadena especificada.
    
    Args:
        session: Sesión activa de SAP GUI.
        id_parcial (str): Parte del ID técnico del objeto (ej: 'TC_1211').
        
    Returns:
        Objeto SAP si se encuentra, de lo contrario None.
    """
    # Iniciamos la búsqueda desde el nivel de usuario para mayor eficiencia
    contenedor_principal = session.findById("wnd[0]/usr")
    
    def buscar_recursivo(objeto_padre):
        try:
            # Verificamos si el objeto actual contiene el ID buscado
            if id_parcial in objeto_padre.Id:
                return objeto_padre
            
            # Si el objeto tiene hijos, exploramos cada uno
            if hasattr(objeto_padre, "Children"):
                for hijo in objeto_padre.Children:
                    resultado = buscar_recursivo(hijo)
                    if resultado:
                        return resultado
        except Exception:
            # Ignorar objetos que no permiten acceso a sus propiedades
            pass
        return None

    return buscar_recursivo(contenedor_principal)




def ObtenerColumnasdf(ruta_archivo: str, ):

    """
    Pruebas obtener columnas de un archivo txt
    """
    df= pd.read_csv(ruta_archivo, dtype=str,sep="|")
    columnas = df.columns.tolist()
    return columnas


def obtener_ultimo_dia_habil_actual():
    """
    Docstring for obtener_ultimo_dia_habil_actual

    # Ejemplo de ejecución
    # resultado = obtener_ultimo_dia_habil_actual()
    # print(resultado)
    """
    # Obtener fecha actual
    hoy = datetime.now()
    anio = hoy.year
    mes = hoy.month
    
    # Obtener el último día del mes
    ultimo_dia_mes = calendar.monthrange(anio, mes)[1]
    fecha = datetime(anio, mes, ultimo_dia_mes)
    
    # Retroceder si es Sábado (5) o Domingo (6)
    while fecha.weekday() > 4:
        fecha -= datetime.timedelta(days=1)
        
    # 4. Formatear como DD.MM.YYYY
    return fecha.strftime('%d.%m.%Y')

def AbrirTransaccion(session, transaccion):
    """session: objeto de SAP GUI
    transaccion: transaccion a buscar
    Realiza la busqueda de la transaccion requerida"""


    logger.info(f"Abrir Transaccion {transaccion}")

    try:
        # WriteLog(
        #     mensaje=f"Abrir Transaccion {transaccion}",
        #     estado="INFO",
        #     nombreTarea="AbrirTransaccion",)
        
        # Validar sesion SAP
        if session is None:

            WriteLog(
                mensaje="Sesion SAP no disponible",
                estado="ERROR",
                nombreTarea="AbrirTransaccion",)
            raise Exception("Sesion SAP no disponible")

        # Abrir transaccion dinamica
        session.findById("wnd[0]/tbar[0]/okcd").text = transaccion
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)

        WriteLog(
            mensaje=f"Transaccion {transaccion} abierta",
            estado="INFO",
            nombreTarea="AbrirTransaccion",
            
        )
        logger.info(f"Transaccion {transaccion} abierta")
        return True
    except Exception as e:
        WriteLog(
            mensaje=f"Error en AbrirTransaccion: {e}",
            estado="ERROR",
            nombreTarea="AbrirTransaccion",
        )

        return False
    
def LeerTXT_SAP_Universal(path: str) -> pd.DataFrame:
    """
    Parser universal para archivos TXT exportados desde SAP ALV con pipes.
    Diseñado para RPA productivo.
    """

    import os

    if not os.path.exists(path):
        raise FileNotFoundError(f"No existe archivo SAP: {path}")

    # --- 1. Leer archivo con fallback encoding ---
    lineas = []
    for enc in ("latin-1", "cp1252", "utf-8"):
        try:
            with open(path, "r", encoding=enc) as f:
                lineas = [l.rstrip("\n") for l in f]
            break
        except:
            continue

    if not lineas:
        raise ValueError("Archivo SAP vacío o no legible")

    # --- 2. Filtrar solo líneas de tabla SAP ---
    lineas_tabla = []
    for l in lineas:
        if not l.startswith("|"):
            continue
        if set(l.replace("|", "").strip()) == {"-"}:
            continue
        lineas_tabla.append(l)

    if not lineas_tabla:
        raise ValueError("No se detectó tabla SAP válida")

    # --- 3. Unificar multiline SAP ---
    filas = []
    buffer = None
    pipe_ref = None

    for linea in lineas_tabla:
        pipes = linea.count("|")

        if pipe_ref is None:
            pipe_ref = pipes
            buffer = linea
            continue

        if pipes == pipe_ref:
            if buffer:
                filas.append(buffer)
            buffer = linea
        else:
            buffer += linea[1:]

    if buffer:
        filas.append(buffer)

    # --- 4. Limpiar filas ---
    data = []
    for f in filas:
        partes = [p.strip() for p in f.split("|")]
        if partes and partes[0] == "":
            partes.pop(0)
        if partes and partes[-1] == "":
            partes.pop()

        # eliminar fila de totales SAP (*)
        if partes and partes[0] == "*":
            continue

        if partes:
            data.append(partes)

    if len(data) < 2:
        raise ValueError("Tabla SAP sin encabezado o sin datos")

    encabezado = data[0]
    cuerpo = data[1:]

    # --- 5. Ajustar longitud filas ---
    ancho = len(encabezado)
    cuerpo_ok = []

    for fila in cuerpo:
        if len(fila) > ancho:
            fila = fila[:ancho]
        elif len(fila) < ancho:
            fila = fila + [""] * (ancho - len(fila))
        cuerpo_ok.append(fila)

    df = pd.DataFrame(cuerpo_ok, columns=[c.strip() for c in encabezado])

    # --- 6. Eliminar encabezados repetidos en medio ---
    df = df[df.iloc[:, 0] != encabezado[0]]

    return df.reset_index(drop=True)



#import numpy as np

def validar_estrategias_sap(df_sap, df_excel):
    # --- A. Limpieza de Precios en SAP ---
    # Convertimos a numero, quitamos puntos, cambiamos coma por punto y a float
    df_sap['Precio_Num'] = pd.to_numeric(
        df_sap['Precio neto'].astype(str)
        .str.replace('.', '', regex=False)
        .str.replace(',', '.', regex=False), 
        errors='coerce'
    ).fillna(0)

    # --- B. Limpieza de Rangos en Excel Corregida ---
    for col in ['Rango Auto min', 'Rango Auto max']:
        df_excel[col] = pd.to_numeric(df_excel[col], errors='coerce').fillna(0)
   
    # --- C. Función de comparación fila por fila ---
    def chequear_fila(fila_sap):
        precio = fila_sap['Precio_Num']
        estr_sap = str(fila_sap['Estr.']).strip().upper()
        
        # Buscamos en el excel el rango que le corresponde
        # El precio debe ser >= Min y <= Max
        match = df_excel[
            (precio >= df_excel['Rango Auto min']) & 
            (precio <= df_excel['Rango Auto max'])
        ]
        
        if not match.empty:
            # estr_teorica = str(match.iloc[0]['ESTRAT']).strip().upper()
            fila_match = match.iloc[0]
            estr_teorica = str(fila_match['ESTRAT']).strip().upper()
            r_min = fila_match['Rango Auto min']
            r_max = fila_match['Rango Auto max']
            
            # DEBUG INDIVIDUAL POR FILA
            # print(f"---> Validando OC {fila_sap['Doc.compr.']}: Precio {precio}")
            # print(f"      Rango Encontrado: {r_min} a {r_max} (Estrategia: {estr_teorica})")
            
            if estr_sap == estr_teorica:
                return "OK"
            else:
                return f"ERROR: Rango {r_min}-{r_max} exige {estr_teorica}"
        else:
            return "FUERA DE RANGO: No coincide con ningun rango del Excel"

    # Aplicamos la lógica a todo el DataFrame de SAP
    df_sap['Resultado_Validacion'] = df_sap.apply(chequear_fila, axis=1)
    
    return df_sap


# def NotificarErroresEstrategia(df_sap_validado, correo_destino):
#     """
#     Filtra las OCs con error y envía un correo con el resumen.
#     """
#     # 1. Filtramos solo los errores para el cuerpo del correo
#     df_errores = df_sap_validado[df_sap_validado['Resultado_Validacion'] != 'OK']
    
#     if df_errores.empty:
#         print("No hay errores que reportar.")
#         return

#     # 2. Creamos el listado en formato HTML para el cuerpo del correo
#     # Convertimos el dataframe a una tabla HTML bonita
#     tabla_html = df_errores[['Doc.compr.', 'Precio neto', 'Estr.', 'Resultado_Validacion']].to_html(index=False, classes='table')

#     asunto = f"Resumen de Errores en Estrategias SAP - {len(df_errores)} hallazgos"

#     style = "<style>table {border-collapse: collapse;} th, td {border: 1px solid black; padding: 5px;} th {background-color: #f2f2f2;}</style>"
#     cuerpo = f"""
#     <html>
#     <head>{style}</head>
#         <body>
#             <h2>Reporte Automático de Validación de Estrategias</h2>
#             <p>Se han detectado discrepancias en las siguientes Órdenes de Compra:</p>
#             {tabla_html}
#             <br>
#             <p>Favor revisar el archivo adjunto para más detalle.</p>
#             <p><i>Atentamente, Robot RIGO</i></p>
#         </body>
#     </html>
#     """

#     # 3. Guardamos el reporte completo a Excel para enviarlo como adjunto
#     ruta_adjunto = "Reporte_Validacion_Estrategias.xlsx"
#     df_sap_validado.to_excel(ruta_adjunto, index=False)

#     # 4. Usamos tu función personalizada para enviar
#     exito = EnviarCorreoPersonalizado(
#         destinatario=correo_destino,
#         asunto=asunto,
#         cuerpo=cuerpo,
#         adjuntos=[ruta_adjunto],
#         nombreTarea="ValidacionSAP"
#     )
    
#     return exito

def impimmirdf(df: pd.DataFrame):
        
        #print(type(df))
        # print("Columnas obtenidas del df de la base de datos:")
        # print(df.columns.tolist())
        # print("Columnas obtenidas del list(df):")
        # print(list(df))
        print("Columnas obtenidas del df.head():")
        print(df.head())
        #print(df.to_string())

        # print("Columnas obtenidas del  df.info()")
        # print(df.info())

def fomatodf(df: pd.DataFrame):
        """
        darle formato a los data frame 
        """
        # Limpiar espacios en los nombres de las columnas
        df.columns = [re.sub(r'\s+', ' ', str(col)).strip() for col in df.columns]
 
        # Identificar y renombrar duplicados
        cols = pd.Series(df.columns)
        for i in cols[cols.duplicated()].unique():
            cols[cols == i] = [f"{i}_{j}" if j != 0 else i for j in range(sum(cols == i))]
        df.columns = cols # Ahora las columnas se llamarán "Nombre 1" (la primera) y "Nombre 1_1" (la segunda)
        # 1. Quitamos filas donde Borrado es "L"
        # 2. Quitamos filas donde Status Lib sea NaN (nulo)
        # 3. Quitamos filas donde Status Lib esté vacío (espacios en blanco)

        df = df[
            (df['Borrado'] != 'L') & 
            (df['Status Lib'].notna()) & 
            (df['Status Lib'].astype(str).str.strip() != '')
        ].copy()
                
        # Filtramos solo las columnas que existan en el DataFrame original #2
        columnas_interes = ['Fecha doc.','Acreedor','Nombre 1','Creado','Estr.', 'Doc.compr.', 'Status Lib', 'Precio neto', ]
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
                "Creado": "first",
                "Estr.": "first", 
                "Status Lib": "first",
                "Precio neto": "sum", # Sumamos el precio neto para cada documento de compra  // STEV : se deja fuera del alcance por ahora.
                "FechaActualizacion": "first",
                "EstadoNotificacion": "first",
                #"CorreoArrendatarios":"first",
                # "Fecha Lib": "first",
                # "Usuario Li": "first",
                # "Fecha Lib.": "first",
                # "Usuario Li": "first"
            }).reset_index()
        
        df['ContadorEnvio']= 0
        df['CorreoArrendatarios'] = 0  

        return df

def descargadataestliberacion (session):
    """
    Pasos en SAP para descargar la data de estrategias de Liberacio, deja un TXT, en el file server.
    """

    try :
        db= Database()
        engine = db.get_engine()
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
        grupoOrgCompras = pd.read_sql_table("Config_Compras", engine, schema="PagoArriendos")
        grupoOrgCompras = grupoOrgCompras['CodigoOrg'].tolist()
        logger.debug(grupoOrgCompras)
        #grupoOrgCompras = ["OC03","OC30","OC02"]# Esto lo puedes traer de la tabla de la base de datos db, parametros 
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
        responsable = pd.read_sql_table("Responsable", engine, schema="PagoArriendos")
        responsable = responsable['Responsable'].tolist()
        #responsable = ["FERNCAMS","ERIIGUZV"] # Esto lo puedes traer de la tabla de la base de datos db, parametros 
        texto_sap = "\r\n".join(responsable)
        pyperclip.copy(texto_sap)
        session.findById("wnd[0]/usr/btn%_R_ERNAM_%_APP_%-VALU_PUSH").press() # Abre ventana responsable de la OC
        session.findById("wnd[1]/tbar[0]/btn[16]").press() # Boton basura, borrar datos
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
               
        
        # Ejecutar búsqueda
        session.findById("wnd[0]/tbar[1]/btn[8]").press() #Ejecutar búsqueda

        # Guardar resultados en txt
        rutaGuardar = fr"{in_config('PathTemp')}\HU08"
        session.findById("wnd[0]/tbar[1]/btn[45]").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = rutaGuardar
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = f"EstrategiasDeLiberacion{fecha_formateada}.txt"
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

    except:
        logger.exception("Error en la descarga de data desde SAP Estrategias de liberacion ")

