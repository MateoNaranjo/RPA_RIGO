# ============================================
# Función Local: validacionME53N
# Autor: Paula Sierra - NetApplications
# Descripcion: Ejecuta ME5A y exporta archivo TXT según estado.
# Ultima modificacion: 24/11/2025
# Propiedad de Colsubsidio
# Cambios: (Si Aplica)
# ============================================
from requests import session
import win32com.client
import traceback
import pandas as pd
import re
import subprocess
import time
import os
from funciones.EscribirLog import WriteLog
import pyautogui
from pyautogui import ImageNotFoundException
from funciones.Login import ObtenerSesionActiva
from typing import List, Literal, Optional
from config.init_config import in_config

from datetime import datetime, timedelta
import calendar

class SapTextEditor:
    """
    Wrapper para el editor de textos SAP (GuiShell - SAPLMMTE).
    Permite leer y modificar texto línea por línea de forma segura.
    #Stev: se prueban varios metodos, pero la mejor opcion es tomar todo el texto y luego setearlo todo de nuevo desde la linea 0, no es recomendable buscar una linea especifica
    # usando EditorTxt.SetUnprotectedTextPart(0,".")
    """

    def __init__(self, session, editor_id):
        """
        Args:
            session: sesión activa SAP GUI
            editor_id (str): ID completo del editor (hasta /shell)
        """
        self.session = session
        self.editor_id = editor_id
        self.shell = self._get_shell()

    def _get_shell(self):
        shell = self.session.findById(self.editor_id)
        if shell.Type != "GuiShell":
            raise Exception("El objeto encontrado no es un GuiShell (Text Editor SAP)")
        return shell

    # ------------------------------------------------------------------
    # LECTURA
    # ------------------------------------------------------------------

    def get_line(self, index):
        """Obtiene el texto completo de una línea."""
        try:
            return self.shell.GetLineText(index)
        except Exception:
            return None

    def get_all_text(self, max_lines=100):
        """
        Obtiene todo el texto del editor SAP sin saltos de línea finales
        ni líneas vacías generadas por el control.
        """
        lines = []


        for i in range(max_lines):
            try:
                line = self.get_line(i)

                if line is None:
                    break

                # Limpia caracteres invisibles pero conserva el contenido
                clean_line = line.rstrip()
                lines.append(clean_line)

            except Exception:
                break

        # Elimina líneas vacías finales
        while lines and lines[-1] == "":
            lines.pop()

        return "\n".join(lines)




    # ------------------------------------------------------------------
    # ESCRITURA
    # ------------------------------------------------------------------

    def set_editable_line(self, index, new_text):
        """Modifica únicamente la parte editable de una línea."""
        try:
            self.shell.SetUnprotectedTextPart(index, new_text)
            return True
        except Exception:
            return False

    def replace_in_text(self, texto: str, replacements: dict):
        """
        Reemplaza textos sobre un string completo, evitando líneas vacías iniciales
        y agregando un mensaje final con el total de cambios.

        Args:
            texto (str): texto original del editor SAP
            replacements (dict): {"SAA": "R3", ...}

        Returns:
            nuevo_texto (str)
            cambios (int)
        #Stev: cambia el texto segun un diccionario de reemplazos y retorna el texto modificado y la cantidad de cambios realizados
        """
        

        if not texto:
            return texto, 0

        # 1. Eliminar SOLO saltos de línea iniciales
        texto = texto.lstrip("\n")
        lineas = texto.splitlines()

        nuevas_lineas = []
        cambios = 0
        CambioExacto = "Stev: No se realizaron cambios exactos" # cambio exacto para el log
        

        for linea in lineas:
            nueva = linea

            for buscar, reemplazar in replacements.items():
                # Reemplazo exacto por línea
                if linea.strip() == buscar:
                    nueva = reemplazar
                    # todo: hacer el apend a la lista para el informe 
                    CambioExacto = (f"[CAMBIO EXACTO] '{linea}' -> '{reemplazar}'")
                else:
                    nueva = nueva.replace(buscar, reemplazar)

            if nueva != linea:
                cambios += 1

            nuevas_lineas.append(nueva)

        # 2. Agregar mensaje final si hubo cambios
        if cambios > 0:
            nuevas_lineas.append("")
            nuevas_lineas.append(
                f"TEXTO MODIFICADO AUTOMÁTICAMENTE POR BOT RPA – CAMBIOS APLICADOS: {cambios}"
            )

        return "\n".join(nuevas_lineas), cambios, CambioExacto

# fin class SapTextEditor:
# fin utilidades

# ===============================================================================================
# Funciones para obtener el ID de los objetos dinamicamnete dependiento del objeto padre
# devuelve el valor de la propiedad o ejecuta la accion deseada
# ===============================================================================================

def set_sap_table_scroll(session, table_id_part, position):
    """
    Busca una tabla por su ID técnico y ajusta su scroll vertical.
    
    Args:
        session: Sesión activa.
        table_id_part (str): Parte única del ID de la tabla (ej: 'TC_1211').
        position (int): Posición a la que queremos mover el scroll.
    """
    usr = session.findById("wnd[0]/usr")
    
    def buscar_tabla(obj):
        try:
            # Buscamos que sea una tabla y que el ID contenga el nombre técnico
            if obj.Type == "GuiTableControl" and table_id_part in obj.Id:
                return obj
            
            for child in obj.Children:
                res = buscar_tabla(child)
                if res: return res
        except Exception:
            pass
        return None

    tabla = buscar_tabla(usr)
    
    if tabla:
        # Ajustamos la posición del scrollbar
        tabla.verticalScrollbar.position = position
    else:
        raise Exception(f"No se encontró la tabla con ID que contenga: {table_id_part}")

def press_GuiButton(session, button_id):
    """
    Presiona un botón (GuiButton) en SAP GUI a partir de su ID lógico.

    Args:
        session: sesión SAP activa
        button_id (str): ID lógico del botón (ej: 'AUTOTEXT002')

    Ejemplo:
        press_GuiButton(session, "AUTOTEXT002")
    """

    if not button_id:
        raise ValueError("button_id es obligatorio")

    usr = session.findById("wnd[0]/usr")
    target_suffix = f"btn%#{button_id}"

    def buscar_boton(obj):
        try:
            if obj.Type == "GuiButton" and obj.Id.endswith(target_suffix):
                return obj

            for child in obj.Children:
                res = buscar_boton(child)
                if res:
                    return res
        except Exception:
            pass
        return None

    boton = buscar_boton(usr)

    if not boton:
        raise Exception(f"No se encontró GuiButton con ID lógico: {button_id}")

    # Press() es seguro incluso si el botón ya fue usado
    boton.Press()

def SetGuiComboBoxkey(session, campo_id, key_value="ZRCR"):
    """
    Selecciona un valor en un GuiComboBox de SAP GUI usando un ID lógico.

    Args:
        session: sesión SAP activa
        campo_id (str): identificador lógico del campo (ej: 'TOPLINE-BSART')
        key_value (str): clave a seleccionar en el combo (ej: 'ZRCR')

    Raises:
        Exception si no se encuentra el GuiComboBox
    Ejemplo de uso:
        SetGuiComboBoxkey(session, "TOPLINE-BSART", "ZRCR")
    """

    if not campo_id:
        raise ValueError("campo_id es obligatorio")

    usr = session.findById("wnd[0]/usr")

    def buscar_combobox(obj):
        try:
            if (
                obj.Type == "GuiComboBox"
                and campo_id in obj.Id
            ):
                return obj

            for child in obj.Children:
                res = buscar_combobox(child)
                if res:
                    return res
        except Exception:
            pass
        return None

    combo = buscar_combobox(usr)

    if not combo:
        raise Exception(f"No se encontró GuiComboBox con ID lógico: {campo_id}")

    # Selección por Key (forma correcta y estable)
    combo.Key = key_value

def set_GuiCabeceraTextField_text(session, campo_id, valor):
    """
    Setea un texto en un GuiCTextField (ME21N / cabecera)
    usando el identificador lógico SAP (ej: 'EKORG', 'EKGRP').

    Args:
        session: sesión activa de SAP GUI
        campo_id (str): ID lógico del campo (ej: 'EKORG')
        valor (str): texto a insertar

    Raises:
        ValueError: si los argumentos son inválidos
        Exception: si no se encuentra el campo
    """

    if not campo_id:
        raise ValueError("campo_id es obligatorio (ej: 'EKORG')")

    if valor is None:
        raise ValueError("valor no puede ser None")

    usr = session.findById("wnd[0]/usr")
    target_suffix = f"-{campo_id}"

    def buscar_ctextfield(obj):
        try:
            if (
                obj.Type == "GuiCTextField"
                and obj.Id.endswith(target_suffix)
            ):
                return obj

            for child in obj.Children:
                res = buscar_ctextfield(child)
                if res:
                    return res
        except Exception:
            pass
        return None

    ctext = buscar_ctextfield(usr)

    if not ctext:
        raise Exception(f"No se encontró GuiCTextField con ID lógico: {campo_id}")

    # En SAP es buena práctica asegurar foco antes de escribir
    try:
        ctext.SetFocus()
    except Exception:
        pass

    ctext.Text = str(valor)


def get_GuiCabeceraTextField_text(session, campo_id):
    """
    Obtiene el texto (.Text.strip()) de un GuiCTextField en ME21N
    a partir del identificador lógico SAP (ej: 'EKORG', 'EKGRP').

    Args:
        session: sesión activa de SAP GUI
        campo_id (str): identificador lógico del campo SAP (ej: 'EKORG')

    Returns:
        str: texto del campo (strip)

    Raises:
        ValueError: si el argumento es inválido
        Exception: si no se encuentra el GuiCTextField
    """

    if not campo_id:
        raise ValueError("campo_id es obligatorio (ej: 'EKORG')")

    usr = session.findById("wnd[0]/usr")
    target_suffix = f"-{campo_id}"

    def buscar_ctextfield(obj):
        try:
            if (
                obj.Type == "GuiCTextField"
                and obj.Id.endswith(target_suffix)
            ):
                return obj

            for child in obj.Children:
                res = buscar_ctextfield(child)
                if res:
                    return res
        except Exception:
            pass
        return None

    ctext = buscar_ctextfield(usr)

    if not ctext:
        raise Exception(f"No se encontró GuiCTextField con ID lógico: {campo_id}")

    return ctext.Text.strip()

#for fila in range(item):
#   precio = get_GuiTextField_text(session, f"NETPR[10,{fila}]")

def get_GuiTextField_text(session, campo_posicion):
    """
    Obtiene el texto de un GuiTextField dentro de un TableControl SAP
    usando una posición lógica (ej: 'NETPR[10,0]').

    Args:
        session: sesión SAP activa
        campo_posicion (str): campo con posición SAP (ej: 'NETPR[10,0]')

    Returns:
        str: texto del campo

    Raises:
        Exception si no se encuentra el objeto
    """

    if not campo_posicion:
        raise ValueError("campo_posicion es obligatorio")

    # Parsear NETPR[10,0]
    match = re.match(r"(.+)\[(\d+),(\d+)\]", campo_posicion)
    if not match:
        raise ValueError("Formato inválido. Use: NETPR[10,0]")

    campo, col, fila = match.groups()
    col = int(col)
    fila = int(fila)

    usr = session.findById("wnd[0]/usr")

    def buscar_textfield(obj):
        try:
            if (
                obj.Type == "GuiTextField"
                and campo in obj.Id
                and obj.Id.endswith(f"[{col},{fila}]")
            ):
                return obj

            for child in obj.Children:
                res = buscar_textfield(child)
                if res:
                    return res
        except Exception:
            pass
        return None

    txt = buscar_textfield(usr)

    if not txt:
        raise Exception(f"No se encontró GuiTextField: {campo_posicion}")

    return txt.Text.strip()

def set_GuiTextField_text(session, campo_posicion, valor):
    """
    Setea el texto de un GuiTextField dentro de un TableControl SAP
    usando posición lógica (ej: 'NETPR[10,0]' o 'MENGE[6,0]').
    Compatible con M21N (MEPO1211).
    """

    if not campo_posicion:
        raise ValueError("campo_posicion es obligatorio")

    if valor is None:
        valor = ""

    # Parseo CAMPO[col,fila]
    match = re.fullmatch(r"([A-Z0-9_]+)\[(\d+),(\d+)\]", campo_posicion.upper())
    if not match:
        raise ValueError("Formato inválido. Use: NETPR[10,0] o MENGE[6,0]")

    campo, col, fila = match.groups()
    col = int(col)
    fila = int(fila)

    usr = session.findById("wnd[0]/usr")
    

    objetivo = f"-{campo}[{col},{fila}]"

    def buscar_textfield(obj):
        try:
            if (
                obj.Type == "GuiTextField"
                and objetivo in obj.Id
            ):
                return obj

            for child in obj.Children:
                res = buscar_textfield(child)
                if res:
                    return res
        except Exception:
            pass
        return None

    txt = buscar_textfield(usr)

    if not txt:
        raise Exception(f"No se encontró GuiTextField SAP: {campo}[{col},{fila}]")

    # Seteo seguro (SAP-friendly)
    txt.SetFocus()
    txt.Text = str(valor)
    txt.CaretPosition = len(txt.Text)
    session.findById("wnd[0]").sendVKey(0)

def set_GuiTextField_Ventana1_text(session, campo_posicion, valor):
    """
    Setea el texto de un GuiTextField dentro de un TableControl SAP
    usando posición lógica (ej: 'NETPR[10,0]' o 'MENGE[6,0]').
    Compatible con M21N (MEPO1211).
    """

    if not campo_posicion:
        raise ValueError("campo_posicion es obligatorio")

    if valor is None:
        valor = ""

    # Parseo CAMPO[col,fila]
    match = re.fullmatch(r"([A-Z0-9_]+)\[(\d+),(\d+)\]", campo_posicion.upper())
    if not match:
        raise ValueError("Formato inválido. Use: NETPR[10,0] o MENGE[6,0]")

    campo, col, fila = match.groups()
    col = int(col)
    fila = int(fila)
    #ventana 1
    usr = session.findById("wnd[1]/usr")
    

    objetivo = f"-{campo}[{col},{fila}]"

    def buscar_textfield(obj):
        try:
            if (
                obj.Type == "GuiCTextField"
                and objetivo in obj.Id
            ):
                return obj

            for child in obj.Children:
                res = buscar_textfield(child)
                if res:
                    return res
        except Exception:
            pass
        return None

    txt = buscar_textfield(usr)

    if not txt:
        raise Exception(f"No se encontró GuiTextField SAP: {campo}[{col},{fila}]")

    # Seteo seguro (SAP-friendly)
    txt.SetFocus()
    txt.Text = str(valor)
    txt.CaretPosition = len(txt.Text)
    session.findById("wnd[1]").sendVKey(0)

def ventana_abierta(session, titulo_parcial):
    """
    Verifica si existe una ventana abierta cuyo título contenga el texto indicado.

    Args:
        session: sesión activa SAP GUI
        titulo_parcial (str): texto a buscar en el título (case-insensitive)

    Returns:
        bool
    """

    titulo_parcial = titulo_parcial.lower()

    for wnd in session.Children:
        try:
            if titulo_parcial in wnd.Text.lower():
                return True
        except Exception:
            pass

    return False

def SelectGuiTab(session, tab_id):
    """
    Selecciona una pestaña (GuiTab) del detalle de posición en ME21N.
    Args:
        session: sesión activa de SAP GUI
        tab_id (str): ID lógico de la pestaña (ej: 'TABIDT14')
    Ejemplos:
        SelectGuiTab(session, "TABIDT14")  # Textos
        SelectGuiTab(session, "TABIDT05")  # Entrega
    """

    if not tab_id:
        raise ValueError("tab_id es obligatorio")

    usr = session.findById("wnd[0]/usr")
    target_suffix = f"tabp{tab_id}"

    def buscar_tab(obj):
        try:
            if obj.Type == "GuiTab" and obj.Id.endswith(target_suffix):
                return obj
            for child in obj.Children:
                res = buscar_tab(child)
                if res:
                    return res
        except Exception:
            pass
        return None

    tab = buscar_tab(usr)

    if not tab:
        raise Exception(f"No se encontró la pestaña GuiTab con ID :{tab_id}")
    # Select() es seguro incluso si ya está seleccionada
    tab.Select()

def boton_existe(session, id):
    """
    Verifica de forma segura si un objeto SAP existe a partir de su ID completo.

    Args:
        session: La sesión activa de SAP GUI.
        id (str): El ID completo del objeto a verificar.

    Returns:
        bool: True si el objeto existe, False en caso contrario.
    """
    try:
        session.findById(id)
        return True
    except:
        return False

def buscar_y_clickear(
    ruta_imagen,
    confidence=0.5,
    intentos=20,
    espera=0.5,
    fail_silently=True,
    log=True
):
    """
    Busca una imagen en pantalla y hace click cuando la encuentra,
    controlando el error si no aparece y permitiendo continuar el flujo.

    Args:
        ruta_imagen (str): Ruta de la imagen a buscar.
        confidence (float): Nivel de confianza para el match (OpenCV).
        intentos (int): Número máximo de intentos.
        espera (float): Tiempo de espera entre intentos (segundos).
        fail_silently (bool): Si True, no lanza excepción al fallar.
        log (bool): Si True, muestra mensajes de estado.

    Returns:
        bool: True si se encontró e hizo click, False si no.
    """
    task_name = "HU4_GeneracionOC"

    for intento in range(1, intentos + 1):
        try:
            pos = pyautogui.locateCenterOnScreen(
                ruta_imagen,
                confidence=confidence
            )

            if pos:
                pyautogui.click(pos)
                #pyautogui.press("enter") # Descomentar si se quiere dar enter tras el click
                if log:
                    WriteLog(
                        mensaje=f"Imagen encontrada y clickeada: {ruta_imagen}",
                        estado="INFO",
                        task_name=task_name,)
                    #print(f"[INFO] Imagen encontrada y clickeada: {ruta_imagen}")
                return True

        except ImageNotFoundException:
            # PyAutoGUI puede lanzar esta excepción en algunas versiones
            #pyautogui.press("enter") # Descomentar si se quiere dar enter tras el click
            pass

        except Exception as e:
            if log:
                 WriteLog(
                        mensaje=f"Error inesperado buscando imagen {ruta_imagen}: {e}",
                        estado="ERROR",
                        task_name=task_name,)
                 #print(f"[ERROR] Error inesperado buscando imagen {ruta_imagen}: {e}")
            if not fail_silently:
                raise

        time.sleep(espera)

    if log:
        WriteLog(
            mensaje=f"Imagen no encontrada tras {intento} intentos: {ruta_imagen}",
            estado="WARNING",
            task_name=task_name,)
        #print(f"[WARNING] Imagen no encontrada tras {intento} intentos: {ruta_imagen}")

    if not fail_silently:
        raise RuntimeError(f"No se encontró la imagen: {ruta_imagen}")

    return False

def clasificar_concepto(concepto: str) -> Literal["PRODUCTO", "SERVICIO"]: 
    """
    Clasifica un concepto como PRODUCTO o SERVICIO
    usando reglas de negocio.
    """

    concepto_upper = concepto.upper()

    # Palabras clave típicas de SERVICIO
    palabras_servicio = [
        "TRANSPORTE",
        "ANIMADOR",
        "LAVADO",
        "COORDINADOR",
        "SERVICIO",
        "MANTENIMIENTO",
        "INSTALACION",
        "REPARACION",
        "LIMPIEZA",
        "ALQUILER",
        "ARRENDAMIENTO",
        "ASEO",
        "REVISION",
        "SOPORTE",
        "CAPACITACION"
    ]

    if any(palabra in concepto_upper for palabra in palabras_servicio):
        return "SERVICIO"
    # Stev: añadir mas reglas si es necesario para producto por palabras clave
    # Regla por descarte (objetos físicos)
    return "PRODUCTO"

def extraer_concepto(texto: str) -> Optional[str]:
    """
    Extrae el valor del campo 'POR CONCEPTO DE:'.
    """
    patron = re.compile(
        r"POR\s+CONCEPTO\s+DE\s*:\s*(.+)",
        re.IGNORECASE
    )

    for linea in texto.splitlines():
        match = patron.search(linea)
        if match:
            return match.group(1).strip()

    return None

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
        # El mensaje de éxito con el número de OC suele aparecer en la barra de estado.
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

def CambiarGrupoCompra(session):
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
            task_name="procesarTablaME5A",)
  
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
            task_name="procesarTablaME5A",)
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

def obtener_importe_por_denominacion(session, nombre_buscado="imp.Saludable"):
    # 1. Identificar la tabla y el scrollbar de forma dinámica
    # Usando TC_1211 como ejemplo para la tabla de condiciones
    tabla = buscar_objeto_por_id_parcial(session, "TC_1211") 
    scrollbar = tabla.verticalScrollbar
    
    encontrado = False
    fila_actual = 0
    total_filas = scrollbar.maximum
    visible_row_count = tabla.visibleRowCount

    while scrollbar.position <= total_filas:
        for i in range(visible_row_count):
            try:
                # Obtenemos el texto de la columna Denominación (VTEXT)
                denominacion = get_GuiTextField_text(session, f"VTEXT[2,{i}]")
                
                if nombre_buscado.lower() in denominacion.lower():
                    # Si coincide, capturamos el valor de la columna Importe (KBETR)
                    # Nota: Debes verificar el índice de columna para Importe en tu SAP
                    importe = get_GuiTextField_text(session, f"KBETR[3,{i}]")
                    return importe
            except Exception:
                # Si falla una fila (ej: fila vacía al final), continuamos
                continue
        
        # 2. Si no se encontró en las visibles, bajar el scroll
        nueva_posicion = scrollbar.position + visible_row_count
        if nueva_posicion > total_filas:
            break # Ya llegamos al final
            
        scrollbar.position = nueva_posicion
        # Importante: Pequeña espera para que SAP refresque los datos internos
        time.sleep(0.5) 
        
    return None


def ObtenerColumnasdf(ruta_archivo: str, ):

    """
    Pruebas obtener columnas de un archivo txt
    """
    df= pd.read_csv(ruta_archivo, dtype=str,sep="|")
    columnas = df.columns.tolist()
    return columnas


def get_importesCondiciones(session, impuesto_buscado="Imp. Saludable IBUE"):
    """
    Obtiene los importes de la pestaña condiciones en ME21N
    segun un impuesto específico.
    
    Args:
        session: Sesión activa de SAP GUI. 
        impuesto_buscado (str): Impuesto a buscar.

    Returns:
        str: Importe del impuesto encontrado o None si no se encuentra.
    
    """

    # Tomar impuesto Saludable en la pestaña de Condiciones 
    SelectGuiTab(session, "TABIDT8")
    bandera = True
    # asegura que empieza en la posición 1 de la tabla de Condiciones
    set_sap_table_scroll(session, "tblSAPLV69ATCTRL_KONDITIONEN", 1)

    while (bandera == True):
        try :
            for i in range(20):  # Revisa las condiciones
                impuestosCondiciones = get_GuiTextField_text(session, f"VTEXT[2,{i}]")
                print(f"Impuesto en la pestaña de condiciones: {impuestosCondiciones}")
                if impuestosCondiciones == impuesto_buscado: 
                    print("Impuesto encontrado:", impuestosCondiciones)
                    bandera=False
                    return get_GuiTextField_text(session, f"KBETR[3,{i}]")  # Retorna el importe asociado
                                        
                elif impuestosCondiciones == "" :
                    print("no encontrado:", impuestosCondiciones)
                    bandera=False
                    return None  # Salir si se encuentra una fila vacía
                    
        except Exception as e:
            SelectGuiTab(session, "TABIDT8")
            set_sap_table_scroll(session, "tblSAPLV69ATCTRL_KONDITIONEN", i)
            print("Error al obtener los impuestos de las condiciones:", str(e))
            #continue


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
        fecha -= timedelta(days=1)
        
    # 4. Formatear como DD.MM.YYYY
    return fecha.strftime('%d.%m.%Y')

def AbrirTransaccion(session, transaccion):
    """session: objeto de SAP GUI
    transaccion: transaccion a buscar
    Realiza la busqueda de la transaccion requerida"""

    try:
        WriteLog(
            mensaje=f"Abrir Transaccion {transaccion}",
            estado="INFO",
            task_name="AbrirTransaccion",)
        
        # Validar sesion SAP
        if session is None:

            WriteLog(
                mensaje="Sesion SAP no disponible",
                estado="ERROR",
                task_name="AbrirTransaccion",)
            raise Exception("Sesion SAP no disponible")

        # Abrir transaccion dinamica
        session.findById("wnd[0]/tbar[0]/okcd").text = transaccion
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)

        WriteLog(
            mensaje=f"Transaccion {transaccion} abierta",
            estado="INFO",
            task_name="AbrirTransaccion",
            
        )
        print(f"Transaccion {transaccion} abierta")
        return True
    except Exception as e:
        WriteLog(
            mensaje=f"Error en AbrirTransaccion: {e}",
            estado="ERROR",
            task_name="AbrirTransaccion",
        )

        return False