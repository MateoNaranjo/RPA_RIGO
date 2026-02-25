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
from config.settings import RUTAS
import pyautogui
from pyautogui import ImageNotFoundException
from funciones.Login import ObtenerSesionActiva
from funciones.GuiShellFunciones import (SapTextEditor,
set_GuiTextField_text,              
get_GuiTextField_text,
buscar_objeto_por_id_parcial,
get_importesCondiciones,
obtener_valor,
extraer_concepto,
obtener_correos,
normalizar_precio_sap, 
clasificar_concepto,
esperar_sap_listo,
buscar_y_clickear, set_sap_table_scroll, 
ventana_abierta,
SelectGuiTab,
MostrarCabecera,
obtener_numero_oc
)
from typing import List, Literal, Optional


def ValidarAjustarSolped(session,item=1):
    """
    Cambia los precios y las cantidades de la Solped segun el texto del "Texto pedido" (textPF.selectedNode ="F01")
    borra los textos adicionales que no se utilizan (textPF.selectedNode ="F02"), hasta el F05

    Args:
        session: sesión SAP activa
        item (int): posiciones que tiene la Solped

    Raises:
        Exception si no se encuentra el objeto
    """


    try:
        
        # Lista de acciones en SAP que sirve de informe
        acciones = []

        for fila in range(item):  #cambiar por item
            
            # Obtiene el Precio de la posicion
            # PrecioPosicion = get_GuiTextField_text(session, f"NETPR[10,{fila}]") #Stev: implementar while para scroll, hacer dinamico el f"NETPR[10,{fila}]"
            PrecioPosicion = get_GuiTextField_text(session, f"NETPR[10,0]")
            PrecioPosicion = normalizar_precio_sap(PrecioPosicion)
            
            # Obtine la Cantidad en la Posicion
            # CantidadPosicion = get_GuiTextField_text(session, f"MENGE[6,{fila}]") #Stev: implementar while para scroll, hacer dinamico el f"MENGE[6,{fila}]"
            CantidadPosicion = get_GuiTextField_text(session, f"MENGE[6,0]")
            #CantidadPosicion = normalizar_precio_sap(CantidadPosicion)

            # Selecbox de la posicion de la solped  ejemplo de guia :  1 [10] 80016676 , LAVADO MANTEL GRANDE 
            PosicionSolped = buscar_objeto_por_id_parcial(session, "cmbDYN_6000-LIST")
            PosicionSolped.key = f"   {fila+1}"

            # Navega a la pestaña de textos
            esperar_sap_listo(session)
            SelectGuiTab(session, "TABIDT14")
            textPF1 = buscar_objeto_por_id_parcial(session, "cntlTEXT_TYPES_0200/shell")
            textPF1.selectedNode = "F01" # Foco en primer Texto IMPORTANTE
            esperar_sap_listo(session)
            EDITOR_ID= buscar_objeto_por_id_parcial(session, "cntlTEXT_EDITOR_0201/shellcont/shell")
        
            esperar_sap_listo(session)
            # obtiene el texto del objeto ├─ Leer textos
            editor = SapTextEditor(session, EDITOR_ID.id)
            texto = editor.get_all_text()

            # Obtiene el valor en el texto (Precio)
            claves = ["VALOR"] # str que busca en el texto
            preciotexto = obtener_valor(texto, claves)
            preciotexto = normalizar_precio_sap(preciotexto)
          
            # Obtiene la cantidad en el texto
            claves = ["CANTIDAD"] # str que busca en el texto
            cantidadtexto = obtener_valor(texto, claves)

            # Obtiene impuestos en el texto
            claves = ["IMPUESTO QUE APLICA"] # str que busca en el texto
            impuestostexto = obtener_valor(texto, claves)

            correosColdubsidio = obtener_correos(texto,"@colsubsidio.com") # ejemplo de uso de la funcion obtener correos
            acciones.append(f"Correos encontrados en el texto de la posicion {fila+1}: {correosColdubsidio}")
            acciones.append(f"Impuestos encontrados en el texto de la posicion {fila+1}: {impuestostexto}")

            concepto = extraer_concepto(texto)
            if concepto:
                tipo = clasificar_concepto(concepto)
                acciones.append(f"{concepto} => {tipo}")

            # Comparacion de Valores de Cantidad
            if cantidadtexto and cantidadtexto.strip() and CantidadPosicion != cantidadtexto:
                set_GuiTextField_text(session, f"MENGE[6,0]", cantidadtexto)
                #print(f"Se mofico posicion :{fila+1}0 Cantidad -> {CantidadPosicion} != {cantidadtexto}")
                acciones.append(f"Posicion {fila+1}0 => CP: {CantidadPosicion} != CT:{cantidadtexto} ⚠️ Se Actualiza Cantidad")                 

            # Comparacion de Valores de Precio
            if preciotexto and str(preciotexto).strip() and PrecioPosicion != preciotexto:
                set_GuiTextField_text(session, f"NETPR[10,0]", preciotexto)
                #print(f"Se mofico posicion :{fila+1}0 Precio -> {PrecioPosicion} != {preciotexto}")
                acciones.append(f"Posicion {fila+1}0 => PP:{PrecioPosicion} != PT:{preciotexto}⚠️ Se Actualiza Precio")
            

            # Realiza los reemplazos en el texto
            reemplazos = {"VENTA SERVICIO": "V1","VENTA PRODUCTO": "V1","GASTO PROPIO SERVICIO": "C2","GASTO PROPIO PRODUCTO": "C2","SAA": "R3","SAA PRODUCTO": "R3"} #"SAA SERVICIO": "R3"
            nuevo_texto,cambios,cambioEcxacto = editor.replace_in_text(texto, reemplazos)

            #Si hay cambios, agrega a la lista de acciones
            if cambios > 0:
                acciones.append(f"Cambios realizados: {cambios} en la posicion :{fila+1}0 en el Texto :{cambioEcxacto}")
            
            # Actualiza el texto en el editor de SAP
            editext=session.findById(EDITOR_ID.id)
            editext.SetUnprotectedTextPart(0,nuevo_texto)

            #Borra los textos de cada editor F02 en adelante
            for i in range(2, 6):  # F02 a F05
                SelectGuiTab(session, "TABIDT14")               
                textPF = buscar_objeto_por_id_parcial(session, "cntlTEXT_TYPES_0200/shell")
                nodo = f"F0{i}"
                textPF.selectedNode = nodo
                editxt = session.findById(EDITOR_ID.id)
                texto = editor.get_all_text()
                if texto :
                    #print("El texto no esta vacío. Procediendo a borrarlo... :"f"F0{i}")
                    editxt.SetUnprotectedTextPart(0,".")

            esperar_sap_listo(session)
            valorImpSaludable = get_importesCondiciones(session)
            if valorImpSaludable:
                acciones.append(f"Impuesto Saludable en la posicion {fila+1}0: {valorImpSaludable}")
            
                   
            # da scroll una posicion hacia abajo para no perder visual de los objetos en la tabla de SAP
            set_sap_table_scroll(session, "TC_1211", fila+1)

            print(f"Primera posicion visible : {get_GuiTextField_text(session, f'EBELP[1,0]')}")
            esperar_sap_listo(session)
        # Devuelve las accines ejecutadas en una lista 
        return acciones

    except Exception as e:
        #todo: canbiar por log
        print(f"\nHa ocurrido un error inesperado durante la ejecución: {e}")
        raise

def AbrirSolped(session, solped, item=2):
    """
    Navega en la GUI de SAP para tomar una Solicitud de Pedido (SOLPED) específica
    y prepararla para la creación de una Orden de Compra.

    Args:
        session: La sesión activa de SAP GUI.
        solped (str): El número de la Solicitud de Pedido a procesar.
        item (int): El número de ítems o posiciones que contiene la SOLPED.

    Raises:
        TimeoutError: Si una ventana esperada de SAP no aparece en el tiempo definido.
        Exception: Captura y relanza errores generales durante la interacción con SAP.
    """
    try:
          
        #esperar_sap_listo(session)
        # Click Variante de Seleccion y selecciona el campo Solicitudes de pedido en la lista
        timeout = time.time() + 25
        ventana= "Solicitudes de pedido"
        while not ventana_abierta(session, ventana):
            if time.time() > timeout:
                raise TimeoutError(f"No se abrió la ventana :{ventana}")
            buscar_y_clickear(rf".\img\vSeleccion.png", confidence=0.8, intentos=5, espera=0.5)
            
            # VarianteSeleccion = buscar_objeto_por_id_parcial(session, "/shell[0]")
            # VarianteSeleccion1= buscar_objeto_por_id_parcial(session, "SELECT")
            # VarianteSeleccion.pressContextButton (VarianteSeleccion1.id)
            # esperar_sap_listo(session)
            # SolicitudesdePedido = buscar_objeto_por_id_parcial(session, ":REQ_QUERY")
            # VarianteSeleccion.selectContextMenuItem (SolicitudesdePedido.id)
            
            time.sleep(2)
            pyautogui.press("s") # selecciona el campo Solicitudes de pedido en la lista

        # ingresa el numero de la solped que va a revisar  #Funciona perfecto
        esperar_sap_listo(session)
        session.findById("wnd[0]/usr/ctxtSP$00026-LOW").text = solped
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        # Navegar hasta la sol.pedido en la lista
        buscar_y_clickear(rf".\img\sol.pedido.png", confidence=0.8, intentos=20, espera=0.5)
        # Despliga los itemns de la solped
        time.sleep(0.5)
        pyautogui.hotkey("right")
        time.sleep(0.5)
        pyautogui.hotkey("down")
        time.sleep(0.5)

        # Selecciona todos los items de la solped revisar variable item para ajustar
        with pyautogui.hold("shift"):
            pyautogui.press("down", presses=item)  # Stev: cantidad de items a bajar articulos de la solped
            time.sleep(0.5)

        # Click en tomar pedido
        buscar_y_clickear(rf".\img\tomar.png", confidence=0.7, intentos=20, espera=0.5)

        #Docstring for MostrarCabecera
        MostrarCabecera()


    except Exception as e:
        print(rf"Error en HU05: {e}", "ERROR")
        raise

