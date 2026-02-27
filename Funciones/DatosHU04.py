import time
from datetime import datetime

def buscar_por_id_parcial(contenedor, id_final):
    """Busca un objeto en SAP cuyo ID termine en una cadena específica."""
    try:
        for hijo in contenedor.Children:
            if hijo.Id.endswith(id_final):
                return hijo
            if hijo.Children.Count > 0:
                resultado = buscar_por_id_parcial(hijo, id_final)
                if resultado:
                    return resultado
    except:
        pass
    return None

def consultar_datos_hu04(sesion, numeroOC):
    # 1. Inicialización de variables
    fecha_creacion = datetime.now()
    dias_mora = 0
    facturada = False
    
    # EL ID EXACTO QUE FUNCIONÓ
    ID_FECHA_TOPLINE = "wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-BEDAT"

    try:
        # 2. Navegación y Carga de OC
        sesion.findById("wnd[0]/tbar[0]/okcd").text = "/nME23N"
        sesion.findById("wnd[0]").sendVKey(0)
        time.sleep(1)
        
        sesion.findById("wnd[0]/tbar[1]/btn[17]").press()
        sesion.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN").text = str(numeroOC)
        sesion.findById("wnd[1]").sendVKey(0)
        
        # Espera de carga y salto de posibles advertencias (Enter)
        time.sleep(1.5)
        sesion.findById("wnd[0]").sendVKey(0) 

        # 3. LECTURA DE FECHA USANDO ID FIJO
        try:
            campo_fecha = sesion.findById(ID_FECHA_TOPLINE)
            # Intentamos .text y si no .value
            fecha_texto = campo_fecha.text.strip() if campo_fecha.text else campo_fecha.value.strip()
            
            if fecha_texto:
                fecha_creacion = datetime.strptime(fecha_texto, "%d.%m.%Y")
                dias_mora = (datetime.now() - fecha_creacion).days
        except Exception as e:
            print(f"   [!] Error al acceder al ID de fecha: {e}")

        # 4. BÚSQUEDA DE FACTURACIÓN (HISTORIAL)
        # El historial sí cambia de posición, por eso usamos búsqueda dinámica aquí
        area_usuario = sesion.findById("wnd[0]/usr")
        grid_historial = buscar_por_id_parcial(area_usuario, "cntlGRIDCONTROL/shellcont/shell")


        
        if grid_historial:
            try:
                for i in range(grid_historial.RowCount):
                    tipo_doc = grid_historial.getCellValue(i, "BEZTP")
                    if tipo_doc in ["RE-L", "Factura", "Fact.rec."]:
                        facturada = True
                        break
            except:
                facturada = False

        return {
            "status": "OK",
            "fecha_sap": fecha_creacion.strftime("%d/%m/%Y"),
            "dias": dias_mora,
            "facturada": facturada
        }

    except Exception as e:
        return {"status": "Error", "detalle": str(e)}