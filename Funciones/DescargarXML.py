from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from datetime import date 
from pathlib import Path
import shutil
import time
import os




def login_colsubsidio(usuario, contraseña, ruta):
    chrome_options = Options()
    chrome_options.add_experimental_option("detach", True)
    
    prefs = {
        "download.default_directory": os.getcwd(),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True, 
        "profile.default_content_setting_values.automatic_downloads": 1,
        "safebrowsing.disable_download_protection": False,
        "download.extensions_to_open": "xml",
    }
    chrome_options.add_experimental_option("prefs", prefs)
    

    # y usamos el driver normalmente
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=chrome_options
    )
    
    wait = WebDriverWait(driver, 25)
    driver.get(ruta)
    driver.maximize_window()

    try:
        # --- Lógica de Usuario y Contraseña ---
        user_input = wait.until(EC.presence_of_element_located((By.ID, "txtUser")))
        user_input.send_keys(usuario)

        pass_input = wait.until(EC.presence_of_element_located((By.ID, "txtPass")))
        pass_input.send_keys(contraseña)

        btn_login = wait.until(EC.element_to_be_clickable((By.ID, "btnIngresarLogin")))
        driver.execute_script("arguments[0].click();", btn_login)

        # 3. Esperar el resultado
        wait.until(
            lambda d: "/Home/Index" in d.current_url or 
            len(d.find_elements(By.ID, "lblError")) > 0 or 
            len(d.find_elements(By.ID, "dvCaptcha")) > 0
        )

        # Caso Exitoso
        if "/Home/Index" in driver.current_url:
            print("Login exitoso")
            return driver # Retornamos el objeto para usarlo fuera

        # Manejo de errores (si no entró a /Home/Index)
        if driver.find_elements(By.ID, "dvCaptcha"):
            print("Captcha detectado")
        
        error_elements = driver.find_elements(By.ID, "lblError")
        if error_elements and error_elements[0].text:
            print(f"Error: {error_elements[0].text}")

        return driver

    except Exception as e:
        print(f"Error inesperado: {e}")
        driver.save_screenshot("error.png")
        return driver
    

def realizar_consulta(contador, driver, oc = None, nro_documento = None):

    fecha_actual = date.today()
    fecha_inicial = date(fecha_actual.year, 1, 1)

    # Formato dd/mm/yyyy
    fecha_final_str = fecha_actual.strftime("%Y/%m/%d")
    fecha_inicial_str = fecha_inicial.strftime("%Y/%m/%d")
    wait = WebDriverWait(driver, 20)
    
    try:
        # 1. Click en el menú "Consultas"
        menu_consultas = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Consultas")))
        menu_consultas.click()
        print("Menú Consultas abierto")
        time.sleep(1) 

        # 2. Click en "Recepción validación previa"
        # Usamos PARTIAL_LINK_TEXT por si hay espacios extra
        opcion_recepcion = wait.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "Recepción validación previa")))
        opcion_recepcion.click()
        print("Entrando a Recepción validación previa...")
        time.sleep(1) 

        # --- Esperar a que cargue el formulario ---
        wait.until(EC.presence_of_element_located((By.ID, "tbNumeroDocumento")))
        time.sleep(1) 

        # 3. Llenar Fechas usando JavaScript (ya que están 'disabled')
        # Esto quita el bloqueo y pone el valor directamente
        script_fechas = """
            var f_ini = document.getElementById('dpFechaInicioExpe');
            var f_fin = document.getElementById('dpFechaFinExpe');
            f_ini.removeAttribute('disabled');
            f_fin.removeAttribute('disabled');
            f_ini.value = arguments[0];
            f_fin.value = arguments[1];
        """
        driver.execute_script(script_fechas, fecha_inicial_str, fecha_final_str)
        print(f"Fechas establecidas: {fecha_inicial_str} - {fecha_final_str}")
        time.sleep(1) 

        # 4. Llenar por OC
        if oc:
            criterio_xpath = "//input[@id='ddlCriterioBusqueda']/preceding-sibling::span"
            criterio_dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, criterio_xpath)))
            criterio_dropdown.click()
            print("Dropdown 'Criterio de Búsqueda' abierto")
            time.sleep(2) 

            # Ahora buscamos la opción "Nro Orden de Compra" en la lista desplegable
            if contador ==1:
                opcion_orden = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[text()='Nro Orden de Compra']")))
                opcion_orden.click()
            elif contador > 1:
                opcion_orden = wait.until( EC.presence_of_element_located((By.XPATH, "//li[@data-offset-index='1']")) )
                driver.execute_script("arguments[0].click();", opcion_orden)

                script_fechas = """
                var f_ini = document.getElementById('dpFechaInicioExpe');
                var f_fin = document.getElementById('dpFechaFinExpe');
                f_ini.removeAttribute('disabled');
                f_fin.removeAttribute('disabled');
                f_ini.value = arguments[0];
                f_fin.value = arguments[1];
                """
                driver.execute_script(script_fechas, fecha_inicial_str, fecha_final_str)
                print(f"Fechas establecidas: {fecha_inicial_str} - {fecha_final_str}")
                time.sleep(1)
            
            print("Criterio cambiado a: Nro Orden de Compra")
            time.sleep(1) 

            # Setea valor de la OC en el campo de "Nro de Busqueda"
            nro_input = driver.find_element(By.ID, "tbNumeroIdentificacionEmisor")
            nro_input.clear()
            nro_input.send_keys(oc)
            print(f"Número de documento ingresado: {oc}")
            time.sleep(2) 
        # 5. Click en el botón Buscar
        # Usamos el selector CSS para el botón que tiene la clase btn-success y el texto 
        time.sleep(1)
        btn_buscar = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn-success.btnfix")))
        driver.execute_script("arguments[0].click();", btn_buscar)
        print("Botón Buscar presionado")
        time.sleep(2) 

    except Exception as e:
        print(f"Error durante la consulta: {e}")
        driver.save_screenshot("error_consulta.png")
        
def descargar_xml_final(driver):
    wait = WebDriverWait(driver, 20)
    try:
        # Esperar a que los resultados carguen y el icono sea visible
        xpath_xml = "//img[@title='Presenta el documento en XML firmado digitalmente']"
        xml_icon = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_xml)))
        time.sleep(2) 
        # Click para descargar
        driver.execute_script("arguments[0].click();", xml_icon)
        print("Descarga activada. Verifica tu carpeta.")
        time.sleep(5) 
        #elemento = driver.find_element(By.CLASS_NAME, "logo-page") 
        #elemento.click()
        
    except Exception as e:
        print(f"Error al descargar: {e}")

def mover_archivos(origen, destino, extension='xml'):
    pathOrigen=Path(origen)
    pathDestino=Path(destino)

    pathDestino.mkdir(parents=True, exist_ok=True)

    for archivo in pathOrigen.glob(f"*.{extension}"):
        if archivo.is_file():
            try:
                shutil.move(str(archivo), pathDestino)
                print(f"[+] Archivo movido: {archivo.name}")
            except Exception as e:
                print(f'[-] error al mover el archivo {e}')



import os
from pathlib import Path

def renombrar_archivo(rutaCarpeta, nuevoNombre, extension="xml"):
    carpeta = Path(rutaCarpeta)

    if not carpeta.exists():
        print(f"[-] La carpeta {carpeta} no existe")
        return

    # Obtener todos los archivos con la extensión indicada
    archivos = list(carpeta.glob(f"*.{extension}"))

    if not archivos:
        print("[-] No se encontraron archivos para renombrar")
        return

    # Ordenar por fecha de modificación y tomar el más reciente
    archivos.sort(key=lambda f: f.stat().st_mtime)
    ultimo_archivo = archivos[-1]

    nombre = f"{nuevoNombre}{ultimo_archivo.suffix}"
    nueva_ruta = carpeta / nombre

    try:
        os.rename(ultimo_archivo, nueva_ruta)
        print(f"[+] {ultimo_archivo.name} -> {nombre}")
    except Exception as e:
        print(f"[-] Error al cambiar el nombre del archivo {ultimo_archivo.name}: {e}")
