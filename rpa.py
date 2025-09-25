"""
Módulo para automatización RPA de descarga de envíos GLS utilizando Selenium.
Este módulo proporciona funciones para interactuar con la interfaz web de GLS
y descargar informes de envíos.
"""
import os
import time
from datetime import datetime, timedelta
import logging
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# Importamos múltiples opciones para gestionar el ChromeDriver
try:
    from webdriver_manager.chrome import ChromeDriverManager
    webdriver_manager_available = True
except ImportError:
    webdriver_manager_available = False

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("rpa_shipments.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("Toolstock-GLS RPA")

def load_config():
    """Cargar variables de entorno desde archivo .env"""
    load_dotenv()
    login = os.getenv('URL_LOGIN')
    shipments = os.getenv('URL_SHIPMENTS')
    username_gls = os.getenv('USERNAME_GLS')
    password_gls = os.getenv('PASSWORD_GLS')
    download_folder = os.getenv('PATH_DOWNLOAD_FOLDER')
    final_folder = os.getenv('PATH_FINAL_FOLDER')
    host_db = os.getenv('HOST_DB')
    port_db = os.getenv('PORT_DB')
    database_db = os.getenv('DATABASE_DB')
    user_db = os.getenv('USER_DB')
    password_db = os.getenv('PASSWORD_DB')
    days_ago = os.getenv('DAYS_AGO')


    CONFIG = {
        "urls": {
            "login": login,
            "shipments": shipments,
        },
        "credentials": {
            "username": username_gls,
            "password": password_gls,
        },
        "paths" : {
            "download_folder": download_folder,
            "final_folder":final_folder,
        },
        "timeouts":{
            "page_load": 10,
            "element_present": 15,
        },
        "selenium":{
            "headless": False,
            "disable_images" : True,
            "chromedriver_path" : "drivers/chromedriver.exe"
        },
        "database":{
            "host": host_db,
            "port":port_db,
            "database_name": database_db,
            "user": user_db,
            "password": password_db,
        },
        "time_ago": days_ago
    }
    return CONFIG

def get_current_date_formatted(config):
    """Devuelve la fecha actual en formato dd/mm/yyyy."""
    return (datetime.now() - timedelta(days=config["time_ago"])).strftime("%d/%m/%Y")
     

def get_date_for_filename(config):
    """Devuelve la fecha actual en formato YYYYMMDD para nombre de archivo."""
    return (datetime.now() - timedelta(days=config["time_ago"])).strftime("%Y%m%d")

     

def setup_selenium_driver(config):
    """Configura y devuelve un WebDriver de Selenium para Chrome."""
    try:
        # Configurar opciones de Chrome
        chrome_options = Options()
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("--disable-notifications")
        
        # Configurar directorio de descargas
        prefs = {
            "download.default_directory": config["paths"]["download_folder"],
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }
        chrome_options.add_experimental_option("prefs", prefs)
        
        # Manejo flexible del ChromeDriver
        try:
            # Método 1: Intentar usar webdriver_manager si está disponible
            if webdriver_manager_available:
                try:
                    logger.info("Intentando inicializar Chrome con webdriver_manager...")
                    driver = webdriver.Chrome(
                        service=Service(ChromeDriverManager().install()),
                        options=chrome_options
                    )
                    logger.info("Chrome inicializado exitosamente con webdriver_manager")
                    
                except Exception as e:
                    logger.warning(f"No se pudo inicializar con webdriver_manager: {e}")
                    raise Exception("Fallo al inicializar con webdriver_manager")
            else:
                raise Exception("webdriver_manager no está disponible")
                
        except Exception:
            # Método 2: Intentar usar el ChromeDriver del PATH
            try:
                logger.info("Intentando inicializar Chrome con ChromeDriver del PATH...")
                driver = webdriver.Chrome(options=chrome_options)
                logger.info("Chrome inicializado exitosamente con ChromeDriver del PATH")
                
            except Exception as e:
                logger.warning(f"No se pudo inicializar con ChromeDriver del PATH: {e}")
                
                # Método 3: Buscar ChromeDriver en una ubicación específica
                try:
                    # Verificar si hay una ruta específica en la configuración
                    chromedriver_path = config.get("selenium", {}).get("chromedriver_path", "")
                    
                    if chromedriver_path and os.path.exists(chromedriver_path):
                        logger.info(f"Intentando inicializar Chrome con ChromeDriver en: {chromedriver_path}")
                        driver = webdriver.Chrome(
                            service=Service(executable_path=chromedriver_path),
                            options=chrome_options
                        )
                        logger.info("Chrome inicializado exitosamente con ChromeDriver específico")
                    else:
                        # Si no hay una ruta específica, intentar ubicaciones comunes
                        possible_paths = [
                            "./chromedriver.exe",  # En el directorio actual
                            "./drivers/chromedriver.exe",
                            "C:/chromedriver.exe",
                        ]
                        
                        for path in possible_paths:
                            if os.path.exists(path):
                                logger.info(f"Intentando inicializar Chrome con ChromeDriver en: {path}")
                                driver = webdriver.Chrome(
                                    service=Service(executable_path=path),
                                    options=chrome_options
                                )
                                logger.info(f"Chrome inicializado exitosamente con ChromeDriver en: {path}")
                                break
                        else:
                            raise Exception("No se encontró ChromeDriver en ninguna ubicación común")
                except Exception as e:
                    logger.error(f"Todos los métodos de inicialización fallaron: {e}")
                    raise Exception("No se pudo inicializar Chrome con ningún método disponible")
        
        # Configurar tiempos de espera globales
        driver.set_page_load_timeout(config["timeouts"]["page_load"])
        
        logger.info("Driver de Selenium inicializado correctamente")
        return driver
    
    except Exception as e:
        logger.error(f"Error al configurar el driver de Selenium: {e}")
        return None

def login_to_gls(driver, config):
    """Realiza el login en GLS usando Selenium."""
    try:
        logger.info(f"Navegando a la página de login: {config['urls']['login']}")
        driver.get(config["urls"]["login"])
        
        # Esperar a que cargue la página de login
        WebDriverWait(driver, config["timeouts"]["element_present"]).until(
            EC.presence_of_element_located((By.ID, "usuario"))
        )
        logger.info("Input con id usuario encontrado")
        
        # Ingresar credenciales si están configuradas
        if config["credentials"]["username"]:
            username_field = driver.find_element(By.ID, "usuario")
            username_field.clear()
            username_field.send_keys(config["credentials"]["username"])
            logger.info("Usuario ingresado")
            
        if config["credentials"]["password"]:
            password_field = driver.find_element(By.ID, "pass")
            password_field.clear()
            password_field.send_keys(config["credentials"]["password"])
            logger.info("Password ingresado")
        
        # Hacer clic en el botón de login
        login_button = driver.find_element(By.ID, "Button1")
        login_button.click()
        
        # Esperar a que cargue la página después del login
        WebDriverWait(driver, config["timeouts"]["element_present"]).until(
            EC.url_contains("Extranet")
        )
        
        logger.info("Login exitoso")
        return True
    except Exception as e:
        logger.error(f"Error durante el login: {e}")
        return False

def navigate_to_shipments(driver, config):
    """Navega a la página de búsqueda de envíos."""
    try:
        logger.info(f"Navegando a la página de envíos: {config['urls']['shipments']}")
        driver.get(config["urls"]["shipments"])
        
        # Esperar a que cargue la página de búsqueda
        WebDriverWait(driver, config["timeouts"]["element_present"]).until(
            EC.presence_of_element_located((By.ID, "fechadesde"))
        )
        
        logger.info("Navegación a página de envíos exitosa")
        return True
    except Exception as e:
        logger.error(f"Error al navegar a la página de envíos: {e}")
        return False

def search_shipments(driver, config):
    """Realiza la búsqueda de envíos para la fecha actual."""
    try:
        logger.info("Realizando búsqueda de envíos")
        current_date = get_current_date_formatted(config)
        
        # Localizar e ingresar fechas
        from_date_field = driver.find_element(By.ID, "fechadesde")
        from_date_field.clear()
        from_date_field.send_keys(current_date)
        
        to_date_field = driver.find_element(By.ID, "fechahasta")
        to_date_field.clear()
        to_date_field.send_keys(current_date)
        
        # Iniciar búsqueda
        search_button = driver.find_element(By.ID, "btBuscar")
        search_button.click()
        
        # Esperar a que se complete la búsqueda (puede variar según la página)
        try:
            # Esperamos a que aparezca el botón de exportar o algún mensaje de "no hay resultados"
            WebDriverWait(driver, config["timeouts"]["element_present"]).until(EC.presence_of_element_located((By.ID, "btXLS")))
            logger.info("Búsqueda completada")
            return True
        except TimeoutException:
            logger.warning("Tiempo de espera agotado al buscar el botón de exportación")
            # Intentamos verificar si hay una tabla de resultados
            try:
                results_table = driver.find_element(By.ID, "envios")
                if results_table:
                    logger.info("Se encontró tabla de resultados, continuando")
                    return True
            except NoSuchElementException:
                logger.warning("No se encontró tabla de resultados")
                return False
    except Exception as e:
        logger.error(f"Error al realizar la búsqueda: {e}")
        return False

def export_to_excel(driver, config):
    """Exporta los resultados de la búsqueda a Excel y estandariza el nombre del archivo."""
    try:
        logger.info("Intentando exportar resultados a Excel")
        
        # Verificar si existe el botón de exportar
        try:
            export_button = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "btXLS"))
            )
            
            # Generamos el nombre del archivo estandarizado que usaremos
            date_str = get_date_for_filename(config)
            standardized_filename = f"GLS_{date_str}.xls"
            final_path = os.path.join(config["paths"]["download_folder"], standardized_filename)
            
            # Si ya existe un archivo con ese nombre, lo eliminamos
            if os.path.exists(final_path):
                os.remove(final_path)
                logger.info(f"Archivo existente eliminado: {final_path}")
            
            # Obtener lista de archivos en la carpeta de descargas antes de exportar
            before_files = set(os.listdir(config["paths"]["download_folder"]))
            
            # Hacer clic en el botón de exportar
            export_button.click()
            logger.info("Botón de exportación presionado")
            
            # Esperar a que aparezca un nuevo archivo en la carpeta de descargas (máximo 30 segundos)
            max_wait_time = 30
            wait_time = 0
            downloaded_file = None
            
            while not downloaded_file and wait_time < max_wait_time:
                time.sleep(1)
                wait_time += 1
                
                # Obtener lista actual de archivos
                current_files = set(os.listdir(config["paths"]["download_folder"]))
                
                # Encontrar archivos nuevos
                new_files = current_files - before_files
               
                # Filtrar solo archivos Excel (xls, xlsx, csv)
                excel_extensions = ['.xls', '.xlsx', '.csv']
                new_excel_files = [f for f in new_files if os.path.splitext(f)[1].lower() in excel_extensions]
                
                # Si encontramos algún archivo nuevo Excel
                if new_excel_files:
    
                    # Ordenar por fecha de modificación (el más reciente primero)
                    newest_file = max(
                        [os.path.join(config["paths"]["download_folder"], f) for f in new_excel_files],
                        key=os.path.getmtime
                    )
                    downloaded_file = newest_file
                    break
            
            if downloaded_file:
                logger.info(f"Archivo descargado detectado: {downloaded_file}")
                
                # Renombrar el archivo al nombre estandarizado
                try:
                    # Verificar si el archivo destino ya existe (en caso de que el archivo bajado ya tenga ese nombre)
                    if downloaded_file != final_path:
                        if os.path.exists(final_path):
                            os.remove(final_path)
                        
                        # Esperar a que el archivo esté completamente descargado y no bloqueado
                        max_rename_attempts = 10
                        for attempt in range(max_rename_attempts):
                            try:
                                os.rename(downloaded_file, final_path)
                                logger.info(f"Archivo renombrado exitosamente a: {final_path}")
                                break
                            except PermissionError:
                                # Si el archivo está siendo usado, esperamos un poco
                                if attempt < max_rename_attempts - 1:
                                    logger.info(f"Archivo todavía en uso, esperando... (intento {attempt+1}/{max_rename_attempts})")
                                    time.sleep(1)
                                else:
                                    raise
                    
                    return final_path
                except Exception as e:
                    logger.error(f"Error al renombrar el archivo: {e}")
                    # Si no pudimos renombrar, devolvemos el archivo original
                    return downloaded_file
            else:
                logger.error("No se detectó ningún archivo descargado")
                return None
            
        except TimeoutException:
            # No hay botón de exportar, posiblemente no hay resultados
            logger.warning("No se encontró el botón de exportar, posiblemente no hay resultados")
            
            # Verificar si hay un mensaje de "no hay resultados"
            try:
                no_results_msg = driver.find_element(By.ID, "envios")
                if no_results_msg:
                    logger.info("No hay resultados para exportar")
            except NoSuchElementException:
                logger.warning("No se encontró mensaje de 'no hay resultados'")
            
            return None
    except Exception as e:
        logger.error(f"Error al exportar a Excel: {e}")
        return None

def process_excel_file(excel_file_path, config):
    """Procesa el archivo descargado y lo convierte a XLSX."""
    try:
        if not excel_file_path or not os.path.exists(excel_file_path):
            logger.error(f"No se puede procesar un archivo que no existe: {excel_file_path}")
            return False
            
        logger.info(f"Procesando archivo descargado: {excel_file_path}")
        date_str = get_date_for_filename(config)
        
        # Nombre del archivo destino
        final_path = os.path.join(config["paths"]["final_folder"], f"{date_str}.xlsx")
        
        # Detectar el formato real del archivo
        # Primero leemos los primeros bytes para determinar si es realmente HTML
        is_html = False
        try:
            with open(excel_file_path, 'r', encoding='utf-8') as f:
                first_chunk = f.read(1000).lower()
                if '<!doctype html>' in first_chunk or '<html' in first_chunk or '<form' in first_chunk:
                    is_html = True
                    logger.info("Archivo detectado como HTML aunque tiene extensión XLS")
        except UnicodeDecodeError:
            # Si falla la decodificación como texto, probablemente es un archivo binario
            is_html = False
        
        # Si el archivo es realmente HTML, usamos pandas con el motor html
        if is_html:
            try:
                import pandas as pd
                
                # Leer el archivo como HTML
                logger.info("Procesando archivo como HTML con pandas")
                tables = pd.read_html(excel_file_path)
                
                if tables and len(tables) > 0:
                    # Por lo general, la tabla principal es la primera
                    df = tables[0]
                    
                    # Guardar como XLSX
                    df.to_excel(final_path, index=False)
                    logger.info(f"Archivo HTML convertido y guardado como XLSX: {final_path}")
                    return True
                else:
                    logger.error("No se encontraron tablas en el archivo HTML")
            except Exception as e:
                logger.error(f"Error al procesar el archivo HTML con pandas: {e}")
                
                # Si pandas falla, intentamos una aproximación alternativa - extraer la tabla con BeautifulSoup
                try:
                    logger.info("Intentando extraer tabla con BeautifulSoup")
                    from bs4 import BeautifulSoup
                    from openpyxl import Workbook  # Importar la librería para trabajar con Excel
                    
                    # Leer el archivo HTML
                    with open(excel_file_path, 'r', encoding='utf-8') as f:
                        soup = BeautifulSoup(f, 'html.parser')
                    
                    # Buscar la tabla (ajustar selectores según el HTML específico de GLS)
                    table = soup.find('table')
                    
                    if table:
                        # Crear un nuevo libro de Excel y seleccionar la hoja activa
                        wb = Workbook()
                        ws = wb.active
                        
                        # Extraer encabezados
                        headers = [th.text.strip() for th in table.find_all('th')]
                        if headers:
                            ws.append(headers)  # Escribir los encabezados en la primera fila
                        
                        # Extraer filas
                        for row in table.find_all('tr'):
                            cells = [td.text.strip() for td in row.find_all('td')]
                            if cells:
                                ws.append(cells)  # Escribir cada fila en el archivo de Excel
                        
                        # Guardar el archivo de Excel
                        wb.save(final_path)
                        
                        logger.info(f"Tabla extraída con BeautifulSoup y guardada como Excel: {final_path}")
                        return True
                    else:
                        logger.error("No se encontró ninguna tabla en el HTML con BeautifulSoup")
                except Exception as e:
                    logger.error(f"Error al extraer tabla con BeautifulSoup: {e}")
                    
                    # Si todo lo demás falla, guardamos una copia del HTML original
                    try:
                        import shutil
                        html_path = os.path.join(config["paths"]["final_folder"], f"GLS_{date_str}.html")
                        shutil.copy2(excel_file_path, html_path)
                        logger.info(f"Se guardó una copia del HTML original: {html_path}")
                        
                        # También podemos intentar abrir el HTML en un navegador y exportarlo manualmente
                        # Esto dependerá de si tienes una implementación específica para esto
                        return False
                    except Exception as copy_error:
                        logger.error(f"Error al copiar el archivo HTML: {copy_error}")
                        return False
        
        # Si no es HTML, intentamos procesarlo como un archivo Excel normal
        else:
            try:
                import pandas as pd
                
                # Intentar con diferentes engines
                logger.info("Intentando procesar como Excel verdadero")
                engines = ['openpyxl', 'xlrd', 'pyxlsb', None]
                
                for engine in engines:
                    try:
                        if engine:
                            logger.info(f"Intentando con motor: {engine}")
                            df = pd.read_excel(excel_file_path, engine=engine)
                        else:
                            logger.info("Intentando sin especificar motor")
                            df = pd.read_excel(excel_file_path)
                        
                        # Si llegamos aquí, la lectura fue exitosa
                        df.to_excel(final_path, index=False)
                        logger.info(f"Archivo convertido y guardado exitosamente: {final_path}")
                        return True
                    except Exception as e:
                        logger.warning(f"Falló con motor {engine}: {e}")
                        continue
                
                # Si llegamos aquí, todos los intentos fallaron
                logger.error("Todos los intentos de leer el archivo como Excel fallaron")
                return False
                
            except Exception as e:
                logger.error(f"Error general al procesar el archivo: {e}")
                return False
            
    except Exception as e:
        logger.error(f"Error al procesar el archivo: {e}")
        return False

def conection_db(config):
    import sqlalchemy
    import urllib.parse
    """ Personal connection string """
    try:
        
        pass_encoded = urllib.parse.quote_plus(config["database"]["password"])
        engine = sqlalchemy.create_engine("mysql+pymysql://" 
                                        + config["database"]["user"] 
                                        + ":" 
                                        + pass_encoded 
                                        + "@" 
                                        + config["database"]["host"]
                                        + ":"
                                        + config["database"]["port"]  
                                        + "/" 
                                        + config["database"]["database_name"],
                                        pool_size=0,
                                        max_overflow=-1)


        return engine

    except Exception as c:
        print(c)

def get_data_ps(config):
    import pandas as pd
    c = conection_db(config)

    df_orders_ps = pd.read_sql("""select 
            CASE WHEN marketplace_order_id IS NULL THEN reference 
            ELSE marketplace_order_id 
            END AS marketplace_order_id, 
        o.id_order as id_order_ps, reference as reference_ps 
        from ps_orders o
            left join (
                select id_order, marketplace_order_id  
                from toolstock_ps.ps_beezup_order
            ) bo
            on bo.id_order = o.id_order
        where current_state in (1, 2, 3, 6, 10, 11, 14, 15, 16, 21)""", c)
    
    return df_orders_ps

def updated_excel(config):
    import pandas as pd
    try:
        path_file = os.path.join(config["paths"]["final_folder"], f"{get_date_for_filename()}.xlsx")

        # Cargar el archivo Excel original
        df_excel = pd.read_excel(path_file)
        
        # Cargar el dataframe de referencia
        df_referencia = get_data_ps(config)
        
        # Crear las nuevas columnas con valores vacíos
        df_excel['id_order_ps'] = ''
        df_excel['reference_ps'] = ''
        
        # Iterar sobre las filas del Excel original
        for index, fila in df_excel.iterrows():
            valor_dpto_dst = fila['DptoDst']
            
            # Buscar coincidencia en el dataframe de referencia
            coincidencia_marketplace = df_referencia[df_referencia['marketplace_order_id'] == valor_dpto_dst]
            coincidencia_reference = df_referencia[df_referencia['reference_ps'] == valor_dpto_dst]
            
            # Si hay coincidencia con marketplace_order_id
            if not coincidencia_marketplace.empty:
                df_excel.at[index, 'id_order_ps'] = coincidencia_marketplace['id_order_ps'].values[0]
                df_excel.at[index, 'reference_ps'] = coincidencia_marketplace['reference_ps'].values[0]
            
            # Si hay coincidencia con reference_ps
            elif not coincidencia_reference.empty:
                df_excel.at[index, 'id_order_ps'] = coincidencia_reference['id_order_ps'].values[0]
                df_excel.at[index, 'reference_ps'] = coincidencia_reference['reference_ps'].values[0]
        
        # Guardar los cambios al archivo Excel
        df_excel.to_excel(path_file, index=False)
        print(f"El archivo '{path_file}' ha sido actualizado con éxito.")

        return True
    
    except Exception as e:
        logger.error(f"Error al actualizar el archivo Excel: {e}")
        return False

def rpa_shipments():
    """Función principal que ejecuta el flujo completo de RPA para envíos GLS."""
    driver = None
    try:
        # Cargar configuración
        config = load_config()
        
        # Configurar el driver de Selenium
        driver = setup_selenium_driver(config)
        if not driver:
            return False
        
        # Realizar login en GLS
        if not login_to_gls(driver, config):
            return False
        
        # Navegar a la página de búsqueda de envíos
        if not navigate_to_shipments(driver, config):
            return False
        
        # Realizar búsqueda de envíos
        if not search_shipments(driver, config):
            return False
        
        # Exportar resultados a Excel
        excel_file_path = export_to_excel(driver, config)
        
        # Si hay archivo para procesar, lo convertimos a XLSX
        if excel_file_path:
            result = process_excel_file(excel_file_path, config)
            
            if result and os.path.exists(excel_file_path):

                # Opcional: eliminar el archivo Excel original después de procesar
                updated_file = updated_excel(config)
                if updated_file and os.path.exists(excel_file_path):
                    try:
                        os.remove(excel_file_path)
                        logger.info(f"Archivo original eliminado: {excel_file_path}")
                    except:
                        logger.warning(f"No se pudo eliminar el archivo original: {excel_file_path}")
            
            return result
        else:
            logger.info("No hay archivos para procesar")
            return True  # Consideramos éxito aunque no haya archivos (puede ser normal)
        
    except Exception as e:
        logger.error(f"Error en el proceso RPA: {e}")
        return False
    finally:
        # Cerrar el driver de Selenium
        if driver:
            try:
                driver.quit()
                logger.info("Driver de Selenium cerrado correctamente")
            except:
                logger.warning("Error al cerrar el driver de Selenium")