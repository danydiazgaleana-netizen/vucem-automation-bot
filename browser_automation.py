"""
Módulo de automatización del navegador para el portal VUCEM.
Soporta modo interactivo (con pausas para validación humana) y modo automático.
"""
import logging
import time
import re
import os
import glob
from pathlib import Path
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    NoSuchWindowException,
    InvalidSessionIdException,
    WebDriverException,
    ElementClickInterceptedException,
    StaleElementReferenceException
)
from selenium.webdriver.chrome.service import Service
from config import Config


class SessionExpiredException(Exception):
    pass


def retry_on_exception(max_retries=3, delay=1, backoff=2):
    def decorator(func):
        def wrapper(*args, **kwargs):
            last_exception = None
            for attempt in range(1, max_retries + 1):
                try:
                    return func(*args, **kwargs)
                except (TimeoutException,
                        NoSuchElementException,
                        WebDriverException,
                        ElementClickInterceptedException,
                        StaleElementReferenceException) as e:
                    logging.warning(f"⚠️ Intento {attempt}/{max_retries} falló en {func.__name__}: {e}")
                    last_exception = e
                    if attempt < max_retries:
                        wait = delay * (backoff ** (attempt - 1))
                        time.sleep(wait)
                    else:
                        logging.error(f"❌ {func.__name__} falló después de {max_retries} intentos.")
                        raise
            raise last_exception
        return wrapper
    return decorator


class VUCEMAutomation:
    SELECTORS = {
        'dropdown_tramites': (By.XPATH, "//a[contains(@class,'dropdown-toggle') and contains(.,'Trámites')]"),
        'solicitudes_nuevas': (By.XPATH, "//ul[contains(@class,'dropdown-menu')]//a[contains(.,'Solicitudes nuevas')]"),
        'secretaria_economia': (By.CSS_SELECTOR, "a[title='Secretaría de Economía']"),
        'tramite_110101': (By.XPATH, "//a[contains(@onclick,'110101')]"),
        'tramite_110101_texto': (By.XPATH, "//a[contains(.,'Certificado de Origen')]"),
        'tramite_ue_texto': (By.XPATH, "//a[contains(.,'UE')]"),
        'tramite_110101_href': (By.XPATH, "//a[contains(@href,'110101')]"),
        'tramite_110101_num': (By.XPATH, "//a[contains(.,'110101')]"),
        'tab_tratados': (By.XPATH, "//a[@href='#tabs-3']"),
        'pais_selector': (By.ID, "solicitud.clavePaisSeleccionado"),
        'tratado_selector': (By.ID, "solicitud.idTratadoAcuerdoSeleccionado"),
        'tratado_opcion_100': (By.XPATH, "//select[@id='solicitud.idTratadoAcuerdoSeleccionado']/option[@value='100']"),
        'criterio_selector': (By.ID, "solicitud.idCriterioCertificadoSeleccionado"),
        'criterio_opcion_b': (By.XPATH, "//select[@id='solicitud.idCriterioCertificadoSeleccionado']/option[contains(text(), 'A partir de materiales originarios')]"),
        'agregar_tratado': (By.XPATH, "//input[@value='Agregar tratado']"),
        'pestania_mercancia': (By.XPATH, "//label[@for='captura.tab.productoCertificado.producto.gobmx']"),
        'nombre_comercial': (By.ID, "solicitud.registroCuestionario.mercanciaAsociada.nombreComercial"),
        'nombre_ingles': (By.ID, "solicitud.registroCuestionario.mercanciaAsociada.nombreIngles"),
        'fraccion': (By.ID, "solicitud.registroCuestionario.mercanciaAsociada.fraccionArancelaria.clave"),
        'precio': (By.ID, "solicitud.registroCuestionario.mercanciaAsociada.precioFrancoFabrica"),
        'boton_carga': (By.ID, "botonCargaArchivo"),
        'campo_archivo': (By.ID, "archivoAdjuntar"),
        'boton_enviar_archivo': (By.ID, "cargaArchivo"),
        'pestania_adicionales': (By.XPATH, "//label[@for='captura.tab.productoCertificado.adicionales.gobmx']"),
        'exportador_autorizado': (By.XPATH, "//input[@name='solicitud.registroCuestionario.solicitaExportadorAutorizado' and @value='true']"),
        'producto_artesanal': (By.XPATH, "//input[@name='solicitud.registroCuestionario.condicionExportador' and @value='PRODUCTO_ARTESANAL']"),
        'entidad': (By.ID, "solicitud.entidadFederativa.entidad.clave"),
        'boton_continuar': (By.ID, "guardarSolicitud"),
        'acuse_titulo': (By.XPATH, "//h1[contains(text(),'Acuse de Recibo')]"),
        'acuse_mensaje': (By.XPATH, "//div[@class='ui-state-highlight ui-corner-all alert alert-success']"),
        'nueva_captura': (By.ID, "showDatosCaptura"),
        'dashboard': (By.XPATH, "//a[contains(@class,'dropdown-toggle') and contains(.,'Trámites')]"),
        'formulario_cargado': (By.XPATH, "//label[@for='captura.tab.solicitante']"),
    }

    def __init__(self):
        self.driver = None
        self.wait = None
        self.modo_simulacion = Config.MODO_SIMULACION
        self._firma_msg_mostrado = False
        self.carpeta_archivos = Config.OUTPUT_DIR
        self.ruta_excel = Config.MASTER_EXCEL_PATH
        self.hoja_excel = "2026"
        if not self.modo_simulacion:
            self._setup_driver()

    def _setup_driver(self):
        options = webdriver.ChromeOptions()
        if Config.HEADLESS:
            options.add_argument("--headless")
        options.add_argument("--start-maximized")
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")

        if hasattr(Config, 'CHROME_DRIVER_PATH') and Config.CHROME_DRIVER_PATH:
            service = Service(executable_path=Config.CHROME_DRIVER_PATH)
            self.driver = webdriver.Chrome(service=service, options=options)
        else:
            self.driver = webdriver.Chrome(options=options)

        self.driver.maximize_window()
        self.wait = WebDriverWait(self.driver, Config.DEFAULT_TIMEOUT)
        logging.info("✅ WebDriver inicializado.")

    # --------------------------------------------------------------
    # FUNCIONES DE EXCEL (con manejo de duplicados)
    # --------------------------------------------------------------
    def _buscar_datos_excel(self, nombre_comercial, modo_interactivo=True):
        """
        Busca el nombre en el Excel. Si hay múltiples coincidencias, permite elegir.
        Retorna (datos, codigo_seleccionado) o (None, None).
        """
        try:
            if not self.ruta_excel.exists():
                logging.error(f"❌ No se encontró el archivo: {self.ruta_excel}")
                return None, None
            df = pd.read_excel(self.ruta_excel, sheet_name=self.hoja_excel, header=None)
            df[1] = df[1].astype(str).str.strip()
            nombre_buscado = nombre_comercial.strip()
            mascara = df[1].str.lower() == nombre_buscado.lower()
            filas = df[mascara]
            if filas.empty:
                logging.warning(f"⚠️ No se encontró el nombre '{nombre_comercial}' en el Excel.")
                return None, None
            if len(filas) == 1:
                fila = filas.iloc[0]
                fraccion = str(fila[5]).strip()
                precio = str(fila[2]).strip()
                num_insumos = int(fila[6]) if fila[6] else 0
                codigo = str(fila[1]).strip()
                logging.info(f"✅ Excel: Fracción={fraccion}, Precio={precio}, Insumos={num_insumos}")
                return {
                    "fraccion": fraccion,
                    "precio": precio,
                    "num_insumos": num_insumos,
                    "codigo": codigo
                }, codigo
            else:
                # Múltiples coincidencias
                print("\n" + "=" * 60)
                print(f"⚠️ Se encontraron {len(filas)} modelos con el nombre '{nombre_comercial}':")
                for idx, (i, row) in enumerate(filas.iterrows(), 1):
                    cod = str(row[1]).strip()
                    print(f"  {idx}. Código: {cod}")
                print("=" * 60)
                if modo_interactivo:
                    while True:
                        try:
                            opcion = int(input("Selecciona el número correcto: "))
                            if 1 <= opcion <= len(filas):
                                fila = filas.iloc[opcion - 1]
                                fraccion = str(fila[5]).strip()
                                precio = str(fila[2]).strip()
                                num_insumos = int(fila[6]) if fila[6] else 0
                                codigo = str(fila[1]).strip()
                                return {
                                    "fraccion": fraccion,
                                    "precio": precio,
                                    "num_insumos": num_insumos,
                                    "codigo": codigo
                                }, codigo
                            else:
                                print("Número fuera de rango.")
                        except ValueError:
                            print("Entrada inválida. Ingresa un número.")
                else:
                    fila = filas.iloc[0]
                    fraccion = str(fila[5]).strip()
                    precio = str(fila[2]).strip()
                    num_insumos = int(fila[6]) if fila[6] else 0
                    codigo = str(fila[1]).strip()
                    logging.info(f"🤖 Modo automático: seleccionado el primero (código {codigo})")
                    return {
                        "fraccion": fraccion,
                        "precio": precio,
                        "num_insumos": num_insumos,
                        "codigo": codigo
                    }, codigo
        except Exception as e:
            logging.error(f"❌ Error al leer Excel: {e}")
            return None, None

    def _limpiar_nombre_para_busqueda(self, nombre):
        nombre_limpio = re.sub(r'[^\w\s-]', ' ', nombre)
        nombre_limpio = re.sub(r'[-\s]+', '_', nombre_limpio)
        return nombre_limpio.strip('_')

    def _contar_filas_csv(self, ruta_archivo):
        try:
            df = pd.read_csv(ruta_archivo, header=None, encoding='utf-8')
            df = df.dropna(how='all')
            return len(df)
        except Exception as e:
            logging.error(f"❌ Error al leer CSV: {e}")
            return 0

    def _seleccionar_archivo_insumos(self, nombre_ingles, datos_excel=None):
        if not nombre_ingles:
            logging.warning("⚠️ No se proporcionó nombre en inglés, no se puede buscar el archivo.")
            return

        try:
            boton_carga = self.wait.until(EC.element_to_be_clickable((By.ID, "botonCargaArchivo")))
            self.driver.execute_script("arguments[0].click();", boton_carga)
            time.sleep(1)

            campo_file = self.wait.until(EC.presence_of_element_located((By.ID, "archivoAdjuntar")))
            self.wait.until(EC.visibility_of(campo_file))

            nombre_normalizado = self._limpiar_nombre_para_busqueda(nombre_ingles)
            logging.info(f"🔍 Buscando archivo para: '{nombre_ingles}' (normalizado: '{nombre_normalizado}')")
            logging.info(f"   📁 Carpeta: {self.carpeta_archivos}")

            archivos_csv = []
            archivos_xlsx = []

            patron_csv = f"PLANTILLA_{nombre_ingles.replace(' ', '_')}.*\\.csv"
            archivos_csv = glob.glob(str(self.carpeta_archivos / patron_csv))

            if not archivos_csv:
                patron_csv_norm = f"PLANTILLA_{nombre_normalizado}.*\\.csv"
                archivos_csv = glob.glob(str(self.carpeta_archivos / patron_csv_norm))

            if not archivos_csv:
                patron_xlsx = f"PLANTILLA_{nombre_ingles.replace(' ', '_')}.*\\.xlsx"
                archivos_xlsx = glob.glob(str(self.carpeta_archivos / patron_xlsx))
                if not archivos_xlsx:
                    patron_xlsx_norm = f"PLANTILLA_{nombre_normalizado}.*\\.xlsx"
                    archivos_xlsx = glob.glob(str(self.carpeta_archivos / patron_xlsx_norm))

            if not archivos_csv and not archivos_xlsx:
                todos = glob.glob(str(self.carpeta_archivos / "PLANTILLA_*"))
                nombre_buscado_limpio = self._limpiar_nombre_para_busqueda(nombre_ingles).lower()
                for arch in todos:
                    nombre_arch = os.path.basename(arch)
                    nombre_sin_prefijo = nombre_arch[len("PLANTILLA_"):]
                    nombre_base = os.path.splitext(nombre_sin_prefijo)[0]
                    nombre_base_limpio = self._limpiar_nombre_para_busqueda(nombre_base).lower()
                    if nombre_base_limpio == nombre_buscado_limpio:
                        if arch.endswith('.csv'):
                            archivos_csv.append(arch)
                        elif arch.endswith('.xlsx'):
                            archivos_xlsx.append(arch)
                        break

            ruta_archivo = None

            if archivos_csv:
                ruta_archivo = archivos_csv[0]
                logging.info(f"✅ Archivo CSV encontrado: {os.path.basename(ruta_archivo)}")
            elif archivos_xlsx:
                xlsx_path = archivos_xlsx[0]
                logging.info(f"📄 Archivo XLSX encontrado: {os.path.basename(xlsx_path)}")
                logging.info("🔄 Convirtiendo XLSX a CSV temporal...")
                try:
                    df = pd.read_excel(xlsx_path, header=None)
                    temp_csv = xlsx_path.replace('.xlsx', '_temp.csv')
                    df.to_csv(temp_csv, index=False, header=False)
                    ruta_archivo = temp_csv
                    logging.info(f"✅ Archivo convertido a: {os.path.basename(temp_csv)}")
                except Exception as e:
                    logging.error(f"❌ Error al convertir XLSX: {e}")
                    return
            else:
                logging.warning(f"⚠️ No se encontró ningún archivo para: {nombre_ingles}")
                return

            if datos_excel and 'num_insumos' in datos_excel and ruta_archivo:
                num_esperado = datos_excel['num_insumos']
                filas_csv = self._contar_filas_csv(ruta_archivo)
                logging.info(f"📊 El archivo tiene {filas_csv} filas. Se esperaban {num_esperado}.")
                if filas_csv != num_esperado:
                    logging.warning("⚠️ ADVERTENCIA: El número de insumos no coincide. Se subirá igual.")

            campo_file.send_keys(os.path.abspath(ruta_archivo))
            logging.info("✅ Archivo seleccionado en el campo.")
            time.sleep(1)

            boton_enviar = self.wait.until(EC.element_to_be_clickable((By.ID, "cargaArchivo")))
            self.driver.execute_script("arguments[0].click();", boton_enviar)
            logging.info("✅ Click en 'Enviar' realizado. Esperando procesamiento...")
            time.sleep(5)

            try:
                self.wait.until(EC.invisibility_of_element_located((By.ID, "archivoAdjuntar")))
                logging.info("✅ Modal de carga cerrado.")
            except:
                logging.warning("⚠️ El modal no se cerró automáticamente, pero la subida podría haberse completado.")

            if ruta_archivo and ruta_archivo.endswith('_temp.csv'):
                try:
                    os.remove(ruta_archivo)
                    logging.info("   🗑️ Archivo temporal eliminado.")
                except:
                    pass

        except Exception as e:
            logging.error(f"❌ Error al seleccionar o enviar el archivo: {e}")
            import traceback
            traceback.print_exc()
            raise

    # --------------------------------------------------------------
    # Login y navegación
    # --------------------------------------------------------------
    def login_with_manual_verification(self):
        if self.modo_simulacion:
            logging.info("🔐 [SIMULACIÓN] Login manual simulado. Presiona Enter para continuar...")
            input()
            return

        logging.info("🌐 Navegando a la página de login de VUCEM...")
        self.driver.get(Config.VUCEM_URL)
        input("🔐 Inicia sesión manualmente, resuelve CAPTCHA/2FA y presiona ENTER cuando estés listo...")

        try:
            self.wait.until(EC.presence_of_element_located(self.SELECTORS['dashboard']))
            logging.info("✅ Login verificado.")
        except TimeoutException:
            logging.warning("⚠️ No se detectó el dashboard, pero se continuará.")

    @retry_on_exception(max_retries=3, delay=1)
    def _click_solicitudes_nuevas(self):
        dropdown = self.wait.until(EC.element_to_be_clickable(self.SELECTORS['dropdown_tramites']))
        dropdown.click()
        self.wait.until(EC.visibility_of_element_located(self.SELECTORS['solicitudes_nuevas']))
        solicitudes = self.wait.until(EC.element_to_be_clickable(self.SELECTORS['solicitudes_nuevas']))
        solicitudes.click()
        logging.info("✅ Click en 'Solicitudes nuevas'")
        self.wait.until(EC.presence_of_element_located(self.SELECTORS['secretaria_economia']))

    @retry_on_exception(max_retries=3, delay=1)
    def _click_secretaria_economia(self):
        se_btn = self.wait.until(EC.element_to_be_clickable(self.SELECTORS['secretaria_economia']))
        se_btn.click()
        logging.info("✅ Click en Secretaría de Economía")

    @retry_on_exception(max_retries=3, delay=2)
    def _click_tramite_110101(self):
        selectores = [
            self.SELECTORS['tramite_110101'],
            self.SELECTORS['tramite_110101_texto'],
            self.SELECTORS['tramite_ue_texto'],
            self.SELECTORS['tramite_110101_href'],
            self.SELECTORS['tramite_110101_num']
        ]
        for selector in selectores:
            try:
                tramite = self.wait.until(EC.presence_of_element_located(selector))
                self.driver.execute_script("arguments[0].scrollIntoView(true);", tramite)
                time.sleep(0.3)
                self.driver.execute_script("arguments[0].click();", tramite)
                logging.info(f"✅ Click en trámite 110101 con selector: {selector}")
                return
            except TimeoutException:
                continue
        raise TimeoutException("No se pudo encontrar el trámite 110101.")

    def _click_tab_tratados(self):
        """
        Hace clic en la pestaña Tratados con espera robusta y reintento.
        """
        try:
            tab_tratados = WebDriverWait(self.driver, 15).until(
                EC.element_to_be_clickable(self.SELECTORS['tab_tratados'])
            )
        except TimeoutException:
            # Si no es clickeable, intentar con presence y clic forzado
            tab_tratados = self.wait.until(EC.presence_of_element_located(self.SELECTORS['tab_tratados']))
            logging.warning("⚠️ Pestaña Tratados no clickeable, forzando clic por JS.")

        self.driver.execute_script("arguments[0].scrollIntoView(true);", tab_tratados)
        time.sleep(0.3)
        self.driver.execute_script("arguments[0].click();", tab_tratados)
        logging.info("✅ Pestaña 'Tratados'")
        self.wait.until(EC.presence_of_element_located(self.SELECTORS['pais_selector']))

    @retry_on_exception(max_retries=3, delay=1)
    def _seleccionar_union_europea(self):
        select_pais = Select(self.wait.until(EC.presence_of_element_located(self.SELECTORS['pais_selector'])))
        select_pais.select_by_visible_text("Union Europea")
        logging.info("✅ Unión Europea seleccionada")

    @retry_on_exception(max_retries=3, delay=1)
    def _seleccionar_tlcuem(self):
        self.wait.until(EC.presence_of_element_located(self.SELECTORS['tratado_opcion_100']))
        select_tratado = Select(self.wait.until(EC.presence_of_element_located(self.SELECTORS['tratado_selector'])))
        select_tratado.select_by_value("100")
        logging.info("✅ TLCUEM seleccionado")

    @retry_on_exception(max_retries=3, delay=1)
    def _seleccionar_criterio_b(self):
        self.wait.until(EC.presence_of_element_located(self.SELECTORS['criterio_opcion_b']))
        select_criterio = Select(self.wait.until(EC.presence_of_element_located(self.SELECTORS['criterio_selector'])))
        select_criterio.select_by_visible_text("A partir de materiales originarios (B)")
        logging.info("✅ Criterio B seleccionado")

    @retry_on_exception(max_retries=3, delay=1)
    def _click_agregar_tratado(self):
        boton = self.wait.until(EC.element_to_be_clickable(self.SELECTORS['agregar_tratado']))
        boton.click()
        logging.info("✅ Tratado agregado")
        time.sleep(2)

    # --------------------------------------------------------------
    # CONFIGURAR TRATADOS (se usa al inicio y después de Nueva Captura)
    # --------------------------------------------------------------
    def _configurar_tratados(self):
        """
        Configura los tratados en la pestaña Tratados.
        """
        logging.info("🔄 Configurando tratados...")

        # 1. Ir a la pestaña Tratados
        for intento in range(3):
            try:
                self._click_tab_tratados()
                break
            except Exception as e:
                logging.warning(f"⚠️ Intento {intento+1} para pestaña Tratados falló: {e}")
                time.sleep(2)
                if intento == 2:
                    raise

        # 2. Seleccionar Unión Europea
        self._seleccionar_union_europea()

        # 3. Seleccionar TLCUEM
        self._seleccionar_tlcuem()

        # 4. Seleccionar Criterio B
        self._seleccionar_criterio_b()

        # 5. Agregar tratado
        self._click_agregar_tratado()

        logging.info("✅ Tratados configurados correctamente.")

    def _click_pestania_mercancia(self):
        pestania = self.wait.until(EC.presence_of_element_located(self.SELECTORS['pestania_mercancia']))
        self.driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(0.3)
        self.driver.execute_script("arguments[0].scrollIntoView(true);", pestania)
        time.sleep(0.3)
        self.driver.execute_script("arguments[0].click();", pestania)
        logging.info("✅ Pestaña 'Datos de la mercancía'")
        self.wait.until(EC.presence_of_element_located(self.SELECTORS['nombre_comercial']))

    def navigate_to_form(self):
        if self.modo_simulacion:
            logging.info("🧭 [SIMULACIÓN] Navegación simulada.")
            return

        logging.info("🧭 Iniciando navegación real al formulario...")
        self._click_solicitudes_nuevas()
        self._click_secretaria_economia()
        self._click_tramite_110101()
        # Configurar tratados por primera vez
        self._configurar_tratados()
        # Ir a la pestaña Datos de la mercancía
        self._click_pestania_mercancia()
        logging.info("✅ Navegación completada. Formulario listo.")

    # --------------------------------------------------------------
    # Métodos auxiliares con reintentos
    # --------------------------------------------------------------
    def _safe_operation(self, func, max_retries=3, delay=1):
        for attempt in range(1, max_retries + 1):
            try:
                return func()
            except (TimeoutException,
                    NoSuchElementException,
                    WebDriverException,
                    ElementClickInterceptedException,
                    StaleElementReferenceException) as e:
                logging.warning(f"⏳ Intento {attempt}/{max_retries} falló: {e}")
                if attempt == max_retries:
                    raise
                time.sleep(delay * attempt)
        return None

    def _fill_nombre_comercial(self, nombre):
        campo = self.wait.until(EC.presence_of_element_located(self.SELECTORS['nombre_comercial']))
        campo.clear()
        campo.send_keys(nombre)
        campo_ingles = self.wait.until(EC.presence_of_element_located(self.SELECTORS['nombre_ingles']))
        campo_ingles.clear()
        campo_ingles.send_keys(nombre)

    def _fill_fraccion_precio(self, fraccion, precio):
        campo_fraccion = self.wait.until(EC.presence_of_element_located(self.SELECTORS['fraccion']))
        campo_fraccion.clear()
        campo_fraccion.send_keys(str(fraccion))
        campo_precio = self.wait.until(EC.presence_of_element_located(self.SELECTORS['precio']))
        campo_precio.clear()
        campo_precio.send_keys(str(precio))

    def _fill_datos_adicionales(self):
        pestania = self.wait.until(EC.element_to_be_clickable(self.SELECTORS['pestania_adicionales']))
        self.driver.execute_script("arguments[0].click();", pestania)
        self.wait.until(EC.presence_of_element_located(self.SELECTORS['exportador_autorizado']))
        checkbox = self.wait.until(EC.element_to_be_clickable(self.SELECTORS['exportador_autorizado']))
        if not checkbox.is_selected():
            self.driver.execute_script("arguments[0].click();", checkbox)
        radio = self.wait.until(EC.element_to_be_clickable(self.SELECTORS['producto_artesanal']))
        if not radio.is_selected():
            self.driver.execute_script("arguments[0].click();", radio)
        select_entidad = Select(self.wait.until(EC.presence_of_element_located(self.SELECTORS['entidad'])))
        select_entidad.select_by_visible_text("MORELOS")

    def _click_continuar(self):
        boton = self.wait.until(EC.element_to_be_clickable(self.SELECTORS['boton_continuar']))
        self.driver.execute_script("arguments[0].scrollIntoView(true);", boton)
        time.sleep(0.3)
        self.driver.execute_script("arguments[0].click();", boton)
        logging.info("✅ Click en 'Continuar' – enviado a firma.")

    def _esperar_firma_y_acuse(self):
        self.wait.until(EC.presence_of_element_located(self.SELECTORS['acuse_titulo']))
        logging.info("✅ Acuse de Recibo detectado.")

    def _extraer_folio(self):
        try:
            mensaje = self.wait.until(EC.presence_of_element_located(self.SELECTORS['acuse_mensaje']))
            texto = mensaje.text
            patron = r'folio\s*<([^>]+)>'
            match = re.search(patron, texto, re.IGNORECASE)
            if match:
                folio = match.group(1)
                logging.info(f"📋 Folio extraído: {folio}")
                return folio
            logging.warning(f"⚠️ No se pudo extraer el folio. Texto del acuse: {texto}")
            return None
        except Exception as e:
            logging.error(f"❌ Error extrayendo folio: {e}")
            return None

    # --- MODIFICADO: Nueva Captura con espera y configuración inmediata ---
    def _click_nueva_captura(self):
        boton = self.wait.until(EC.element_to_be_clickable(self.SELECTORS['nueva_captura']))
        self.driver.execute_script("arguments[0].click();", boton)

        # Esperar que el formulario se reinicie (pestaña solicitante visible)
        self.wait.until(EC.presence_of_element_located(self.SELECTORS['formulario_cargado']))
        logging.info("✅ Formulario reiniciado (Nueva Captura).")

        # Esperar a que la pestaña de tratados esté clickeable
        try:
            WebDriverWait(self.driver, 15).until(
                EC.element_to_be_clickable(self.SELECTORS['tab_tratados'])
            )
            logging.info("✅ Pestaña Tratados está lista después de Nueva Captura.")
        except TimeoutException:
            logging.warning("⚠️ Pestaña Tratados no se volvió clickeable después de 15s, pero se continuará.")

        # --- INMEDIATAMENTE después de Nueva Captura, configurar tratados para el siguiente ciclo ---
        # Esto es clave: en lugar de esperar a que process_model lo haga, lo hacemos aquí.
        self._configurar_tratados()

        # Una vez configurados, ir a la pestaña de mercancía para que el siguiente modelo esté listo
        self._click_pestania_mercancia()

    # --------------------------------------------------------------
    # Procesamiento principal
    # --------------------------------------------------------------
    def process_model(self, model_data: dict, modo_automatico: bool = False):
        codigo = model_data.get('codigo', 'DESCONOCIDO')
        nombre = model_data.get('nombre', codigo)
        fraccion = model_data.get('fraccion', '')
        precio_factura = model_data.get('precio_factura', 0)
        ruta_csv = model_data.get('csv_path', None)

        logging.info(f"📤 Procesando modelo {codigo}")

        if self.modo_simulacion:
            logging.info(f"   [SIMULACIÓN] Modelo {codigo} procesado (sin navegador).")
            time.sleep(0.3)
            return

        # NOTA: Ya no llamamos a _configurar_tratados() aquí porque se hizo en _click_nueva_captura
        # Pero por si acaso, verificamos que la pestaña de mercancía esté activa.
        # Si no, la activamos.
        try:
            self.driver.find_element(*self.SELECTORS['nombre_comercial'])
        except:
            self._click_pestania_mercancia()

        try:
            if not modo_automatico:
                # ---- MODO INTERACTIVO ----
                logging.info("Esperando campo 'Nombre comercial'...")
                campo_comercial = self.wait.until(EC.presence_of_element_located(self.SELECTORS['nombre_comercial']))
                self.driver.execute_script("arguments[0].scrollIntoView(true);", campo_comercial)
                time.sleep(0.5)

                print("\n" + "=" * 60)
                print(">>> ESCRIBE EL NOMBRE COMERCIAL EN EL NAVEGADOR <<<")
                print("    Escribe el modelo, verifica que sea correcto.")
                print("    Luego, PRESIONA ENTER EN ESTA CONSOLA para continuar.")
                print("=" * 60 + "\n")
                input()

                nombre = campo_comercial.get_attribute("value")
                if not nombre:
                    logging.warning("⚠️ No se detectó texto en el campo. Se omite este modelo.")
                    return

                logging.info(f"✅ Texto leído: '{nombre}'")

                campo_ingles = self.wait.until(EC.presence_of_element_located(self.SELECTORS['nombre_ingles']))
                campo_ingles.clear()
                campo_ingles.send_keys(nombre)

                datos_excel, codigo_excel = self._buscar_datos_excel(nombre, modo_interactivo=True)
                if datos_excel:
                    self._fill_fraccion_precio(datos_excel['fraccion'], datos_excel['precio'])
                    if codigo_excel:
                        codigo = codigo_excel
                        logging.info(f"🔍 Usando código del Excel: {codigo}")
                else:
                    logging.warning("⚠️ No se encontraron datos en Excel para el nombre ingresado.")
                    if fraccion and precio_factura:
                        self._fill_fraccion_precio(fraccion, precio_factura)

                self._seleccionar_archivo_insumos(nombre, datos_excel)

            else:
                # ---- MODO AUTOMÁTICO ----
                logging.info(f"🤖 Modo automático: usando nombre '{nombre}' para {codigo}")
                self._fill_nombre_comercial(nombre)
                self._fill_fraccion_precio(fraccion, precio_factura)

                if ruta_csv and Path(ruta_csv).exists():
                    self._upload_insumos_file(ruta_csv)
                else:
                    logging.warning(f"⚠️ No se encontró CSV para {codigo}, se omite carga.")

            # ---- DATOS ADICIONALES ----
            self._safe_operation(self._fill_datos_adicionales)

            # ---- PAUSA DE REVISIÓN (solo en modo interactivo) ----
            if not modo_automatico:
                print("\n" + "=" * 70)
                print("👀 REVISA LOS DATOS EN EL NAVEGADOR")
                print("   Verifica nombre, fracción, precio, archivo cargado y datos adicionales.")
                print("   Si falta un proveedor o algo es incorrecto, puedes corregirlo ahora.")
                print("\n   Cuando todo esté correcto, PRESIONA ENTER para continuar.")
                print("   (El bot hará clic en 'Continuar' automáticamente)")
                print("=" * 70 + "\n")
                input()
            else:
                logging.info("🤖 Modo automático: enviando a firma sin revisión manual.")

            # ---- CONTINUAR ----
            self._safe_operation(self._click_continuar)

            # ---- FIRMA ELECTRÓNICA ----
            if not self._firma_msg_mostrado:
                logging.info("🔐 Esperando firma electrónica manual...")
                if modo_automatico:
                    print("\n" + "=" * 70)
                    print("🤖 MODO AUTOMÁTICO: Se han enviado modelos a firma.")
                    print("   Realiza la firma electrónica en el navegador (sube .cer/.key, contraseña).")
                    print("   Presiona ENTER cuando el ACUSE sea visible para continuar.")
                    print("=" * 70 + "\n")
                    input()
                else:
                    input(">>> Realiza la firma (sube .cer/.key, contraseña, clic en 'Firmar') y espera el Acuse. Presiona ENTER cuando el acuse sea visible...")
                self._firma_msg_mostrado = True
            else:
                logging.info("⏳ Esperando acuse para este modelo...")
                print("\n⏳ Presiona ENTER cuando el ACUSE sea visible en el navegador para este modelo...")
                input()

            # ---- ESPERAR ACUSE ----
            self._esperar_firma_y_acuse()

            # ---- EXTRAER FOLIO Y NUEVA CAPTURA ----
            folio = self._extraer_folio()
            self._click_nueva_captura()  # Aquí se configuran tratados para el siguiente ciclo

            logging.info(f"✅ Modelo {codigo} completado. Folio: {folio}")
            return folio

        except (InvalidSessionIdException, NoSuchWindowException) as e:
            logging.error(f"❌ Sesión perdida en {codigo}: {e}")
            raise SessionExpiredException("La sesión de Selenium ha expirado.") from e
        except Exception as e:
            logging.error(f"❌ Error en {codigo}: {e}")
            raise

    def _upload_insumos_file(self, ruta_csv):
        boton_carga = self.wait.until(EC.element_to_be_clickable((By.ID, "botonCargaArchivo")))
        self.driver.execute_script("arguments[0].click();", boton_carga)
        campo_file = self.wait.until(EC.visibility_of_element_located((By.ID, "archivoAdjuntar")))
        campo_file.send_keys(str(ruta_csv))
        boton_enviar = self.wait.until(EC.element_to_be_clickable((By.ID, "cargaArchivo")))
        self.driver.execute_script("arguments[0].click();", boton_enviar)
        self.wait.until(EC.invisibility_of_element_located((By.ID, "archivoAdjuntar")))
        logging.info("✅ Archivo de insumos cargado.")

    def close(self):
        if self.driver:
            self.driver.quit()
            logging.info("🛑 Browser cerrado.")