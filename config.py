from pathlib import Path
from selenium.webdriver.common.by import By

class Config:
    BASE_DIR = Path(__file__).resolve().parent

    DATA_DIR = BASE_DIR / "data"
    OUTPUT_DIR = BASE_DIR / "output"
    LOGS_DIR = BASE_DIR / "logs"

    # Archivos de entrada (nombres genéricos)
    MASTER_EXCEL_PATH = DATA_DIR / "modelos_maestro.xlsx"
    BOM_EXCEL_PATH = DATA_DIR / "lista_materiales.xlsx"

    # URLs de VUCEM (para modo real)
    VUCEM_URL = "https://www.ventanillaunica.gob.mx/vucem/Ingreso.html"
    LOGIN_URL = VUCEM_URL

    # Selenium
    DEFAULT_TIMEOUT = 30
    HEADLESS = False
    CHROME_DRIVER_PATH = None

    # Modo simulación (True para pruebas, False para producción)
    MODO_SIMULACION = True   # <--- IMPORTANTE: mantener en True para GitHub

    # Selectores (definidos pero no validados en producción aún)
    SELECTORS = {
        'nombre_comercial': (By.ID, "solicitud.registroCuestionario.mercanciaAsociada.nombreComercial"),
        'nombre_ingles': (By.ID, "solicitud.registroCuestionario.mercanciaAsociada.nombreIngles"),
        'fraccion': (By.ID, "solicitud.registroCuestionario.mercanciaAsociada.fraccionArancelaria.clave"),
        'precio': (By.ID, "solicitud.registroCuestionario.mercanciaAsociada.precioFrancoFabrica"),
    }

    MASTER_COLUMNS = {
        'codigo': 1,
        'precio_factura': 2,
    }

    BOM_LM_COLUMNS = {
        'codigo': 0,
        'insumo_desc': 3
    }

    BOM_PROVEEDORES_COLUMNS = {
        'nombre_insumo': 0,
        'proveedor': 1
    }

    @classmethod
    def ensure_directories(cls):
        cls.DATA_DIR.mkdir(parents=True, exist_ok=True)
        cls.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        cls.LOGS_DIR.mkdir(parents=True, exist_ok=True)