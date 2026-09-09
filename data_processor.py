"""
data_processor.py — Procesador de datos VUCEM
Autor: Daniela Diaz Galeana

Flujo completo en cada ejecución:
  1. Lee BOM (lista de materiales) → construye mapa de insumos por modelo
  2. Actualiza col G del archivo ALTA directamente (igual que llenar_columna_G.py)
  3. Lee el ALTA ya actualizado → genera plantillas CSV para cada modelo
"""
import csv
import shutil
import logging
from collections import defaultdict
from pathlib import Path
from openpyxl import load_workbook

try:
    from config_local import Config
except ImportError:
    from config import Config

# ============================================================
# DICCIONARIO DE CATEGORÍAS
# ============================================================
CATEGORIAS = {
    "dona":           ["dona para pelo"],
    "yute":           ["tela yute", "yute"],
    "tela":           ["tela"],
    "esponja":        ["esponja"],
    "blister":        ["blister"],
    "hilo":           ["hilo"],
    "multifilamento": ["multifilamento"],
    "elastico":       ["elastico", "elástico"],
    "epoxico":        ["epoxico", "epóxico", "pintura epoxica"],
    "suaje":          ["suaje"],
    "broche":         ["broche"],
    "flor":           ["rosa", "flor"],
    "botones":        ["botones", "boton"],
    "argolla":        ["argolla"],
    "guantes":        ["guantes"],
    "sombrero":       ["sombrero"],
    "adhesivo":       ["kola loka", "adhesivo", "pegamento"],
    "pigmento":       ["pigmento"],
    "peluche":        ["peluche"],
    "cuerda":         ["cuerda", "cordon", "cordón"],
    "encaje":         ["punta de encaje", "encaje"],
    "sujeta":         ["sujeta", "sujetador", "sujetadocumentos"],
    "lazo":           ["lazo"],
}

# Filas fijas del proceso base
# En producción (config_local.py) se usan los insumos reales de látex.
# En GitHub (config.py) se usan nombres genéricos sanitizados.
FILAS_FIJAS = getattr(Config, 'FILAS_FIJAS_LATEX', [
    ("Insumo Fijo A", "Proveedor Genérico A", "Proveedor Genérico A"),
    ("Insumo Fijo B", "Proveedor Genérico B", "Proveedor Genérico B"),
    ("Insumo Fijo C", "Proveedor Genérico C", "Proveedor Genérico C"),
])

CHARS_INVALIDOS = '<>:"/\\|?*'


# ============================================================
# FUNCIONES AUXILIARES
# ============================================================

def identificar_categoria_y_keyword(desc_lower: str):
    """
    Devuelve (cat_id, keyword) si la descripción contiene alguna
    palabra clave de las categorías, o (None, None) si no.
    Usa 'in' para detectar categorías en cualquier posición del texto,
    compatible con descripciones tipo 'MODELO CATEGORIA 01' y
    tipo 'Hilo Blanco Para Coser'.
    """
    for cat_id, keywords in CATEGORIAS.items():
        for kw in keywords:
            if kw in desc_lower:
                return cat_id, kw
    return None, None


def obtener_clave_dedup(desc_original: str, keyword: str) -> tuple:
    """
    Clave de deduplicación = (primeros 5 chars después de la keyword,
    últimos 3 chars de la descripción).
    Distingue variantes del mismo insumo: 'Hilo Negro' vs 'Hilo Blanco'.
    """
    desc_lower = desc_original.lower()
    idx = desc_lower.find(keyword)
    primeros5 = desc_lower[idx + len(keyword): idx + len(keyword) + 5]
    ultimos3  = desc_original.strip()[-3:].strip()
    return (primeros5, ultimos3)


def normalizar_nombre(insumo: str) -> str:
    txt = str(insumo).strip()
    txt_lower = txt.lower()
    if "epoxico" in txt_lower or "pintura epoxica" in txt_lower:
        return "Pintura Epoxica"
    if "kola loka" in txt_lower and "20 gms" in txt_lower:
        return "Pegamento instantaneo Industrial bote 20 gms"
    return txt


def nombre_seguro(texto: str) -> str:
    return "".join("_" if c in CHARS_INVALIDOS else c for c in texto).replace(" ", "_")


def escribir_csv(ruta: Path, filas: list):
    with open(ruta, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, delimiter=",", lineterminator="\r\n")
        for fila in filas:
            writer.writerow(fila)


# ============================================================
# CLASE PROCESADORA PRINCIPAL
# ============================================================

class DataProcessor:
    def __init__(self):
        self.master_file = Config.MASTER_EXCEL_PATH
        self.bom_file    = Config.BOM_EXCEL_PATH
        self.output_dir  = Config.OUTPUT_DIR
        self.output_dir.mkdir(parents=True, exist_ok=True)

        self.proveedores_map    = {}
        self.insumos_por_modelo = defaultdict(dict)   # code12 → {clave_dedup: nombre_norm}
        self.col_g_por_modelo   = defaultdict(int)    # code12 → número de combinaciones únicas

        self.col_codigo  = Config.MASTER_COLUMNS['codigo']
        self.col_precio  = Config.MASTER_COLUMNS['precio_factura']
        self.col_fraccion = 5   # col F = CODIGO ARANCEL
        self.col_g_idx   = 6   # col G = INSUMOS PIEZAS (base 0)

        self.col_bom_codigo  = Config.BOM_LM_COLUMNS['codigo']
        self.col_bom_insumo  = Config.BOM_LM_COLUMNS['insumo_desc']
        self.col_prov_nombre   = Config.BOM_PROVEEDORES_COLUMNS['nombre_insumo']
        self.col_prov_proveedor = Config.BOM_PROVEEDORES_COLUMNS['proveedor']

        self.modelos_procesados = []

    # ----------------------------------------------------------
    # PASO 1 — Cargar proveedores desde BASE DE DATOS
    # ----------------------------------------------------------
    def _cargar_proveedores(self):
        logging.info("Cargando proveedores desde BOM...")
        wb = load_workbook(self.bom_file, read_only=True)
        ws = wb["BASE DE DATOS"]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[self.col_prov_nombre]:
                nombre_db = normalizar_nombre(str(row[self.col_prov_nombre]).strip())
                proveedor = str(row[self.col_prov_proveedor] if row[self.col_prov_proveedor] else "S/P").strip()
                self.proveedores_map[nombre_db] = proveedor
        wb.close()
        logging.info(f"✅ {len(self.proveedores_map)} proveedores cargados.")

    # ----------------------------------------------------------
    # PASO 2 — Cargar insumos desde hoja LM y calcular col G
    # ----------------------------------------------------------
    def _cargar_insumos_y_calcular_g(self):
        logging.info("Cargando insumos desde BOM y calculando col G...")
        wb = load_workbook(self.bom_file, read_only=True)
        ws = wb["LM"]

        combinaciones = defaultdict(set)   # code12 → set de (cat_id, primeros5, ultimos3)

        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row[self.col_bom_codigo]:
                continue
            code12 = str(row[self.col_bom_codigo]).strip()[:12]
            desc   = str(row[self.col_bom_insumo]).strip() if len(row) > self.col_bom_insumo else ""
            if not desc:
                continue

            cat_id, keyword = identificar_categoria_y_keyword(desc.lower())
            if cat_id:
                clave_g    = (cat_id,) + obtener_clave_dedup(desc, keyword)
                clave_dict = (cat_id, obtener_clave_dedup(desc, keyword))

                # Para col G: deduplicar por (cat + primeros5 + ultimos3)
                combinaciones[code12].add(clave_g)

                # Para plantillas: guardar nombre normalizado (sin duplicados)
                nombre_norm = normalizar_nombre(desc)
                flat_key = (cat_id, obtener_clave_dedup(desc, keyword))
                if flat_key not in self.insumos_por_modelo[code12]:
                    self.insumos_por_modelo[code12][flat_key] = nombre_norm

        wb.close()

        # Col G = base 3 + número de combinaciones únicas
        for code12, combs in combinaciones.items():
            self.col_g_por_modelo[code12] = 3 + len(combs)

        logging.info(f"✅ {len(self.insumos_por_modelo)} modelos con insumos cargados.")

    # ----------------------------------------------------------
    # PASO 3 — Actualizar col G en el archivo ALTA
    # ----------------------------------------------------------
    def _actualizar_col_g_en_alta(self):
        logging.info("Actualizando col G en archivo ALTA...")
        wb = load_workbook(self.master_file)
        ws = wb["2026"]
        actualizados = 0

        for row_idx in range(2, ws.max_row + 1):
            celda_b = ws.cell(row=row_idx, column=self.col_codigo + 1)
            if not celda_b.value:
                continue
            code12   = str(celda_b.value).strip()[:12]
            nuevo_g  = self.col_g_por_modelo.get(code12, 3)
            ws.cell(row=row_idx, column=self.col_g_idx + 1).value = nuevo_g
            actualizados += 1

        wb.save(self.master_file)
        wb.close()
        logging.info(f"✅ Col G actualizada en {actualizados} filas del ALTA.")

    # ----------------------------------------------------------
    # PASO 4 — Generar plantilla CSV para un modelo
    # ----------------------------------------------------------
    def _generar_plantilla(self, row: tuple) -> tuple:
        codigo_full   = row[self.col_codigo]
        if not codigo_full:
            return None, None, None, None
        codigo_full   = str(codigo_full).strip()
        precio_factura = row[self.col_precio]
        fraccion       = str(row[self.col_fraccion]).strip() if len(row) > self.col_fraccion else ""

        if not precio_factura or precio_factura == 0:
            logging.warning(f"⏭️ {codigo_full}: precio vacío, omitido.")
            return None, None, None, None

        codigo_12      = codigo_full[:12]
        insumos_dict   = self.insumos_por_modelo.get(codigo_12, {})
        insumos_lista  = list(insumos_dict.values())

        # Usar siempre col_g calculado desde el BOM — ignora el valor del Excel
        col_g = self.col_g_por_modelo.get(codigo_12, 3)

        # fraccion 95059099 = máscaras de látex → base 3 + filas fijas reales
        # fraccion 62101001 = disfraces/vestuario → base 0, sin filas fijas,
        #                     solo insumos de categorías detectadas en el BOM
        if fraccion == "62101001":
            base          = 0
            incluir_fijas = False
        else:
            base          = 3
            incluir_fijas = True

        try:
            precio_unitario = f"{precio_factura * 0.70 / col_g:.2f}"
        except Exception:
            logging.error(f"❌ {codigo_full}: error calculando precio unitario.")
            return None, None, None, None

        filas_csv = []
        if incluir_fijas:
            for insumo, prov_b, prov_c in FILAS_FIJAS:
                filas_csv.append([insumo, prov_b, prov_c, "", precio_unitario, "", "", "", "s"])

        for insumo in insumos_lista:
            proveedor = self.proveedores_map.get(insumo, "PROVEEDOR NO ENCONTRADO")
            filas_csv.append([insumo, proveedor, proveedor, "", precio_unitario, "", "", "", "s"])

        # Recalcular col_g real basado en las filas generadas (fuente de verdad)
        col_g = len(filas_csv)

        return codigo_full, filas_csv, col_g, fraccion

    # ----------------------------------------------------------
    # PROCESO PRINCIPAL
    # ----------------------------------------------------------
    def process(self):
        if not self.master_file.exists():
            raise FileNotFoundError(f"No se encontró el archivo ALTA: {self.master_file}")
        if not self.bom_file.exists():
            raise FileNotFoundError(f"No se encontró el archivo BOM: {self.bom_file}")

        # 1. Cargar datos del BOM
        self._cargar_proveedores()
        self._cargar_insumos_y_calcular_g()

        # 2. Actualizar col G en el ALTA (en disco)
        self._actualizar_col_g_en_alta()

        # 3. Leer el ALTA ya actualizado y generar plantillas
        logging.info("Generando plantillas CSV desde ALTA actualizado...")
        wb = load_workbook(self.master_file, data_only=True)
        ws = wb["2026"]

        generados = 0
        omitidos  = 0
        errores   = 0

        for row in ws.iter_rows(min_row=2, values_only=True):
            codigo_full, filas_csv, col_g, fraccion = self._generar_plantilla(row)
            if codigo_full is None:
                if row[self.col_codigo]:
                    if not row[self.col_precio] or row[self.col_precio] == 0:
                        omitidos += 1
                    else:
                        errores += 1
                continue

            # El nombre del CSV incluye el código completo para evitar
            # que dos modelos con el mismo nombre (ej. tallas distintas)
            # se sobreescriban mutuamente.
            base_nombre = f"PLANTILLA_{nombre_seguro(codigo_full)}"
            ruta        = self.output_dir / f"{base_nombre}.csv"
            escribir_csv(ruta, filas_csv)
            generados += 1
            logging.info(
                f"✅ Generado: {ruta.name} "
                f"({len(filas_csv)} filas, col_G={col_g}, fracción={fraccion})"
            )

            self.modelos_procesados.append({
                'codigo':   codigo_full,
                'nombre':   codigo_full,
                'csv_path': ruta,
                'data': {
                    'codigo':         codigo_full,
                    'nombre':         codigo_full,
                    'precio_factura': row[self.col_precio],
                    'fraccion':       fraccion,
                    'col_g':          col_g,
                    'precio_unitario': filas_csv[0][4] if filas_csv else "",
                    'insumos':        [f[0] for f in filas_csv],
                    'csv_path':       ruta,
                }
            })

        wb.close()

        logging.info("📊 Resumen final:")
        logging.info(f"   ✅ Plantillas generadas: {generados}")
        logging.info(f"   ⏭️ Omitidas (sin precio): {omitidos}")
        logging.info(f"   ❌ Errores (inconsistencia): {errores}")
        logging.info(f"   📁 Carpeta: '{self.output_dir}'")

        return self.modelos_procesados
# dp
