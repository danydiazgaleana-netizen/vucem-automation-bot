"""
Genera datos de ejemplo sanitizados para el proyecto VUCEM Automation Bot.
Estos archivos son ficticios y no contienen información real de Grupo REV.
"""
import pandas as pd
import random
from pathlib import Path
from config import Config

# Asegurar directorios
Config.ensure_directories()

# Listas de datos ficticios
NOMBRES_MODELOS = [
    "Producto Ejemplo A", "Producto Ejemplo B", "Producto Ejemplo C",
    "Producto Ejemplo D", "Producto Ejemplo E", "Producto Ejemplo F",
    "Producto Ejemplo G", "Producto Ejemplo H", "Producto Ejemplo I",
    "Producto Ejemplo J"
]

INSUMOS_FICTICIOS = [
    "Insumo X", "Insumo Y", "Insumo Z", "Insumo W",
    "Material A", "Material B", "Componente C", "Componente D"
]

PROVEEDORES_FICTICIOS = [
    "Proveedor Alpha", "Proveedor Beta", "Proveedor Gamma",
    "Proveedor Delta", "Proveedor Epsilon"
]

def generar_maestro():
    """Genera el archivo maestro (modelos_maestro.xlsx)"""
    data = []
    for i in range(1, 21):  # 20 modelos de ejemplo
        codigo = f"000000{100000 + i:06d}"
        nombre = random.choice(NOMBRES_MODELOS) + f" {i}"
        precio = round(random.uniform(10, 200), 2)
        fraccion = random.choice(["95059099", "62101001", "69139099"])
        # Para algunos modelos, col_G puede ser 3 o 0 según fracción
        col_g = 3 if fraccion == "95059099" else (0 if fraccion == "62101001" else 3)
        data.append({
            "CODIGO-MODELO": codigo,
            "NOMBRE_MODELO": nombre,
            "PRECIO FACTURA": precio,
            "D": 0.3,
            "E": precio * 0.7,
            "CODIGO ARANCEL": fraccion,
            "INSUMOS PIEZAS": col_g,
            "Precio unitario": precio * 0.7 / max(col_g, 1),
            "I": "",
            "J": "",
            "K": "",
            "L": ""
        })
    df = pd.DataFrame(data)
    output_path = Config.MASTER_EXCEL_PATH
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name="2026", index=False)
    print(f"✅ Maestro generado: {output_path}")

def generar_bom():
    """Genera el archivo BOM (lista_materiales.xlsx)"""
    # Hoja LM: insumos por modelo
    lm_data = []
    for i in range(1, 21):
        codigo = f"000000{100000 + i:06d}"
        # Cada modelo tiene entre 2 y 5 insumos
        num_insumos = random.randint(2, 5)
        insumos_seleccionados = random.sample(INSUMOS_FICTICIOS, min(num_insumos, len(INSUMOS_FICTICIOS)))
        for insumo in insumos_seleccionados:
            lm_data.append({
                "CODIGO-MODELO": codigo,
                "B": "",
                "C": "",
                "INSUMO": insumo,
                "E": "",
                "F": ""
            })
    df_lm = pd.DataFrame(lm_data)

    # Hoja BASE DE DATOS: proveedores
    proveedores_data = []
    for insumo in INSUMOS_FICTICIOS:
        proveedor = random.choice(PROVEEDORES_FICTICIOS)
        proveedores_data.append({
            "INSUMO": insumo,
            "PROVEEDOR": proveedor
        })
    df_prov = pd.DataFrame(proveedores_data)

    output_path = Config.BOM_EXCEL_PATH
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_lm.to_excel(writer, sheet_name="LM", index=False)
        df_prov.to_excel(writer, sheet_name="BASE DE DATOS", index=False)
    print(f"✅ BOM generado: {output_path}")

if __name__ == "__main__":
    generar_maestro()
    generar_bom()
    print("🎉 Datos de ejemplo generados correctamente.")