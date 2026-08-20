
# 🚀 VUCEM Automation Bot

![Python](https://img.shields.io/badge/Python-3.10+-blue)
![Selenium](https://img.shields.io/badge/Selenium-4.0+-green)
![License](https://img.shields.io/badge/License-MIT-yellow)

## 📖 Overview

**VUCEM Automation Bot** es un sistema de automatización RPA (Robotic Process Automation) diseñado para optimizar el alta de modelos en el portal VUCEM (Ventanilla Digital Mexicana de Comercio Exterior). Reemplaza el proceso manual de 15-20 minutos por modelo con un flujo automatizado que reduce el tiempo a menos de **3 minutos por lote**.

Este proyecto fue desarrollado para gestionar certificados de origen TLCUEM, procesando listas de materiales (BOM) y archivos maestros en Excel para generar plantillas CSV listas para VUCEM.

---

## ✨ Características principales

- ✅ **Automatización total** – navega por los menús de VUCEM, completa formularios y carga archivos sin intervención humana.
- ✅ **Detección inteligente de categorías** – clasifica insumos usando un diccionario de palabras clave y calcula la columna G dinámicamente.
- ✅ **Modo simulación** – ejecuta el flujo completo sin conexión a VUCEM, ideal para demostraciones y pruebas.
- ✅ **Interfaz interactiva** – menú con colores y tablas para seleccionar modelos, reintentar fallos o exportar reportes.
- ✅ **Robustez** – reintentos automáticos, checkpoints (`StateManager`) y logging profesional.
- ✅ **Arquitectura modular** – separación clara de responsabilidades (configuración, procesamiento de datos, automatización del navegador).
- ✅ **Seguridad** – el repositorio no contiene datos sensibles ni información confidencial.

---

## 🏗️ Arquitectura
config.py → Configuración centralizada (rutas, URLs, selectores)
data_processor.py → Lógica de negocio: lectura de Excel, categorización, cálculo de G y exportación CSV
data_manager.py → Orquestación de la preparación de datos
browser_automation.py → Automatización de Selenium: login, navegación, carga de archivos
state_manager.py → Gestión de checkpoints y recuperación de estado
main.py → Punto de entrada con menú interactivo

text

---

## 📁 Estructura del proyecto
.
├── data/ → Archivos Excel de entrada (se generan con generate_sample_data.py)
├── output/ → Plantillas CSV generadas
├── logs/ → Logs de ejecución
├── config.py → Configuración centralizada
├── data_processor.py → Lógica de negocio
├── data_manager.py → Orquestación de datos
├── browser_automation.py → Automatización Selenium
├── state_manager.py → Checkpoints y recuperación
├── main.py → Punto de entrada
├── generate_sample_data.py → Generador de datos de ejemplo
├── requirements.txt → Dependencias
└── README.md → Este archivo

text

---

## 🛠️ Instalación

### 1. Clonar el repositorio

```bash
git clone https://github.com/danydiazgaleana-netizen/vucem-automation-bot.git
cd vucem-automation-bot
2. Instalar dependencias
bash
pip install -r requirements.txt
3. Descargar ChromeDriver
Asegúrate de tener ChromeDriver instalado y en tu PATH, o bien, el script lo gestiona automáticamente con webdriver_manager.

4. Generar datos de ejemplo (opcional)
bash
python generate_sample_data.py
Esto creará archivos Excel ficticios en la carpeta data/ para probar el bot sin usar datos reales.

▶️ Uso
Modo simulación (recomendado para pruebas)
Abre config.py y asegúrate de que MODO_SIMULACION = True.

Ejecuta:

bash
python main.py
El bot generará las plantillas CSV y simulará el procesamiento sin abrir el navegador.

Modo real (producción)
Abre config.py y cambia MODO_SIMULACION = False.

Ajusta las URLs de VUCEM si es necesario.

Coloca tus archivos Excel reales en data/ (con los nombres y estructura esperada).

Ejecuta:

bash
python main.py
El bot abrirá el navegador, te pedirá iniciar sesión manualmente y luego procesará los modelos automáticamente.

⚙️ Configuración
Edita config.py para personalizar:

MODO_SIMULACION → True (simulación) o False (producción)

MASTER_EXCEL_PATH → ruta al archivo maestro

BOM_EXCEL_PATH → ruta al archivo de lista de materiales (BOM)

SELECTORS → selectores XPath o CSS de VUCEM (ajustar si el portal cambia)

VUCEM_URL → URL del portal de VUCEM

🔐 Privacidad y datos
Este repositorio no contiene datos reales de Grupo REV ni de ninguna empresa. Los archivos Excel de ejemplo son generados con datos ficticios. Si utilizas datos reales, asegúrate de no subirlos a version control (están ignorados por .gitignore).

📈 Métricas de impacto
Tiempo manual por modelo: 15-20 minutos

Tiempo automatizado por lote: < 3 minutos

Modelos procesados por lote: 187+

Tasa de éxito: ~95% (con reintentos automáticos)

🧪 Tecnologías utilizadas
Python 3.10+

Selenium – automatización del navegador

Pandas / OpenPyXL – procesamiento de datos

Rich – interfaz de terminal mejorada

WebDriverWait – esperas explícitas (sin time.sleep)

Jinja2 – generación de reportes HTML

🤝 Contribuciones
Las contribuciones son bienvenidas. Si deseas mejorar el proyecto, por favor:

Haz fork del repositorio

Crea una rama para tu feature (git checkout -b feature/nueva-funcionalidad)

Haz commit de tus cambios (git commit -m 'Añadir nueva funcionalidad')

Haz push a la rama (git push origin feature/nueva-funcionalidad)

Abre un Pull Request

📄 Licencia
MIT License. Consulta el archivo LICENSE para más detalles.

👤 Autor
Daniela Diaz Galeana

LinkedIn

GitHub

🙏 Agradecimientos
Portal VUCEM por el desafío técnico

Comunidad de código abierto por las herramientas utilizadas

Empresa Grupo REV por el contexto real de la operación

📞 Contacto
Si tienes preguntas o deseas colaborar, puedes contactarme a través de LinkedIn o GitHub.

⭐ Si te gusta este proyecto, no olvides darle una estrella en GitHub! ⭐

