# VUCEM Automation Bot

![Python](https://img.shields.io/badge/Python-3.10+-blue)
![Selenium](https://img.shields.io/badge/Selenium-4.0+-green)
![License](https://img.shields.io/badge/License-MIT-yellow)

## 🚀 Overview

This project automates the generation and upload of model templates to the VUCEM (Mexican Customs) portal. It replaces manual, error-prone data entry with a robust, modular pipeline.

**Key Features:**
- **Zero manual input per cycle** – scans Excel files and generates CSVs automatically.
- **Smart category detection** – uses a dictionary of keywords to classify insumos.
- **Dynamic column G calculation** – `G = base + number of unique categories per model`.
- **Interactive insumo selection** – choose which insumos to upload for each model.
- **Simulation mode** – run without connecting to VUCEM, perfect for portfolio demos.
- **Clean architecture** – separated concerns (config, data processing, browser automation).
- **Explicit waits** – no `time.sleep()`, uses `WebDriverWait` for reliability.
- **Professional logging** – detailed logs for debugging and auditing.

## 🏗️ Architecture
config.py → Centralized configuration (paths, URLs, column mappings)
data_processor.py → Business logic: reads Excel, categorizes, calculates G, exports CSVs
data_manager.py → Orchestrates data preparation
browser_automation.py → Selenium automation: login, upload, field filling
state_manager.py → Checkpoints and state recovery
main.py → Entry point, ties everything together

text

## 📁 Project Structure
.
├── data/ → Place input Excel files here (or generate with sample script)
├── output/ → Generated CSV templates
├── logs/ → Application logs
├── config.py
├── data_processor.py
├── data_manager.py
├── browser_automation.py
├── state_manager.py
├── main.py
├── generate_sample_data.py → Generate sample Excel files
├── requirements.txt
└── README.md

text

## 🛠️ Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/vucem-automation-bot.git
   cd vucem-automation-bot
Install dependencies:

bash
pip install -r requirements.txt
Download ChromeDriver and place it in your PATH.

Generate sample data (for portfolio/testing):

bash
python generate_sample_data.py
This creates two Excel files in the data/ folder with fake data.

(Optional) Place your real Excel files in data/ (overwrite the sample ones).

▶️ Usage
Simulation Mode (No browser, for portfolio)
Set MODO_SIMULACION = True in config.py. Then run:

bash
python main.py
The bot will process the data and simulate the upload, printing logs only.

Real Mode (With browser)
Set MODO_SIMULACION = False and ensure the VUCEM URLs are correct in config.py. Then:

bash
python main.py
The bot will open the browser, wait for you to log in manually, and then proceed to upload the templates.

⚙️ Configuration
Edit config.py to:

Change file paths.

Toggle simulation mode (MODO_SIMULACION).

Set headless mode (HEADLESS).

Modify timeout values.

📈 Performance
Processes 187 models in ~5 minutes (excluding manual login).

Generates consistent, error-free CSVs ready for VUCEM.

🔒 Data Privacy
Important: This repository contains no real business data. All Excel files and sample data are fictional and generated for demonstration purposes. If you use real data, ensure you do not commit them to version control.

🤝 Contributing
Pull requests are welcome. For major changes, please open an issue first.

📄 License
MIT

👤 Author
Daniela Diaz | 777 132 3165 | linkedin.com/in/daniela-diaz-galeana-76589b314

🙏 Acknowledgments
VUCEM portal for the challenge.

Open-source libraries that made this possible.
