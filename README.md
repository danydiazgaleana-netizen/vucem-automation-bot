# 🚀 VUCEM Automation Bot

[![Python](https://img.shields.io/badge/Python-3.10+-blue)](https://python.org)
[![Selenium](https://img.shields.io/badge/Selenium-4.0+-green)](https://selenium.dev)
[![License](https://img.shields.io/badge/License-MIT-yellow)](LICENSE)

## 📖 Overview

**VUCEM Automation Bot** is an end-to-end RPA (Robotic Process Automation) system built to eliminate manual data entry in Mexico's foreign trade portal VUCEM (Ventanilla Digital Mexicana de Comercio Exterior).

What used to take **45 minutes for 19 models** now runs in under **3 minutes per batch** — a 90%+ time reduction.

The bot handles certificate-of-origin registration (TLCUEM), reads Bill of Materials (BOM) Excel files, dynamically calculates customs data, generates CSV templates, and automates the entire VUCEM web form submission — including login detection, dropdown selection, and file upload.

---

## ✨ Key Features

- ✅ **Full browser automation** — navigates VUCEM menus, fills forms, and uploads files without human intervention
- ✅ **Smart category detection** — classifies materials using a keyword dictionary and calculates column G dynamically from the BOM
- ✅ **Auto-updating BOM pipeline** — reads the latest Bill of Materials on every run, recalculates G values, and regenerates all CSV templates automatically
- ✅ **Simulation mode** — runs the full pipeline without connecting to VUCEM, ideal for demos and testing
- ✅ **Interactive terminal UI** — color-coded menu with tables to select models, retry failures, or export reports
- ✅ **Resilient execution** — automatic retries with exponential backoff, checkpoints via `StateManager`, and structured logging
- ✅ **Modular architecture** — clean separation of concerns across config, data processing, browser automation, and state management
- ✅ **Privacy-first** — repository contains zero sensitive or company data; real config lives in `config_local.py` (git-ignored)

---

## 🏗️ Architecture

```
config.py            → Centralized config (paths, URLs, selectors) — simulation mode
config_local.py      → Local production config — NEVER committed (git-ignored)
data_processor.py    → Business logic: BOM reading, category detection, G calculation, CSV export
data_manager.py      → Data orchestration and deduplication
browser_automation.py → Selenium automation: login detection, navigation, form filling, file upload
state_manager.py     → Checkpoint management, retry logic, HTML/Excel report generation
main.py              → Entry point with interactive menu (Rich UI)
```

---

## 📁 Project Structure

```
vucem-automation-bot/
├── data/                    → Excel input files (git-ignored; use sample data for testing)
│   └── Sample/              → Sample Excel files for simulation mode
├── output/                  → Generated CSV templates (git-ignored)
├── logs/                    → Execution logs (git-ignored)
├── config.py                → Public config (simulation mode, generic paths)
├── config_local.py          → ⚠️ Local production config — NOT included in repo
├── data_processor.py        → Core processing logic
├── data_manager.py          → Data orchestration
├── browser_automation.py    → Selenium VUCEM automation
├── state_manager.py         → Checkpoints and reporting
├── main.py                  → Interactive entry point
├── requirements.txt         → Dependencies
├── .gitignore               → Excludes sensitive files and generated outputs
└── README.md                → This file
```

---

## 🛠️ Installation

### 1. Clone the repository

```bash
git clone https://github.com/danydiazgaleana-netizen/vucem-automation-bot.git
cd vucem-automation-bot
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. ChromeDriver

ChromeDriver is managed automatically via `webdriver_manager` — no manual installation needed.

### 4. Prepare sample data (simulation mode)

Place sample Excel files in the `data/` folder matching the structure expected by `config.py`:
- `modelos_maestro.xlsx` — master model list
- `lista_materiales_bom.xlsx` — Bill of Materials

---

## ▶️ Usage

### Simulation mode (recommended for testing)

Make sure `config.py` has `MODO_SIMULACION = True`, then:

```bash
python main.py
```

The bot will process the BOM, calculate G values, generate CSV templates, and display the interactive menu — all without opening a browser.

### Production mode (real VUCEM)

Create `config_local.py` with your real file paths and `MODO_SIMULACION = False`. The bot automatically detects this file:

```bash
python main.py
# Log will show: ⚙️ Configuración activa: LOCAL (producción real)
```

The browser opens, you log in manually, and the bot takes over from there.

---

## ⚙️ Configuration

The bot uses a **dual-config pattern** for safety:

| File | Purpose | In repo? |
|---|---|---|
| `config.py` | Public config, simulation mode, generic paths | ✅ Yes |
| `config_local.py` | Production config, real paths, `MODO_SIMULACION=False` | ❌ Never |

`main.py` automatically loads `config_local.py` if it exists, falling back to `config.py`:

```python
try:
    from config_local import Config   # production
except ImportError:
    from config import Config         # simulation / GitHub
```

---

## 📊 Pipeline Flow

Each run executes this pipeline automatically:

```
1. Read BOM (lista_materiales.xlsx)
      ↓
2. Detect material categories using keyword dictionary
      ↓
3. Calculate col G = base + unique (category, variant) combinations
      ↓
4. Update col G directly in the master Excel file
      ↓
5. Generate CSV templates for each model
      ↓
6. Open browser → manual login → bot navigates VUCEM
      ↓
7. Interactive menu: select model → bot fills form → submit
      ↓
8. Checkpoint saved → next model
```

---

## 📈 Impact Metrics

| Metric | Before | After |
|---|---|---|
| Time for 19 models | ~45 minutes | ~3 minutes |
| Manual data entry errors | Frequent | Zero |
| Models processed per run | 19 | 187+ |
| Success rate | — | ~95% (with auto-retry) |

---

## 🧪 Tech Stack

| Tool | Purpose |
|---|---|
| Python 3.10+ | Core language |
| Selenium 4+ | Browser automation |
| OpenPyXL | Excel read/write |
| Pandas | Data validation and reporting |
| Rich | Terminal UI (tables, colors, panels) |
| Jinja2 | HTML report generation |
| webdriver_manager | Automatic ChromeDriver management |

---

## 🔐 Privacy & Security

This repository contains **zero sensitive data**:
- No real Excel files (git-ignored)
- No credentials or API keys
- No company-specific configuration
- `config_local.py` is git-ignored and never committed

Sample data uses generic fictitious values for demonstration purposes.

---

## 🤝 Contributing

Contributions are welcome. Please:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/your-feature`)
3. Commit your changes (`git commit -m 'Add your feature'`)
4. Push to the branch (`git push origin feature/your-feature`)
5. Open a Pull Request

---

## 📄 License

MIT License — see [LICENSE](LICENSE) for details.

---

## 👤 Author

**Daniela Diaz Galeana**

[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-blue)](https://linkedin.com/in/daniela-diaz-galeana)
[![GitHub](https://img.shields.io/badge/GitHub-Follow-black)](https://github.com/danydiazgaleana-netizen)

---

## 📞 Contact

Questions or collaboration? Reach out via [LinkedIn](https://linkedin.com/in/daniela-diaz-galeana) or [GitHub](https://github.com/danydiazgaleana-netizen).

---

⭐ If this project helped you, please give it a star on GitHub!
