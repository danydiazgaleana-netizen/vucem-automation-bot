import json
import logging
from pathlib import Path
from datetime import datetime
import pandas as pd
from config import Config

class StateManager:
    def __init__(self, state_file: Path = None, fail_file: Path = None):
        self.state_file = state_file or Config.BASE_DIR / "processed_models.json"
        self.fail_file = fail_file or Config.BASE_DIR / "failed_models.json"
        self.report_file = Config.LOGS_DIR / f"reporte_{datetime.now().strftime('%Y%m%d')}.html"
        self.completed = set()                    # solo códigos
        self.completed_details = {}               # código -> {'folio': str, 'timestamp': str}
        self.failures = []
        self._load_state()
        self._load_failures()

    def _load_state(self):
        if self.state_file.exists():
            try:
                with open(self.state_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.completed = set(data.get('completed', []))
                    self.completed_details = data.get('completed_details', {})
                logging.info(f"📂 Estado cargado: {len(self.completed)} modelos ya procesados.")
            except Exception as e:
                logging.error(f"⚠️ Error al leer estado: {e}. Se inicia desde cero.")
                self.completed = set()
                self.completed_details = {}
        else:
            self.completed = set()
            self.completed_details = {}

    def _save_state(self):
        temp = self.state_file.with_suffix('.tmp')
        try:
            data = {
                'completed': list(self.completed),
                'completed_details': self.completed_details
            }
            with open(temp, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            if self.state_file.exists():
                self.state_file.unlink()
            temp.rename(self.state_file)
        except Exception as e:
            logging.error(f"❌ Error guardando estado: {e}")

    def _load_failures(self):
        if self.fail_file.exists():
            try:
                with open(self.fail_file, 'r', encoding='utf-8') as f:
                    self.failures = json.load(f)
                logging.info(f"📂 Fallos previos cargados: {len(self.failures)} modelos.")
            except Exception:
                self.failures = []
        else:
            self.failures = []

    def _save_failures(self):
        temp = self.fail_file.with_suffix('.tmp')
        try:
            with open(temp, 'w', encoding='utf-8') as f:
                json.dump(self.failures, f, indent=2)
            if self.fail_file.exists():
                self.fail_file.unlink()
            temp.rename(self.fail_file)
        except Exception as e:
            logging.error(f"❌ Error guardando fallos: {e}")

    def mark_completed(self, codigo: str, folio: str = ""):
        """Marca un modelo como completado y guarda el folio y timestamp."""
        self.completed.add(codigo)
        self.completed_details[codigo] = {
            'folio': folio if folio else '',
            'timestamp': datetime.now().isoformat()
        }
        self._save_state()

    def is_completed(self, codigo: str) -> bool:
        return codigo in self.completed

    def mark_failed(self, codigo: str, error: str, traceback: str = ""):
        for f in self.failures:
            if f['codigo'] == codigo:
                return
        self.failures.append({
            'codigo': codigo,
            'timestamp': datetime.now().isoformat(),
            'error': error,
            'traceback': traceback[:500]
        })
        self._save_failures()

    def get_pending(self, models: list) -> list:
        completed_set = self.completed
        failed_codes = {f['codigo'] for f in self.failures}
        pending = [m for m in models if m['codigo'] not in completed_set and m['codigo'] not in failed_codes]
        skipped = len(models) - len(pending)
        if skipped:
            logging.info(f"⏭️ {skipped} modelos ya procesados o en fallos. Saltando.")
        return pending

    def get_summary(self):
        return {
            'processed': len(self.completed),
            'failed': len(self.failures)
        }

    def get_completed_details(self) -> dict:
        """Retorna el diccionario de detalles de modelos completados."""
        return self.completed_details

    def export_report_excel(self):
        """Exporta un archivo Excel con los modelos completados y sus folios."""
        if not self.completed_details:
            logging.warning("⚠️ No hay modelos completados para exportar.")
            return None

        data = []
        for codigo, detalles in self.completed_details.items():
            data.append({
                'Código': codigo,
                'Folio': detalles.get('folio', ''),
                'Fecha': detalles.get('timestamp', '')
            })
        df = pd.DataFrame(data)

        # Guardar en Excel
        excel_path = Config.BASE_DIR / f"reporte_folios_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        df.to_excel(excel_path, index=False, sheet_name='Modelos Procesados')
        logging.info(f"📄 Reporte Excel generado: {excel_path}")
        return excel_path

    def generar_reporte_html(self):
        """Genera un reporte HTML con el resumen del lote (incluye folios)."""
        try:
            from jinja2 import Template
        except ImportError:
            logging.warning("⚠️ Jinja2 no instalado. No se generará reporte HTML.")
            return

        template_str = """
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <title>Reporte de Procesamiento VUCEM</title>
            <style>
                body { font-family: Arial, sans-serif; margin: 20px; }
                h1 { color: #2c3e50; }
                .summary { display: flex; gap: 20px; flex-wrap: wrap; }
                .card { border: 1px solid #ccc; padding: 15px; border-radius: 8px; min-width: 150px; }
                .success { background: #d4edda; border-color: #28a745; }
                .failed { background: #f8d7da; border-color: #dc3545; }
                table { border-collapse: collapse; width: 100%; margin-top: 20px; }
                th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
                th { background-color: #f2f2f2; }
                .footer { margin-top: 30px; font-size: 0.9em; color: #777; }
            </style>
        </head>
        <body>
            <h1>Reporte de Procesamiento VUCEM</h1>
            <p>Fecha: {{ fecha }}</p>
            <div class="summary">
                <div class="card success">
                    <strong>✅ Éxitos</strong><br>
                    {{ success_count }}
                </div>
                <div class="card failed">
                    <strong>❌ Fallos</strong><br>
                    {{ fail_count }}
                </div>
                <div class="card">
                    <strong>⏱️ Tiempo total</strong><br>
                    {{ tiempo_total }}
                </div>
            </div>

            <h2>Modelos procesados con éxito</h2>
            {% if success_list %}
            <table>
                <tr><th>Código</th><th>Folio</th><th>Timestamp</th></tr>
                {% for item in success_list %}
                <tr><td>{{ item.codigo }}</td><td>{{ item.folio }}</td><td>{{ item.timestamp }}</td></tr>
                {% endfor %}
            </table>
            {% else %}
            <p>No hay modelos exitosos.</p>
            {% endif %}

            <h2>Modelos fallidos</h2>
            {% if fail_list %}
            <table>
                <tr><th>Código</th><th>Error</th><th>Timestamp</th></tr>
                {% for item in fail_list %}
                <tr><td>{{ item.codigo }}</td><td>{{ item.error }}</td><td>{{ item.timestamp }}</td></tr>
                {% endfor %}
            </table>
            {% else %}
            <p>No hay modelos fallidos.</p>
            {% endif %}

            <div class="footer">
                Generado por VUCEM Automation Bot - {{ fecha }}
            </div>
        </body>
        </html>
        """
        template = Template(template_str)
        success_list = []
        for codigo in self.completed:
            detalles = self.completed_details.get(codigo, {})
            success_list.append({
                'codigo': codigo,
                'folio': detalles.get('folio', ''),
                'timestamp': detalles.get('timestamp', '')
            })
        fail_list = self.failures

        html = template.render(
            fecha=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            success_count=len(self.completed),
            fail_count=len(self.failures),
            tiempo_total="N/A",
            success_list=success_list,
            fail_list=fail_list
        )
        with open(self.report_file, 'w', encoding='utf-8') as f:
            f.write(html)
        logging.info(f"📄 Reporte HTML generado: {self.report_file}")