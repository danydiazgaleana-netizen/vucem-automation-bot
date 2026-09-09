from typing import List, Dict
import logging

try:
    from config_local import Config
except ImportError:
    from config import Config

from data_processor import DataProcessor


class DataManager:
    def __init__(self):
        Config.ensure_directories()
        self.processor = DataProcessor()
        self.models = []

    def _deduplicate_models(self, models: List[Dict]) -> List[Dict]:
        """
        Deduplica por código COMPLETO (no por nombre).
        Dos modelos con el mismo nombre pero distinto código (ej. tallas distintas)
        son modelos diferentes y deben procesarse por separado.
        Si el código completo se repite (mismo código, mismo precio), ahí sí
        es un duplicado real y se elimina la segunda ocurrencia.
        """
        seen = set()
        unique = []
        for m in models:
            key = m.get('codigo')  # código completo incluyendo nombre
            if key in seen:
                logging.warning(
                    f"⚠️ Código duplicado: '{key}' — se elimina la segunda ocurrencia."
                )
                continue
            seen.add(key)
            unique.append(m)
        if len(unique) < len(models):
            logging.info(
                f"🧹 Desduplicación: {len(models)} → {len(unique)} modelos únicos."
            )
        return unique

    def prepare_models(self) -> List[Dict]:
        modelos_procesados = self.processor.process()
        modelos_procesados = self._deduplicate_models(modelos_procesados)
        self.models = []
        for m in modelos_procesados:
            self.models.append({
                'codigo':   m['codigo'],
                'nombre':   m.get('nombre', m['codigo']),
                'csv_path': m.get('csv_path'),
                'data':     m.get('data', {})
            })
        logging.info(f"📦 {len(self.models)} modelos listos para el bot.")
        return self.models
# dm
# dm
