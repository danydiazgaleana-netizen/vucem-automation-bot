from typing import List, Dict
import logging
from config import Config
from data_processor import DataProcessor

class DataManager:
    def __init__(self):
        Config.ensure_directories()
        self.processor = DataProcessor()
        self.models = []

    def _deduplicate_models(self, models: List[Dict]) -> List[Dict]:
        seen = set()
        unique = []
        for m in models:
            key = m.get('codigo')
            if key in seen:
                logging.warning(f"⚠️ Modelo duplicado detectado: {key} - se elimina la segunda ocurrencia.")
                continue
            seen.add(key)
            unique.append(m)
        if len(unique) < len(models):
            logging.info(f"🧹 Desduplicación: {len(models)} → {len(unique)} modelos únicos.")
        return unique

    def prepare_models(self) -> List[Dict]:
        modelos_procesados = self.processor.process()
        modelos_procesados = self._deduplicate_models(modelos_procesados)
        self.models = []
        for m in modelos_procesados:
            self.models.append({
                'codigo': m['codigo'],
                'nombre': m.get('nombre', m['codigo']),
                'csv_path': m.get('csv_path'),
                'data': m.get('data', {})
            })
        logging.info(f"📦 {len(self.models)} modelos listos para el bot.")
        return self.models