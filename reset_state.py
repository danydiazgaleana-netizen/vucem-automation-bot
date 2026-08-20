import os
from pathlib import Path

for f in ["processed_models.json", "failed_models.json"]:
    p = Path(f)
    if p.exists():
        p.unlink()
        print(f"🗑️ Eliminado: {f}")
print("✅ Estado reseteado.")