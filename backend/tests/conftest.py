"""Test setup: the backend modules import each other as top-level modules."""

import os
import sys
from pathlib import Path

BACKEND_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BACKEND_DIR))

# Keep tests hermetic: never pick up a developer's real .env values.
for var in ("EMAIL_SENDER", "EMAIL_PASSWORD", "ANTHROPIC_API_KEY", "AUTO_SEND_ENABLED", "EXCEL_SOURCE_PATH"):
    os.environ.pop(var, None)
