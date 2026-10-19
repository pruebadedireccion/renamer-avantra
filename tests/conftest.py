# -*- coding: utf-8 -*-
"""Configuracion comun de pytest: agrega src/ al path."""

import sys
from pathlib import Path

SRC = Path(__file__).resolve().parent.parent / "src"
sys.path.insert(0, str(SRC))