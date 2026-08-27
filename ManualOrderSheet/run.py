# -*- coding: utf-8 -*-
"""수기운용지시서 생성 실행 진입점.

    python run.py 20260825
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from src.generate import main

if __name__ == "__main__":
    main()
