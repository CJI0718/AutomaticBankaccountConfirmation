"""리소스 경로 — 일반 실행과 PyInstaller(.exe) 번들 양쪽에서 동작.

PyInstaller onefile은 번들을 임시폴더(sys._MEIPASS)에 풀고 거기에 configs/ 를 둔다
(빌드 시 --add-data "configs;configs"). frozen 여부에 따라 기준 경로를 바꾼다.
"""
from __future__ import annotations

import sys
from pathlib import Path


def base_dir() -> Path:
    if getattr(sys, "frozen", False):  # PyInstaller 번들
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
    return Path(__file__).resolve().parent.parent


CONFIG_DIR = base_dir() / "configs"
