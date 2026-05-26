"""1회성 셋업 스크립트: venv 생성 + 의존성 설치 (Windows / Linux / macOS 공통)."""
from __future__ import annotations

import os
import subprocess
import sys
import venv
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
VENV_DIR = ROOT / ".venv"
REQ = ROOT / "requirements.txt"


def venv_python() -> Path:
    if os.name == "nt":
        return VENV_DIR / "Scripts" / "python.exe"
    return VENV_DIR / "bin" / "python"


def main() -> int:
    print(f"[setup] root        : {ROOT}")
    print(f"[setup] venv        : {VENV_DIR}")
    print(f"[setup] python host : {sys.version.split()[0]}")

    if not VENV_DIR.exists():
        print("[setup] 가상환경 생성 중...")
        builder = venv.EnvBuilder(with_pip=True, clear=False, upgrade=False)
        builder.create(str(VENV_DIR))
    else:
        print("[setup] 가상환경 이미 존재 → 재사용")

    py = venv_python()
    if not py.exists():
        print(f"[ERROR] venv python 미발견: {py}")
        return 1

    print("[setup] pip 업그레이드...")
    rc = subprocess.call([str(py), "-m", "pip", "install", "--upgrade", "pip"])
    if rc != 0:
        return rc

    if REQ.exists():
        print(f"[setup] 의존성 설치: {REQ.name}")
        rc = subprocess.call([str(py), "-m", "pip", "install", "-r", str(REQ)])
        if rc != 0:
            return rc
    else:
        print("[WARN] requirements.txt 없음")

    print("[setup] 검증: import langgraph, openai, pandas, openpyxl")
    rc = subprocess.call([
        str(py), "-c",
        "import langgraph, openai, pandas, openpyxl; print('OK', langgraph.__version__)"
    ])
    if rc != 0:
        print("[ERROR] 패키지 검증 실패")
        return rc

    print("\n[OK] 셋업 완료")
    return 0


if __name__ == "__main__":
    sys.exit(main())
