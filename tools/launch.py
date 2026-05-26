"""실행 진입점: 환경 점검 후 categorizer 실행 (엑셀 경로는 categorizer.py에서 하드코딩)."""
from __future__ import annotations

import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))


def main() -> int:
    if not os.environ.get("OPENAI_API_KEY"):
        print("[ERROR] OPENAI_API_KEY 환경변수가 필요합니다.")
        return 3

    from categorizer import run, EXCEL_PATH
    import json

    if not EXCEL_PATH.exists():
        print(f"[ERROR] 하드코딩된 엑셀 파일이 없습니다: {EXCEL_PATH}")
        print(f"        해당 경로에 파일을 배치하거나 categorizer.py 의 EXCEL_PATH 상수를 수정하세요.")
        return 2

    result = run()
    print(json.dumps(result, ensure_ascii=False, indent=2, default=str))
    return 0


if __name__ == "__main__":
    sys.exit(main())
