"""실행 진입점: 셋업 상태 점검 후 categorizer 실행."""
from __future__ import annotations

import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))


def main(argv: list[str]) -> int:
    if not argv:
        print("사용법: python tools/launch.py file1.xlsx [file2.xlsx ...]")
        return 64

    if not os.environ.get("OPENAI_API_KEY"):
        print("[ERROR] OPENAI_API_KEY 환경변수가 필요합니다.")
        return 3

    # 파일 존재 검증
    missing = [p for p in argv if not Path(p).exists()]
    if missing:
        print(f"[ERROR] 파일 없음: {missing}")
        return 2

    from categorizer import run
    import json

    result = run(argv)
    print(json.dumps(result, ensure_ascii=False, indent=2, default=str))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
