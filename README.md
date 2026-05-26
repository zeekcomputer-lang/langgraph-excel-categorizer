# LangGraph Excel Categorizer

엑셀 파일을 순차 읽어 **기준 JSON 기반 LLM 분류** + **전체 점검(집계+검증)** 까지 수행하는 LangGraph 파이프라인.

## 핵심 로직

### 분류 기준 (B안)
- `categories.json` 파일에 카테고리 정의 (이름/설명/예시)
- LangGraph 시작 시 1회 로드 → LLM **system prompt에 주입**
- LLM이 매 항목을 기준 JSON 참조하여 카테고리 배정 (`temperature=0.1`)

### 전체 점검 (옵션 c)
모든 파일 처리 완료 후 두 단계 점검:
1. **aggregate** (집계 리포트) — 카테고리별 건수, 파일별 분포, default 비율, 미허용 카테고리 검출, 경고
2. **review** (LLM 검증) — 전체 분류 결과를 LLM에 다시 보여주고 오분류 의심 항목 플래그

## 입력 스펙
- 엑셀 컬럼 (하드코딩, `categorizer.py` 상단 상수):
  - `진행단계` (컬럼1)
  - `파일명` (컬럼2)
- 두 컬럼 결합 → 1 data row → LLM 입력

## 엑셀 엔진 (2중)
1. **xlwings** (1차) — Windows + Excel
2. **pandas + openpyxl** (폴백)

## 파이프라인
```
init(load criteria)
  ↓
load_file → read_rows → compare → categorize(LLM, criteria)
   ↑                                    │
   └── 다음 파일 있으면 loop ────────────┘
                                        │
                                  aggregate(집계)
                                        ↓
                                  review(LLM 검증)
                                        ↓
                                    finalize → END
```

---

## Windows 실행
```cmd
windows\setup.bat
set OPENAI_API_KEY=sk-...
windows\start.bat data1.xlsx data2.xlsx
```

## Linux / macOS
```bash
python3 tools/setup.py
export OPENAI_API_KEY=sk-...
.venv/bin/python tools/launch.py data1.xlsx data2.xlsx
```

---

## 출력 형식
```json
{
  "categorized": {
    "data1.xlsx": {
      "진행중": [ {"진행단계": "작성중", "파일명": "report_v2.docx"} ],
      "완료":   [ ... ],
      "동일(변경없음)": [ ... ],
      "_meta_변경이력": [ {"before": {...}, "after": {...}} ]
    }
  },
  "aggregate": {
    "total_items": 42,
    "by_category": {"진행중": 18, "완료": 12, "대기": 7, "기타": 5},
    "by_file": {"data1.xlsx": {"진행중": 10, "완료": 5}},
    "default_category_share": 0.119,
    "unknown_categories": [],
    "warnings": []
  },
  "review": {
    "flags": [
      {
        "index": 17,
        "current_category": "완료",
        "suggested_category": "진행중",
        "reason": "진행단계 '검토중'은 description 상 '진행중' 에 해당",
        "source_file": "data1.xlsx",
        "item": {"진행단계": "검토중", "파일명": "spec_v3.xlsx"}
      }
    ],
    "summary": "총 42건 중 1건 오분류 의심"
  }
}
```

## 기준 JSON (`categories.json`)
```json
{
  "categories": [
    {
      "name": "진행중",
      "description": "현재 작업이 수행되고 있는 상태",
      "examples": [{"진행단계": "작성중", "파일명": "report_v2.docx"}]
    }
  ],
  "default_category": "기타",
  "review": {
    "enabled": true,
    "instruction": "오분류 의심 항목 식별"
  }
}
```
- 사용자가 직접 수정 (향후 하드코딩)
- `name`, `description`, `examples` 모두 LLM이 참조
- `default_category`: 어디에도 부합하지 않는 항목의 fallback
- `review.enabled=false` 로 LLM 검증 단계 스킵 가능

## 비교 로직 (이전 파일과 누적 비교)
- 두 컬럼 값 완전 동일 → `unchanged`
- `파일명` 동일 + `진행단계` 다름 → `changed` (before/after 보존)
- 그 외 → `new` (LLM 분류 대상)

## 디렉터리
```
langgraph-excel-categorizer/
├── categorizer.py                # 파이프라인 본체
├── categories.json               # 분류 기준 (사용자 수정)
├── requirements.txt
├── requirements-windows.txt
├── README.md
├── .gitignore
├── tools/
│   ├── setup.py
│   └── launch.py
└── windows/
    ├── setup.bat
    └── start.bat
```

## 환경변수
| 이름 | 필수 | 설명 |
|---|---|---|
| `OPENAI_API_KEY` | ✔ | OpenAI API Key |
| `OPENAI_MODEL` | ✖ | 기본 `gpt-4o-mini` |
