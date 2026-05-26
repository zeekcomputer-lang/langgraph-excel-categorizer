# LangGraph Excel Categorizer (단일 엑셀 / AsyncOpenAI)

엑셀 1개 파일을 읽어 **기준 JSON 기반 LLM 분류** + **전체 점검(집계+검증)** 까지 수행하는 LangGraph 비동기 파이프라인.

## 핵심 사양

| 항목 | 값 |
|---|---|
| 엑셀 파일 | **1개 하드코딩** (`data/input.xlsx`, `categorizer.py` `EXCEL_PATH` 상수) |
| 컬럼 | `진행단계`, `파일명` (하드코딩) |
| 엑셀 엔진 | xlwings 1차 → pandas+openpyxl 폴백 |
| LLM 클라이언트 | `AsyncOpenAI` (base_url + default_headers) |
| 호출 형태 | `response = await client.chat.completions.create(...)` |
| 분류 방식 | B안 — `categories.json` 을 system prompt 주입 |
| 점검 | aggregate(집계) + review(LLM 검증) |

## 파이프라인
```
init → read_rows → compare → categorize(LLM async)
                                    ↓
                              aggregate → review(LLM async) → finalize → END
```

## AsyncOpenAI 설정
`categorizer.py` 상단:
```python
OPENAI_BASE_URL = os.getenv("OPENAI_BASE_URL")   # 예: "https://api.openai.com/v1"
OPENAI_API_KEY  = os.getenv("OPENAI_API_KEY")
OPENAI_MODEL    = os.getenv("OPENAI_MODEL", "gpt-4o-mini")

DEFAULT_HEADERS = {
    # "X-Project-Id": "langgraph-excel-categorizer",
    # "X-Tenant": "default",
}
```
- 하드코딩하거나 환경변수로 주입
- `OPENAI_EXTRA_HEADERS` (JSON 문자열) 로도 추가 헤더 병합 가능

클라이언트 생성:
```python
client = AsyncOpenAI(
    api_key=OPENAI_API_KEY,
    base_url=OPENAI_BASE_URL,
    default_headers=DEFAULT_HEADERS,
)

# 호출
response = await client.chat.completions.create(
    model=OPENAI_MODEL,
    response_format={"type": "json_object"},
    messages=[...],
    temperature=0.1,
)
```

---

## Windows 실행
```cmd
windows\setup.bat
set OPENAI_API_KEY=sk-...
REM 선택: set OPENAI_BASE_URL=https://your-proxy/v1
windows\start.bat
```

## Linux / macOS
```bash
python3 tools/setup.py
export OPENAI_API_KEY=sk-...
.venv/bin/python tools/launch.py
```

> 엑셀 경로는 CLI 인자가 아닌 **`categorizer.py` 의 `EXCEL_PATH` 상수**에서 변경합니다.
> 기본 경로: `<repo>/data/input.xlsx`

---

## 출력 형식
```json
{
  "categorized": {
    "input.xlsx": {
      "진행중": [ {"진행단계": "작성중", "파일명": "report_v2.docx"} ],
      "완료":   [ ... ],
      "동일(변경없음)": [ ... ],
      "_meta_변경이력": [ ... ]
    }
  },
  "aggregate": {
    "total_items": 42,
    "by_category": {"진행중": 18, "완료": 12, "대기": 7, "기타": 5},
    "default_category_share": 0.119,
    "unknown_categories": [],
    "warnings": []
  },
  "review": {
    "flags": [
      {"index": 17, "current_category": "완료",
       "suggested_category": "진행중", "reason": "...",
       "source_file": "input.xlsx", "item": {...}}
    ],
    "summary": "..."
  }
}
```

## 비교 로직 (이전 누적과의 비교)
파이프라인 호출 시 `run(previous_rows=[...])` 또는 `await run_async(previous_rows=[...])`로 이전 데이터 전달:
- 두 컬럼 값 완전 동일 → `unchanged`
- `파일명` 동일 + `진행단계` 다름 → `changed` (before/after 보존)
- 그 외 → `new` (LLM 분류 대상)

이전 데이터 미전달 시 모든 행이 `new` 처리.

## 기준 JSON (`categories.json`)
```json
{
  "categories": [
    {"name": "진행중", "description": "...", "examples": [...]},
    {"name": "완료",   "description": "...", "examples": [...]}
  ],
  "default_category": "기타",
  "review": {"enabled": true, "instruction": "..."}
}
```

## 디렉터리
```
langgraph-excel-categorizer/
├── categorizer.py                # 파이프라인 + AsyncOpenAI
├── categories.json               # 분류 기준
├── data/                         # ← 여기에 input.xlsx 배치
├── requirements.txt
├── requirements-windows.txt
├── README.md
├── tools/{setup,launch}.py
└── windows/{setup,start}.bat
```

## 환경변수
| 이름 | 필수 | 설명 |
|---|---|---|
| `OPENAI_API_KEY` | ✔ | API Key |
| `OPENAI_BASE_URL` | ✖ | OpenAI 호환 엔드포인트 |
| `OPENAI_MODEL` | ✖ | 기본 `gpt-4o-mini` |
| `OPENAI_EXTRA_HEADERS` | ✖ | JSON 문자열 형식의 추가 헤더 |
