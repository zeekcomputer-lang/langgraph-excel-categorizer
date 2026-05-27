# LangGraph Excel Categorizer (단일 엑셀 / AsyncOpenAI / **GPT-OSS 대응**)

엑셀 1개 파일을 읽어 **기준 JSON 기반 LLM 분류** + **전체 점검(집계+검증)** 까지 수행하는 LangGraph 비동기 파이프라인.

> **주의:** 본 파이프라인은 **GPT-OSS 계열(오픈소스 / 로컬)** 모델 사용을 전제로 한다.
> `response_format={"type":"json_object"}` 인자가 미지원이므로 사용하지 않으며,
> **프롬프트 가이드 + robust JSON extractor + 재시도 로직**으로 JSON 응답을 강제한다.

## 핵심 사양

| 항목 | 값 |
|---|---|
| 엑셀 파일 | **1개 하드코딩** (`data/input.xlsx`, `categorizer.py` `EXCEL_PATH` 상수) |
| 컬럼 | `진행단계`, `파일명` (하드코딩) |
| 엑셀 엔진 | xlwings 1차 → pandas+openpyxl 폴백 |
| LLM 클라이언트 | `AsyncOpenAI` (base_url + default_headers) |
| 대상 모델 | **GPT-OSS** (`gpt-oss-20b` 기본, OpenAI 호환 API 노출) |
| 호출 형태 | `response = await client.chat.completions.create(...)` — **`response_format` 인자 미사용** |
| JSON 강제 | 프롬프트 규약 + `extract_json()` (코드펜스/잡담 제거) + 최대 2회 재시도 |
| 분류 방식 | B안 — `categories.json` 을 system prompt 주입 |
| 점검 | aggregate(집계) + review(LLM 검증) |

## 파이프라인
```
init → read_rows → compare → categorize(LLM async)
                                    ↓
                              aggregate → review(LLM async) → finalize → END
```

## AsyncOpenAI 설정 (GPT-OSS 호환 엔드포인트)
`categorizer.py` 상단:
```python
OPENAI_BASE_URL = os.getenv("OPENAI_BASE_URL")   # 예: "http://localhost:8000/v1" (vLLM/Ollama/LM Studio 등)
OPENAI_API_KEY  = os.getenv("OPENAI_API_KEY")    # GPT-OSS 게이트웨이 토큰
OPENAI_MODEL    = os.getenv("OPENAI_MODEL", "gpt-oss-20b")

DEFAULT_HEADERS = {
    # "X-Project-Id": "langgraph-excel-categorizer",
    # "X-Tenant": "default",
}
```
- 하드코딩하거나 환경변수로 주입
- `OPENAI_EXTRA_HEADERS` (JSON 문자열) 로도 추가 헤더 병합 가능

클라이언트 생성 및 호출 (response_format 인자 없음):
```python
client = AsyncOpenAI(
    api_key=OPENAI_API_KEY,
    base_url=OPENAI_BASE_URL,
    default_headers=DEFAULT_HEADERS,
)

# 호출 — response_format 미사용 (GPT-OSS 미지원)
response = await client.chat.completions.create(
    model=OPENAI_MODEL,
    messages=[...],   # system 프롬프트에 'JSON-only' 규약이 자동 첨부됨
    temperature=0.1,
)
raw = response.choices[0].message.content
parsed = extract_json(raw, expect="object")   # 코드펜스/잡담 제거 + 균형 스캔
```

### JSON 강제 전략 (GPT-OSS 대응)
`llm_chat_json()` 단일 진입점이 아래 3단계로 JSON 응답을 확보한다:

1. **프롬프트 규약 자동 첨부** — system 메시지 끝에 `[출력 규약 — 엄수]` 블록 추가.
   - "오직 하나의 JSON", "코드펜스/설명 금지", "첫 문자는 `{` 또는 `[`", "후행 콤마 금지"
2. **`extract_json()` 파서** — 다음 순서로 최대한 구제:
   - (a) 그대로 `json.loads`
   - (b) ` ```json ... ``` ` 코드펜스 내부
   - (c) 중괄호 균형 스캔으로 첫 최상위 블록 추출
3. **재시도** — 파싱 실패 시 모델의 직전 응답을 assistant 메시지로 넘기고,
   user 메시지로 "JSON 만 다시 출력하라" 요청. 기본 **최대 2회**.

모두 실패 시 `RuntimeError` 로 파이프라인이 멈춘다 (silent 실패 없음).

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
