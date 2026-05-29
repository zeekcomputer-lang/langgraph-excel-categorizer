# HANDOFF.md — 다음 AI Agent 인수인계 문서

> **프로젝트:** langgraph-excel-categorizer
> **GitHub:** https://github.com/zeekcomputer-lang/langgraph-excel-categorizer
> **로컬:** `~/.openclaw/workspace/projects/langgraph-excel-categorizer/`
> **최종 업데이트:** 2026-05-29
> **상태:** 코드 작성 완료 / 실행 검증 보류 (사용자 환경 gpt-oss 기동 + 입력 엑셀 배치 필요)

---

## §0. 30초 요약

단일 엑셀 파일을 읽어 **`categories.json` 기준으로 LLM이 분류** + **집계(aggregate)** + **LLM 검증(review)** 까지 수행하는 LangGraph 비동기 파이프라인. AsyncOpenAI + GPT-OSS 전제. `response_format` 미사용, **프롬프트 가드 + extract_json 3단 파서 + 재시도** 로 JSON 응답 강제. 단일 파일 `categorizer.py` 633줄.

**이 repo의 LLM 호출 패턴(`llm_chat_json` + `extract_json` + HARDCODE placeholder)이 프로젝트 간 공유 표준임.** `deep-doc-pipeline`의 `src/llm.py`도 이 패턴으로 정렬되어 있다.

다음 작업자가 가장 먼저 할 것:
1. 이 문서 §0~3 통독 (5분)
2. `categories.json` 확인 — 분류 기준 전체 정의 (사용자가 직접 편집)
3. `data/input.xlsx` 배치 + `OPENAI_API_KEY`/`OPENAI_BASE_URL` 설정 후 실행 검증

---

## §1. 파일 지도

```
projects/langgraph-excel-categorizer/
├── README.md                  ★ 실행 가이드 + AsyncOpenAI 설정 + JSON 강제 전략 문서
├── HANDOFF.md                 ★ 이 문서
├── LESSONS.md                 누적 교훈
├── categorizer.py             ★ 전체 파이프라인 단일 파일 (633줄)
├── categories.json            분류 기준 JSON (5 카테고리 + default + review)
├── requirements.txt           langgraph, openai, pandas, openpyxl, pydantic
├── requirements-windows.txt   Windows 전용 (xlwings + pywin32 추가)
├── .gitignore
├── data/
│   └── .gitkeep               ★ input.xlsx 미배치 상태
├── tools/
│   ├── setup.py               1회성 venv + pip install (크로스 플랫폼)
│   └── launch.py              실행 진입점 (환경 점검 → categorizer.run())
└── windows/
    ├── setup.bat              Windows 셋업 wrapper
    └── start.bat              Windows 실행 wrapper
```

---

## §2. 아키텍처 요약

### 파이프라인 흐름
```
init → read_rows → compare → categorize(LLM async)
                                    ↓
                              aggregate → review(LLM async) → finalize → END
```

### 노드 설명

| 노드 | LLM? | 역할 |
|------|------|------|
| `node_init` | ✗ | `categories.json` 로드, 엑셀 경로 설정 |
| `node_read_rows` | ✗ | 엑셀 읽기 (xlwings 1차 → pandas 폴백) |
| `node_compare` | ✗ | 이전 실행 결과와 diff (new/changed/unchanged) |
| `node_categorize` | ✔ | `llm_categorize_with_criteria()` — 행별 카테고리 배정 |
| `node_aggregate` | ✗ | 카테고리별 집계 + 경고(default 비율 과다, 미등록 카테고리) |
| `node_review` | ✔ | `llm_review_classification()` — 오분류 의심 항목 플래그 |
| `node_finalize` | ✗ | 패스 스루 |

### LLM 호출 구조 (이 repo가 표준)

```python
# 1. HARDCODE placeholder 패턴
OPENAI_BASE_URL = os.getenv("OPENAI_BASE_URL") or None  # <-- HARDCODE
OPENAI_API_KEY  = os.getenv("OPENAI_API_KEY") or None    # <-- HARDCODE
OPENAI_MODEL    = os.getenv("OPENAI_MODEL", "gpt-oss-20b")

# 2. DEFAULT_HEADERS + OPENAI_EXTRA_HEADERS 병합
DEFAULT_HEADERS: Dict[str, str] = { ... }

# 3. AsyncOpenAI 클라이언트
client: AsyncOpenAI = make_client()

# 4. JSON 강제 단일 진입점
async def llm_chat_json(*, system, user, expect, temperature, max_retries) -> Any:
    # (a) 프롬프트 가드 자동 첨부 (출력 규약 + JSON-only 규칙)
    # (b) client.chat.completions.create — response_format 미사용
    # (c) extract_json() 3단 파서 (raw → 코드펜스 → 균형 스캔)
    # (d) 실패 시 직전 응답을 assistant로 넘겨 재요청 (기본 2회)
```

**다른 프로젝트에서 LLM 호출 신규 작성 시 이 4단계 패턴을 따를 것.**

---

## §3. 핵심 설계 결정

### 3-1. response_format 미사용
GPT-OSS(vLLM/Ollama/LM Studio 등)는 `response_format={"type":"json_object"}` 미지원.
프롬프트 규약 + 3단 파서 + 재시도로 대체.

### 3-2. 분류 기준 외부화
`categories.json` — 5개 카테고리(대기/진행중/완료/보류/오류) + default(기타) + review 설정.
코드 수정 없이 기준만 교체하면 다른 도메인 분류에 즉시 적용 가능.

### 3-3. 엑셀 읽기 2엔진
xlwings(COM 기반, Windows DRM 엑셀 대응) → pandas+openpyxl(크로스 플랫폼 폴백).
Linux에서는 자동으로 pandas 경로.

### 3-4. compare 노드 (diff)
`previous_rows`를 주입하면 변경분만 LLM 분류 호출. 미주입 시 전체 행 분류.
반복 실행 시 API 비용 절감 목적.

### 3-5. review 노드 (교차 검증)
`categories.json`의 `review.enabled: true` 시 활성.
분류 완료 결과 전체를 LLM에 넘겨 오분류 의심 항목 플래그.

---

## §4. 절대 준수 사항

1. **`response_format` 인자 사용 금지** — GPT-OSS 미지원. 프롬프트 가드 + `extract_json()` 로만 JSON 강제.
2. **LangChain LLM 래퍼 금지** — `openai.AsyncOpenAI` 직접 사용.
3. **`llm_chat_json()` 단일 진입점** 유지 — LLM 호출은 이 함수를 통해서만. 직접 `client.chat.completions.create` 금지.
4. **엑셀 경로 하드코딩** — `EXCEL_PATH` 상수. 변경 시 README 동시 갱신.
5. **컬럼 이름 하드코딩** — `COL_STEP`, `COL_FILE`. 실제 엑셀에 맞게 교체 필요.

---

## §5. 진행 현황

### ✅ 완료
- [x] 엑셀 읽기 2엔진 (xlwings + pandas 폴백)
- [x] LangGraph 비동기 파이프라인 7노드 (init → finalize)
- [x] `categories.json` 기준 B안 (system prompt에 JSON 주입)
- [x] aggregate 집계 + LLM review 검증
- [x] GPT-OSS 대응 — response_format 제거, extract_json 3단 파서
- [x] HARDCODE placeholder 패턴 (base_url/api_key/model/headers)
- [x] Windows 셋업/실행 분리 (setup.bat / start.bat)
- [x] 크로스 플랫폼 Python 셋업 (tools/setup.py)
- [x] GitHub 업로드

### ⏸️ 보류 (사용자 환경 측)
- [ ] `data/input.xlsx` 배치 — 현재 `.gitkeep`만 존재
- [ ] `OPENAI_BASE_URL` / `OPENAI_API_KEY` 설정 (환경변수 또는 코드 내 하드코딩)
- [ ] 실제 LLM 호출 실행 검증
- [ ] 실행 시 발견되는 이슈 디버깅

### 🟡 미반영 권장 보강
- [ ] `previous_rows` 영속 저장소 (JSON 파일) — 현재는 매 실행마다 전체 분류
- [ ] 출력 파일 저장 (JSON/Excel) — 현재 stdout print만
- [ ] 대용량 행(1000+) 시 배치 분할 호출 (토큰 한도 대비)
- [ ] 단위 테스트 (pytest)
- [ ] categories.json 스키마 검증 (Pydantic)

---

## §6. 다음 AI Agent 첫 작업 시나리오

### 시나리오 A: 사용자가 "실행했는데 에러 남"
1. 에러 메시지에서 `[llm][retry]` / `extract_json` / `ValueError` 키워드 확인
2. JSON 파싱 실패 → 모델이 한국어 설명을 섞어 응답할 가능성 → 프롬프트 가드 강화 고려
3. `OPENAI_BASE_URL` 미설정 → OpenAI 공식 API로 요청 → 과금/인증 에러

### 시나리오 B: "다른 엑셀 컬럼으로 바꾸고 싶다"
1. `categorizer.py` 상단 `COL_STEP`, `COL_FILE` 수정
2. `categories.json`의 examples도 새 컬럼에 맞게 갱신
3. `node_compare`의 `_row_key` 로직 확인 (모든 컬럼 기준 JSON 직렬화)

### 시나리오 C: "분류 카테고리를 바꾸고 싶다"
1. `categories.json` 만 수정하면 됨 — 코드 변경 불필요
2. `default_category` 도 함께 고려
3. `review.instruction` 에 새 기준에 맞는 검증 지시 추가

### 시나리오 D: "결과를 엑셀로 저장하고 싶다"
1. `node_finalize` 에 openpyxl/pandas 기반 엑셀 쓰기 로직 추가
2. 카테고리 시트 분할 또는 원본+카테고리 컬럼 병합 방식 선택

---

## §7. 빠른 디버깅 체크리스트

| 증상 | 1순위 의심 | 확인 방법 |
|------|----------|----------|
| `FileNotFoundError: data/input.xlsx` | 입력 파일 미배치 | `data/` 폴더에 엑셀 배치 |
| `OPENAI_API_KEY 환경변수가 필요합니다` | 환경변수 미설정 | `export OPENAI_API_KEY=...` |
| `LLM JSON 응답 확보 실패` | 모델이 JSON 외 텍스트 출력 | `llm_chat_json` retry 로그 확인, 프롬프트 가드 강화 |
| `기준 JSON에 없는 카테고리 발견` | LLM이 기준 외 카테고리 생성 | aggregate 경고 로그 확인, 프롬프트에 허용 목록 강조 |
| `default 카테고리 비율 과다` | 모델의 분류 정확도 부족 | categories.json examples 보강, 모델 변경 검토 |
| xlwings import 실패 (Linux) | 정상 — pandas 폴백 동작 | `[read] xlwings 실패 → pandas 폴백` 로그 확인 |

---

## §8. 참고 명령어 모음

```bash
# 로컬 진입
cd ~/.openclaw/workspace/projects/langgraph-excel-categorizer

# 셋업 (Linux/macOS)
python3 tools/setup.py

# 셋업 (Windows)
windows\setup.bat

# 환경변수 설정
export OPENAI_API_KEY="sk-..."
export OPENAI_BASE_URL="http://localhost:8000/v1"
export OPENAI_MODEL="gpt-oss-20b"

# 실행 (Linux/macOS)
.venv/bin/python tools/launch.py

# 실행 (Windows)
windows\start.bat

# AST 문법 검증
python3 -c "import ast; ast.parse(open('categorizer.py').read()); print('OK')"

# GitHub 동기화
git add -A && git commit -m "..." && git push
```

---

## §9. 크로스 프로젝트 참조

이 repo의 LLM 호출 패턴은 다음 프로젝트에서 **표준 출처**로 참조됨:

| 프로젝트 | 파일 | 정렬 상태 |
|---------|------|----------|
| [deep-doc-pipeline](https://github.com/zeekcomputer-lang/deep-doc-pipeline) | `src/llm.py` | ✅ v1.1-r1 (`442f9ba`, 2026-05-29) |

신규 LLM 파이프라인 프로젝트 시작 시, 이 repo의 `llm_chat_json` + `extract_json` + HARDCODE placeholder 패턴을 복사하여 시작할 것.

---

## §10. 다음 AI Agent 첫 5분 체크리스트

- [ ] 이 문서(`HANDOFF.md`) §0~4 읽기
- [ ] `categories.json` 열어 분류 기준 확인
- [ ] `categorizer.py` 상단 50줄 (설정 상수) + `llm_chat_json` 함수 읽기
- [ ] `git log --oneline` 커밋 히스토리 확인
- [ ] 사용자 첫 메시지에서 시나리오 A/B/C/D 분류
- [ ] 코드 수정 시 AST 검증 후 커밋
- [ ] 사용자에게 한국어, 사무적 톤, 간결하게 응답

---

_본 문서는 다음 AI Agent의 빠른 인계를 위해 작성됨. 추가 작업 시 §5, §7을 갱신할 것._
