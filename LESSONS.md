# LESSONS.md — 누적 교훈 인덱스

> 다음 AI Agent가 같은 실수를 반복하지 않도록 정리한 교훈 카드.
> 이 프로젝트의 LLM 호출 패턴이 다른 프로젝트의 표준 출처임 — 변경 시 영향 범위 확인 필수.

## 인덱스

| ID | 분류 | 요약 |
|----|------|------|
| L-001 | 호환성 | response_format 미사용이 GPT-OSS 표준 — 프롬프트 가드 + 파서로 대체 |
| L-002 | 아키텍처 | 단일 파일 구조의 장단점 — 633줄까지는 유지, 이후 분리 |
| L-003 | 설계 | 분류 기준 외부화(JSON) → 코드 변경 없이 도메인 전환 가능 |
| L-004 | 환경 | xlwings COM은 Windows 전용 — Linux에서는 pandas 폴백 자동 |
| L-005 | 크로스PJT | 이 repo의 LLM 호출 패턴이 표준 — 수정 시 downstream 확인 |

---

## L-001: response_format 미사용이 GPT-OSS 환경의 표준

**상황:**
GPT-OSS(vLLM/Ollama/LM Studio)는 `response_format={"type":"json_object"}` 미지원.
`client.beta.chat.completions.parse` (Structured Outputs)도 미지원.

**대응:**
1. `response_format` 인자 자체를 코드에서 완전히 제거
2. JSON 출력 규약을 system 프롬프트 끝에 자동 첨부 (프롬프트 가드)
3. `extract_json()` 3단 폴백 파서: (a) raw → (b) 코드펜스 → (c) 균형 스캔
4. 실패 시 직전 응답을 assistant 메시지로 넘겨 "JSON만 다시 출력" 재요청

**원칙:**
> `response_format`은 보너스이지 필수가 아님. LLM API 호출 신규 작성 시 **미사용 전제**로 설계하고, 프롬프트 + 파서 + 재시도로 JSON을 강제할 것.

---

## L-002: 단일 파일 구조 (633줄)의 장단점

**장점:**
- 인수인계 시 파일 1개만 읽으면 전체 파악 가능
- import 경로 오류 없음
- IDE에서 한 파일 내 전문 검색 용이

**단점:**
- 700줄 초과 시 가독성 급감 (함수 경계 파악 어려움)
- 테스트 시 개별 함수 import 번거로움
- 다수 LLM 호출 패턴이 섞이면 리팩터 난이도 증가

**원칙:**
> 700줄 이하면 단일 파일 유지. 초과 시 `llm.py`(호출) / `nodes.py`(노드) / `io.py`(I/O) 3파일 분리 고려. deep-doc-pipeline이 분리 구조의 참조 예시.

---

## L-003: 분류 기준 외부화의 가치

**상황:**
초기 설계에서 카테고리를 코드 내 dict로 하드코딩하는 A안 vs 외부 JSON(B안) 선택지.

**결정:**
B안 (categories.json) 채택 — 코드 변경 없이 분류 기준 교체 가능.

**구조:**
```json
{
  "categories": [{"name": "...", "description": "...", "examples": [...]}],
  "default_category": "기타",
  "review": {"enabled": true, "instruction": "..."}
}
```

**원칙:**
> LLM에 주입되는 **판단 기준**은 코드가 아닌 외부 설정으로 분리할 것. 프롬프트 엔지니어링 수정이 코드 커밋 없이 가능해짐.

---

## L-004: xlwings COM 제약

**상황:**
xlwings는 Excel COM Automation 기반 — Windows + Excel 설치 필수.
Linux/macOS에서는 `ImportError` 또는 COM 미지원.

**대응:**
`read_excel()` 함수에서 xlwings 시도 → 실패 시 pandas+openpyxl 자동 폴백.
DRM 보호 엑셀은 Windows xlwings로만 열 수 있음 — 이 경우 Linux 실행 불가.

**원칙:**
> 엑셀 읽기는 항상 2엔진(xlwings → pandas) 폴백 구조. DRM 엑셀 여부를 사전에 확인하고, DRM이면 Windows 실행을 안내할 것.

---

## L-005: 이 repo가 LLM 호출 표준 출처

**상황:**
2026-05-29 기준, `deep-doc-pipeline/src/llm.py`가 이 repo의 패턴으로 정렬됨 (`442f9ba`).
이 repo의 `llm_chat_json` + `extract_json` + HARDCODE placeholder가 **프로젝트 간 공유 표준**.

**영향 범위:**

| 이 repo에서 변경하면 | 확인해야 할 곳 |
|-------------------|-------------|
| `extract_json()` 파서 로직 | deep-doc-pipeline `src/llm.py` |
| HARDCODE placeholder 패턴 | deep-doc-pipeline `src/llm.py` |
| `llm_chat_json()` 시그니처 | deep-doc-pipeline은 `structured_call()`이므로 직접 영향 없음 (패턴만 참조) |

**원칙:**
> 이 repo의 LLM 호출 패턴을 변경할 때는, HANDOFF.md §9의 downstream 목록을 확인하고 동기화 필요 여부를 판단할 것.

---

## 신규 교훈 추가 시 규칙

1. ID 부여: 다음 번호 (L-006, L-007, ...)
2. 인덱스 표 상단에 행 추가
3. 상세 카드는 ID 순서대로 본문 하단 추가
4. 분류 카테고리: 호환성 / 아키텍처 / 설계 / 환경 / 크로스PJT
5. 각 카드 구조: 상황 → 대응 → 원칙 (또는 영향 범위)

---

_본 문서는 이 프로젝트뿐 아니라 향후 LangGraph/LLM 파이프라인 작업에 참조 가능._
