"""
LangGraph 기반 단일 엑셀 분류 파이프라인 (AsyncOpenAI)

흐름:
  init → read_rows → compare(이전 누적) → categorize(LLM)
                                              │
                                        aggregate → review(LLM) → finalize → END

특징:
  - 엑셀 파일 1개 하드코딩 (EXCEL_PATH)
  - categories.json 기준을 system prompt에 주입 (B안)
  - aggregate(집계) + LLM review(검증) 으로 전체 점검 (옵션 c)
  - AsyncOpenAI 클라이언트 (base_url + default_headers)

엑셀 읽기:
  1차 xlwings → 2차 pandas+openpyxl 폴백

컬럼 (하드코딩):
  - "진행단계"
  - "파일명"
"""

from __future__ import annotations

import asyncio
import json
import os
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, TypedDict

from openai import AsyncOpenAI
from langgraph.graph import StateGraph, END


# ──────────────────────────────────────────────────────────────────────────────
# 0. 하드코딩 상수 (사용자가 직접 수정)
# ──────────────────────────────────────────────────────────────────────────────
ROOT = Path(__file__).resolve().parent

# 입력 엑셀 파일 1개 (하드코딩)
EXCEL_PATH = ROOT / "data" / "input.xlsx"

# 컬럼명 (하드코딩)
COL_STEP = "진행단계"
COL_FILE = "파일명"

# 분류 기준
CATEGORIES_JSON = ROOT / "categories.json"

# ─────────────────────────────────────────────────────────────────────────
# OpenAI 호환 엔드포인트 설정
# ─────────────────────────────────────────────────────────────────────────
# [HARDCODE HERE] 아래 placeholder 를 실제 값으로 교체하면 환경변수 없이 동작합니다.
# 환경변수(OPENAI_BASE_URL / OPENAI_API_KEY / OPENAI_MODEL) 가 설정되어 있으면 그 값이 우선.
# 두 방식 다 사용 가능하며, 아래 None / 빈 문자열을 실제 값으로 바꾸면 하드코딩 완료.

# 예) OPENAI_BASE_URL = "https://api.openai.com/v1"
# 예) OPENAI_BASE_URL = "https://your-internal-gateway.example.com/v1"
OPENAI_BASE_URL: Optional[str] = os.getenv("OPENAI_BASE_URL") or None  # <-- HARDCODE

# 예) OPENAI_API_KEY = "sk-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
OPENAI_API_KEY: Optional[str] = os.getenv("OPENAI_API_KEY") or None  # <-- HARDCODE

# 예) OPENAI_MODEL = "gpt-oss-20b"  /  "gpt-oss-120b"  /  "내부-모델-id"
# 본 파이프라인은 GPT-OSS(오픈소스/로컬) 모델 사용을 전제로 함.
# → response_format 인자 미지원. JSON 출력은 프롬프트 + 파서로 강제.
OPENAI_MODEL: str = os.getenv("OPENAI_MODEL", "gpt-oss-20b")  # <-- HARDCODE

# ─────────────────────────────────────────────────────────────────────────
# 추가 HTTP 헤더 (AsyncOpenAI default_headers 로 주입됨)
# ─────────────────────────────────────────────────────────────────────────
# [HARDCODE HERE] 아래 dict 의 주석을 해제하고 실제 헤더 값으로 교체하세요.
# 환경변수 OPENAI_EXTRA_HEADERS (JSON 문자열) 를 추가로 병합할 수 있습니다.
DEFAULT_HEADERS: Dict[str, str] = {
    # "Authorization":   "Bearer <YOUR_TOKEN_HERE>",          # (필요 시) 별도 인증 헤더
    # "X-API-Key":       "<YOUR_API_KEY_HERE>",               # (필요 시) 게이트웨이 API Key
    # "X-Project-Id":    "<YOUR_PROJECT_ID>",                 # 프로젝트 식별
    # "X-Tenant":        "<YOUR_TENANT_ID>",                  # 멀티 테넌트
    # "X-Request-Source": "langgraph-excel-categorizer",      # 호출 출처 태깅
    # "User-Agent":      "langgraph-excel-categorizer/1.0",
}
# 환경변수 병합 (선택)
_extra = os.getenv("OPENAI_EXTRA_HEADERS")
if _extra:
    try:
        DEFAULT_HEADERS.update(json.loads(_extra))
    except json.JSONDecodeError:
        print(f"[WARN] OPENAI_EXTRA_HEADERS JSON 파싱 실패: {_extra}")


# ──────────────────────────────────────────────────────────────────────────────
# 1. AsyncOpenAI 클라이언트
# ──────────────────────────────────────────────────────────────────────────────
def make_client() -> AsyncOpenAI:
    kwargs: Dict[str, Any] = {}
    if OPENAI_API_KEY:
        kwargs["api_key"] = OPENAI_API_KEY
    if OPENAI_BASE_URL:
        kwargs["base_url"] = OPENAI_BASE_URL
    if DEFAULT_HEADERS:
        kwargs["default_headers"] = DEFAULT_HEADERS
    return AsyncOpenAI(**kwargs)


client: AsyncOpenAI = make_client()


# ──────────────────────────────────────────────────────────────────────────────
# 2. State
# ──────────────────────────────────────────────────────────────────────────────
class PipelineState(TypedDict):
    excel_path: str

    current_rows: List[Dict[str, Any]]
    previous_rows: List[Dict[str, Any]]

    new_rows: List[Dict[str, Any]]
    changed_rows: List[Dict[str, Any]]
    unchanged_rows: List[Dict[str, Any]]

    # 단일 파일이므로 {파일명: {카테고리: [...]}} 구조 유지
    categorized: Dict[str, Dict[str, List[Dict[str, Any]]]]
    read_engine: str

    criteria: Dict[str, Any]
    aggregate_report: Dict[str, Any]
    review_report: Dict[str, Any]


# ──────────────────────────────────────────────────────────────────────────────
# 3. 기준 로더
# ──────────────────────────────────────────────────────────────────────────────
def load_criteria() -> Dict[str, Any]:
    if not CATEGORIES_JSON.exists():
        raise FileNotFoundError(f"categories.json 없음: {CATEGORIES_JSON}")
    with open(CATEGORIES_JSON, "r", encoding="utf-8") as f:
        return json.load(f)


# ──────────────────────────────────────────────────────────────────────────────
# 4. 엑셀 읽기 (xlwings + pandas 폴백)
# ──────────────────────────────────────────────────────────────────────────────
def read_excel_xlwings(path: str) -> List[Dict[str, Any]]:
    import xlwings as xw

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        wb = app.books.open(str(Path(path).resolve()), read_only=True, update_links=False)
        try:
            sht = wb.sheets[0]
            data = sht.used_range.value
            if not data:
                return []
            if not isinstance(data[0], list):
                data = [data] if isinstance(data, list) else [[data]]
            headers = [str(h).strip() if h is not None else "" for h in data[0]]
            rows: List[Dict[str, Any]] = []
            for raw in data[1:]:
                if raw is None:
                    continue
                if not isinstance(raw, list):
                    raw = [raw]
                row = {headers[i]: raw[i] for i in range(min(len(headers), len(raw)))}
                if any(v is not None and str(v).strip() != "" for v in row.values()):
                    rows.append(row)
            return rows
        finally:
            wb.close()
    finally:
        app.quit()


def read_excel_pandas(path: str) -> List[Dict[str, Any]]:
    import pandas as pd

    df = pd.read_excel(path, engine="openpyxl")
    df = df.where(pd.notnull(df), None)
    return df.to_dict(orient="records")


def read_excel(path: str) -> tuple[List[Dict[str, Any]], str]:
    try:
        return read_excel_xlwings(path), "xlwings"
    except Exception as e:
        print(f"[read] xlwings 실패 ({type(e).__name__}: {e}) → pandas 폴백")
        return read_excel_pandas(path), "pandas"


def project_columns(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    return [{COL_STEP: r.get(COL_STEP), COL_FILE: r.get(COL_FILE)} for r in rows]


# ──────────────────────────────────────────────────────────────────────────────
# 5-0. JSON Extractor — GPT-OSS 응답에서 JSON 강제 추출
# ──────────────────────────────────────────────────────────────────────────────
# GPT-OSS 계열은 `response_format={"type":"json_object"}` 인자를 지원하지 않음.
# 따라서 모델 응답에 코드펜스/설명문/잡담이 섞일 수 있으며,
# 아래 extractor 가 다단계 폴백으로 첫 번째 유효 JSON 객체를 추출한다.
#
# 폴백 순서:
#   1) 그대로 json.loads
#   2) ```json ... ``` 또는 ``` ... ``` 코드펜스 내부
#   3) 중괄호 균형 스캔으로 첫 최상위 {...} 블록 추출
#   4) 모두 실패 시 ValueError
# ──────────────────────────────────────────────────────────────────────────────
_CODE_FENCE = re.compile(r"```(?:json)?\s*(\{.*?\}|\[.*?\])\s*```", re.DOTALL | re.IGNORECASE)


def _scan_balanced_json(text: str) -> Optional[str]:
    """중괄호/대괄호 균형 스캔으로 첫 최상위 JSON 블록 추출."""
    start = -1
    open_ch = ""
    close_ch = ""
    for i, ch in enumerate(text):
        if ch in "{[":
            start = i
            open_ch = ch
            close_ch = "}" if ch == "{" else "]"
            break
    if start < 0:
        return None
    depth = 0
    in_str = False
    esc = False
    for i in range(start, len(text)):
        ch = text[i]
        if in_str:
            if esc:
                esc = False
            elif ch == "\\":
                esc = True
            elif ch == '"':
                in_str = False
            continue
        if ch == '"':
            in_str = True
            continue
        if ch == open_ch:
            depth += 1
        elif ch == close_ch:
            depth -= 1
            if depth == 0:
                return text[start:i + 1]
    return None


def extract_json(text: str, *, expect: str = "object") -> Any:
    """GPT-OSS 응답에서 JSON 추출. expect='object' | 'array'.
    실패 시 ValueError(원문 일부 포함) 발생.
    """
    if text is None:
        raise ValueError("LLM 응답이 None")
    text = text.strip()
    if not text:
        raise ValueError("LLM 응답이 빈 문자열")

    # 1) 그대로 파싱
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # 2) 코드펜스
    m = _CODE_FENCE.search(text)
    if m:
        try:
            return json.loads(m.group(1))
        except json.JSONDecodeError:
            pass

    # 3) 균형 스캔
    block = _scan_balanced_json(text)
    if block:
        try:
            return json.loads(block)
        except json.JSONDecodeError:
            pass

    preview = text[:300].replace("\n", "\\n")
    raise ValueError(f"LLM 응답에서 JSON 추출 실패 (expect={expect}): {preview!r}")


async def llm_chat_json(
    *,
    system: str,
    user: str,
    expect: str = "object",
    temperature: float = 0.1,
    max_retries: int = 2,
) -> Any:
    """GPT-OSS 호환 JSON-강제 LLM 호출.

    - response_format 인자를 사용하지 않음 (GPT-OSS 미지원).
    - 프롬프트에 JSON-only 가이드를 명시하고, 응답은 extract_json 으로 파싱.
    - 파싱 실패 시 '직전 출력은 JSON 이 아니었음. JSON 만 다시 출력하라' 메시지로 재시도.
    """
    json_guard = (
        "\n\n[출력 규약 — 엄수]\n"
        f"- 응답은 오직 하나의 JSON {expect} 만 출력한다.\n"
        "- 코드펜스(```), 설명, 머리말/꼬리말, 사고 과정 출력 금지.\n"
        "- 응답의 첫 문자는 '{' 또는 '[' 이어야 한다.\n"
        "- 키/문자열은 모두 큰따옴표 사용. 후행 콤마 금지.\n"
    )
    messages: List[Dict[str, str]] = [
        {"role": "system", "content": system + json_guard},
        {"role": "user", "content": user},
    ]

    last_err: Optional[Exception] = None
    last_raw: str = ""
    for attempt in range(max_retries + 1):
        response = await client.chat.completions.create(
            model=OPENAI_MODEL,
            messages=messages,
            temperature=temperature,
        )
        last_raw = response.choices[0].message.content or ""
        try:
            return extract_json(last_raw, expect=expect)
        except ValueError as e:
            last_err = e
            # 재시도: 직전 응답을 보여주고 JSON-only 재요청
            messages.append({"role": "assistant", "content": last_raw})
            messages.append({
                "role": "user",
                "content": (
                    "직전 응답은 유효한 JSON 이 아니다. "
                    f"동일 작업을 다시 수행하되, 오직 하나의 JSON {expect} 만 출력하라. "
                    "코드펜스/설명/머리말 모두 금지."
                ),
            })
            print(f"[llm][retry {attempt + 1}/{max_retries}] JSON 파싱 실패: {e}")

    raise RuntimeError(f"LLM JSON 응답 확보 실패 ({max_retries}회 재시도 후): {last_err}")


# ──────────────────────────────────────────────────────────────────────────────
# 5. LLM 호출 (async)
# ──────────────────────────────────────────────────────────────────────────────
async def llm_categorize_with_criteria(
    rows: List[Dict[str, Any]],
    criteria: Dict[str, Any],
) -> Dict[str, List[Dict[str, Any]]]:
    if not rows:
        return {}

    allowed = [c["name"] for c in criteria["categories"]] + [criteria.get("default_category", "기타")]
    payload = [
        {"index": i, "data": f"[{r.get(COL_STEP)}] {r.get(COL_FILE)}",
         "raw": {COL_STEP: r.get(COL_STEP), COL_FILE: r.get(COL_FILE)}}
        for i, r in enumerate(rows)
    ]

    sys_prompt = (
        "당신은 엄격한 데이터 분류기입니다. 아래 '분류 기준 JSON'을 절대 기준으로 사용하여 "
        "각 항목을 정확히 하나의 카테고리에 배정하세요.\n\n"
        "규칙:\n"
        "1. 카테고리 이름은 반드시 기준 JSON의 categories[].name 또는 default_category 만 사용.\n"
        "2. 어떤 카테고리에도 명확히 부합하지 않으면 default_category 로 분류.\n"
        "3. examples 와 description 을 모두 참고.\n"
        f"4. 허용 카테고리 목록: {allowed}\n\n"
        "응답 형식 (JSON 객체 1개, 그 외 텍스트 절대 금지):\n"
        '{"카테고리명": [index 정수, ...], ...}\n'
        "예) {\"진행중\": [0, 2], \"완료\": [1], \"기타\": [3]}\n\n"
        "[분류 기준 JSON]\n"
        + json.dumps(criteria, ensure_ascii=False, indent=2)
    )

    mapping: Dict[str, List[int]] = await llm_chat_json(
        system=sys_prompt,
        user=json.dumps(payload, ensure_ascii=False),
        expect="object",
        temperature=0.1,
    )

    if not isinstance(mapping, dict):
        raise ValueError(f"분류 응답이 객체가 아님: {type(mapping).__name__}")

    result: Dict[str, List[Dict[str, Any]]] = {}
    for cat, idx_list in mapping.items():
        if not isinstance(idx_list, list):
            continue
        result[str(cat)] = [rows[i] for i in idx_list if isinstance(i, int) and 0 <= i < len(rows)]
    return result


async def llm_review_classification(
    categorized: Dict[str, Dict[str, List[Dict[str, Any]]]],
    criteria: Dict[str, Any],
) -> Dict[str, Any]:
    flat = []
    index_map = {}
    idx = 0
    for fname, buckets in categorized.items():
        for cat, items in buckets.items():
            if cat.startswith("_meta_"):
                continue
            for it in items:
                flat.append({
                    "index": idx,
                    "source_file": fname,
                    "category": cat,
                    "진행단계": it.get(COL_STEP),
                    "파일명": it.get(COL_FILE),
                })
                index_map[idx] = (fname, cat, it)
                idx += 1

    if not flat:
        return {"flags": [], "summary": "분류된 항목 없음"}

    instruction = criteria.get("review", {}).get("instruction", "")
    sys_prompt = (
        "당신은 분류 검증 감사관입니다. 아래 '분류 기준 JSON'과 '분류 결과 전체'를 비교하여 "
        "오분류로 의심되는 항목을 식별하세요.\n\n"
        f"감사 지시: {instruction}\n\n"
        "응답 형식 (JSON 객체):\n"
        "{\n"
        '  "flags": [\n'
        '    {"index": <int>, "current_category": "...", "suggested_category": "...", "reason": "..."}\n'
        "  ],\n"
        '  "summary": "전체 검증 요약 문장"\n'
        "}\n\n"
        "[분류 기준 JSON]\n"
        + json.dumps(criteria, ensure_ascii=False, indent=2)
    )

    parsed = await llm_chat_json(
        system=sys_prompt,
        user=json.dumps(flat, ensure_ascii=False),
        expect="object",
        temperature=0.1,
    )

    if not isinstance(parsed, dict):
        parsed = {"flags": [], "summary": f"검증 응답 형식 오류: {type(parsed).__name__}"}

    enriched = []
    for flag in parsed.get("flags", []) or []:
        if not isinstance(flag, dict):
            continue
        i = flag.get("index")
        if isinstance(i, int) and i in index_map:
            fname, cat, item = index_map[i]
            enriched.append({**flag, "source_file": fname, "item": item})
    parsed["flags"] = enriched
    return parsed


# ──────────────────────────────────────────────────────────────────────────────
# 6. 노드 (모두 async — AsyncOpenAI 사용 노드 포함)
# ──────────────────────────────────────────────────────────────────────────────
async def node_init(state: PipelineState) -> PipelineState:
    state["criteria"] = load_criteria()
    state["excel_path"] = str(EXCEL_PATH)
    print(f"[init] excel={EXCEL_PATH.name} categories={[c['name'] for c in state['criteria']['categories']]}")
    return state


async def node_read_rows(state: PipelineState) -> PipelineState:
    if not Path(state["excel_path"]).exists():
        raise FileNotFoundError(f"엑셀 파일 없음: {state['excel_path']}")
    raw, engine = read_excel(state["excel_path"])
    state["current_rows"] = project_columns(raw)
    state["read_engine"] = engine
    print(f"[read] {Path(state['excel_path']).name}: {len(state['current_rows'])} rows ({engine})")
    return state


def _row_key(row: Dict[str, Any]) -> str:
    return json.dumps(row, sort_keys=True, ensure_ascii=False, default=str)


async def node_compare(state: PipelineState) -> PipelineState:
    prev = state["previous_rows"]
    curr = state["current_rows"]
    prev_keys = {_row_key(r) for r in prev}
    prev_by_id = {r.get(COL_FILE): r for r in prev if r.get(COL_FILE) is not None}

    new_rows, changed_rows, unchanged_rows = [], [], []
    for r in curr:
        rk = _row_key(r)
        fid = r.get(COL_FILE)
        if rk in prev_keys:
            unchanged_rows.append(r)
        elif fid is not None and fid in prev_by_id:
            changed_rows.append({"before": prev_by_id[fid], "after": r})
        else:
            new_rows.append(r)

    state["new_rows"] = new_rows
    state["changed_rows"] = changed_rows
    state["unchanged_rows"] = unchanged_rows
    print(f"[compare] new={len(new_rows)} changed={len(changed_rows)} unchanged={len(unchanged_rows)}")
    return state


async def node_categorize(state: PipelineState) -> PipelineState:
    file_name = Path(state["excel_path"]).name
    bucket: Dict[str, List[Dict[str, Any]]] = {}

    targets = state["new_rows"] + [c["after"] for c in state["changed_rows"]]
    if targets:
        bucket = await llm_categorize_with_criteria(targets, state["criteria"])

    if state["unchanged_rows"]:
        bucket.setdefault("동일(변경없음)", []).extend(state["unchanged_rows"])
    if state["changed_rows"]:
        bucket.setdefault("_meta_변경이력", []).extend(state["changed_rows"])

    state["categorized"][file_name] = bucket
    print(f"[categorize] {file_name}: {sorted(bucket.keys())}")
    return state


async def node_aggregate(state: PipelineState) -> PipelineState:
    by_category: Dict[str, int] = {}
    by_file_category: Dict[str, Dict[str, int]] = {}
    total = 0
    meta_count = 0

    for fname, buckets in state["categorized"].items():
        by_file_category[fname] = {}
        for cat, items in buckets.items():
            n = len(items)
            by_file_category[fname][cat] = n
            if cat.startswith("_meta_"):
                meta_count += n
                continue
            by_category[cat] = by_category.get(cat, 0) + n
            total += n

    allowed = {c["name"] for c in state["criteria"]["categories"]}
    default_cat = state["criteria"].get("default_category", "기타")
    unknown = [c for c in by_category if c not in allowed and c != default_cat and c != "동일(변경없음)"]
    default_share = by_category.get(default_cat, 0) / total if total else 0.0

    report = {
        "total_items": total,
        "meta_items": meta_count,
        "by_category": dict(sorted(by_category.items(), key=lambda x: -x[1])),
        "by_file": by_file_category,
        "default_category_share": round(default_share, 3),
        "unknown_categories": unknown,
        "warnings": [],
    }
    if unknown:
        report["warnings"].append(f"기준 JSON에 없는 카테고리 발견: {unknown}")
    if default_share > 0.3:
        report["warnings"].append(f"default 카테고리 비율 과다 ({default_share:.1%})")

    state["aggregate_report"] = report
    print(f"[aggregate] total={total} by_category={report['by_category']}")
    for w in report["warnings"]:
        print(f"[aggregate][WARN] {w}")
    return state


async def node_review(state: PipelineState) -> PipelineState:
    if not state["criteria"].get("review", {}).get("enabled", True):
        state["review_report"] = {"skipped": True}
        return state
    report = await llm_review_classification(state["categorized"], state["criteria"])
    state["review_report"] = report
    print(f"[review] flags={len(report.get('flags', []))} summary={report.get('summary', '')[:80]}")
    return state


async def node_finalize(state: PipelineState) -> PipelineState:
    return state


# ──────────────────────────────────────────────────────────────────────────────
# 7. 그래프
# ──────────────────────────────────────────────────────────────────────────────
def build_graph():
    g = StateGraph(PipelineState)
    g.add_node("init", node_init)
    g.add_node("read_rows", node_read_rows)
    g.add_node("compare", node_compare)
    g.add_node("categorize", node_categorize)
    g.add_node("aggregate", node_aggregate)
    g.add_node("review", node_review)
    g.add_node("finalize", node_finalize)

    g.set_entry_point("init")
    g.add_edge("init", "read_rows")
    g.add_edge("read_rows", "compare")
    g.add_edge("compare", "categorize")
    g.add_edge("categorize", "aggregate")
    g.add_edge("aggregate", "review")
    g.add_edge("review", "finalize")
    g.add_edge("finalize", END)
    return g.compile()


# ──────────────────────────────────────────────────────────────────────────────
# 8. 실행 진입점 (async)
# ──────────────────────────────────────────────────────────────────────────────
async def run_async(previous_rows: Optional[List[Dict[str, Any]]] = None) -> Dict[str, Any]:
    app = build_graph()
    init: PipelineState = {
        "excel_path": "",
        "current_rows": [],
        "previous_rows": previous_rows or [],
        "new_rows": [],
        "changed_rows": [],
        "unchanged_rows": [],
        "categorized": {},
        "read_engine": "",
        "criteria": {},
        "aggregate_report": {},
        "review_report": {},
    }
    final = await app.ainvoke(init)
    return {
        "categorized": final["categorized"],
        "aggregate": final["aggregate_report"],
        "review": final["review_report"],
    }


def run(previous_rows: Optional[List[Dict[str, Any]]] = None) -> Dict[str, Any]:
    """동기 wrapper."""
    return asyncio.run(run_async(previous_rows))


if __name__ == "__main__":
    result = run()
    print(json.dumps(result, ensure_ascii=False, indent=2, default=str))
