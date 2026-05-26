"""
LangGraph 기반 엑셀 행 반복 읽기 + LLM 카테고라이즈 + 전체 점검 파이프라인

흐름:
  load_file → read_rows → compare → categorize(LLM, B안)
       ↑                                  │
       └── 다음 파일 있으면 loop ──────────┘
                                          │
                                  aggregate(집계) → review(LLM 검증) → finalize → END

분류 방식 (B안):
  - categories.json 의 기준을 LLM system prompt에 주입
  - LLM이 매 항목을 기준 JSON 참조하여 카테고리 배정

점검 방식 (옵션 c):
  - aggregate: 카테고리별 건수/총합/누락 통계
  - review:    LLM에 전체 분류 결과 전달 → 오분류 의심 플래그

엑셀 읽기:
  1차: xlwings (Excel COM, DRM 호환)
  2차: pandas + openpyxl (폴백)

컬럼 (하드코딩):
  - "진행단계"
  - "파일명"
"""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any, Dict, List, Optional, TypedDict

from openai import OpenAI
from langgraph.graph import StateGraph, END


# ──────────────────────────────────────────────────────────────────────────────
# 0. 하드코딩 컬럼명 + 기준 JSON 경로
# ──────────────────────────────────────────────────────────────────────────────
COL_STEP = "진행단계"
COL_FILE = "파일명"

ROOT = Path(__file__).resolve().parent
CATEGORIES_JSON = ROOT / "categories.json"


# ──────────────────────────────────────────────────────────────────────────────
# 1. State
# ──────────────────────────────────────────────────────────────────────────────
class PipelineState(TypedDict):
    files: List[str]
    file_idx: int

    current_file: Optional[str]
    current_rows: List[Dict[str, Any]]
    previous_rows: List[Dict[str, Any]]

    new_rows: List[Dict[str, Any]]
    changed_rows: List[Dict[str, Any]]
    unchanged_rows: List[Dict[str, Any]]

    categorized: Dict[str, Dict[str, List[Dict[str, Any]]]]
    read_engine: str

    # 분류 기준 (B안 — categories.json 로드 결과)
    criteria: Dict[str, Any]

    # 점검 결과 (옵션 c)
    aggregate_report: Dict[str, Any]
    review_report: Dict[str, Any]


# ──────────────────────────────────────────────────────────────────────────────
# 2. 분류 기준 로더
# ──────────────────────────────────────────────────────────────────────────────
def load_criteria() -> Dict[str, Any]:
    if not CATEGORIES_JSON.exists():
        raise FileNotFoundError(f"categories.json 없음: {CATEGORIES_JSON}")
    with open(CATEGORIES_JSON, "r", encoding="utf-8") as f:
        return json.load(f)


# ──────────────────────────────────────────────────────────────────────────────
# 3. 엑셀 읽기 엔진 (xlwings 1차 + pandas 폴백)
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
        rows = read_excel_xlwings(path)
        return rows, "xlwings"
    except Exception as e:
        print(f"[read] xlwings 실패 ({type(e).__name__}: {e}) → pandas 폴백")
        rows = read_excel_pandas(path)
        return rows, "pandas"


def project_columns(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    return [{COL_STEP: r.get(COL_STEP), COL_FILE: r.get(COL_FILE)} for r in rows]


# ──────────────────────────────────────────────────────────────────────────────
# 4. OpenAI 클라이언트
# ──────────────────────────────────────────────────────────────────────────────
client = OpenAI()
MODEL = os.getenv("OPENAI_MODEL", "gpt-4o-mini")


def llm_categorize_with_criteria(
    rows: List[Dict[str, Any]],
    criteria: Dict[str, Any],
) -> Dict[str, List[Dict[str, Any]]]:
    """B안: categories.json 전체를 system prompt에 주입 → LLM이 기준 참조하여 분류."""
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
        "응답 형식 (JSON 객체 하나):\n"
        '{"카테고리명": [index 정수, ...], ...}\n\n'
        "[분류 기준 JSON]\n"
        + json.dumps(criteria, ensure_ascii=False, indent=2)
    )

    resp = client.chat.completions.create(
        model=MODEL,
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": sys_prompt},
            {"role": "user", "content": json.dumps(payload, ensure_ascii=False)},
        ],
        temperature=0.1,
    )
    raw = resp.choices[0].message.content or "{}"
    mapping: Dict[str, List[int]] = json.loads(raw)

    result: Dict[str, List[Dict[str, Any]]] = {}
    for cat, idx_list in mapping.items():
        result[cat] = [rows[i] for i in idx_list if isinstance(i, int) and 0 <= i < len(rows)]
    return result


def llm_review_classification(
    categorized: Dict[str, Dict[str, List[Dict[str, Any]]]],
    criteria: Dict[str, Any],
) -> Dict[str, Any]:
    """전체 분류 결과를 LLM에 보여주고 오분류 의심 항목 검출."""
    # 평탄화: [(파일, 카테고리, 진행단계, 파일명), ...]
    flat = []
    idx = 0
    index_map = {}
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

    resp = client.chat.completions.create(
        model=MODEL,
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": sys_prompt},
            {"role": "user", "content": json.dumps(flat, ensure_ascii=False)},
        ],
        temperature=0.1,
    )
    raw = resp.choices[0].message.content or "{}"
    parsed = json.loads(raw)

    # 플래그에 원본 컨텍스트 보강
    enriched_flags = []
    for flag in parsed.get("flags", []):
        i = flag.get("index")
        if isinstance(i, int) and i in index_map:
            fname, cat, item = index_map[i]
            enriched_flags.append({
                **flag,
                "source_file": fname,
                "item": item,
            })
    parsed["flags"] = enriched_flags
    return parsed


# ──────────────────────────────────────────────────────────────────────────────
# 5. 노드
# ──────────────────────────────────────────────────────────────────────────────
def node_init(state: PipelineState) -> PipelineState:
    state["criteria"] = load_criteria()
    print(f"[init] 기준 로드: categories={[c['name'] for c in state['criteria']['categories']]}")
    return state


def node_load_file(state: PipelineState) -> PipelineState:
    state["current_file"] = state["files"][state["file_idx"]]
    return state


def node_read_rows(state: PipelineState) -> PipelineState:
    raw, engine = read_excel(state["current_file"])
    state["current_rows"] = project_columns(raw)
    state["read_engine"] = engine
    print(f"[read] {Path(state['current_file']).name}: {len(state['current_rows'])} rows ({engine})")
    return state


def _row_key(row: Dict[str, Any]) -> str:
    return json.dumps(row, sort_keys=True, ensure_ascii=False, default=str)


def node_compare(state: PipelineState) -> PipelineState:
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


def node_categorize(state: PipelineState) -> PipelineState:
    file_name = Path(state["current_file"]).name
    bucket: Dict[str, List[Dict[str, Any]]] = {}

    targets = state["new_rows"] + [c["after"] for c in state["changed_rows"]]
    if targets:
        bucket = llm_categorize_with_criteria(targets, state["criteria"])

    if state["unchanged_rows"]:
        bucket.setdefault("동일(변경없음)", []).extend(state["unchanged_rows"])
    if state["changed_rows"]:
        bucket.setdefault("_meta_변경이력", []).extend(state["changed_rows"])

    state["categorized"][file_name] = bucket
    print(f"[categorize] {file_name}: {sorted(bucket.keys())}")
    return state


def node_accumulate(state: PipelineState) -> PipelineState:
    state["previous_rows"] = state["previous_rows"] + state["current_rows"]
    state["file_idx"] += 1
    state["current_file"] = None
    state["current_rows"] = []
    state["new_rows"] = []
    state["changed_rows"] = []
    state["unchanged_rows"] = []
    return state


def node_aggregate(state: PipelineState) -> PipelineState:
    """집계 리포트: 카테고리별 건수, 총합, 누락(default 카테고리 비율) 등."""
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
    unknown_categories = [c for c in by_category if c not in allowed and c != default_cat and c != "동일(변경없음)"]
    default_share = by_category.get(default_cat, 0) / total if total else 0.0

    report = {
        "total_items": total,
        "meta_items": meta_count,
        "by_category": dict(sorted(by_category.items(), key=lambda x: -x[1])),
        "by_file": by_file_category,
        "default_category_share": round(default_share, 3),
        "unknown_categories": unknown_categories,
        "warnings": [],
    }
    if unknown_categories:
        report["warnings"].append(f"기준 JSON에 없는 카테고리 발견: {unknown_categories}")
    if default_share > 0.3:
        report["warnings"].append(f"default 카테고리 비율 과다 ({default_share:.1%}) — 기준 보강 권장")

    state["aggregate_report"] = report
    print(f"[aggregate] total={total} categories={list(report['by_category'].keys())}")
    if report["warnings"]:
        for w in report["warnings"]:
            print(f"[aggregate][WARN] {w}")
    return state


def node_review(state: PipelineState) -> PipelineState:
    """LLM 검증: 전체 분류 결과 점검 → 오분류 의심 플래그."""
    if not state["criteria"].get("review", {}).get("enabled", True):
        state["review_report"] = {"skipped": True}
        return state
    report = llm_review_classification(state["categorized"], state["criteria"])
    state["review_report"] = report
    flags = report.get("flags", [])
    print(f"[review] flags={len(flags)} / summary={report.get('summary', '')[:80]}")
    return state


def node_finalize(state: PipelineState) -> PipelineState:
    return state


def route_after_accumulate(state: PipelineState) -> str:
    return "load_file" if state["file_idx"] < len(state["files"]) else "aggregate"


# ──────────────────────────────────────────────────────────────────────────────
# 6. 그래프 빌드
# ──────────────────────────────────────────────────────────────────────────────
def build_graph():
    g = StateGraph(PipelineState)
    g.add_node("init", node_init)
    g.add_node("load_file", node_load_file)
    g.add_node("read_rows", node_read_rows)
    g.add_node("compare", node_compare)
    g.add_node("categorize", node_categorize)
    g.add_node("accumulate", node_accumulate)
    g.add_node("aggregate", node_aggregate)
    g.add_node("review", node_review)
    g.add_node("finalize", node_finalize)

    g.set_entry_point("init")
    g.add_edge("init", "load_file")
    g.add_edge("load_file", "read_rows")
    g.add_edge("read_rows", "compare")
    g.add_edge("compare", "categorize")
    g.add_edge("categorize", "accumulate")
    g.add_conditional_edges(
        "accumulate",
        route_after_accumulate,
        {"load_file": "load_file", "aggregate": "aggregate"},
    )
    g.add_edge("aggregate", "review")
    g.add_edge("review", "finalize")
    g.add_edge("finalize", END)
    return g.compile()


# ──────────────────────────────────────────────────────────────────────────────
# 7. 실행 진입점
# ──────────────────────────────────────────────────────────────────────────────
def run(file_paths: List[str]) -> Dict[str, Any]:
    app = build_graph()
    init: PipelineState = {
        "files": file_paths,
        "file_idx": 0,
        "current_file": None,
        "current_rows": [],
        "previous_rows": [],
        "new_rows": [],
        "changed_rows": [],
        "unchanged_rows": [],
        "categorized": {},
        "read_engine": "",
        "criteria": {},
        "aggregate_report": {},
        "review_report": {},
    }
    final = app.invoke(init)
    return {
        "categorized": final["categorized"],
        "aggregate": final["aggregate_report"],
        "review": final["review_report"],
    }


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("사용법: python categorizer.py file1.xlsx file2.xlsx ...")
        sys.exit(1)
    result = run(sys.argv[1:])
    print(json.dumps(result, ensure_ascii=False, indent=2, default=str))
