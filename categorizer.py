"""
LangGraph 기반 엑셀 행 반복 읽기 + LLM 카테고라이즈 파이프라인

흐름:
  load_file → read_rows → compare → categorize(LLM) → accumulate → (next file?) loop → finalize

엑셀 읽기:
  1차: xlwings (Excel COM, DRM 호환) — Windows + Excel 설치 필요
  2차: pandas + openpyxl (폴백)

컬럼 (하드코딩, 사용자가 추후 직접 수정):
  - "진행단계"
  - "파일명"

요구사항:
  pip install -r requirements.txt
  Windows + Excel 설치 권장 (xlwings 사용)
  pip install -r requirements-windows.txt  # xlwings, pywin32

환경변수:
  OPENAI_API_KEY=sk-...
  (선택) OPENAI_MODEL=gpt-4o-mini
"""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any, Dict, List, Optional, TypedDict

from openai import OpenAI
from langgraph.graph import StateGraph, END


# ──────────────────────────────────────────────────────────────────────────────
# 0. 하드코딩 컬럼명 (사용자가 직접 수정)
# ──────────────────────────────────────────────────────────────────────────────
COL_STEP = "진행단계"   # 컬럼1
COL_FILE = "파일명"     # 컬럼2


# ──────────────────────────────────────────────────────────────────────────────
# 1. State 정의
# ──────────────────────────────────────────────────────────────────────────────
class PipelineState(TypedDict):
    files: List[str]                       # 처리할 엑셀 파일 경로 목록
    file_idx: int                          # 현재 처리 중 인덱스

    current_file: Optional[str]
    current_rows: List[Dict[str, Any]]     # 이번 파일에서 읽은 행
    previous_rows: List[Dict[str, Any]]    # 직전까지 누적된 행

    new_rows: List[Dict[str, Any]]
    changed_rows: List[Dict[str, Any]]
    unchanged_rows: List[Dict[str, Any]]

    # 파일명 기준 카테고라이즈 결과: {엑셀파일명: {카테고리: [행, ...]}}
    categorized: Dict[str, Dict[str, List[Dict[str, Any]]]]

    read_engine: str                       # 사용된 엔진 ("xlwings" | "pandas")


# ──────────────────────────────────────────────────────────────────────────────
# 2. 엑셀 읽기 엔진 (xlwings 1차 + pandas 폴백)
# ──────────────────────────────────────────────────────────────────────────────
def read_excel_xlwings(path: str) -> List[Dict[str, Any]]:
    """xlwings로 엑셀 읽기. Excel COM 필요 (Windows + Excel 설치)."""
    import xlwings as xw

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        wb = app.books.open(str(Path(path).resolve()), read_only=True, update_links=False)
        try:
            sht = wb.sheets[0]
            used = sht.used_range
            data = used.value  # 2D list (헤더 포함)
            if not data:
                return []
            if not isinstance(data[0], list):
                # 단일 행/열 케이스
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
    """pandas + openpyxl 폴백 읽기."""
    import pandas as pd

    df = pd.read_excel(path, engine="openpyxl")
    df = df.where(pd.notnull(df), None)
    return df.to_dict(orient="records")


def read_excel(path: str) -> tuple[List[Dict[str, Any]], str]:
    """xlwings 1차 → 실패시 pandas. (rows, engine_used) 반환."""
    try:
        rows = read_excel_xlwings(path)
        return rows, "xlwings"
    except Exception as e:
        print(f"[read] xlwings 실패 ({type(e).__name__}: {e}) → pandas 폴백")
        rows = read_excel_pandas(path)
        return rows, "pandas"


def project_columns(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """하드코딩 두 컬럼만 추출. 누락 컬럼은 None."""
    result = []
    for r in rows:
        result.append({
            COL_STEP: r.get(COL_STEP),
            COL_FILE: r.get(COL_FILE),
        })
    return result


# ──────────────────────────────────────────────────────────────────────────────
# 3. OpenAI 클라이언트
# ──────────────────────────────────────────────────────────────────────────────
client = OpenAI()
MODEL = os.getenv("OPENAI_MODEL", "gpt-4o-mini")


def llm_categorize(rows: List[Dict[str, Any]], categories_hint: Optional[List[str]] = None) -> Dict[str, List[Dict[str, Any]]]:
    """진행단계+파일명 두 컬럼을 합쳐 하나의 data로 카테고라이즈."""
    if not rows:
        return {}

    payload = [
        {"index": i, "data": f"[{r.get(COL_STEP)}] {r.get(COL_FILE)}"}
        for i, r in enumerate(rows)
    ]

    hint = categories_hint or ["대기", "진행중", "완료", "오류", "기타"]
    sys_prompt = (
        "당신은 데이터 분류기입니다. 각 항목의 data 문자열(진행단계+파일명 결합)을 보고 카테고리에 배치하세요. "
        "응답은 반드시 JSON 객체 하나여야 합니다. "
        '형식: {"카테고리명": [index 정수, ...], ...} '
        f"가능하면 다음 카테고리를 우선 사용: {hint}. 필요시 새 카테고리를 만들어도 됩니다."
    )

    resp = client.chat.completions.create(
        model=MODEL,
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": sys_prompt},
            {"role": "user", "content": json.dumps(payload, ensure_ascii=False)},
        ],
        temperature=0.2,
    )
    raw = resp.choices[0].message.content or "{}"
    mapping: Dict[str, List[int]] = json.loads(raw)

    result: Dict[str, List[Dict[str, Any]]] = {}
    for cat, idx_list in mapping.items():
        result[cat] = [rows[i] for i in idx_list if isinstance(i, int) and 0 <= i < len(rows)]
    return result


# ──────────────────────────────────────────────────────────────────────────────
# 4. 노드 함수
# ──────────────────────────────────────────────────────────────────────────────
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
    """이전 누적과 비교. 키 컬럼 = 파일명(COL_FILE)."""
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
        bucket = llm_categorize(targets)

    if state["unchanged_rows"]:
        bucket.setdefault("동일(변경없음)", []).extend(state["unchanged_rows"])
    if state["changed_rows"]:
        bucket.setdefault("_meta_변경이력", []).extend(state["changed_rows"])

    state["categorized"][file_name] = bucket
    print(f"[categorize] {file_name}: {len(bucket)} categories")
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


def node_finalize(state: PipelineState) -> PipelineState:
    return state


def route_after_accumulate(state: PipelineState) -> str:
    return "load_file" if state["file_idx"] < len(state["files"]) else "finalize"


# ──────────────────────────────────────────────────────────────────────────────
# 5. 그래프 빌드
# ──────────────────────────────────────────────────────────────────────────────
def build_graph():
    g = StateGraph(PipelineState)
    g.add_node("load_file", node_load_file)
    g.add_node("read_rows", node_read_rows)
    g.add_node("compare", node_compare)
    g.add_node("categorize", node_categorize)
    g.add_node("accumulate", node_accumulate)
    g.add_node("finalize", node_finalize)

    g.set_entry_point("load_file")
    g.add_edge("load_file", "read_rows")
    g.add_edge("read_rows", "compare")
    g.add_edge("compare", "categorize")
    g.add_edge("categorize", "accumulate")
    g.add_conditional_edges(
        "accumulate",
        route_after_accumulate,
        {"load_file": "load_file", "finalize": "finalize"},
    )
    g.add_edge("finalize", END)
    return g.compile()


# ──────────────────────────────────────────────────────────────────────────────
# 6. 실행 진입점
# ──────────────────────────────────────────────────────────────────────────────
def run(file_paths: List[str]) -> Dict[str, Dict[str, List[Dict[str, Any]]]]:
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
    }
    final = app.invoke(init)
    return final["categorized"]


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("사용법: python categorizer.py file1.xlsx file2.xlsx ...")
        sys.exit(1)
    result = run(sys.argv[1:])
    print(json.dumps(result, ensure_ascii=False, indent=2, default=str))
