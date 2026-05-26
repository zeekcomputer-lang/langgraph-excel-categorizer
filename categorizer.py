"""
LangGraph 기반 엑셀 행 반복 읽기 + LLM 카테고라이즈 파이프라인

흐름:
  load_file → read_rows → compare → categorize(LLM) → accumulate → (next file?) loop → finalize

요구사항:
  pip install langgraph langchain-core openai pandas openpyxl pydantic

환경변수:
  OPENAI_API_KEY=sk-...
  (선택) OPENAI_MODEL=gpt-4o-mini
"""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any, Dict, List, Optional, TypedDict

import pandas as pd
from openai import OpenAI
from langgraph.graph import StateGraph, END


# ──────────────────────────────────────────────────────────────────────────────
# 1. State 정의
# ──────────────────────────────────────────────────────────────────────────────
class PipelineState(TypedDict):
    # 입력
    files: List[str]                       # 처리할 엑셀 파일 경로 목록
    file_idx: int                          # 현재 처리 중 인덱스

    # 진행 데이터
    current_file: Optional[str]
    current_rows: List[Dict[str, Any]]     # 이번 파일에서 읽은 행
    previous_rows: List[Dict[str, Any]]    # 직전까지 누적된 행 (비교 기준)

    # 비교 결과
    new_rows: List[Dict[str, Any]]         # 신규(이전엔 없던) 행
    changed_rows: List[Dict[str, Any]]     # 변경된 행
    unchanged_rows: List[Dict[str, Any]]   # 동일 행

    # 카테고라이즈 결과: {파일명: {카테고리: [행, ...]}}
    categorized: Dict[str, Dict[str, List[Dict[str, Any]]]]


# ──────────────────────────────────────────────────────────────────────────────
# 2. OpenAI 클라이언트
# ──────────────────────────────────────────────────────────────────────────────
client = OpenAI()  # OPENAI_API_KEY 환경변수 사용
MODEL = os.getenv("OPENAI_MODEL", "gpt-4o-mini")


def llm_categorize(rows: List[Dict[str, Any]], categories_hint: Optional[List[str]] = None) -> Dict[str, List[Dict[str, Any]]]:
    """LLM에게 행 묶음을 카테고리별로 분류 요청."""
    if not rows:
        return {}

    hint = categories_hint or ["중요", "일반", "보류", "기타"]
    sys_prompt = (
        "당신은 데이터 분류기입니다. 주어진 행 목록을 의미 단위로 카테고리에 배치하세요. "
        "응답은 반드시 JSON 객체 하나여야 합니다. "
        '형식: {"카테고리명": [원본 행 인덱스(0-based 정수), ...], ...} '
        f"가능하면 다음 카테고리를 우선 사용: {hint}. 필요시 새 카테고리를 만들어도 됩니다."
    )
    user_payload = {"rows": rows}

    resp = client.chat.completions.create(
        model=MODEL,
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": sys_prompt},
            {"role": "user", "content": json.dumps(user_payload, ensure_ascii=False)},
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
# 3. 노드 함수
# ──────────────────────────────────────────────────────────────────────────────
def node_load_file(state: PipelineState) -> PipelineState:
    """다음 파일을 큐에서 꺼내 현재 파일로 설정."""
    idx = state["file_idx"]
    state["current_file"] = state["files"][idx]
    return state


def node_read_rows(state: PipelineState) -> PipelineState:
    """엑셀 파일에서 행을 dict 리스트로 읽는다."""
    path = state["current_file"]
    df = pd.read_excel(path)
    # NaN을 None으로 정리
    rows = df.where(pd.notnull(df), None).to_dict(orient="records")
    state["current_rows"] = rows
    return state


def _row_key(row: Dict[str, Any]) -> str:
    """행을 해시 가능한 키로 직렬화 (간단 비교용)."""
    return json.dumps(row, sort_keys=True, ensure_ascii=False, default=str)


def node_compare(state: PipelineState) -> PipelineState:
    """이전 누적 데이터와 비교하여 new / changed / unchanged 분리."""
    prev = state["previous_rows"]
    curr = state["current_rows"]

    prev_keys = {_row_key(r) for r in prev}

    # 키(id 컬럼) 후보 탐색: id / ID / 코드 / key 가 있으면 변경 감지에 활용
    id_col = None
    if curr:
        for c in ("id", "ID", "Id", "code", "코드", "key"):
            if c in curr[0]:
                id_col = c
                break

    prev_by_id = {r.get(id_col): r for r in prev} if id_col else {}

    new_rows, changed_rows, unchanged_rows = [], [], []
    for r in curr:
        rk = _row_key(r)
        if rk in prev_keys:
            unchanged_rows.append(r)
        elif id_col and r.get(id_col) in prev_by_id:
            changed_rows.append({"before": prev_by_id[r[id_col]], "after": r})
        else:
            new_rows.append(r)

    state["new_rows"] = new_rows
    state["changed_rows"] = changed_rows
    state["unchanged_rows"] = unchanged_rows
    return state


def node_categorize(state: PipelineState) -> PipelineState:
    """LLM으로 신규 + 변경 행을 카테고라이즈 (unchanged는 별도 묶음)."""
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
    return state


def node_accumulate(state: PipelineState) -> PipelineState:
    """현재 행들을 다음 비교용 누적 데이터에 추가."""
    state["previous_rows"] = state["previous_rows"] + state["current_rows"]
    state["file_idx"] += 1
    # 일회용 필드 초기화
    state["current_file"] = None
    state["current_rows"] = []
    state["new_rows"] = []
    state["changed_rows"] = []
    state["unchanged_rows"] = []
    return state


def node_finalize(state: PipelineState) -> PipelineState:
    """종료 노드. 별도 처리 없음 (반환만)."""
    return state


# ──────────────────────────────────────────────────────────────────────────────
# 4. 조건부 라우팅
# ──────────────────────────────────────────────────────────────────────────────
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
