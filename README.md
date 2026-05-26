# LangGraph Excel Categorizer

엑셀 파일을 순차 읽어 이전 누적 데이터와 비교 후 LLM(OpenAI SDK)으로 카테고라이즈하는 LangGraph 파이프라인.

## 입력 스펙 (하드코딩)
- 엑셀 시트의 두 컬럼만 사용:
  - **`진행단계`** (컬럼1)
  - **`파일명`** (컬럼2)
- 두 컬럼은 결합되어 **1개 data** 로 LLM에 전달
- 컬럼명은 `categorizer.py` 상단 `COL_STEP`, `COL_FILE` 상수에서 직접 수정

## 엑셀 읽기 엔진 (2중 구동)
1. **xlwings** (1차) — Excel COM, DRM/매크로 호환. Windows + Excel 설치 필요
2. **pandas + openpyxl** (폴백) — xlwings 실패 시 자동 전환

## 파이프라인
```
load_file → read_rows(xlwings|pandas) → compare → categorize(LLM) → accumulate
     ↑                                                                 │
     └─────────────── 다음 파일 있으면 loop ─────────────────────────────┘
                                                                       │
                                                                 finalize → END
```

---

## Windows 실행

### 1) 1회성 셋업 (자동으로 `requirements-windows.txt` 설치 → xlwings 포함)
```cmd
windows\setup.bat
```

### 2) API Key
```cmd
set OPENAI_API_KEY=sk-...
```

### 3) 실행
```cmd
windows\start.bat data1.xlsx data2.xlsx
```

> ⚠️ xlwings는 Microsoft Excel 설치가 필요합니다. Excel이 없으면 자동 pandas 폴백.

---

## Linux / macOS 실행
```bash
python3 tools/setup.py            # requirements.txt 만 (xlwings 제외)
export OPENAI_API_KEY=sk-...
.venv/bin/python tools/launch.py data1.xlsx data2.xlsx
```
> Linux/macOS는 xlwings 불가 → pandas 단독 사용

---

## 출력 형식
```json
{
  "data1.xlsx": {
    "진행중": [ {"진행단계": "...", "파일명": "..."} ],
    "완료":   [ {...} ],
    "동일(변경없음)": [ ... ],
    "_meta_변경이력": [ {"before": {...}, "after": {...}} ]
  },
  "data2.xlsx": { ... }
}
```

## 비교 로직
- 두 컬럼 값 완전 동일 → `unchanged`
- `파일명` 동일 + `진행단계` 다름 → `changed` (before/after 보존)
- 그 외 → `new` (LLM 카테고라이즈 대상)

## 디렉터리
```
langgraph-excel-categorizer/
├── categorizer.py                # LangGraph 본체 + 2중 엑셀 엔진
├── requirements.txt              # 공통
├── requirements-windows.txt      # + xlwings, pywin32
├── README.md
├── .gitignore
├── tools/
│   ├── setup.py
│   └── launch.py
└── windows/
    ├── setup.bat
    └── start.bat
```

## 컬럼명 변경 방법
`categorizer.py` 상단:
```python
COL_STEP = "진행단계"   # ← 원하는 컬럼명으로
COL_FILE = "파일명"     # ← 원하는 컬럼명으로
```

## 환경변수
| 이름 | 필수 | 설명 |
|---|---|---|
| `OPENAI_API_KEY` | ✔ | OpenAI API Key |
| `OPENAI_MODEL` | ✖ | 기본 `gpt-4o-mini` |
