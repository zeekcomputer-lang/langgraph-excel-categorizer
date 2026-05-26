# LangGraph Excel Categorizer

엑셀 파일들을 순차 읽어 이전 누적 데이터와 비교 후 LLM(OpenAI SDK)으로 카테고라이즈하는 LangGraph 파이프라인.

## 파이프라인 구조
```
load_file → read_rows → compare → categorize(LLM) → accumulate
     ↑                                                  │
     └──────── 다음 파일 있으면 loop ────────────────────┘
                                                        │
                                                  finalize → END
```

## 요구사항
- Python 3.10+
- OpenAI API Key

---

## Windows 실행 (가상환경 자동 생성)

### 1) 1회성 셋업
```cmd
windows\setup.bat
```
- `.venv\` 생성 + `requirements.txt` 설치 + import 검증

### 2) API Key 등록
```cmd
set OPENAI_API_KEY=sk-...
```
또는 PowerShell:
```powershell
$env:OPENAI_API_KEY = "sk-..."
```

### 3) 실행
```cmd
windows\start.bat data1.xlsx data2.xlsx data3.xlsx
```

---

## Linux / macOS 실행
```bash
python3 tools/setup.py
export OPENAI_API_KEY=sk-...
.venv/bin/python tools/launch.py data1.xlsx data2.xlsx
```

또는 직접:
```bash
pip install -r requirements.txt
python categorizer.py data1.xlsx data2.xlsx
```

---

## 출력 형식
```json
{
  "data1.xlsx": {
    "중요": [ {...} ],
    "일반": [ {...} ],
    "동일(변경없음)": [ ... ],
    "_meta_변경이력": [ {"before": {...}, "after": {...}} ]
  },
  "data2.xlsx": { ... }
}
```

## 비교 로직
- 행 전체 직렬화 동일 → `unchanged`
- `id`/`code`/`key` 컬럼 동일 + 내용 다름 → `changed` (before/after 보존)
- 그 외 → `new` (LLM 카테고라이즈 대상)

## 디렉터리
```
langgraph-excel-categorizer/
├── categorizer.py          # LangGraph 파이프라인 본체
├── requirements.txt
├── README.md
├── .gitignore
├── tools/
│   ├── setup.py            # 1회성 셋업 (venv + pip install)
│   └── launch.py           # 실행 진입점 (검증 + run)
└── windows/
    ├── setup.bat           # tools/setup.py wrapper
    └── start.bat           # tools/launch.py wrapper
```

## 환경변수
| 이름 | 필수 | 설명 |
|---|---|---|
| `OPENAI_API_KEY` | ✔ | OpenAI API Key |
| `OPENAI_MODEL` | ✖ | 기본 `gpt-4o-mini` |
