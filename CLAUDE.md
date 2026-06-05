# 물리치료 집계 리포트 — Claude Code 지시서

## 프로젝트 개요
병원 물리치료실 처방 데이터를 자동 집계하여 Excel 리포트 + HTML 대시보드를 생성하는 자동화 시스템.

## 폴더 구조
```
pt_report/
├── data/
│   ├── weekly/          ← 주별 원본 데이터 넣는 곳 (.xls/.xlsx/.XLS/.XLSX 모두 인식)
│   │   └── _done/       ← 처리 완료 파일 보관 + 누적 집계 시 자동 포함
│   │       └── YYYY-MM/ ← 월별 하위 폴더 자동 생성
│   └── monthly/         ← 월별 전체 데이터 넣는 곳
│       └── _done/
│           └── YYYY-MM/
├── output/
│   └── YYYY-MM/
│       ├── reports/     ← 생성된 Excel 리포트
│       │   └── _archive/← 이전 결과물 보관
│       └── dashboard/   ← 생성된 HTML 대시보드
│           └── _archive/
├── scripts/
│   ├── process.py       ← 핵심 데이터 처리 로직
│   ├── excel_report.py  ← Excel 생성
│   ├── html_dashboard.py← HTML 생성
│   └── validator.py     ← 주별/월별 정합성 검증
├── app.py           ← GUI (tkinter, Windows 전용)
├── run.py           ← 메인 실행 진입점 (CLI / 스케줄러)
├── run_silent.vbs   ← Windows 작업 스케줄러용 무창 실행
├── setup.bat        ← 설치 (Python, 패키지, 스케줄 등록)
├── build.bat        ← PyInstaller EXE 빌드 (의존성 설치 포함)
├── uninstall.bat    ← 스케줄 삭제
├── requirements.txt ← 런타임 의존성 (pyinstaller 제외)
├── CLAUDE.md        ← 이 파일
└── run.log          ← 실행 로그 (자동 생성)
```

## 입력 데이터 컬럼 (필수)
| 컬럼명 | 설명 |
|--------|------|
| 환자성명 | 환자 이름 |
| 등록번호 | 환자 등록번호 (중복 제거 기준) |
| 처방의 | 담당 의사 |
| 접수일자 | 진료 날짜 |
| 병동 | 입원 병동 (없으면 외래) |
| 병실 | 입원 병실 (없으면 외래) |

## 핵심 규칙

### 데이터 처리
- **중복 제거 기준**: 처방의 + 등록번호 + 접수일자 + 구분(입원/외래)
- **입원/외래 분류**: 병동 OR 병실 중 하나라도 있으면 입원, 둘 다 없으면 외래
- **공휴일**: `workalendar` 라이브러리로 자동 계산 (없으면 고정 공휴일 fallback)
- **근로자의날(5/1)**: workalendar에 포함되지 않아 `get_holidays()`에서 별도 추가
- **엑셀 엔진**: `.xls` → `xlrd`, `.xlsx` → `openpyxl` (모든 읽기 함수에서 명시적 지정)
- **파일 탐색**: `_glob_excel()`은 `iterdir() + suffix.lower()` — `.XLS`/`.XLSX` 대소문자 무관 인식

### 주별 집계 누적 로직 (`run.py: _collect_weekly_files`)
- **탐색 범위**: `data/weekly/` + `data/weekly/_done/YYYY-MM/` 동시 탐색
- **우선순위**: 같은 파일명이 `weekly/`와 `_done/` 양쪽에 있으면 **`weekly/` 파일 무조건 우선** (사용자가 방금 넣은 파일 = 최신 데이터)
- **중복 제거**: `_HHMMSS` 타임스탬프 접미사 파일(`파일명_235542.XLS`)은 베이스 파일과 동일 취급, 1개만 사용
- **파일 이동**: 처리 성공 후 `data/weekly/`의 해당 월 파일 **전체** `_done/YYYY-MM/`으로 이동
- **중복 이동 방지**: 같은 이름·크기 파일이 이미 `_done/`에 있으면 소스만 삭제 (타임스탬프 복사본 생성 안 함)

### 월별 집계 동작 (`run.py: _run_monthly`)
- **`data/monthly/` 파일 있음**: 해당 파일로 집계 → `_collect_weekly_files`로 주별 전체 합산과 정합성 비교
- **`data/monthly/` 파일 없음**: `weekly/` + `_done/YYYY-MM/` 누적 주별 데이터로 자동 대체 집계 (파일 이동 없음)
- **정합성 불일치 시**: `log.warning` 레벨 → GUI에서 노란색(warn 태그)으로 표시 + "원본 확인" 안내
- 주별 파일도 없으면: "처리할 파일 없음" 로그 후 종료

### 파일 및 출력
- **출력 타임스탬프**: `%Y%m%d_%H%M%S` (초 단위 — 같은 분에 주별·월별 실행해도 파일명 충돌 없음)
- **Excel 시트명**: 처방의별 탭은 `doc[:31]` truncate — Excel 31자 제한 대응
- **HTML CDN 의존성**: Chart.js · Google Fonts를 외부 CDN에서 로드. 인터넷 미연결 시 차트 미표시·폰트 fallback

## 자동 실행 스케줄
| 구분 | 시각 | 태스크명 |
|------|------|---------|
| 주별 | 매주 월요일 12:00 | `PT_Weekly` |
| 월별 | 매월 1일 12:00 | `PT_Monthly` |

스케줄 등록은 `Register-ScheduledTask` PowerShell cmdlet 사용 (로케일 무관).

**스케줄 자동 등록 동작 (app.py)**
- 앱 시작 0.9초 후 백그라운드 스레드에서 `_do_register()` 실행 (메인 스레드 블로킹 없음)
- PS1 스크립트: 이미 등록 → `exit 2`(조용히 종료), 신규 등록 → `exit 0`
- `exit 0`일 때만 "✓ 스케줄 등록 완료" 로그 출력 — **첫 실행 시에만 표시**
- `subprocess.CREATE_NO_WINDOW` 플래그 — 콘솔 창 완전 숨김
- 로그 메시지는 `self._q.put()`으로 큐 경유 (스레드 안전)

**실행 흐름 (배포 exe 기준)**
```
Windows 작업 스케줄러
  → run_silent.vbs
  → ptreport.exe --auto   ← GUI 없이 run_auto() 실행 후 종료
  → run.log 에 기록
```
`--auto` 없이 실행하면 GUI 모드로 시작.

## 실행 방법
```bash
# 1. 패키지 설치 (최초 1회, setup.bat 으로 대체 가능)
pip install pandas openpyxl xlrd apscheduler workalendar

# 2. 파일 직접 처리
python run.py data/weekly/파일명.xlsx

# 3. 강제 전체 처리 (중복 무시)
python run.py --force

# 4. GUI 버튼용 직접 호출
python run.py --weekly
python run.py --monthly

# 5. 정합성 검증만
python run.py --validate

# 6. APScheduler 데몬 모드 (창 켜놔야 함)
python run.py --schedule
```

## 출력 결과
- **Excel**: `output/YYYY-MM/reports/집계리포트_{기간}_{타임스탬프}.xlsx`
  - 시트: 월별 요약 / 주별 집계 / 일별 상세 / 처방의별 탭 / (정합성 검증)
- **HTML**: `output/YYYY-MM/dashboard/대시보드_{기간}_{타임스탬프}.html`
  - 월별 / 주별(전체+처방의별) / 일별(전체+처방의별) 탭
  - 다크/라이트 모드, 반응형

## 작업 요청 시 참고
- 로직 변경 → `scripts/process.py`
- Excel 스타일/구조 변경 → `scripts/excel_report.py`
- HTML 대시보드 변경 → `scripts/html_dashboard.py`
- 정합성 검증 변경 → `scripts/validator.py`
- GUI 변경 → `app.py`
- 스케줄/실행 흐름 변경 → `run.py`
- 새 컬럼/시트 추가 시 세 파일(`process.py`, `excel_report.py`, `html_dashboard.py`) 모두 확인 필요

## 인코딩 정책
- Python 소스: UTF-8
- `.bat` 파일: ASCII(영문)만 사용 — cmd.exe CP949 환경과 충돌 방지
- `run.py` 실행 시 `sys.stdout/stderr.reconfigure(encoding='utf-8')` 선언
- PS1 임시 파일: `utf-8-sig`(BOM) — PowerShell 5 UTF-8 인식 보장

## PyInstaller 빌드 참고
- `build.bat` 실행 시 `ERROR: Hidden import 'process' not found` 등 4개 오류 → `--add-data scripts;scripts` + `app.py`의 `sys.path.insert(BUNDLE/'scripts')` 조합으로 런타임 정상 동작, 무시해도 됨
- `WARNING: Hidden import "jinja2" not found!` → jinja2 미사용, 무시
- `WARNING: Failed to collect submodules for 'workalendar.tests'` → 테스트 모듈만 누락, 기능 영향 없음
- `build.bat` [1/4] 단계에서 의존성 패키지를 먼저 설치한 뒤 빌드 — `xlrd` 미설치로 XLS 처리 불가 문제 방지
