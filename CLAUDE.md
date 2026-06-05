# 물리치료 집계 리포트 — Claude Code 지시서

## 프로젝트 개요
병원 물리치료실 처방 데이터를 자동 집계하여 Excel 리포트 + HTML 대시보드를 생성하는 자동화 시스템.

## 폴더 구조
```
pt_report/
├── data/
│   ├── weekly/          ← 주별 원본 데이터 (xlsx) 넣는 곳
│   │   └── _done/       ← 처리 완료 파일 자동 이동
│   └── monthly/         ← 월별 전체 데이터 (xlsx) 넣는 곳
│       └── _done/
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
├── build.bat        ← PyInstaller EXE 빌드
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
- **중복 제거 기준**: 처방의 + 등록번호 + 접수일자 + 구분(입원/외래)
- **입원/외래 분류**: 병동 OR 병실 중 하나라도 있으면 입원, 둘 다 없으면 외래
- **공휴일**: `workalendar` 라이브러리로 자동 계산 (없으면 고정 공휴일 fallback)
- **근로자의날(5/1)**: workalendar에 포함되지 않아 `get_holidays()`에서 별도 추가
- **엑셀 엔진**: `.xls` → `xlrd`, `.xlsx` → `openpyxl` (모든 읽기 함수에서 명시적 지정)

## 자동 실행 스케줄
| 구분 | 시각 | 태스크명 |
|------|------|---------|
| 주별 | 매주 월요일 12:00 | `PT_Weekly` |
| 월별 | 매월 1일 12:00 | `PT_Monthly` |

스케줄 등록은 `Register-ScheduledTask` PowerShell cmdlet 사용 (로케일 무관).

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
