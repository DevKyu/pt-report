"""
process.py — 물리치료 집계 데이터 처리 핵심 로직
Python 3.8+ 호환
"""
from __future__ import annotations
import pandas as pd
import datetime
import json
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

# ── 공휴일 ───────────────────────────────────────────────────────────────
_holiday_cache: Dict[int, set] = {}

def get_holidays(year: int) -> set:
    """한국 공휴일 (workalendar 사용, 캐시, 설/추석 포함)"""
    if year in _holiday_cache:
        return _holiday_cache[year]

    holidays: set = set()
    try:
        from workalendar.asia import SouthKorea
        for d, _ in SouthKorea().holidays(year):
            holidays.add(d)
    except Exception:
        for m, d in [(1,1),(3,1),(5,5),(6,6),(8,15),(10,3),(10,9),(12,25)]:
            try:
                holidays.add(datetime.date(year, m, d))
            except ValueError:
                pass

    holidays.add(datetime.date(year, 5, 1))  # 근로자의날
    _holiday_cache[year] = holidays
    return holidays

def is_red_day(d: datetime.date) -> bool:
    return d.weekday() == 5 or d in get_holidays(d.year)

def get_dayname(d: datetime.date) -> str:
    return ['월','화','수','목','금','토','일'][d.weekday()]

def get_week_label(d: datetime.date) -> str:
    first_day    = d.replace(day=1)
    first_monday = first_day - datetime.timedelta(days=first_day.weekday())
    return f"{d.month}월 {(d - first_monday).days // 7 + 1}주차"

# ── 분류 ─────────────────────────────────────────────────────────────────
def _classify_inout(row: pd.Series) -> str:
    for col in ('병동', '병실'):
        val = str(row.get(col, '')).strip()
        if val and val.lower() not in ('nan', 'none', ''):
            return '입원'
    return '외래'

def _normalize_reg(val: Any) -> str:
    digits = ''.join(c for c in str(val) if c.isdigit())
    return digits.zfill(8) if digits else '00000000'

def _normalize_doctor(val: Any) -> str:
    return ' '.join(str(val).split())

# ── 연월 캐시 ─────────────────────────────────────────────────────────────
_ym_cache: Dict[Path, Optional[Tuple[int, int]]] = {}

def get_file_year_month(filepath: Path) -> Optional[Tuple[int, int]]:
    """파일의 데이터 연월 반환 (캐시 적용)"""
    fp = filepath.resolve()
    if fp in _ym_cache:
        return _ym_cache[fp]

    result: Optional[Tuple[int, int]] = None

    m = re.search(r'(20\d{2})[.\-_](\d{1,2})', filepath.stem)
    if m:
        yr, mo = int(m.group(1)), int(m.group(2))
        if 1 <= mo <= 12:
            result = (yr, mo)

    if result is None:
        try:
            engine = 'xlrd' if fp.suffix.lower() == '.xls' else 'openpyxl'
            df = pd.read_excel(fp, sheet_name=0, usecols=['접수일자'], nrows=100, engine=engine)
            df['접수일자'] = pd.to_datetime(df['접수일자'], errors='coerce').dt.date
            valid = df['접수일자'].dropna()
            if not valid.empty:
                d = valid.iloc[0]
                result = (d.year, d.month)
        except Exception:
            pass

    _ym_cache[fp] = result
    return result

def clear_ym_cache() -> None:
    _ym_cache.clear()

# ── 처방의 이름 이슈 감지 ─────────────────────────────────────────────────
def detect_doctor_name_issues(df: pd.DataFrame) -> List[str]:
    """
    처방의명 잠재적 오류 감지 후 경고 목록 반환
    - 앞뒤 공백, 중간 중복 공백 → 정규화로 해결
    - 유사 이름 쌍 감지 (편집거리 1)
    """
    warnings: List[str] = []
    doctors = df['처방의'].dropna().unique().tolist()

    for doc in doctors:
        normalized = _normalize_doctor(doc)
        if normalized == '':
            warnings.append(f"처방의명이 공백뿐인 행 발견 → 집계에서 제외됨 (원본: '{doc}')")
        elif doc != normalized:
            warnings.append(f"처방의명 공백 이슈: '{doc}' → '{normalized}' 로 처리됨")

    normalized_docs = [_normalize_doctor(d) for d in doctors]
    unique_docs = list(set(normalized_docs))
    for i in range(len(unique_docs)):
        for j in range(i+1, len(unique_docs)):
            if _edit_distance(unique_docs[i], unique_docs[j]) == 1:
                warnings.append(
                    f"유사 처방의명 발견 (오탈자 확인 필요): "
                    f"'{unique_docs[i]}' vs '{unique_docs[j]}'"
                )
    return warnings

def _edit_distance(a: str, b: str) -> int:
    """레벤슈타인 거리 (짧은 이름에만 적용)"""
    if abs(len(a) - len(b)) > 1:
        return 99
    if len(a) > 10 or len(b) > 10:
        return 99
    m, n = len(a), len(b)
    dp = list(range(n + 1))
    for i in range(1, m + 1):
        prev = dp[:]
        dp[0] = i
        for j in range(1, n + 1):
            dp[j] = prev[j-1] if a[i-1] == b[j-1] else \
                    1 + min(prev[j], dp[j-1], prev[j-1])
    return dp[n]

# ── 처방의 표시 순서 / 별칭 설정 ────────────────────────────────────────────
DEFAULT_DOCTOR_ORDER = [
    '김기택', '신재흥', '조명국', '이인수', '김범석',
    '강영훈', '홍태민', '김용찬', '서광욱',
]

def _project_root() -> Path:
    """exe(번들)/스크립트 실행 모두에서 실제 설치 경로(루트)를 반환"""
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent

def get_doctor_config_path() -> Path:
    return _project_root() / 'doctor_settings.json'

def load_doctor_config() -> List[Dict[str, str]]:
    """
    처방의 표시 순서/별칭 설정 로드 → [{'name': 원본명, 'alias': 별칭}, ...]
    파일이 없으면 DEFAULT_DOCTOR_ORDER 기준으로 새로 생성
    """
    path = get_doctor_config_path()
    if path.exists():
        try:
            raw = json.loads(path.read_text(encoding='utf-8'))
            return [
                {'name': item['name'], 'alias': item.get('alias', '')}
                for item in raw if item.get('name')
            ]
        except Exception:
            pass
    config = [{'name': n, 'alias': ''} for n in DEFAULT_DOCTOR_ORDER]
    save_doctor_config(config)
    return config

def save_doctor_config(config: List[Dict[str, str]]) -> None:
    path = get_doctor_config_path()
    path.write_text(
        json.dumps(config, ensure_ascii=False, indent=2), encoding='utf-8'
    )

def _apply_doctor_order(dedup: pd.DataFrame) -> pd.DataFrame:
    """
    설정 파일 기준으로 처방의 별칭을 적용하고 표시 순서(Categorical)를 부여한다.
    설정에 없는 새 이름은 가나다순으로 뒤에 추가하고 설정 파일에도 반영한다.
    """
    dedup = dedup.copy()
    config = load_doctor_config()
    alias_map = {e['name']: e['alias'] for e in config if e.get('alias')}
    if alias_map:
        dedup['처방의'] = dedup['처방의'].replace(alias_map)

    # 처방의 값이 빈 문자열인 행은 집계에서 제외 (공백 셀 등 — 정합성 검증에서 차이로 드러남)
    dedup = dedup[dedup['처방의'] != '']

    known: List[str] = []
    seen = set()
    for e in config:
        disp = e.get('alias') or e['name']
        if disp not in seen:
            seen.add(disp)
            known.append(disp)

    present = sorted(dedup['처방의'].unique().tolist())
    new_names = sorted(set(present) - seen)
    if new_names:
        save_doctor_config(config + [{'name': n, 'alias': ''} for n in new_names])
        known = known + new_names

    dedup['처방의'] = pd.Categorical(dedup['처방의'], categories=known, ordered=True)
    return dedup

# ── 공통 집계 ─────────────────────────────────────────────────────────────
def _aggregate(dedup: pd.DataFrame) -> Dict[str, Any]:
    """중복 제거된 DataFrame에서 일별/주별/요약 집계 수행"""
    dedup = _apply_doctor_order(dedup)

    min_date: datetime.date = dedup['접수일자'].min()
    year_month: Tuple[int, int] = (min_date.year, min_date.month)
    period: str = f"{min_date.strftime('%Y년 %m월')} 집계"

    daily = (
        dedup.groupby(['처방의', '접수일자', '구분'], observed=True)
        .size().unstack(fill_value=0).reset_index()
    )
    daily.columns.name = None
    for col in ('입원', '외래'):
        if col not in daily.columns:
            daily[col] = 0
    daily['합계'] = daily['입원'] + daily['외래']
    daily['요일'] = daily['접수일자'].apply(get_dayname)
    daily['주차'] = daily['접수일자'].apply(get_week_label)
    daily['휴일'] = daily['접수일자'].apply(is_red_day)

    weeks_ordered = list(dict.fromkeys(
        get_week_label(d) for d in sorted(dedup['접수일자'].unique())
    ))

    weekly = (
        daily.groupby(['처방의', '주차'], observed=True)
        .agg(입원=('입원', 'sum'), 외래=('외래', 'sum'), 합계=('합계', 'sum'))
        .reset_index()
    )

    summary = (
        dedup.groupby(['처방의', '구분'], observed=True)
        .size().unstack(fill_value=0).reset_index()
    )
    summary.columns.name = None
    for col in ('입원', '외래'):
        if col not in summary.columns:
            summary[col] = 0
    summary['합계'] = summary['입원'] + summary['외래']
    summary = summary[['처방의', '입원', '외래', '합계']]

    all_daily = (
        daily.groupby('접수일자')
        .agg(입원=('입원', 'sum'), 외래=('외래', 'sum'), 합계=('합계', 'sum'))
        .reset_index()
    )
    all_daily['요일'] = all_daily['접수일자'].apply(get_dayname)
    all_daily['주차'] = all_daily['접수일자'].apply(get_week_label)
    all_daily['휴일'] = all_daily['접수일자'].apply(is_red_day)

    all_weekly = (
        daily.groupby('주차')
        .agg(입원=('입원', 'sum'), 외래=('외래', 'sum'), 합계=('합계', 'sum'))
        .reindex(weeks_ordered).reset_index()
    )

    return {
        'raw':        dedup,
        'summary':    summary,
        'daily':      daily,
        'weekly':     weekly,
        'all_daily':  all_daily,
        'all_weekly': all_weekly,
        'docs':       summary['처방의'].tolist(),
        'weeks':      weeks_ordered,
        'period':     period,
        'year_month': year_month,
        'warnings':   [],
    }

# ── 단일 파일 처리 ────────────────────────────────────────────────────────
def load_and_process(filepath: "str | Path") -> Dict[str, Any]:
    """
    엑셀 파일을 읽어 모든 집계 데이터 반환
    여러 달 데이터가 섞여 있으면 가장 많은 달 기준으로 처리하고 경고 발생
    """
    fp = Path(filepath)

    try:
        engine = 'xlrd' if fp.suffix.lower() == '.xls' else 'openpyxl'
        df = pd.read_excel(fp, sheet_name=0, engine=engine)
    except PermissionError:
        raise PermissionError(
            f"파일이 열려 있습니다: {fp.name} — Excel을 닫고 다시 시도하세요."
        )
    except Exception as e:
        raise ValueError(f"파일 읽기 실패: {fp.name} — {e}")

    required = ['환자성명', '처방의', '접수일자', '등록번호']
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"필수 컬럼 누락: {missing} ({fp.name})")
    if df.empty:
        raise ValueError(f"빈 파일: {fp.name}")

    df['구분'] = df.apply(_classify_inout, axis=1)
    df = df.dropna(subset=['처방의']).copy()

    orig_doctors = df[['처방의']].copy()  # 정규화 전 원본 (경고 감지용)
    df['처방의']  = df['처방의'].apply(_normalize_doctor)
    df['접수일자'] = pd.to_datetime(df['접수일자'], errors='coerce').dt.date
    df = df.dropna(subset=['접수일자'])
    df['등록번호_norm'] = df['등록번호'].apply(_normalize_reg)

    if df.empty:
        raise ValueError(f"유효 데이터 없음: {fp.name}")

    warnings = detect_doctor_name_issues(orig_doctors)

    # 여러 달 데이터 감지
    ym_counts = df.groupby(df['접수일자'].apply(
        lambda d: (d.year, d.month)
    )).size()

    if len(ym_counts) > 1:
        months_str = ', '.join(f"{y}-{m:02d}" for (y, m) in ym_counts.index)
        dominant   = ym_counts.idxmax()
        warnings.append(
            f"여러 달 데이터 혼재 감지: {months_str} "
            f"→ 가장 많은 {dominant[0]}-{dominant[1]:02d} 기준으로 처리"
        )
        df = df[df['접수일자'].apply(lambda d: (d.year, d.month)) == dominant]

    dedup = df.drop_duplicates(
        subset=['처방의', '등록번호_norm', '접수일자', '구분']
    )

    result = _aggregate(dedup)
    result['warnings'] = warnings
    result['filepath'] = fp
    return result


# ── 여러 파일 합산 처리 ───────────────────────────────────────────────────
def load_and_merge(filepaths: list) -> Dict[str, Any]:
    """
    같은 달 파일 여러 개를 읽어 합산 처리
    1주차, 2주차 ... 파일을 각각 넣으면 자동으로 합쳐서 하나의 데이터로 반환

    Args:
        filepaths: 같은 달 xlsx/xls 파일 경로 목록
    Returns:
        load_and_process() 와 동일한 구조의 dict
    """
    if not filepaths:
        raise ValueError("처리할 파일 없음")

    if len(filepaths) == 1:
        return load_and_process(filepaths[0])

    all_warnings: List[str] = []
    frames: List[pd.DataFrame] = []

    for fp in filepaths:
        fp = Path(fp)
        try:
            engine = 'xlrd' if fp.suffix.lower() == '.xls' else 'openpyxl'
            df = pd.read_excel(fp, sheet_name=0, engine=engine)
        except PermissionError:
            raise PermissionError(f"파일이 열려 있습니다: {fp.name}")
        except Exception as e:
            raise ValueError(f"파일 읽기 실패: {fp.name} — {e}")

        required = ['환자성명', '처방의', '접수일자', '등록번호']
        missing = [c for c in required if c not in df.columns]
        if missing:
            raise ValueError(f"필수 컬럼 누락: {missing} ({fp.name})")

        df['_source'] = fp.name
        frames.append(df)

    merged = pd.concat(frames, ignore_index=True)

    merged['구분'] = merged.apply(_classify_inout, axis=1)
    merged = merged.dropna(subset=['처방의']).copy()

    orig_doctors = merged[['처방의']].copy()  # 정규화 전 원본 (경고 감지용)
    merged['처방의']   = merged['처방의'].apply(_normalize_doctor)
    merged['접수일자'] = pd.to_datetime(merged['접수일자'], errors='coerce').dt.date
    merged = merged.dropna(subset=['접수일자'])
    merged['등록번호_norm'] = merged['등록번호'].apply(_normalize_reg)

    if merged.empty:
        raise ValueError("합산 후 유효 데이터 없음")

    ym_counts = merged.groupby(
        merged['접수일자'].apply(lambda d: (d.year, d.month))
    ).size()

    if len(ym_counts) > 1:
        months_str = ', '.join(f"{y}-{m:02d}" for (y, m) in ym_counts.index)
        dominant   = ym_counts.idxmax()
        all_warnings.append(
            f"여러 달 데이터 혼재: {months_str} → {dominant[0]}-{dominant[1]:02d} 기준"
        )
        merged = merged[
            merged['접수일자'].apply(lambda d: (d.year, d.month)) == dominant
        ]

    all_warnings += detect_doctor_name_issues(orig_doctors)

    dedup = merged.drop_duplicates(
        subset=['처방의', '등록번호_norm', '접수일자', '구분']
    )

    source_files = [Path(fp).name for fp in filepaths]
    all_warnings.insert(0, f"합산 파일 {len(filepaths)}개: {', '.join(source_files)}")

    result = _aggregate(dedup)
    result['warnings']   = all_warnings
    result['filepath']   = Path(filepaths[-1])
    result['filepaths']  = [Path(fp) for fp in filepaths]
    return result
