"""
run.py — 물리치료 집계 자동화 메인
Python 3.8+ 호환

사용법:
  python run.py                  # 오늘 날짜 기준 자동 실행
  python run.py 파일.xlsx        # 특정 파일 처리
  python run.py --force          # 전체 강제 실행 (날짜/중복 무시)
  python run.py --weekly         # 주별만 즉시 실행 (GUI 버튼용)
  python run.py --monthly        # 월별만 즉시 실행 (GUI 버튼용)
  python run.py --validate       # 정합성 검증만
  python run.py --schedule       # 스케줄러 데몬 모드 (창 켜놔야 함)
  ※ 창 없는 자동 실행은 Windows 작업 스케줄러 + run_silent.vbs 사용
"""
from __future__ import annotations

import sys
import logging
import logging.handlers
import datetime
import traceback
import time
from pathlib import Path
from typing import Optional, Set, Tuple

# Windows cmd.exe 환경에서 한글 로그가 깨지지 않도록 UTF-8 강제
for _stream in (sys.stdout, sys.stderr):
    if _stream and hasattr(_stream, 'reconfigure'):
        try:
            _stream.reconfigure(encoding='utf-8')
        except Exception:
            pass

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / 'scripts'))

DATA_WEEKLY   = ROOT / 'data' / 'weekly'
DATA_MONTHLY  = ROOT / 'data' / 'monthly'
PROCESSED_LOG = ROOT / 'processed.log'
LOG_FILE      = ROOT / 'run.log'

# ── 로깅 (10MB 로테이션, 3개 보관) ──────────────────────────────────────
def _setup_logging() -> logging.Logger:
    fmt = logging.Formatter(
        '%(asctime)s [%(levelname)s] %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    logger = logging.getLogger('pt_report')
    if logger.handlers:
        return logger
    logger.setLevel(logging.INFO)

    if sys.stdout is not None:
        ch = logging.StreamHandler(sys.stdout)
        ch.setFormatter(fmt)
        logger.addHandler(ch)

    try:
        fh = logging.handlers.RotatingFileHandler(
            LOG_FILE, maxBytes=10*1024*1024, backupCount=3, encoding='utf-8'
        )
        fh.setFormatter(fmt)
        logger.addHandler(fh)
    except Exception:
        pass
    return logger

log = _setup_logging()

# ── 지연 임포트 ───────────────────────────────────────────────────────────
def _get_process():
    from process import load_and_process, get_file_year_month, load_and_merge
    return load_and_process, get_file_year_month, load_and_merge

def _get_generators():
    from excel_report import generate_excel
    from html_dashboard import generate_html
    return generate_excel, generate_html

# ── output 경로 ───────────────────────────────────────────────────────────
def get_output_dirs(year: int, month: int) -> Tuple[Path, Path]:
    ym = f'{year}-{month:02d}'
    reports   = ROOT / 'output' / ym / 'reports'
    dashboard = ROOT / 'output' / ym / 'dashboard'
    reports.mkdir(parents=True, exist_ok=True)
    dashboard.mkdir(parents=True, exist_ok=True)
    return reports, dashboard

def _move_to_done(filepath: Path) -> None:
    """처리 완료 파일을 data/_done/YYYY-MM/ 으로 이동 (원본 보존)"""
    try:
        import shutil
        from process import get_file_year_month
        ym = get_file_year_month(filepath)
        if ym:
            done_dir = filepath.parent / '_done' / f'{ym[0]}-{ym[1]:02d}'
        else:
            done_dir = filepath.parent / '_done'
        done_dir.mkdir(parents=True, exist_ok=True)
        dest = done_dir / filepath.name
        # 같은 이름 파일이 이미 있으면 타임스탬프 붙여서 저장
        if dest.exists():
            stamp = datetime.datetime.now().strftime('%H%M%S')
            dest  = done_dir / f"{filepath.stem}_{stamp}{filepath.suffix}"
        shutil.move(str(filepath), str(dest))
        log.info(f"원본 파일 이동: {filepath.name} → data/_done/")
    except Exception as e:
        log.warning(f"원본 파일 이동 실패 (파일은 그대로): {e}")


def _cleanup_old_outputs(reports_dir: Path, dashboard_dir: Path) -> None:
    """
    같은 기간 이전 결과물을 _archive/ 로 이동 (최신 1개만 유지)
    """
    archive_r = reports_dir   / '_archive'
    archive_d = dashboard_dir / '_archive'

    for folder, archive, suffix in [
        (reports_dir,   archive_r, '.xlsx'),
        (dashboard_dir, archive_d, '.html'),
    ]:
        files = sorted(
            [f for f in folder.glob(f'*{suffix}') if not f.name.startswith('정합성')],
            key=lambda p: p.stat().st_mtime
        )
        if len(files) > 1:
            archive.mkdir(exist_ok=True)
            for old_file in files[:-1]:  # 최신 1개 제외하고 이동
                try:
                    old_file.rename(archive / old_file.name)
                except Exception:
                    pass

# ── 처리 완료 기록 ────────────────────────────────────────────────────────
def _load_processed() -> Set[str]:
    """오늘 처리 완료한 '파일경로|수정시간' 집합 반환"""
    if not PROCESSED_LOG.exists():
        return set()
    today = datetime.date.today().isoformat()
    result = set()
    try:
        for line in PROCESSED_LOG.read_text(encoding='utf-8').splitlines():
            parts = line.strip().split('|')
            if len(parts) >= 2 and parts[0] == today:
                # 경로|수정시간 또는 구버전 경로만인 경우 모두 지원
                key = '|'.join(parts[1:])
                result.add(key)
    except Exception:
        pass
    return result

def _is_already_processed(filepath: Path, processed: Set[str]) -> bool:
    """파일이 오늘 이미 처리됐는지 확인 (경로+수정시간 기준)"""
    try:
        mtime = int(filepath.stat().st_mtime)
        key = f"{filepath.resolve()}|{mtime}"
        # 새 형식: 경로|수정시간
        if key in processed:
            return True
        # 구버전 호환: 경로만
        if str(filepath.resolve()) in processed:
            return True
    except Exception:
        pass
    return False

def _mark_processed(filepath: Path) -> None:
    """경로 + 파일 수정시간 기록 (같은 경로 파일 교체도 감지)"""
    today = datetime.date.today().isoformat()
    try:
        mtime = int(filepath.stat().st_mtime)
        with open(PROCESSED_LOG, 'a', encoding='utf-8') as f:
            f.write(f"{today}|{filepath.resolve()}|{mtime}\n")
    except Exception as e:
        log.warning(f"처리 기록 저장 실패: {e}")

# ── 파일 유효성 ───────────────────────────────────────────────────────────
def is_valid_xlsx(filepath: Path) -> bool:
    if not filepath.is_file() or filepath.suffix.lower() != '.xlsx':
        return False
    if filepath.stat().st_size < 1024:
        return False
    try:
        import openpyxl
        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        wb.close()
        return True
    except Exception:
        return False

# ── 파일 탐색 ─────────────────────────────────────────────────────────────
def _glob_excel(folder: Path) -> list:
    """xls + xlsx 모두 탐색 (수정시간 내림차순)"""
    files = list(folder.glob('*.xlsx')) + list(folder.glob('*.xls'))
    return sorted(files, key=lambda p: p.stat().st_mtime, reverse=True)
def _find_latest_for_month(
    folder: Path, year: int, month: int, processed: Set[str]
) -> Optional[Path]:
    _, get_ym, _ = _get_process()
    candidates = [
        f for f in sorted(folder.glob('*.xlsx'),
                          key=lambda p: p.stat().st_mtime, reverse=True)
        if is_valid_xlsx(f)
        and not _is_already_processed(f, processed)
        and get_ym(f) == (year, month)
    ]
    return candidates[0] if candidates else None

def _find_all_for_month(
    folder: Path, year: int, month: int, processed: Set[str]
) -> list:
    """같은 달 미처리 파일 전부 반환 (주별 합산용, 오래된 순)"""
    _, get_ym, _ = _get_process()
    candidates = [
        f for f in sorted(folder.glob('*.xlsx'),
                          key=lambda p: p.stat().st_mtime)  # 오래된 것부터
        if is_valid_xlsx(f)
        and not _is_already_processed(f, processed)
        and get_ym(f) == (year, month)
    ]
    return candidates

def _find_any_for_month(folder: Path, year: int, month: int) -> Optional[Path]:
    _, get_ym, _ = _get_process()
    candidates = [
        f for f in sorted(folder.glob('*.xlsx'),
                          key=lambda p: p.stat().st_mtime, reverse=True)
        if is_valid_xlsx(f) and get_ym(f) == (year, month)
    ]
    return candidates[0] if candidates else None

def _find_all_year_months() -> list:
    _, get_ym, _ = _get_process()
    ym_set = set()
    for folder in (DATA_WEEKLY, DATA_MONTHLY):
        for f in folder.glob('*.xlsx'):
            if is_valid_xlsx(f):
                ym = get_ym(f)
                if ym:
                    ym_set.add(ym)
    return sorted(ym_set)

# ── 날짜 판단 ─────────────────────────────────────────────────────────────
def _get_today_mode(today: datetime.date) -> Optional[str]:
    if today.weekday() == 6:
        log.info(f"일요일({today}) — 스킵")
        return None
    if today.day in (1, 2):
        return 'monthly'
    if today.weekday() in (0, 1):
        return 'weekly'
    log.info(f"실행 대상 날짜 아님({today}) — 스킵")
    return None

def _get_target_ym(today: datetime.date, mode: str) -> Tuple[int, int]:
    if mode == 'monthly':
        prev = today.replace(day=1) - datetime.timedelta(days=1)
        return prev.year, prev.month
    return today.year, today.month

# ── 단일 파일 처리 (재시도 포함) ─────────────────────────────────────────

def process_file(
    filepath: Path,
    mark: bool = True,
    counterpart: Optional[Path] = None,
    retry: bool = True,
) -> bool:
    """
    filepath:    처리할 파일
    mark:        완료 기록 여부
    counterpart: 정합성 비교 파일
    retry:       파일 잠김 시 30분 후 1회 재시도
    """
    load_and_process, _, load_and_merge = _get_process()
    generate_excel, generate_html = _get_generators()

    log.info(f"처리 시작: {filepath.name}")

    for attempt in range(2):  # 최대 2회 시도
        try:
            data = load_and_process(filepath)
            break
        except PermissionError as e:
            if attempt == 0 and retry:
                log.warning(f"파일 잠김 — 30분 후 재시도: {filepath.name}")
                time.sleep(30 * 60)
                continue
            log.error(str(e))
            return False
        except ValueError as e:
            log.error(f"데이터 오류: {e}")
            return False
        except Exception as e:
            log.error(f"처리 실패: {filepath.name} — {e}")
            log.debug(traceback.format_exc())
            return False
    else:
        return False

    # 경고 로깅
    for w in data.get('warnings', []):
        log.warning(w)

    year, month = data['year_month']
    period      = data['period']
    stamp       = datetime.datetime.now().strftime('%Y%m%d_%H%M')

    reports_dir, dashboard_dir = get_output_dirs(year, month)
    excel_path = reports_dir   / f"집계리포트_{period}_{stamp}.xlsx"
    html_path  = dashboard_dir / f"대시보드_{period}_{stamp}.html"

    # 정합성 검증
    validation_result = None
    if counterpart and is_valid_xlsx(counterpart):
        try:
            load_fn, _, _ = _get_process()
            other_data = load_fn(counterpart)
            if other_data['year_month'] == (year, month):
                from validator import validate
                # 항상 (주별누적, 월별) 순서
                if counterpart.parent == DATA_MONTHLY:
                    validation_result = validate(data, other_data)
                else:
                    validation_result = validate(other_data, data)
                status = '✓ 일치' if validation_result['ok'] else \
                         f"⚠ 불일치 {len(validation_result['detail'])}건"
                log.info(f"정합성 검증: {status}")
            else:
                log.info("정합성 검증 스킵 (연월 불일치)")
        except Exception as e:
            log.warning(f"정합성 검증 실패 (리포트는 생성 계속): {e}")

    try:
        generate_excel(data, excel_path, validation_result=validation_result)
        generate_html(data, html_path)
    except Exception as e:
        log.error(f"출력 생성 실패: {e}")
        log.debug(traceback.format_exc())
        return False

    # 이전 결과물 archive 이동 (최신 1개 유지)
    _cleanup_old_outputs(reports_dir, dashboard_dir)

    # 처리 완료 원본 파일 → data/_done/ 으로 이동
    if mark:
        _mark_processed(filepath)
        _move_to_done(filepath)

    log.info(f"✓ 완료: {period} → output/{year}-{month:02d}/")
    return True

# ── 주별 / 월별 실행 ──────────────────────────────────────────────────────
def _merge_weekly_files(year: int, month: int) -> Optional[Path]:
    """
    data/weekly/ 의 같은 달 파일 전부를 합산한 임시 파일 생성
    처리 완료(_done/)된 파일 + 현재 파일 모두 포함
    """
    import pandas as pd
    import tempfile
    _, get_ym, _ = _get_process()

    # 현재 + _done/ 폴더에서 같은 달 파일 모두 수집
    all_files = []
    for search_dir in [DATA_WEEKLY, DATA_WEEKLY / '_done' / f'{year}-{month:02d}']:
        if search_dir.exists():
            for f in sorted(search_dir.glob('*.xlsx'), key=lambda p: p.stat().st_mtime):
                if is_valid_xlsx(f) and get_ym(f) == (year, month):
                    all_files.append(f)

    if not all_files:
        return None

    if len(all_files) == 1:
        return all_files[0]  # 파일 1개면 그대로 반환

    # 여러 파일 합산
    log.info(f"주별 파일 {len(all_files)}개 합산: {[f.name for f in all_files]}")
    dfs = []
    for f in all_files:
        try:
            _engine = 'xlrd' if f.suffix.lower() == '.xls' else 'openpyxl'
            dfs.append(pd.read_excel(f, sheet_name=0, engine=_engine))
        except Exception as e:
            log.warning(f"합산 중 파일 읽기 실패 (스킵): {f.name} — {e}")

    if not dfs:
        return None

    merged = pd.concat(dfs, ignore_index=True)

    # 임시 파일로 저장 (처리 후 자동 삭제)
    tmp = tempfile.NamedTemporaryFile(
        suffix='.xlsx', delete=False,
        dir=DATA_WEEKLY,
        prefix=f'_merged_{year}{month:02d}_'
    )
    tmp.close()
    tmp_path = Path(tmp.name)
    try:
        merged.to_excel(tmp_path, index=False)
    except Exception:
        try:
            tmp_path.unlink()
        except Exception:
            pass
        raise
    return tmp_path


def _run_weekly(year: int, month: int, processed: Set[str], mark: bool = True) -> None:
    log.info(f"=== 주별 집계 {year}-{month:02d} ===")

    # 새로 추가된 파일이 있는지 확인
    new_fp = _find_latest_for_month(DATA_WEEKLY, year, month, processed)
    if not new_fp:
        log.info("새로 처리할 파일 없음 (없거나 이미 처리됨)")
        return

    # 같은 달 전체 파일 합산 (누적 리포트)
    merged_fp = _merge_weekly_files(year, month)
    is_temp = merged_fp is not None and merged_fp.name.startswith('_merged_')

    try:
        target_fp = merged_fp if merged_fp else new_fp
        counterpart = _find_any_for_month(DATA_MONTHLY, year, month)
        # 합산 파일은 mark/move 하지 않고 새 파일(new_fp)만 기록/이동
        ok = process_file(target_fp, mark=False, counterpart=counterpart)
        if ok and mark:
            _mark_processed(new_fp)
            _move_to_done(new_fp)
    finally:
        # 임시 합산 파일 삭제
        if is_temp and merged_fp and merged_fp.exists():
            try:
                merged_fp.unlink()
            except Exception:
                pass

def _run_monthly(year: int, month: int, processed: Set[str], mark: bool = True) -> None:
    log.info(f"=== 월별 집계 {year}-{month:02d} ===")
    fp = _find_latest_for_month(DATA_MONTHLY, year, month, processed)
    if not fp:
        log.info("data/monthly/ 파일 없음 → data/weekly/ 최신 파일로 대체")
        fp = _find_latest_for_month(DATA_WEEKLY, year, month, processed)
    if not fp:
        log.info("처리할 파일 없음")
        return
    counterpart = _find_any_for_month(DATA_WEEKLY, year, month) \
                  if fp.parent == DATA_MONTHLY else None
    process_file(fp, mark=mark, counterpart=counterpart)

# ── 자동 실행 ─────────────────────────────────────────────────────────────
def run_auto(force: bool = False) -> None:
    today     = datetime.date.today()
    processed = set() if force else _load_processed()

    if force:
        ym_list = _find_all_year_months()
        if not ym_list:
            log.info("처리할 파일 없음")
            return
        log.info(f"강제 실행: {[f'{y}-{m:02d}' for y, m in ym_list]}")
        for year, month in ym_list:
            _run_weekly(year, month, processed, mark=False)
            _run_monthly(year, month, processed, mark=False)
        return

    mode = _get_today_mode(today)
    if mode is None:
        return

    year, month = _get_target_ym(today, mode)
    if mode == 'monthly':
        _run_monthly(year, month, processed)
    else:
        _run_weekly(year, month, processed)

# ── 정합성 검증 단독 ──────────────────────────────────────────────────────
def run_validate_only() -> None:
    today = datetime.date.today()
    mode  = 'monthly' if today.day in (1, 2) else 'weekly'
    year, month = _get_target_ym(today, mode)

    log.info(f"=== 정합성 검증 {year}-{month:02d} ===")

    weekly_fp  = _find_any_for_month(DATA_WEEKLY,  year, month)
    monthly_fp = _find_any_for_month(DATA_MONTHLY, year, month)

    if not weekly_fp:
        log.warning(f"data/weekly/ 에 {year}-{month:02d} 파일 없음")
        return
    if not monthly_fp:
        log.warning(f"data/monthly/ 에 {year}-{month:02d} 파일 없음")
        return

    try:
        load_and_process, _, load_and_merge = _get_process()
        from validator import validate, add_validation_sheet
        import openpyxl

        w_data = load_and_process(weekly_fp)
        m_data = load_and_process(monthly_fp)

        for w in w_data.get('warnings', []) + m_data.get('warnings', []):
            log.warning(w)

        result = validate(w_data, m_data)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = '커버'
        ws['A1'] = (f"정합성 검증 — {year}-{month:02d}  "
                    f"({datetime.datetime.now().strftime('%Y-%m-%d %H:%M')})")
        add_validation_sheet(wb, result)

        reports_dir, _ = get_output_dirs(year, month)
        stamp = datetime.datetime.now().strftime('%Y%m%d_%H%M')
        out = reports_dir / f"정합성검증_{year}{month:02d}_{stamp}.xlsx"
        wb.save(out)

        status = '✓ 일치' if result['ok'] else f"⚠ 불일치 {len(result['detail'])}건"
        log.info(f"검증 완료: {status} → {out.name}")

    except Exception as e:
        log.error(f"검증 실패: {e}")
        log.debug(traceback.format_exc())

# ── 스케줄러 데몬 ─────────────────────────────────────────────────────────
def run_schedule_mode() -> None:
    try:
        from apscheduler.schedulers.blocking import BlockingScheduler
        from apscheduler.triggers.cron import CronTrigger
    except ImportError:
        log.error("apscheduler 미설치:\n  pip install apscheduler")
        sys.exit(1)

    scheduler = BlockingScheduler(timezone='Asia/Seoul')
    scheduler.add_job(run_auto, CronTrigger(day_of_week='mon', hour=12),
                      id='weekly', misfire_grace_time=600, coalesce=True)
    scheduler.add_job(run_auto, CronTrigger(day='1', hour=12),
                      id='monthly', misfire_grace_time=600, coalesce=True)

    log.info("스케줄러 시작 — 주별 월요일 12:00 / 월별 1일 12:00")
    log.info("※ 창 없이 실행하려면 이 모드 대신 Windows 작업 스케줄러 사용")
    try:
        scheduler.start()
    except (KeyboardInterrupt, SystemExit):
        log.info("스케줄러 종료")

# ── 메인 ──────────────────────────────────────────────────────────────────
def main() -> None:
    # 데이터 폴더 없으면 자동 생성
    DATA_WEEKLY.mkdir(parents=True, exist_ok=True)
    DATA_MONTHLY.mkdir(parents=True, exist_ok=True)

    args = sys.argv[1:]

    if '--schedule' in args:
        run_schedule_mode()
    elif '--validate' in args:
        run_validate_only()
    elif '--weekly' in args:
        # GUI 주별 버튼용: 오늘 기준 or 가장 최근 연월 주별 처리
        today = datetime.date.today()
        ym_list = _find_all_year_months()
        year, month = ym_list[-1] if ym_list else (today.year, today.month)
        _run_weekly(year, month, set(), mark=True)
    elif '--monthly' in args:
        # GUI 월별 버튼용: 오늘 기준 or 가장 최근 연월 월별 처리
        today = datetime.date.today()
        ym_list = _find_all_year_months()
        year, month = ym_list[-1] if ym_list else (today.year, today.month)
        _run_monthly(year, month, set(), mark=True)
    elif '--force' in args:
        remaining = [a for a in args if a != '--force']
        if remaining:
            fp = Path(remaining[0])
            if not is_valid_xlsx(fp):
                log.error(f"유효하지 않은 파일: {fp}")
                sys.exit(1)
            process_file(fp, mark=False)
        else:
            run_auto(force=True)
    elif args and not args[0].startswith('--'):
        fp = Path(args[0])
        if not is_valid_xlsx(fp):
            log.error(f"유효하지 않은 파일: {fp}")
            sys.exit(1)
        process_file(fp)
    else:
        run_auto()

if __name__ == '__main__':
    main()
