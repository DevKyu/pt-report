"""
app.py - PT Report GUI (Windows)
- 열고 닫는 용도 (트레이 없음)
- 즉시 실행 / 결과 열기 / 폴더 열기
- 스케줄러는 별도로 백그라운드 동작
"""
from __future__ import annotations

import os
import sys
import ctypes
import threading
import queue
import datetime
import subprocess
import logging
import logging.handlers
import tkinter as tk
from tkinter import messagebox
from pathlib import Path

# ── 경로 설정 (exe/py 모두 대응) ─────────────────────────────────────────
if getattr(sys, 'frozen', False):
    ROOT   = Path(sys.executable).parent
    BUNDLE = Path(sys._MEIPASS)
    sys.path.insert(0, str(BUNDLE))
    sys.path.insert(0, str(BUNDLE / 'scripts'))
else:
    ROOT   = Path(__file__).resolve().parent
    BUNDLE = ROOT
    sys.path.insert(0, str(ROOT / 'scripts'))

# ── 로거 선점 초기화 ──────────────────────────────────────────────────────
# exec_module 로 run.py 를 로드하면 run.py 의 _setup_logging 이 실행된다.
# 그 시점 __file__ 은 MEIPASS(임시폴더)이므로 LOG_FILE 경로가 틀려진다.
# 여기서 먼저 핸들러를 달아두면 _setup_logging 이 "이미 핸들러 있음" 판단 후 스킵한다.
_pt_logger = logging.getLogger('pt_report')
if not _pt_logger.handlers:
    _pt_logger.setLevel(logging.INFO)
    try:
        _fh = logging.handlers.RotatingFileHandler(
            ROOT / 'run.log', maxBytes=10 * 1024 * 1024, backupCount=3, encoding='utf-8'
        )
        _fh.setFormatter(logging.Formatter(
            '%(asctime)s [%(levelname)s] %(message)s', datefmt='%Y-%m-%d %H:%M:%S'
        ))
        _pt_logger.addHandler(_fh)
    except Exception:
        pass

# ── 색상 / 폰트 ──────────────────────────────────────────────────────────
C = {
    'bg':      '#EFF3F8',
    'surface': '#FFFFFF',
    'border':  '#CBD5E1',
    'primary': '#1E3A5F',
    'text':    '#1A202C',
    'muted':   '#64748B',
    'success': '#15803D',
    'error':   '#B91C1C',
    'accent':  '#1D4ED8',
    'btn_wk':  '#1D4ED8',
    'btn_mo':  '#0F766E',
    'btn_res': '#15803D',
    'btn_dim': '#64748B',
}
F      = ('Malgun Gothic', 10)
FB     = ('Malgun Gothic', 10, 'bold')
FTITLE = ('Malgun Gothic', 12, 'bold')
FS     = ('Malgun Gothic', 9)
FLOG   = ('Consolas', 9)


def _darker(h: str, a: int = 25) -> str:
    try:
        r, g, b = int(h[1:3],16), int(h[3:5],16), int(h[5:7],16)
        return f'#{max(0,r-a):02x}{max(0,g-a):02x}{max(0,b-a):02x}'
    except Exception:
        return h


class GUILogHandler(logging.Handler):
    def __init__(self, q: queue.Queue):
        super().__init__()
        self._q = q

    def emit(self, record):
        msg = record.getMessage()
        tag = ('ok'   if '완료' in msg else
               'warn' if record.levelno >= logging.WARNING else
               'err'  if record.levelno >= logging.ERROR else 'info')
        self._q.put(('log', tag, msg))


class PTApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title('물리치료 집계')
        self.root.configure(bg=C['bg'])
        self.root.resizable(False, False)

        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass

        self._q       = queue.Queue()
        self._running = False
        self._pending = {}

        for d in ('data/weekly', 'data/monthly', 'output'):
            (ROOT / d).mkdir(parents=True, exist_ok=True)

        self._build()
        self._refresh()
        # 창 닫으면 그냥 종료
        self.root.protocol('WM_DELETE_WINDOW', self.root.destroy)
        self.root.after(300, self._poll)
        self.root.after(900, self._auto_register)

    # ── UI ───────────────────────────────────────────────────────────────
    def _build(self):
        r = self.root

        # 헤더
        hdr = tk.Frame(r, bg=C['primary'], pady=11)
        hdr.pack(fill='x')
        tk.Label(hdr, text='물리치료 집계 자동화',
                 font=FTITLE, bg=C['primary'], fg='white').pack()
        tk.Label(hdr, text='처방 데이터 집계  ·  Excel + 대시보드 자동 생성',
                 font=FS, bg=C['primary'], fg='#93B8D8').pack()

        body = tk.Frame(r, bg=C['bg'], padx=14, pady=12)
        body.pack(fill='both', expand=True)

        # 주별 / 월별 2열
        row1 = tk.Frame(body, bg=C['bg'])
        row1.pack(fill='x', pady=(0,10))
        row1.columnconfigure(0, weight=1, uniform='c')
        row1.columnconfigure(1, weight=1, uniform='c')

        wk = self._card(row1, 0, 0, padr=5)
        tk.Label(wk, text='📁  주별 집계', font=FB,
                 bg=C['surface'], fg=C['primary']).pack(anchor='w')
        self._lbl_wk = tk.Label(wk, text='아직 실행 안 됨',
                                  font=FS, bg=C['surface'], fg=C['muted'])
        self._lbl_wk.pack(anchor='w', pady=(3,8))
        self._btn_wk = self._mkbtn(wk, '▶  주별 실행',
                                    self._run_wk, C['btn_wk'], big=True)
        self._btn_wk.pack(fill='x', pady=(0,5))
        self._mkbtn(wk, '📂  폴더 열기',
                    lambda: self._openfolder(ROOT/'data'/'weekly'),
                    C['btn_dim']).pack(fill='x')

        mo = self._card(row1, 0, 1, padl=5)
        tk.Label(mo, text='📋  월별 집계', font=FB,
                 bg=C['surface'], fg=C['primary']).pack(anchor='w')
        self._lbl_mo = tk.Label(mo, text='아직 실행 안 됨',
                                  font=FS, bg=C['surface'], fg=C['muted'])
        self._lbl_mo.pack(anchor='w', pady=(3,8))
        self._btn_mo = self._mkbtn(mo, '▶  월별 실행',
                                    self._run_mo, C['btn_mo'], big=True)
        self._btn_mo.pack(fill='x', pady=(0,5))
        self._mkbtn(mo, '📂  폴더 열기',
                    lambda: self._openfolder(ROOT/'data'/'monthly'),
                    C['btn_dim']).pack(fill='x')

        # 결과 열기
        res_o = tk.Frame(body, bg=C['surface'],
                         highlightbackground=C['border'], highlightthickness=1)
        res_o.pack(fill='x', pady=(0,10))
        res = tk.Frame(res_o, bg=C['surface'], padx=12, pady=10)
        res.pack(fill='x')
        top = tk.Frame(res, bg=C['surface'])
        top.pack(fill='x', pady=(0,6))
        tk.Label(top, text='✅  결과 열기', font=FB,
                 bg=C['surface'], fg=C['primary']).pack(side='left')
        self._lbl_res = tk.Label(top, text='결과 없음',
                                   font=FS, bg=C['surface'], fg=C['muted'])
        self._lbl_res.pack(side='right')
        br = tk.Frame(res, bg=C['surface'])
        br.pack(fill='x')
        br.columnconfigure((0,1,2), weight=1, uniform='rb')
        self._mkbtn(br, '📊  Excel',
                    self._open_excel, C['btn_res']).grid(
                    row=0, column=0, sticky='ew', padx=(0,4))
        self._mkbtn(br, '🌐  대시보드',
                    self._open_html, C['btn_res']).grid(
                    row=0, column=1, sticky='ew', padx=4)
        self._mkbtn(br, '📁  결과폴더',
                    lambda: self._openfolder(ROOT/'output'),
                    C['btn_dim']).grid(row=0, column=2, sticky='ew', padx=(4,0))

        # 로그 (전체 너비)
        log_o = tk.Frame(body, bg=C['surface'],
                         highlightbackground=C['border'], highlightthickness=1)
        log_o.pack(fill='x', pady=(0,8))
        logf = tk.Frame(log_o, bg=C['surface'], padx=12, pady=10)
        logf.pack(fill='x')
        lh = tk.Frame(logf, bg=C['surface'])
        lh.pack(fill='x', pady=(0,4))
        tk.Label(lh, text='📝  실행 로그', font=FB,
                 bg=C['surface'], fg=C['primary']).pack(side='left')
        self._mkbtn(lh, '파일 열기', self._open_log,
                    C['btn_dim']).pack(side='right', padx=(4,0))
        self._mkbtn(lh, '지우기', self._clear_log,
                    C['btn_dim']).pack(side='right')

        self._log_txt = tk.Text(logf, height=5, font=FLOG,
                                 bg='#1A202C', fg='#E2E8F0',
                                 relief='flat', state='disabled', wrap='word')
        self._log_txt.pack(fill='x')
        self._log_txt.tag_config('ok',   foreground='#4ADE80')
        self._log_txt.tag_config('warn', foreground='#FCD34D')
        self._log_txt.tag_config('err',  foreground='#F87171')
        self._log_txt.tag_config('info', foreground='#7DD3FC')

        # 자동 스케줄 안내 (한 줄)
        sch_o = tk.Frame(body, bg=C['surface'],
                         highlightbackground=C['border'], highlightthickness=1)
        sch_o.pack(fill='x')
        sch = tk.Frame(sch_o, bg=C['surface'], padx=12, pady=8)
        sch.pack(fill='x')
        tk.Label(sch, text='🕐',
                 font=F, bg=C['surface'], fg=C['primary']).pack(side='left', padx=(0,6))
        tk.Label(sch, text='주별  월요일  12:00',
                 font=FS, bg=C['surface'], fg=C['text']).pack(side='left', padx=(0,14))
        tk.Label(sch, text='월별  1일  12:00',
                 font=FS, bg=C['surface'], fg=C['text']).pack(side='left', padx=(0,14))
        tk.Label(sch, text='PC 켜져 있으면 자동 처리',
                 font=FS, bg=C['surface'], fg=C['muted']).pack(side='left')

        r.update_idletasks()
        r.update()
        h = max(r.winfo_reqheight() + 20, 580)
        r.geometry(f'540x{h}')
        r.minsize(540, 560)

    # ── 위젯 헬퍼 ────────────────────────────────────────────────────────
    def _card(self, parent, row, col, padl=0, padr=0):
        outer = tk.Frame(parent, bg=C['surface'],
                         highlightbackground=C['border'], highlightthickness=1)
        outer.grid(row=row, column=col, sticky='nsew', padx=(padl, padr))
        inner = tk.Frame(outer, bg=C['surface'], padx=12, pady=10)
        inner.pack(fill='both', expand=True)
        return inner

    def _mkbtn(self, parent, text, cmd, color, big=False):
        b = tk.Button(parent, text=text, command=cmd,
                      font=FB if big else F,
                      bg=color, fg='white', relief='flat',
                      pady=9 if big else 5, cursor='hand2',
                      activebackground=_darker(color),
                      activeforeground='white')
        b.bind('<Enter>', lambda e: b.config(bg=_darker(color)))
        b.bind('<Leave>', lambda e: b.config(bg=color))
        return b

    # ── 상태 갱신 ────────────────────────────────────────────────────────
    def _refresh(self):
        log_path = ROOT / 'run.log'
        lw = lm = None
        if log_path.exists():
            try:
                lines = log_path.read_text(
                    encoding='utf-8', errors='replace').splitlines()
                for line in reversed(lines):
                    if lw is None and '주별 집계' in line and '===' in line:
                        lw = line[:19]
                    if lm is None and '월별 집계' in line and '===' in line:
                        lm = line[:19]
                    if lw and lm:
                        break
            except Exception:
                pass

        def fmt(t): return f'✓ 마지막: {t}' if t else '아직 실행 안 됨'
        self._lbl_wk.config(text=fmt(lw),
                             fg=C['success'] if lw else C['muted'])
        self._lbl_mo.config(text=fmt(lm),
                             fg=C['success'] if lm else C['muted'])

        f = self._latest('reports', '집계리포트_*.xlsx')
        if f:
            mtime = datetime.datetime.fromtimestamp(
                f.stat().st_mtime).strftime('%Y-%m-%d %H:%M')
            self._lbl_res.config(
                text=f'{f.parent.parent.name}  처리: {mtime}',
                fg=C['success'])
        else:
            self._lbl_res.config(text='결과 없음', fg=C['muted'])

    # ── 실행 ─────────────────────────────────────────────────────────────
    def _run_wk(self):
        self._execute('weekly', self._btn_wk, self._lbl_wk, '주별 집계')

    def _run_mo(self):
        self._execute('monthly', self._btn_mo, self._lbl_mo, '월별 집계')

    def _execute(self, mode: str, btn, lbl, label: str):
        if self._running:
            messagebox.showinfo('알림', '현재 처리 중입니다.')
            return

        # 파일 존재 확인 (대소문자 무관: .xls/.XLS/.xlsx/.XLSX 모두 인식)
        def _ls_excel(d: Path):
            return [f for f in d.iterdir()
                    if f.is_file() and f.suffix.lower() in ('.xlsx', '.xls')
                    and f.stat().st_size > 1024]

        folder = ROOT / 'data' / ('weekly' if mode == 'weekly' else 'monthly')
        xlsx = _ls_excel(folder)
        if not xlsx and mode == 'monthly':
            xlsx = _ls_excel(ROOT / 'data' / 'weekly')
        if not xlsx:
            lbl.config(text='⚠ 처리할 파일 없음', fg=C['error'])
            messagebox.showwarning('파일 없음',
                f'data/{("weekly" if mode=="weekly" else "monthly")}/ 폴더에\n'
                f'.xlsx 파일을 넣고 다시 실행하세요.')
            return

        self._running = True
        orig_c = btn.cget('bg')
        orig_t = btn.cget('text')
        btn.config(state='disabled', bg='#94A3B8', text='⏳ 처리 중...')
        lbl.config(text='처리 중...', fg=C['accent'])
        self._pending = {'btn': btn, 'lbl': lbl, 'label': label,
                         'color': orig_c, 'text': orig_t}
        self._logwrite(f'{label} 시작', 'info')

        def worker():
            gui_handler = GUILogHandler(self._q)
            logger = logging.getLogger('pt_report')
            logger.addHandler(gui_handler)
            try:
                import importlib.util
                run_path = (BUNDLE / 'run.py') if getattr(sys, 'frozen', False) \
                           else (ROOT / 'run.py')
                spec = importlib.util.spec_from_file_location('pt_run', run_path)
                run_mod = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(run_mod)

                # 경로 덮어쓰기 (exe 환경 대응)
                run_mod.ROOT          = ROOT
                run_mod.DATA_WEEKLY   = ROOT / 'data' / 'weekly'
                run_mod.DATA_MONTHLY  = ROOT / 'data' / 'monthly'
                run_mod.PROCESSED_LOG = ROOT / 'processed.log'
                run_mod.LOG_FILE      = ROOT / 'run.log'

                ym_list = run_mod._find_all_year_months()
                if not ym_list:
                    self._q.put(('log', 'warn', '처리할 파일 없음'))
                    self._q.put(('done', False))
                    return

                ok = True
                for year, month in ym_list:
                    try:
                        if mode == 'weekly':
                            run_mod._run_weekly(year, month, set(), mark=True)
                        else:
                            run_mod._run_monthly(year, month, set(), mark=True)
                    except Exception as e:
                        self._q.put(('log', 'err', str(e)))
                        ok = False

                self._q.put(('done', ok))
            except Exception as e:
                import traceback
                self._q.put(('log', 'err', f'오류: {e}'))
                self._q.put(('log', 'err',
                             traceback.format_exc().splitlines()[-1]))
                self._q.put(('done', False))
            finally:
                logger.removeHandler(gui_handler)

        threading.Thread(target=worker, daemon=True).start()

    # ── 폴링 ─────────────────────────────────────────────────────────────
    def _poll(self):
        try:
            while True:
                item = self._q.get_nowait()
                if item[0] == 'log':
                    self._logwrite(item[2], item[1])
                else:
                    ok = item[1]
                    p  = self._pending
                    ts = datetime.datetime.now().strftime('%m/%d %H:%M')
                    p['btn'].config(state='normal',
                                    bg=p['color'], text=p['text'])
                    if ok:
                        p['lbl'].config(
                            text=f'✓ 완료  {ts}', fg=C['success'])
                        self._refresh()
                    else:
                        p['lbl'].config(
                            text='⚠ 오류  로그 확인', fg=C['error'])
                    self._running = False
        except queue.Empty:
            pass
        self.root.after(300, self._poll)

    # ── 파일 열기 ─────────────────────────────────────────────────────────
    def _openfolder(self, p: Path):
        p.mkdir(parents=True, exist_ok=True)
        os.startfile(str(p))

    def _open_excel(self):
        f = self._latest('reports', '집계리포트_*.xlsx')
        if f: os.startfile(str(f))
        else: messagebox.showinfo('알림', 'Excel 결과가 없습니다.\n먼저 집계를 실행하세요.')

    def _open_html(self):
        f = self._latest('dashboard', '대시보드_*.html')
        if f: os.startfile(str(f))
        else: messagebox.showinfo('알림', '대시보드가 없습니다.\n먼저 집계를 실행하세요.')

    def _latest(self, sub, pat):
        out = ROOT / 'output'
        if not out.exists(): return None
        for d in sorted([x for x in out.iterdir()
                          if x.is_dir() and x.name[:4].isdigit()],
                        reverse=True):
            files = [f for f in (d/sub).glob(pat)
                     if '_archive' not in str(f)] \
                    if (d/sub).exists() else []
            if files:
                return max(files, key=lambda p: p.stat().st_mtime)
        return None

    def _open_log(self):
        p = ROOT / 'run.log'
        if p.exists(): os.startfile(str(p))
        else: messagebox.showinfo('알림', 'run.log가 아직 없습니다.')

    def _logwrite(self, msg: str, tag: str = 'info'):
        self._log_txt.config(state='normal')
        ts = datetime.datetime.now().strftime('%H:%M:%S')
        self._log_txt.insert('end', f'[{ts}] {msg}\n', tag)
        self._log_txt.see('end')
        self._log_txt.config(state='disabled')

    def _clear_log(self):
        self._log_txt.config(state='normal')
        self._log_txt.delete('1.0', 'end')
        self._log_txt.config(state='disabled')

    # ── 스케줄러 자동 등록 (첫 실행 시) ──────────────────────────────────
    def _auto_register(self):
        """백그라운드 스레드에서 실행 — 메인 스레드 블로킹 방지"""
        threading.Thread(target=self._do_register, daemon=True).start()

    def _do_register(self):
        vbs = ROOT / 'run_silent.vbs'
        old_names = [
            'PT_Weekly', 'PT_Monthly',
            'PT_Weekly_MON_12', 'PT_Weekly_MON_18',
            'PT_Weekly_TUE_12', 'PT_Weekly_TUE_18',
            'PT_Weekly_12', 'PT_Weekly_18',
            'PT_Monthly_10', 'PT_Monthly_12', 'PT_Monthly_15', 'PT_Monthly_18',
            'PT_Monthly_10B', 'PT_Monthly_12B', 'PT_Monthly_15B', 'PT_Monthly_18B',
        ]
        unregister = '\n'.join(
            f"Unregister-ScheduledTask -TaskName '{n}' -Confirm:$false -ErrorAction SilentlyContinue"
            for n in old_names
        )
        vbs_ps  = str(vbs).replace("'", "''")
        exe_ps  = str(ROOT / 'ptreport.exe').replace("'", "''")
        work_ps = str(ROOT).replace("'", "''")
        # 이미 등록된 경우 스킵, 미등록이면 등록 — PowerShell 1회 호출
        ps_script = f"""if (Get-ScheduledTask -TaskName 'PT_Weekly' -ErrorAction SilentlyContinue) {{ exit 0 }}
{unregister}
$act = New-ScheduledTaskAction -Execute 'wscript.exe' -Argument ('"' + '{vbs_ps}' + '"')
$twk = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At '12:00'
$tmo = New-ScheduledTaskTrigger -Monthly -DaysOfMonth 1 -At '12:00'
$cfg = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Minutes 30) -MultipleInstances IgnoreNew
$pri = New-ScheduledTaskPrincipal -UserId $Env:USERNAME -LogonType Interactive -RunLevel Limited
Register-ScheduledTask -TaskName 'PT_Weekly'  -Action $act -Trigger $twk -Settings $cfg -Principal $pri -Force | Out-Null
Register-ScheduledTask -TaskName 'PT_Monthly' -Action $act -Trigger $tmo -Settings $cfg -Principal $pri -Force | Out-Null
$ws = New-Object -ComObject WScript.Shell
$sc = $ws.CreateShortcut([Environment]::GetFolderPath('Desktop') + '\\PT Report.lnk')
$sc.TargetPath = '{exe_ps}'
$sc.WorkingDirectory = '{work_ps}'
$sc.Description = 'PT Report - 물리치료 집계 자동화'
$sc.Save()
"""
        tmp_dir = Path(os.environ.get('TEMP') or os.environ.get('TMP') or str(Path.home()))
        ps_file = tmp_dir / f'pt_sched_{os.getpid()}.ps1'
        ok = False
        try:
            ps_file.write_text(ps_script, encoding='utf-8-sig')
            # CREATE_NO_WINDOW: windowed 앱에서 PowerShell 콘솔 창이 뜨지 않도록 억제
            result = subprocess.run(
                ['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass',
                 '-File', str(ps_file)],
                capture_output=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            ok = result.returncode == 0
            if ok:
                self._q.put(('log', 'ok', '✓ 스케줄 등록 완료'))
        except Exception as e:
            self._q.put(('log', 'err', f'스케줄 등록 오류: {e}'))
        finally:
            try:
                ps_file.unlink()
            except Exception:
                pass
        if not ok:
            self._q.put(('log', 'warn', '⚠ 스케줄 등록 실패'))

    def run(self):
        self._logwrite(
            '시작  —  data/ 폴더에 파일을 넣고 실행 버튼을 클릭하세요.',
            'info')
        self.root.mainloop()


if __name__ == '__main__':
    if '--auto' in sys.argv:
        # 스케줄러 무창 실행 — GUI 없이 run_auto() 만 실행 후 종료
        import importlib.util as _ilu
        _run_path = (BUNDLE / 'run.py') if getattr(sys, 'frozen', False) \
                    else (ROOT / 'run.py')
        _spec = _ilu.spec_from_file_location('pt_run', _run_path)
        _mod  = _ilu.module_from_spec(_spec)
        _spec.loader.exec_module(_mod)
        _mod.ROOT          = ROOT
        _mod.DATA_WEEKLY   = ROOT / 'data' / 'weekly'
        _mod.DATA_MONTHLY  = ROOT / 'data' / 'monthly'
        _mod.PROCESSED_LOG = ROOT / 'processed.log'
        _mod.LOG_FILE      = ROOT / 'run.log'
        _mod.run_auto()
    else:
        PTApp().run()
