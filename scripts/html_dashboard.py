from __future__ import annotations
from typing import Optional, Union
"""
html_dashboard.py — HTML 대시보드 생성
process.py의 load_and_process() 결과를 받아 HTML 파일 생성
"""
import json
from pathlib import Path

def _to_js(data) -> str:
    return json.dumps(data, ensure_ascii=False)

def generate_html(data: dict, output_path: Union[str, Path]) -> Path:
    """
    집계 데이터를 받아 HTML 대시보드 생성
    
    Args:
        data: process.load_and_process() 반환값
        output_path: 저장 경로
    
    Returns:
        생성된 파일 Path
    """
    summary  = data['summary']
    daily    = data['daily']
    weekly   = data['weekly']
    all_daily   = data['all_daily']
    all_weekly  = data['all_weekly']
    docs     = data['docs']
    weeks    = data['weeks']
    period   = data['period']
    min_date = data['raw']['접수일자'].min()
    max_date = data['raw']['접수일자'].max()
    date_range = f"{min_date.strftime('%Y-%m-%d')} ~ {max_date.strftime('%Y-%m-%d')}"

    # JS 데이터 직렬화
    summary_js = _to_js([{'d':str(r['처방의']),'i':int(r['입원']),'o':int(r['외래']),'t':int(r['합계'])} for _,r in summary.iterrows()])
    docs_js    = _to_js(docs)
    weeks_js   = _to_js(weeks)

    # 처방의별 일별
    daily_js_dict = {}
    for doc in docs:
        dd = daily[daily['처방의']==doc].sort_values('접수일자')
        daily_js_dict[doc] = [
            {'dt':str(r['접수일자']),'dy':str(r['요일']),'wk':str(r['주차']),
             'i':int(r['입원']),'o':int(r['외래']),'t':int(r['합계']),
             'r':1 if r['휴일'] else 0}
            for _,r in dd.iterrows()
        ]
    daily_js = _to_js(daily_js_dict)

    # 처방의별 주별
    weekly_js_dict = {}
    for doc in docs:
        wg = weekly[weekly['처방의']==doc]
        weekly_js_dict[doc] = [
            {'wk':str(r['주차']),'i':int(r['입원']),'o':int(r['외래']),'t':int(r['합계'])}
            for _,r in wg.iterrows()
        ]
    weekly_js = _to_js(weekly_js_dict)

    # 전체
    all_daily_js = _to_js([
        {'dt':str(r['접수일자']),'dy':str(r['요일']),'wk':str(r['주차']),
         'i':int(r['입원']),'o':int(r['외래']),'t':int(r['합계']),
         'r':1 if r['휴일'] else 0}
        for _,r in all_daily.iterrows()
    ])
    all_weekly_js = _to_js([
        {'wk':str(r['주차']),'i':int(r['입원']),'o':int(r['외래']),'t':int(r['합계'])}
        for _,r in all_weekly.iterrows()
    ])

    html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0">
<title>물리치료 집계 대시보드 — {period}</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;700;900&family=DM+Mono:wght@400;500&display=swap');
:root{{--bg:#0f1923;--sur:#162030;--sur2:#1d2d42;--brd:#263548;--txt:#e8edf5;--muted:#7a90aa;--dim:#3a5068;--in:#4e9af1;--out:#f4a22d;--tot:#5bc87a;--red:#f87171;--wk-fg:#a78bfa;--wk-bg:rgba(167,139,250,.08)}}
[data-theme="light"]{{--bg:#eef2f8;--sur:#fff;--sur2:#e5ecf5;--brd:#c8d4e4;--txt:#1a2340;--muted:#6a7a98;--dim:#aab8cc;--in:#2563eb;--out:#b45309;--tot:#166534;--red:#dc2626;--wk-fg:#7c3aed;--wk-bg:rgba(124,58,237,.07)}}
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:'Noto Sans KR',sans-serif;background:var(--bg);color:var(--txt);min-height:100vh;transition:background .25s,color .25s}}
.hdr{{background:var(--sur);border-bottom:1px solid var(--brd);padding:15px 32px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100;box-shadow:0 1px 8px rgba(0,0,0,.12)}}
.hdr-l h1{{font-size:16px;font-weight:700}}.hdr-l p{{font-size:11px;color:var(--muted);margin-top:2px}}
.hdr-r{{display:flex;align-items:center;gap:8px}}
.badge{{background:var(--sur2);border:1px solid var(--brd);border-radius:6px;padding:4px 10px;font-family:'DM Mono',monospace;font-size:10px;color:var(--in)}}
.tbtn{{background:var(--sur2);border:1px solid var(--brd);border-radius:6px;padding:4px 10px;font-family:'Noto Sans KR',sans-serif;font-size:11px;color:var(--muted);cursor:pointer;transition:all .2s;white-space:nowrap}}
.tbtn:hover{{border-color:var(--in);color:var(--in)}}
.wrap{{max-width:1360px;margin:0 auto;padding:20px 32px}}
.kpis{{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:20px}}
.kpi{{background:var(--sur);border:1px solid var(--brd);border-radius:10px;padding:14px 18px;position:relative;overflow:hidden}}
.kpi::after{{content:'';position:absolute;top:0;left:0;right:0;height:3px;border-radius:10px 10px 0 0}}
.kpi.in::after{{background:var(--in)}}.kpi.out::after{{background:var(--out)}}.kpi.tot::after{{background:var(--tot)}}
.kpi-l{{font-size:9px;color:var(--muted);letter-spacing:.8px;text-transform:uppercase;margin-bottom:4px}}
.kpi-v{{font-size:30px;font-weight:900;font-family:'DM Mono',monospace;line-height:1}}
.kpi.in .kpi-v{{color:var(--in)}}.kpi.out .kpi-v{{color:var(--out)}}.kpi.tot .kpi-v{{color:var(--tot)}}
.kpi-s{{font-size:9px;color:var(--muted);margin-top:3px}}
.vtabs{{display:flex;gap:2px;background:var(--sur2);padding:3px;border-radius:8px;width:fit-content;margin-bottom:16px}}
.vt{{background:transparent;border:none;border-radius:6px;padding:6px 18px;font-family:'Noto Sans KR',sans-serif;font-size:13px;font-weight:500;color:var(--muted);cursor:pointer;transition:all .18s}}
.vt.active{{background:var(--sur);color:var(--txt);box-shadow:0 1px 4px rgba(0,0,0,.2)}}
.vp{{display:none}}.vp.active{{display:block;animation:fi .2s ease}}
@keyframes fi{{from{{opacity:0;transform:translateY(3px)}}to{{opacity:1;transform:translateY(0)}}}}
.sec{{font-size:10px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:var(--dim);margin-bottom:10px;padding-bottom:6px;border-bottom:1px solid var(--brd)}}
.cg{{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:20px}}
.cc{{background:var(--sur);border:1px solid var(--brd);border-radius:10px;padding:14px 16px}}
.cc.full{{grid-column:1/-1}}
.ct{{font-size:12px;font-weight:600;color:var(--txt);margin-bottom:10px}}
.cw{{position:relative;height:230px}}.cw.tall{{height:280px}}
.doc-area{{background:var(--sur);border:1px solid var(--brd);border-radius:10px;overflow:hidden;margin-bottom:20px}}
.dtw{{display:flex;align-items:center;border-bottom:1px solid var(--brd);background:var(--sur2);padding:0 12px;overflow-x:auto;-webkit-overflow-scrolling:touch;scrollbar-width:none}}
.dtw::-webkit-scrollbar{{display:none}}
.dtab{{background:transparent;border:none;border-bottom:2px solid transparent;padding:10px 14px;font-family:'Noto Sans KR',sans-serif;font-size:12px;font-weight:500;color:var(--muted);cursor:pointer;transition:all .15s;white-space:nowrap;flex-shrink:0}}
.dtab:hover{{color:var(--txt)}}.dtab.active{{color:var(--in);border-bottom-color:var(--in);font-weight:600}}
.dpw{{padding:16px}}
.dp{{display:none}}.dp.active{{display:block;animation:fi .18s ease}}
.tw{{overflow-x:auto;border-radius:8px;border:1px solid var(--brd)}}
table{{width:100%;border-collapse:collapse;font-size:12px;table-layout:fixed}}
thead tr{{background:var(--sur2)}}
thead th{{padding:8px 12px;font-weight:600;color:var(--muted);font-size:10px;letter-spacing:.6px;text-transform:uppercase;border-bottom:1px solid var(--brd);white-space:nowrap;text-align:center}}
th.lft{{text-align:left}}
tbody tr{{border-bottom:1px solid var(--brd);transition:background .1s}}
tbody tr:last-child{{border-bottom:none}}
tbody tr:hover:not(.wsep):not(.stot){{background:var(--sur2)!important}}
td{{padding:7px 12px;text-align:center;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}
td.lft{{text-align:left}}td.nm{{text-align:left;font-weight:600;color:var(--txt)}}
td.dt-td{{font-family:'DM Mono',monospace;font-size:11px;color:var(--muted);text-align:left}}
td.dt-td.red{{color:var(--red)!important}}td.wk-td{{font-size:11px;color:var(--wk-fg);text-align:left}}
td.n{{font-family:'DM Mono',monospace;font-size:12px;font-weight:500}}
td.ni{{color:var(--in)}}td.no{{color:var(--out)}}td.nt{{color:var(--tot);font-weight:700}}
tr.wsep td{{background:var(--wk-bg)!important;color:var(--wk-fg)!important;font-size:10px;font-weight:700;letter-spacing:.4px;padding:4px 12px}}
tr.stot td{{background:var(--sur2)!important;font-weight:700;border-top:1px solid var(--brd)}}
tr.stot td.ni{{color:var(--in)!important}}tr.stot td.no{{color:var(--out)!important}}tr.stot td.nt{{color:var(--tot)!important}}
.wcards{{display:grid;grid-template-columns:repeat(5,1fr);gap:8px;margin-bottom:14px}}
.wcard{{background:var(--sur);border:1px solid var(--brd);border-radius:8px;padding:10px 12px}}
.wcard-l{{font-size:9px;color:var(--muted);letter-spacing:.3px;margin-bottom:6px;font-weight:600;text-transform:uppercase}}
.wcard-r{{display:flex;align-items:baseline;gap:4px;margin-bottom:2px}}
.wc-tag{{font-size:9px;color:var(--muted);width:20px}}
.wc-i{{font-size:17px;font-weight:800;color:var(--in);font-family:'DM Mono',monospace;line-height:1}}
.wc-o{{font-size:17px;font-weight:800;color:var(--out);font-family:'DM Mono',monospace;line-height:1}}
.wc-div{{height:1px;background:var(--brd);margin:5px 0}}.wc-t{{font-size:11px;color:var(--muted);font-family:'DM Mono',monospace}}
.wk-tg{{display:grid;grid-template-columns:1fr 1fr;gap:12px}}
.leg{{display:flex;gap:12px;margin-bottom:8px}}.li{{display:flex;align-items:center;gap:4px;font-size:11px;color:var(--muted)}}.ld{{width:7px;height:7px;border-radius:50%}}
@media(max-width:1024px){{.wrap{{padding:16px 20px}}.hdr{{padding:13px 20px}}.cg{{grid-template-columns:1fr}}.cc.full{{grid-column:1}}.wcards{{grid-template-columns:repeat(3,1fr)}}.wk-tg{{grid-template-columns:1fr}}}}
@media(max-width:768px){{.hdr{{flex-wrap:wrap;gap:8px}}.hdr-r{{width:100%;justify-content:space-between}}.kpis{{gap:8px}}.kpi-v{{font-size:24px}}.vtabs{{width:100%}}.vt{{flex:1;padding:6px 4px;font-size:12px;text-align:center}}.wcards{{grid-template-columns:repeat(2,1fr)}}.cw{{height:190px}}.cw.tall{{height:220px}}.dtab{{padding:8px 10px;font-size:11px}}table{{font-size:11px}}td,thead th{{padding:6px 8px}}}}
@media(max-width:480px){{.kpis{{grid-template-columns:1fr}}.kpi-v{{font-size:28px}}.wcards{{grid-template-columns:1fr 1fr}}.hdr-l h1{{font-size:14px}}.wrap{{padding:12px 12px}}}}
.footer{{text-align:center;padding:14px;font-size:10px;color:var(--muted);border-top:1px solid var(--brd);margin-top:4px}}
</style>
</head>
<body>
<div class="hdr">
  <div class="hdr-l"><h1>물리치료 진료 집계 대시보드</h1><p>처방의별 입원 / 외래 현황 · 날짜별 중복 제거 기준</p></div>
  <div class="hdr-r">
    <div class="badge">{date_range}</div>
    <button class="tbtn" onclick="toggleTheme()" id="themeBtn">☀ 라이트 모드</button>
  </div>
</div>
<div class="wrap">
  <div class="kpis" id="KPI"></div>
  <div class="vtabs">
    <button class="vt active" onclick="sv('monthly',this)">월별 요약</button>
    <button class="vt" onclick="sv('weekly',this)">주별 집계</button>
    <button class="vt" onclick="sv('daily',this)">일별 상세</button>
  </div>
  <div class="vp active" id="vp-monthly">
    <div class="sec">처방의별 현황</div>
    <div class="cg">
      <div class="cc"><div class="ct">입원 / 외래 비교</div><div class="cw"><canvas id="cBar"></canvas></div></div>
      <div class="cc"><div class="ct">전체 비율</div><div class="cw"><canvas id="cDnt"></canvas></div></div>
      <div class="cc full">
        <div class="ct">처방의별 일별 추이</div>
        <div class="leg"><div class="li"><div class="ld" style="background:var(--in)"></div>입원</div><div class="li"><div class="ld" style="background:var(--out)"></div>외래</div></div>
        <div class="doc-area" style="margin-top:4px"><div class="dtw" id="trTabs"></div><div class="dpw" id="trPanels"></div></div>
      </div>
    </div>
    <div class="sec">처방의별 월 집계</div>
    <div class="tw"><table>
      <colgroup><col style="width:120px"><col style="width:80px"><col style="width:80px"><col style="width:80px"></colgroup>
      <thead><tr><th class="lft">처방의</th><th>입원</th><th>외래</th><th>합계</th></tr></thead>
      <tbody id="sTbody"></tbody>
    </table></div>
  </div>
  <div class="vp" id="vp-weekly">
    <div class="sec">주차별 전체 현황</div>
    <div class="wcards" id="wCards"></div>
    <div class="doc-area"><div class="dtw" id="wkTabs"></div><div class="dpw" id="wkPanels"></div></div>
  </div>
  <div class="vp" id="vp-daily">
    <div class="doc-area"><div class="dtw" id="dlTabs"></div><div class="dpw" id="dlPanels"></div></div>
  </div>
</div>
<div class="footer">물리치료 집계 리포트 · 날짜별 중복 제거 기준 · 빨간 날짜 = 토요일 / 공휴일 · 생성: {period}</div>
<script>
const SUMMARY={summary_js};
const DOCS={docs_js};
const WEEKS={weeks_js};
const DAILY={daily_js};
const WEEKLY={weekly_js};
const ALL_DAILY={all_daily_js};
const ALL_WEEKLY={all_weekly_js};
let dark=true;const CR={{}};
function toggleTheme(){{dark=!dark;document.documentElement.setAttribute('data-theme',dark?'':'light');document.getElementById('themeBtn').textContent=dark?'☀ 라이트 모드':'🌙 다크 모드';const mc=dark?'#7a90aa':'#6a7a98',gc=dark?'#263548':'#c8d4e4';Object.values(CR).forEach(c=>{{if(!c)return;['x','y'].forEach(ax=>{{if(c.options.scales?.[ax]){{c.options.scales[ax].ticks.color=mc;c.options.scales[ax].grid.color=gc;}}}});if(c.options.plugins?.legend)c.options.plugins.legend.labels.color=mc;c.update();}});}}
const tI=SUMMARY.reduce((a,b)=>a+b.i,0),tO=SUMMARY.reduce((a,b)=>a+b.o,0);
document.getElementById('KPI').innerHTML=`<div class="kpi in"><div class="kpi-l">총 입원</div><div class="kpi-v">${{tI.toLocaleString()}}</div><div class="kpi-s">날짜별 중복 제거</div></div><div class="kpi out"><div class="kpi-l">총 외래</div><div class="kpi-v">${{tO.toLocaleString()}}</div><div class="kpi-s">날짜별 중복 제거</div></div><div class="kpi tot"><div class="kpi-l">전체 합계</div><div class="kpi-v">${{(tI+tO).toLocaleString()}}</div><div class="kpi-s">처방의 ${{DOCS.length}}명</div></div>`;
function sv(id,btn){{document.querySelectorAll('.vt').forEach(b=>b.classList.remove('active'));document.querySelectorAll('.vp').forEach(p=>p.classList.remove('active'));btn.classList.add('active');document.getElementById('vp-'+id).classList.add('active');}}
function co(ex={{}}){{const mc=dark?'#7a90aa':'#6a7a98',gc=dark?'#263548':'#c8d4e4';return{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{labels:{{color:mc,font:{{family:'Noto Sans KR',size:11}}}}}},  ...( ex.plugins||{{}})}},scales:{{x:{{ticks:{{color:mc,font:{{family:'Noto Sans KR',size:10}}}},grid:{{color:gc}}}},y:{{ticks:{{color:mc,stepSize:1,callback:v=>Number.isInteger(v)?v:''}},grid:{{color:gc}}}}}}, ...ex}};}}
CR.bar=new Chart(document.getElementById('cBar'),{{type:'bar',data:{{labels:DOCS,datasets:[{{label:'입원',data:SUMMARY.map(d=>d.i),backgroundColor:'rgba(78,154,241,.72)',borderRadius:4}},{{label:'외래',data:SUMMARY.map(d=>d.o),backgroundColor:'rgba(244,162,45,.72)',borderRadius:4}}]}},options:co()}});
CR.dnt=new Chart(document.getElementById('cDnt'),{{type:'doughnut',data:{{labels:['입원','외래'],datasets:[{{data:[tI,tO],backgroundColor:['rgba(78,154,241,.8)','rgba(244,162,45,.8)'],borderWidth:0,hoverOffset:6}}]}},options:{{responsive:true,maintainAspectRatio:false,cutout:'65%',plugins:{{legend:{{position:'bottom',labels:{{color:'#7a90aa',font:{{family:'Noto Sans KR',size:12}},padding:14}}}},tooltip:{{callbacks:{{label:ctx=>` ${{ctx.label}}: ${{ctx.raw.toLocaleString()}}명 (${{(ctx.raw/(tI+tO)*100).toFixed(1)}}%)`}}}}}}}}}});
function mkTabs(tid,pid,items,onAct){{const te=document.getElementById(tid),pe=document.getElementById(pid);items.forEach((k,i)=>{{const btn=document.createElement('button');btn.className='dtab'+(i===0?' active':'');btn.textContent=k;btn.dataset.key=k;btn.onclick=()=>{{te.querySelectorAll('.dtab').forEach(b=>b.classList.remove('active'));btn.classList.add('active');pe.querySelectorAll('.dp').forEach(p=>p.classList.remove('active'));const t=pe.querySelector(`.dp[data-key="${{k}}"]`);if(t)t.classList.add('active');if(onAct)onAct(k);}};te.appendChild(btn);const p=document.createElement('div');p.className='dp'+(i===0?' active':'');p.dataset.key=k;pe.appendChild(p);}});}}
mkTabs('trTabs','trPanels',DOCS,renderTrend);
function renderTrend(doc){{const p=document.querySelector(`#trPanels .dp[data-key="${{doc}}"]`);if(CR['tr'+doc])return;const id='trc'+doc.replace(/[^\\w]/g,'');p.innerHTML=`<div class="cw tall"><canvas id="${{id}}"></canvas></div>`;const dd=DAILY[doc];CR['tr'+doc]=new Chart(document.getElementById(id),{{type:'bar',data:{{labels:dd.map(r=>r.dt.slice(5)+'('+r.dy+')'),datasets:[{{label:'입원',data:dd.map(r=>r.i),backgroundColor:'rgba(78,154,241,.72)',borderRadius:3,stack:'s'}},{{label:'외래',data:dd.map(r=>r.o),backgroundColor:'rgba(244,162,45,.72)',borderRadius:3,stack:'s'}}]}},options:co({{plugins:{{legend:{{display:false}}}}}})}}); }}
renderTrend(DOCS[0]);
const st=document.getElementById('sTbody');SUMMARY.forEach(d=>{{st.innerHTML+=`<tr><td class="nm">${{d.d}}</td><td class="n ni">${{d.i}}</td><td class="n no">${{d.o}}</td><td class="n nt">${{d.t}}</td></tr>`;}});st.innerHTML+=`<tr class="stot"><td class="nm">합계</td><td class="n ni">${{tI}}</td><td class="n no">${{tO}}</td><td class="n nt">${{tI+tO}}</td></tr>`;
const wc=document.getElementById('wCards');WEEKS.forEach(wk=>{{const wI=DOCS.reduce((a,d)=>{{const r=WEEKLY[d].find(x=>x.wk===wk);return a+(r?r.i:0);}},0),wO=DOCS.reduce((a,d)=>{{const r=WEEKLY[d].find(x=>x.wk===wk);return a+(r?r.o:0);}},0);wc.innerHTML+=`<div class="wcard"><div class="wcard-l">${{wk}}</div><div class="wcard-r"><span class="wc-tag">입원</span><span class="wc-i">${{wI}}</span></div><div class="wcard-r"><span class="wc-tag">외래</span><span class="wc-o">${{wO}}</span></div><div class="wc-div"></div><div class="wc-t">총 ${{wI+wO}}</div></div>`;}});
mkTabs('wkTabs','wkPanels',['전체',...DOCS],k=>{{if(k==='전체'&&!CR.allWk)renderWkAll();}});
function renderWkAll(){{const p=document.querySelector('#wkPanels .dp[data-key="전체"]');const sI=ALL_WEEKLY.reduce((a,b)=>a+b.i,0),sO=ALL_WEEKLY.reduce((a,b)=>a+b.o,0);let rows=ALL_WEEKLY.map(r=>`<tr><td class="wk-td">${{r.wk}}</td><td class="n ni">${{r.i}}</td><td class="n no">${{r.o}}</td><td class="n nt">${{r.t}}</td></tr>`).join('');rows+=`<tr class="stot"><td class="nm">합계</td><td class="n ni">${{sI}}</td><td class="n no">${{sO}}</td><td class="n nt">${{sI+sO}}</td></tr>`;p.innerHTML=`<div class="wk-tg"><div class="cc" style="margin:0"><div class="ct">주차별 입원/외래 추이</div><div class="cw"><canvas id="cAllWk"></canvas></div></div><div class="tw" style="align-self:start"><table><colgroup><col style="width:130px"><col style="width:80px"><col style="width:80px"><col style="width:80px"></colgroup><thead><tr><th class="lft">주차</th><th>입원</th><th>외래</th><th>합계</th></tr></thead><tbody>${{rows}}</tbody></table></div></div>`;const mc=dark?'#7a90aa':'#6a7a98',gc=dark?'#263548':'#c8d4e4';CR.allWk=new Chart(document.getElementById('cAllWk'),{{type:'bar',data:{{labels:ALL_WEEKLY.map(r=>r.wk),datasets:[{{label:'입원',data:ALL_WEEKLY.map(r=>r.i),backgroundColor:'rgba(78,154,241,.72)',borderRadius:4,stack:'s'}},{{label:'외래',data:ALL_WEEKLY.map(r=>r.o),backgroundColor:'rgba(244,162,45,.72)',borderRadius:4,stack:'s'}}]}},options:co()}});}}
renderWkAll();
DOCS.forEach(doc=>{{const p=document.querySelector(`#wkPanels .dp[data-key="${{doc}}"]`);const wd=WEEKLY[doc]||[],sI=wd.reduce((a,b)=>a+b.i,0),sO=wd.reduce((a,b)=>a+b.o,0);let rows=wd.map(r=>`<tr><td class="wk-td">${{r.wk}}</td><td class="n ni">${{r.i}}</td><td class="n no">${{r.o}}</td><td class="n nt">${{r.t}}</td></tr>`).join('');rows+=`<tr class="stot"><td class="nm">합계</td><td class="n ni">${{sI}}</td><td class="n no">${{sO}}</td><td class="n nt">${{sI+sO}}</td></tr>`;p.innerHTML=`<div class="tw"><table><colgroup><col style="width:130px"><col style="width:80px"><col style="width:80px"><col style="width:80px"></colgroup><thead><tr><th class="lft">주차</th><th>입원</th><th>외래</th><th>합계</th></tr></thead><tbody>${{rows}}</tbody></table></div>`;}});
mkTabs('dlTabs','dlPanels',['전체',...DOCS],null);
(function(){{const p=document.querySelector('#dlPanels .dp[data-key="전체"]');let rows='',pw='';const sI=ALL_DAILY.reduce((a,b)=>a+b.i,0),sO=ALL_DAILY.reduce((a,b)=>a+b.o,0);ALL_DAILY.forEach(r=>{{if(r.wk!==pw){{rows+=`<tr class="wsep"><td colspan="4">${{r.wk}}</td></tr>`;pw=r.wk;}}rows+=`<tr><td class="dt-td${{r.r?' red':''}}">${{r.dt}} (${{r.dy}})</td><td class="n ni">${{r.i}}</td><td class="n no">${{r.o}}</td><td class="n nt">${{r.t}}</td></tr>`;}});rows+=`<tr class="stot"><td class="nm">합계</td><td class="n ni">${{sI}}</td><td class="n no">${{sO}}</td><td class="n nt">${{sI+sO}}</td></tr>`;p.innerHTML=`<div class="tw"><table><colgroup><col style="width:170px"><col style="width:80px"><col style="width:80px"><col style="width:80px"></colgroup><thead><tr><th class="lft">날짜</th><th>입원</th><th>외래</th><th>합계</th></tr></thead><tbody>${{rows}}</tbody></table></div>`;}})();
DOCS.forEach(doc=>{{const p=document.querySelector(`#dlPanels .dp[data-key="${{doc}}"]`);const dd=DAILY[doc];let rows='',pw='';const sI=dd.reduce((a,b)=>a+b.i,0),sO=dd.reduce((a,b)=>a+b.o,0);dd.forEach(r=>{{if(r.wk!==pw){{rows+=`<tr class="wsep"><td colspan="4">${{r.wk}}</td></tr>`;pw=r.wk;}}rows+=`<tr><td class="dt-td${{r.r?' red':''}}">${{r.dt}} (${{r.dy}})</td><td class="n ni">${{r.i}}</td><td class="n no">${{r.o}}</td><td class="n nt">${{r.t}}</td></tr>`;}});rows+=`<tr class="stot"><td class="nm">합계</td><td class="n ni">${{sI}}</td><td class="n no">${{sO}}</td><td class="n nt">${{sI+sO}}</td></tr>`;p.innerHTML=`<div class="tw"><table><colgroup><col style="width:170px"><col style="width:80px"><col style="width:80px"><col style="width:80px"></colgroup><thead><tr><th class="lft">날짜</th><th>입원</th><th>외래</th><th>합계</th></tr></thead><tbody>${{rows}}</tbody></table></div>`;}});
</script>
</body>
</html>"""

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"[HTML] 저장 완료: {output_path}")
    return output_path