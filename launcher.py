# -*- coding: utf-8 -*-
"""
장애 주간보고 생성기 - 로컬 웹 런처
====================================
표준 라이브러리만 사용(추가 설치 없음). 내 PC 안에서만(localhost) 도는 UI.
- 엑셀 업데이트: 원본을 '기존파일명_날짜.xlsx'로 백업하고 원본을 인플레이스 갱신
- PPT 생성: 지난주 장애공유 PPT를 템플릿 겸 비교본으로 사용, 결과는 지난주 PPT 폴더에 생성
- 파일/폴더 선택: Windows 네이티브 대화상자(PowerShell WinForms) 호출

실행: 장애보고_실행.bat  (또는  <venv>\python.exe launcher.py)
"""

import base64
import datetime as dt
import json
import os
import subprocess
import sys
import threading
import time
import urllib.parse
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import appconfig         # 회사 고유값(파일명 등)은 report_config.json 에서 로드
import update_excel      # noqa: 프리징(exe) 시 함께 번들 + --child 모드에서 직접 호출
import make_ppt          # noqa

FROZEN = getattr(sys, "frozen", False)
BASE_DIR = os.path.dirname(sys.executable) if FROZEN else os.path.dirname(os.path.abspath(__file__))


def _child_cmd(mode):
    """하위 작업 실행 커맨드. exe로 프리징되면 exe 자신을 --child 로 재실행."""
    if FROZEN:
        return [sys.executable, "--child", mode]
    return [sys.executable, os.path.abspath(__file__), "--child", mode]

DEFAULTS = {
    "date": dt.date.today().strftime("%Y-%m-%d"),
    "src": os.path.join(BASE_DIR, "data", appconfig.EXCEL_NAME),
    "prev": os.path.join(BASE_DIR, appconfig.TEMPLATE_NAME),
}

_run_lock = threading.Lock()   # 중복 실행 방지
_last_ping = time.time()       # 웹 페이지 heartbeat 마지막 수신 시각
_bye_at = 0.0                  # 탭 닫힘(pagehide) 신호 수신 시각
_HEARTBEAT_TIMEOUT = 90        # 이 시간(초) 동안 ping 이 없으면(백그라운드 등) 안전망 종료

PAGE = r"""<!DOCTYPE html>
<html lang="ko"><head><meta charset="utf-8"><title>장애 주간보고 생성기</title>
<style>
  :root{--card:#fff;--line:#e2e8f0;--ink:#0f172a;--sub:#64748b;--brand:#1e40af;
        --brand2:#2563eb;--accent:#f1f5f9;}
  *{box-sizing:border-box} body{margin:0;font-family:'Segoe UI','Malgun Gothic',sans-serif;
    background:#f8fafc;color:var(--ink);}
  .wrap{max-width:900px;margin:0 auto;padding:28px 20px 60px;}
  h1{font-size:22px;margin:0 0 4px;} .desc{color:var(--sub);font-size:13px;margin:0 0 20px;}
  .card{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:22px;
        box-shadow:0 1px 3px rgba(0,0,0,.05);margin-bottom:18px;}
  label{display:block;font-size:13px;font-weight:600;margin:14px 0 6px;}
  .row{display:flex;gap:8px;}
  input[type=text],input[type=date]{flex:1;padding:9px 11px;border:1px solid var(--line);
    border-radius:9px;font-size:13px;font-family:inherit;color:var(--ink);}
  input:focus{outline:2px solid #bfdbfe;border-color:var(--brand2);}
  .browse{padding:9px 12px;border:1px solid var(--line);background:var(--accent);
    border-radius:9px;font-size:12px;cursor:pointer;white-space:nowrap;} .browse:hover{background:#e2e8f0;}
  .btns{display:flex;gap:10px;flex-wrap:wrap;margin-top:22px;}
  button.run{border:0;border-radius:10px;padding:12px 18px;font-size:14px;font-weight:700;
    cursor:pointer;color:#fff;} .b1{background:var(--brand);} .b2{background:var(--brand2);} .ball{background:#0e7490;}
  button.run:hover{filter:brightness(1.08);} button.run:disabled{opacity:.5;cursor:not-allowed;}
  .ghost{background:#fff;color:var(--sub);border:1px solid var(--line);border-radius:10px;
    padding:10px 16px;font-size:13px;cursor:pointer;} .ghost:hover{background:#f1f5f9;}
  #log{background:#0b1020;color:#d1e0ff;font-family:Consolas,monospace;font-size:12.5px;line-height:1.55;
    border-radius:12px;padding:16px;height:340px;overflow:auto;white-space:pre-wrap;border:1px solid #1e293b;}
  #log .err{color:#fca5a5;} #log .ok{color:#86efac;} #log .hd{color:#93c5fd;}
  .status{font-size:13px;margin:10px 0;min-height:20px;}
  .results{display:flex;gap:10px;flex-wrap:wrap;margin-top:10px;}
  .results a{font-size:13px;background:#ecfdf5;border:1px solid #a7f3d0;color:#065f46;padding:9px 13px;
    border-radius:9px;text-decoration:none;cursor:pointer;}
  .hint{font-size:12px;color:var(--sub);margin-top:6px;}
  .topbar{display:flex;justify-content:space-between;align-items:center;}
</style></head>
<body><div class="wrap">
  <div class="topbar">
    <div><h1>📊 장애 주간보고 생성기</h1>
    <p class="desc">기준일과 파일을 지정하고 실행하세요. 내 PC에서만 동작합니다.</p></div>
    <button class="ghost" onclick="quitApp()">■ 종료</button>
  </div>

  <div class="card">
    <label>기준일</label>
    <div class="row"><input type="date" id="date"></div>
    <div class="hint" id="xlnote"></div>

    <label>원본 엑셀 (데이터 시트) — 갱신 시 같은 폴더에 '_날짜' 백업 후 원본을 덮어씁니다</label>
    <div class="row"><input type="text" id="src"><button class="browse" onclick="browse('src','file','excel')">찾아보기</button></div>

    <label>지난주 장애공유 PPT — 템플릿 겸 비교본. 이번주 PPT는 이 파일 폴더에 생성됩니다</label>
    <div class="row"><input type="text" id="prev"><button class="browse" onclick="browse('prev','file','pptx')">찾아보기</button></div>

    <div class="btns">
      <button class="run b1" id="btnExcel" onclick="run('excel')">① 엑셀 업데이트</button>
      <button class="run b2" id="btnPpt" onclick="run('ppt')">② PPT 생성</button>
      <button class="run ball" id="btnBoth" onclick="run('both')">전체 실행 (①→②)</button>
    </div>
    <div class="hint">전체 실행: 엑셀을 갱신한 뒤 그 결과로 PPT까지 한 번에 생성합니다.</div>
  </div>

  <div class="card">
    <div class="status" id="status">대기 중…</div>
    <div id="log"></div>
    <div class="results" id="results"></div>
  </div>
</div>

<script>
const $=id=>document.getElementById(id);
let lastExcelDate='';
async function loadDefaults(){
  const d=await (await fetch('/defaults')).json();
  for(const k of ['date','src','prev']) $(k).value=d[k]||'';
  refreshExcelDate();
}
async function refreshExcelDate(){
  // 엑셀 기준일은 '표시/비교용'으로만 읽는다. 날짜 입력칸은 오늘 날짜 기본을 유지(덮어쓰지 않음).
  const src=$('src').value;
  if(!src){ lastExcelDate=''; $('xlnote').textContent=''; return; }
  try{
    const r=await (await fetch('/exceldate?src='+encodeURIComponent(src))).json();
    lastExcelDate=r.date||'';
    $('xlnote').textContent = lastExcelDate ? ('현재 엑셀 기준일: '+lastExcelDate) : '';
  }catch(e){ lastExcelDate=''; $('xlnote').textContent=''; }
}
async function browse(target,kind,flt){
  const cur=encodeURIComponent($(target).value||'');
  try{
    const r=await (await fetch(`/browse?kind=${kind}&flt=${flt}&cur=${cur}`)).json();
    if(r.path){
      $(target).value=r.path;
      $(target).dispatchEvent(new Event('input',{bubbles:true}));  // 리페인트 유도
      try{ window.focus(); }catch(e){}
      if(target==='src') refreshExcelDate();
    }
  }catch(e){ alert('파일 선택 창을 열 수 없습니다. 경로를 직접 입력하세요.'); }
}
function setBusy(b){ for(const id of ['btnExcel','btnPpt','btnBoth']) $(id).disabled=b; }
let es=null;
function run(mode){
  // PPT만 생성인데 화면 기준일이 엑셀 기준일과 다르면: 엑셀을 먼저 그 날짜로 갱신할지 물어봄
  if(mode==='ppt' && lastExcelDate && $('date').value && $('date').value!==lastExcelDate){
    const ok=confirm('선택한 기준일('+$('date').value+')로 엑셀을 먼저 업데이트한 뒤 PPT를 생성할까요?\n\n'
      +'[확인] 엑셀 업데이트('+$('date').value+') → PPT 생성\n'
      +'[취소] 엑셀 기준일('+lastExcelDate+') 그대로 PPT만 생성');
    if(ok){ mode='both'; }
    else { alert('엑셀 기준일('+lastExcelDate+')에 맞춰 PPT를 생성합니다.'); $('date').value=lastExcelDate; }
  }
  const p=new URLSearchParams({mode,date:$('date').value,src:$('src').value,prev:$('prev').value});
  $('log').innerHTML=''; $('results').innerHTML=''; setBusy(true); $('status').textContent='실행 중…';
  es=new EventSource('/run?'+p.toString());
  es.onmessage=e=>addLog(e.data);
  es.addEventListener('done',e=>{
    const r=JSON.parse(e.data); es.close(); setBusy(false);
    if(r.status==='ok'){ $('status').innerHTML='✅ 완료되었습니다.'; refreshExcelDate();
      const res=$('results'); res.innerHTML='';
      if(r.excel) res.appendChild(link('📗 엑셀 열기',r.excel));
      if(r.ppt) res.appendChild(link('📕 PPT 열기',r.ppt));
      if(r.excel_dir) res.appendChild(link('📁 엑셀 폴더',r.excel_dir));
      if(r.ppt_dir && r.ppt_dir!==r.excel_dir) res.appendChild(link('📁 PPT 폴더',r.ppt_dir));
    } else $('status').innerHTML='❌ 오류가 발생했습니다. 로그를 확인하세요.';
  });
  es.onerror=()=>{ if(es) es.close(); setBusy(false); };
}
function addLog(line){
  const d=document.createElement('div');
  if(/오류|ERROR|Traceback|Permission/i.test(line)) d.className='err';
  else if(/완료|저장 완료|✅/.test(line)) d.className='ok';
  else if(/────|===|\[슬라이드|\[요약|\[사업부|\[전체/.test(line)) d.className='hd';
  d.textContent=line; const log=$('log'); log.appendChild(d); log.scrollTop=log.scrollHeight;
}
function link(text,path){ const a=document.createElement('a'); a.textContent=text;
  a.onclick=()=>fetch('/open?path='+encodeURIComponent(path)); return a; }
function quitApp(){ if(confirm('런처를 종료할까요?')){ fetch('/quit');
  document.body.innerHTML='<div class="wrap"><h1>종료되었습니다.</h1><p class="desc">이 탭을 닫으셔도 됩니다.</p></div>'; } }
// heartbeat: 탭이 열려있는 동안 주기적으로 ping. 탭을 닫으면 pagehide로 /bye → 런처가 곧 종료.
// (새로고침 시엔 재접속하며 ping 이 와서 종료 취소됨)
fetch('/ping').catch(()=>{});
setInterval(()=>{ fetch('/ping').catch(()=>{}); }, 5000);
window.addEventListener('pagehide',()=>{ try{ fetch('/bye',{keepalive:true}); }catch(e){} });
loadDefaults();
</script>
</body></html>"""


def _ctypes_open_file(flt, initial_dir):
    """Windows '열기' 대화상자를 Python 내장 ctypes(comdlg32.GetOpenFileNameW)로 직접 호출.
    별도 프로세스/Add-Type 없이 in-process 로 떠서 매우 빠르다. 실패 시 예외 → 호출부에서 PS 폴백.
    hwndOwner 를 현재 포그라운드(=브라우저)로 주어 대화상자가 앞에 뜨고, 닫힐 때 포커스가 브라우저로 복귀."""
    import ctypes
    from ctypes import wintypes
    u32 = ctypes.windll.user32
    comdlg32 = ctypes.windll.comdlg32
    u32.GetForegroundWindow.restype = wintypes.HWND
    comdlg32.GetOpenFileNameW.restype = wintypes.BOOL

    filters = {
        "excel": "Excel 파일\0*.xlsx;*.xlsm\0모든 파일\0*.*\0\0",
        "pptx": "PowerPoint\0*.pptx;*.potx\0모든 파일\0*.*\0\0",
    }
    filt = filters.get(flt, "모든 파일\0*.*\0\0")

    class OFN(ctypes.Structure):
        _fields_ = [
            ("lStructSize", wintypes.DWORD), ("hwndOwner", wintypes.HWND),
            ("hInstance", wintypes.HINSTANCE), ("lpstrFilter", wintypes.LPCWSTR),
            ("lpstrCustomFilter", wintypes.LPWSTR), ("nMaxCustFilter", wintypes.DWORD),
            ("nFilterIndex", wintypes.DWORD), ("lpstrFile", wintypes.LPWSTR),
            ("nMaxFile", wintypes.DWORD), ("lpstrFileTitle", wintypes.LPWSTR),
            ("nMaxFileTitle", wintypes.DWORD), ("lpstrInitialDir", wintypes.LPCWSTR),
            ("lpstrTitle", wintypes.LPCWSTR), ("Flags", wintypes.DWORD),
            ("nFileOffset", wintypes.WORD), ("nFileExtension", wintypes.WORD),
            ("lpstrDefExt", wintypes.LPCWSTR), ("lCustData", ctypes.c_void_p),
            ("lpfnHook", ctypes.c_void_p), ("lpTemplateName", wintypes.LPCWSTR),
            ("pvReserved", ctypes.c_void_p), ("dwReserved", wintypes.DWORD),
            ("FlagsEx", wintypes.DWORD),
        ]

    buf = ctypes.create_unicode_buffer(2048)
    ofn = OFN()
    ofn.lStructSize = ctypes.sizeof(OFN)
    try:
        ofn.hwndOwner = u32.GetForegroundWindow()
    except Exception:
        pass
    ofn.lpstrFilter = filt
    ofn.lpstrFile = ctypes.cast(buf, wintypes.LPWSTR)
    ofn.nMaxFile = 2048
    if initial_dir:
        ofn.lpstrInitialDir = initial_dir
    ofn.Flags = 0x00080000 | 0x00001000 | 0x00000800 | 0x00000008  # EXPLORER|FILEMUSTEXIST|PATHMUSTEXIST|NOCHANGEDIR
    ok = comdlg32.GetOpenFileNameW(ctypes.byref(ofn))
    if ok:
        return buf.value
    err = comdlg32.CommDlgExtendedError()
    if err:
        raise OSError("GetOpenFileNameW error 0x%x" % err)
    return ""     # 사용자 취소


def _native_browse(kind, flt, cur):
    """파일/폴더 선택. 파일은 ctypes 내장 대화상자(빠름) 우선, 실패 시 PowerShell 폴백. 폴더는 PowerShell."""
    initial = cur or ""
    if initial and os.path.isfile(initial):
        initial = os.path.dirname(initial)
    if not initial or not os.path.isdir(initial):
        initial = BASE_DIR
    if os.name == "nt" and kind != "folder":
        try:
            path = _ctypes_open_file(flt, initial)
            if path == "":
                return ""                       # 사용자 취소
            if os.path.exists(path):
                return path
        except Exception:
            pass                                # ctypes 실패 → PowerShell 폴백
    return _ps_browse(kind, flt, initial)


def _ps_browse(kind, flt, initial):
    """PowerShell WinForms 대화상자(폴백/폴더용). 느리지만 확실."""
    ini = initial.replace("'", "''")
    if kind == "folder":
        body = (f"$d = New-Object System.Windows.Forms.FolderBrowserDialog\n"
                f"$d.SelectedPath = '{ini}'\n"
                f"$res = $d.ShowDialog($owner)\n"
                f"if($res -eq [System.Windows.Forms.DialogResult]::OK){{ [Console]::Out.Write($d.SelectedPath) }}\n")
    else:
        filt = {
            "excel": "Excel 파일 (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|모든 파일 (*.*)|*.*",
            "pptx": "PowerPoint (*.pptx;*.potx)|*.pptx;*.potx|모든 파일 (*.*)|*.*",
        }.get(flt, "모든 파일 (*.*)|*.*")
        body = (f"$d = New-Object System.Windows.Forms.OpenFileDialog\n"
                f"$d.InitialDirectory = '{ini}'\n"
                f"$d.Filter = '{filt}'\n"
                f"$res = $d.ShowDialog($owner)\n"
                f"if($res -eq [System.Windows.Forms.DialogResult]::OK){{ [Console]::Out.Write($d.FileName) }}\n")
    # 대화상자를 최상위/포그라운드로 강제(ALT 키 트릭 + SetForegroundWindow)해 뒤에 숨지 않게 함
    fg = (
        'Add-Type @"\n'
        'using System;\n'
        'using System.Runtime.InteropServices;\n'
        'public class FG {\n'
        '  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);\n'
        '  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();\n'
        '  [DllImport("user32.dll")] public static extern void keybd_event(byte b, byte s, uint f, IntPtr e);\n'
        '}\n'
        '"@\n'
    )
    script = (
        "$ProgressPreference = 'SilentlyContinue'\n"
        "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8\n"
        "Add-Type -AssemblyName System.Windows.Forms | Out-Null\n"
        + fg +
        "$prev = [FG]::GetForegroundWindow()\n"     # 대화상자 열기 전 포그라운드(=브라우저) 저장
        "$owner = New-Object System.Windows.Forms.Form\n"
        "$owner.TopMost = $true; $owner.ShowInTaskbar = $false; $owner.Opacity = 0\n"
        "$owner.StartPosition = 'CenterScreen'\n"
        "$owner.Show(); $owner.Activate()\n"
        "[FG]::keybd_event(0x12,0,0,[IntPtr]::Zero); [FG]::keybd_event(0x12,0,2,[IntPtr]::Zero)\n"
        "[FG]::SetForegroundWindow($owner.Handle) | Out-Null\n"
        + body
        + "$owner.Close()\n"
        # 대화상자가 가져갔던 포커스를 브라우저로 복귀 → 선택 즉시 화면 반영(리페인트)
        "[FG]::keybd_event(0x12,0,0,[IntPtr]::Zero); [FG]::keybd_event(0x12,0,2,[IntPtr]::Zero)\n"
        "if($prev -ne [IntPtr]::Zero){ [FG]::SetForegroundWindow($prev) | Out-Null }\n"
    )
    enc = base64.b64encode(script.encode("utf-16-le")).decode()
    # CREATE_NO_WINDOW: 콘솔 없는(windowed) 부모에서 powershell 이 새 콘솔을 할당하지 않게 함
    # (이게 없으면 콘솔 상속 실패로 매번 콘솔 생성 → 파일선택이 수 초씩 느려짐)
    flags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
    try:
        r = subprocess.run(["powershell", "-NoProfile", "-Sta", "-EncodedCommand", enc],
                           capture_output=True, text=True, encoding="utf-8",
                           errors="replace", timeout=300, creationflags=flags)
        return (r.stdout or "").strip()
    except Exception:
        return ""


def _cmd_paths(qs):
    date = qs.get("date", DEFAULTS["date"])
    src = os.path.abspath(qs.get("src", DEFAULTS["src"]))
    prev = os.path.abspath(qs.get("prev", DEFAULTS["prev"]))
    stamp = date.replace("-", "")
    src_dir = os.path.dirname(src)
    prev_dir = os.path.dirname(prev)
    base = os.path.splitext(os.path.basename(src))[0]
    backup = os.path.join(src_dir, f"{base}_{stamp}.xlsx")
    out_ppt = os.path.join(prev_dir, f"{stamp}_장애공유.pptx")
    excel_cmd = _child_cmd("excel") + ["--src", src, "--date", date]     # 인플레이스(백업 후 원본 덮어쓰기)
    ppt_cmd = _child_cmd("ppt") + ["--xlsx", src, "--prev", prev,
                                   "--date", date, "--out", out_ppt]     # 템플릿=prev
    return {"date": date, "src": src, "prev": prev, "backup": backup,
            "out_ppt": out_ppt, "src_dir": src_dir, "prev_dir": prev_dir,
            "excel_cmd": excel_cmd, "ppt_cmd": ppt_cmd}


def _run_stream(cmd, emit):
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"
    flags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0   # 하위 콘솔창 깜빡임 방지
    try:
        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                text=True, encoding="utf-8", errors="replace",
                                bufsize=1, env=env, cwd=BASE_DIR, creationflags=flags)
    except Exception as e:
        emit(f"[오류] 실행 시작 실패: {e}")
        return False
    for line in proc.stdout:
        emit(line.rstrip("\r\n"))
    proc.wait()
    return proc.returncode == 0


def _heartbeat_watchdog():
    """웹 페이지가 살아있는 동안 /ping 을 주기적으로 보낸다.
    - 탭을 닫으면 pagehide 로 /bye 가 오고, 잠깐 뒤 새 ping 이 없으면 즉시 종료(새로고침은 재연결되어 생존).
    - /bye 를 못 받아도 ping 이 오래 끊기면(백그라운드 스로틀 고려) 안전망으로 종료."""
    while True:
        time.sleep(2)
        now = time.time()
        if _bye_at and now - _bye_at > 3 and _last_ping <= _bye_at:
            os._exit(0)          # 탭 닫힘 → 빠른 종료
        if now - _last_ping > _HEARTBEAT_TIMEOUT:
            os._exit(0)          # 안전망


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _send(self, code, ctype, body):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.wfile.write(body)

    def do_GET(self):
        u = urllib.parse.urlparse(self.path)
        qs = {k: v[0] for k, v in urllib.parse.parse_qs(u.query).items()}
        if u.path == "/":
            self._send(200, "text/html; charset=utf-8", PAGE)
        elif u.path == "/defaults":
            self._send(200, "application/json", json.dumps(DEFAULTS))
        elif u.path == "/browse":
            path = _native_browse(qs.get("kind", "file"), qs.get("flt", ""), qs.get("cur", ""))
            self._send(200, "application/json", json.dumps({"path": path}))
        elif u.path == "/exceldate":
            src = os.path.abspath(qs.get("src", "") or DEFAULTS["src"])
            d = make_ppt.read_basedate(src) if os.path.exists(src) else None
            self._send(200, "application/json",
                       json.dumps({"date": d.strftime("%Y-%m-%d") if d else ""}))
        elif u.path == "/open":
            p = qs.get("path", "")
            try:
                if p and os.path.exists(p):
                    os.startfile(p)  # noqa (Windows)
                    self._send(200, "application/json", json.dumps({"ok": True}))
                else:
                    self._send(404, "application/json", json.dumps({"ok": False}))
            except Exception as e:
                self._send(500, "application/json", json.dumps({"ok": False, "err": str(e)}))
        elif u.path == "/run":
            self._run(qs)
        elif u.path == "/ping":
            global _last_ping, _bye_at
            _last_ping = time.time()
            _bye_at = 0.0          # 살아있음 → 닫힘 신호 취소(새로고침 대비)
            self._send(200, "application/json", json.dumps({"ok": True}))
        elif u.path == "/bye":
            globals()["_bye_at"] = time.time()   # 탭 닫힘 신호
            self._send(200, "application/json", json.dumps({"ok": True}))
        elif u.path == "/quit":
            self._send(200, "application/json", json.dumps({"ok": True}))
            threading.Timer(0.4, lambda: os._exit(0)).start()
        else:
            self._send(404, "text/plain", "not found")

    def _run(self, qs):
        self.send_response(200)
        self.send_header("Content-Type", "text/event-stream; charset=utf-8")
        self.send_header("Cache-Control", "no-cache")
        self.end_headers()

        def emit(line):
            self.wfile.write(f"data: {line}\n\n".encode("utf-8"))
            self.wfile.flush()

        c = _cmd_paths(qs)
        mode = qs.get("mode", "both")
        result = {"status": "error", "excel": None, "ppt": None,
                  "excel_dir": c["src_dir"], "ppt_dir": c["prev_dir"]}

        if not _run_lock.acquire(blocking=False):
            try:
                emit("[알림] 이미 실행 중입니다. 완료 후 다시 시도하세요.")
                self.wfile.write(f"event: done\ndata: {json.dumps(result)}\n\n".encode("utf-8"))
                self.wfile.flush()
            except Exception:
                pass
            return
        try:
            ok = True
            if not os.path.exists(c["src"]):
                emit(f"[오류] 원본 엑셀을 찾을 수 없습니다: {c['src']}"); ok = False
            if ok and mode in ("ppt", "both") and not os.path.exists(c["prev"]):
                emit(f"[오류] 지난주 PPT를 찾을 수 없습니다: {c['prev']}"); ok = False

            if ok and mode in ("excel", "both"):
                emit("──────── ① 엑셀 업데이트 시작 ────────")
                ok = _run_stream(c["excel_cmd"], emit)
                if ok:
                    result["excel"] = c["src"]

            if ok and mode in ("ppt", "both"):
                emit("──────── ② PPT 생성 시작 ────────")
                ok = _run_stream(c["ppt_cmd"], emit)
                if ok:
                    result["ppt"] = c["out_ppt"]

            result["status"] = "ok" if ok else "error"
        except Exception as e:
            try:
                emit(f"[오류] {e}")
            except Exception:
                pass
        finally:
            _run_lock.release()
        try:
            self.wfile.write(f"event: done\ndata: {json.dumps(result)}\n\n".encode("utf-8"))
            self.wfile.flush()
        except Exception:
            pass


def main():
    # 하위 작업 모드: exe/py 자신을 재실행해 update_excel / make_ppt 를 직접 수행(로그는 stdout으로 스트리밍)
    if len(sys.argv) >= 2 and sys.argv[1] == "--child":
        # windowed(--noconsole) exe에서는 sys.stdout/err 가 None → 부모가 준 PIPE(fd 1/2)에 재연결.
        # 그 외엔 UTF-8로 강제(cp949 방지) → 부모가 UTF-8로 읽어 한글 로그 정상.
        for _fd, _name in ((1, "stdout"), (2, "stderr")):
            try:
                _s = getattr(sys, _name, None)
                if _s is None:
                    setattr(sys, _name, os.fdopen(_fd, "w", encoding="utf-8",
                                                  errors="replace", buffering=1))
                else:
                    _s.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass
        mode = sys.argv[2] if len(sys.argv) >= 3 else ""
        rest = sys.argv[3:]
        if mode == "excel":
            sys.argv = ["update_excel"] + rest
            update_excel.main()
        elif mode == "ppt":
            sys.argv = ["make_ppt"] + rest
            make_ppt.main()
        return

    # windowed(--noconsole) exe: 부모(서버) 프로세스는 stdout 이 None → print() 안전하게 무시
    if sys.stdout is None or sys.stderr is None:
        class _Null:
            def write(self, *a):
                return 0
            def flush(self):
                pass
        if sys.stdout is None:
            sys.stdout = _Null()
        if sys.stderr is None:
            sys.stderr = _Null()

    no_ui = "--no-ui" in sys.argv
    port_arg = 0
    if "--port" in sys.argv:
        try:
            port_arg = int(sys.argv[sys.argv.index("--port") + 1])
        except Exception:
            port_arg = 0
    httpd = ThreadingHTTPServer(("127.0.0.1", port_arg), Handler)
    url = f"http://127.0.0.1:{httpd.server_address[1]}/"
    print("=" * 60)
    print("  장애 주간보고 생성기 실행 중")
    print(f"  브라우저: {url}")
    print("  (웹페이지 '종료' 또는 브라우저 탭을 닫으면 끝납니다)")
    print("=" * 60)
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    if not no_ui:
        global _last_ping
        _last_ping = time.time()
        # 브라우저 탭이 닫히면(ping 끊김) 스스로 종료
        threading.Thread(target=_heartbeat_watchdog, daemon=True).start()
        time.sleep(0.4)
        try:
            webbrowser.open(url)
        except Exception:
            pass
    threading.Event().wait()


if __name__ == "__main__":
    main()
