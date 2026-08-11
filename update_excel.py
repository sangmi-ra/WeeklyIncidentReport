# -*- coding: utf-8 -*-
"""
장애 주간보고 - 엑셀 데이터 업데이트 프로그램
================================================

원본 엑셀(data 폴더의 원본 파일; 파일명은 report_config.json 참조)을 바탕으로 아래 작업을 수행하고,
원본을 '기존파일명_날짜.xlsx'로 백업한 뒤 원본을 그 자리에서(인플레이스) 갱신한다.

수행 작업
  1) "요약" 시트 B1(기준일)을 지정 날짜(기본: 오늘)로 설정
       -> B2 =IF(B1="",TODAY(),B1) 가 해당 날짜로 계산됨
  2) 모든 피벗 테이블 새로고침(요약 시트 하단 피벗 차트 포함) + 전체 재계산
  3) "사업부별" / "사업부별_계열사별" 시트의 월별 장애 건수 카운트 수식을
     기준일이 속한 달까지 채운다. (없는 달은 동일 패턴 + 동일 서식으로 추가)
  4) 위 데이터를 참조하는 라인 차트들의 계열 범위를 기준일 달까지로 확장한다.

집계 기준: "데이터" 시트의 [실제시작] 월 (표의 연도=A열, 월=G열, 사업부=Q열, 계열사여부=AJ열)

요구사항: Windows + Microsoft Excel 설치 + pywin32
사용법:
    python update_excel.py                 # 기준일 = 오늘
    python update_excel.py --date 2026-08-05
    python update_excel.py --src "...xlsx" --out "...xlsx"
"""

import argparse
import datetime as dt
import os
import re
import shutil
import sys

import win32com.client as win32

import appconfig   # 회사 고유값(사업부명·파일명)은 report_config.json 에서 로드

# 엑셀 PasteSpecial 상수(서식만 붙여넣기). 지연바인딩이라 숫자 리터럴로 사용.
XL_PASTE_FORMATS = -4122

SUMMARY_SHEET = "요약"
BASE_YEAR = 2024  # 월별 카운트 표가 2024년 1월부터 시작

# ─────────────────────────────────────────────────────────────────────────────
# 시트 레이아웃 정의
#   - 각 블록은 header_row(월 라벨) 바로 아래 data_row(카운트 수식) 구조.
#   - start_col: 2024년 1월에 해당하는 컬럼 번호.
#   - extra_filter: 모든 수식에 공통으로 곱해지는 추가 조건(계열사별 시트용).
#   - year_header_row: 우측 블록 상단의 연도 라벨 행(없으면 None).
# ─────────────────────────────────────────────────────────────────────────────
def _right_blocks(divmap):
    """{사업부명: 데이터행} → [{header_row(=데이터행-1), data_row, division}] (사업부명은 config에서)"""
    return [{"header_row": row - 1, "data_row": row, "division": name}
            for name, row in divmap.items()]


SHEETS = [
    {
        "name": "사업부별",
        "left_start_col": 2,    # B
        "right_start_col": 41,  # AO
        "extra_filter": "",
        "year_header_row": 2,
        "left": {"header_row": 22, "data_row": 23},   # 전사
        "right": _right_blocks(appconfig.DIVISIONS_MAIN),
    },
    {
        "name": "사업부별_계열사별",
        "left_start_col": 1,    # A
        "right_start_col": 39,  # AM
        "extra_filter": '*(데이터!$AJ:$AJ="Y")',   # AJ='계열사여부' = Y (계열사 장애만)
        "year_header_row": None,
        "left": {"header_row": 16, "data_row": 17},   # 전사(계열사)
        "right": _right_blocks(appconfig.DIVISIONS_AFFILIATE),
    },
]

_RANGE_RE = re.compile(r"\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)")


# ─────────────────────────────────────────────────────────────────────────────
# 유틸
# ─────────────────────────────────────────────────────────────────────────────
def log(msg):
    print(msg, flush=True)


def col_letter(idx):
    s = ""
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def col_index_from_letter(letters):
    idx = 0
    for ch in letters:
        idx = idx * 26 + (ord(ch) - 64)
    return idx


def excel_serial(d):
    """날짜 -> 엑셀 날짜 일련번호(1899-12-30 기준).
    pywin32 로 datetime 을 넘길 때 발생하는 시간대(UTC) 변환 오류를 피하기 위해 사용."""
    if isinstance(d, dt.datetime):
        d = d.date()
    return (d - dt.date(1899, 12, 30)).days


def months_until(target_year, target_month):
    """(BASE_YEAR,1) 부터 (target_year,target_month) 까지의 (연,월) 목록"""
    out = []
    y, m = BASE_YEAR, 1
    while (y, m) <= (target_year, target_month):
        out.append((y, m))
        m += 1
        if m > 12:
            m, y = 1, y + 1
    return out


def col_for_month(start_col, year, month):
    """블록 시작 컬럼(2024-01 기준)에서 (year,month)에 해당하는 컬럼 번호"""
    return start_col + (year - BASE_YEAR) * 12 + (month - 1)


def make_formula(year, month, division, extra_filter):
    """월별 건수 SUMPRODUCT 배열수식 문자열 (Excel .FormulaArray 용, US-English)"""
    mm = f"{month:02d}"
    terms = f'(데이터!$A:$A={year})*(데이터!$G:$G="{mm}")'
    if division:
        terms += f'*(데이터!$Q:$Q="{division}")'
    terms += extra_filter or ""
    return f"=SUMPRODUCT({terms})"


# ─────────────────────────────────────────────────────────────────────────────
# 한 블록 채우기 (헤더 + 수식 + 신규 컬럼 서식 복사)
# ─────────────────────────────────────────────────────────────────────────────
def fill_block(ws, header_row, data_row, start_col, division, extra_filter,
               year_header_row, target_year, target_month, name):
    new_cols = []
    for (y, m) in months_until(target_year, target_month):
        c = col_for_month(start_col, y, m)
        is_new = ws.Cells(data_row, c).Formula in (None, "", 0)
        ws.Cells(header_row, c).Value = f"{m}월"
        ws.Cells(data_row, c).FormulaArray = make_formula(y, m, division, extra_filter)
        # 우측 블록 연도 라벨(1월 컬럼 상단) 보정
        if year_header_row and m == 1:
            try:
                yc = ws.Cells(year_header_row, c)
                if yc.Value in (None, ""):
                    yc.Value = y
                    # 이전 연도 라벨의 서식 복사(12칸 왼쪽)
                    prev = ws.Cells(year_header_row, c - 12)
                    prev.Copy()
                    yc.PasteSpecial(Paste=XL_PASTE_FORMATS)
            except Exception:
                pass
        if is_new:
            new_cols.append(c)

    # 신규 컬럼에 기존 월 셀의 서식을 그대로 복사 (헤더행+수식행)
    if new_cols:
        first_new, last_new = new_cols[0], new_cols[-1]
        template = first_new - 1  # 직전(기존) 월 컬럼 = 서식 원본
        try:
            src = ws.Range(ws.Cells(header_row, template), ws.Cells(data_row, template))
            dst = ws.Range(ws.Cells(header_row, first_new), ws.Cells(data_row, last_new))
            src.Copy()
            dst.PasteSpecial(Paste=XL_PASTE_FORMATS)
            fmt_note = f"(서식 원본 {col_letter(template)}열 복사)"
        except Exception as e:
            fmt_note = f"(서식 복사 실패: {e})"
        cols_txt = ", ".join(col_letter(c) for c in new_cols)
        log(f"    [{name}] 신규 월 추가: {cols_txt}  {fmt_note}")
    else:
        log(f"    [{name}] 신규 추가 없음")
    return new_cols


# ─────────────────────────────────────────────────────────────────────────────
# 차트 계열 범위 확장 (기존 시작 컬럼은 유지, 끝만 기준월까지 확장)
# ─────────────────────────────────────────────────────────────────────────────
def update_charts_for_sheet(ws, cfg, target_year, target_month):
    # 데이터행 -> (헤더행, 블록 시작 컬럼)
    rowmap = {}
    L = cfg["left"]
    rowmap[L["data_row"]] = (L["header_row"], cfg["left_start_col"])
    for b in cfg["right"]:
        rowmap[b["data_row"]] = (b["header_row"], cfg["right_start_col"])

    n = ws.ChartObjects().Count
    log(f"    차트 개수: {n}")
    for i in range(1, n + 1):
        chart = ws.ChartObjects(i).Chart
        try:
            sc = chart.SeriesCollection()
            scount = sc.Count
        except Exception:
            continue
        for si in range(1, scount + 1):
            s = sc(si)
            try:
                f = s.Formula  # =SERIES(name, xvals, yvals, order)
            except Exception:
                continue
            refs = _RANGE_RE.findall(f)
            if len(refs) < 2:
                continue
            v = refs[-1]  # 마지막 범위 = Values
            v_start = col_index_from_letter(v[0])
            v_end = col_index_from_letter(v[2])
            v_row = int(v[1])
            if v_row not in rowmap:
                continue
            header_row, block_start = rowmap[v_row]
            target_col = col_for_month(block_start, target_year, target_month)
            end_col = max(target_col, v_end)  # 차트를 줄이지는 않음
            try:
                s.XValues = ws.Range(ws.Cells(header_row, v_start), ws.Cells(header_row, end_col))
                s.Values = ws.Range(ws.Cells(v_row, v_start), ws.Cells(v_row, end_col))
                log(f"      차트{i} 계열{si}: {col_letter(v_start)}{v_row}:{col_letter(end_col)}{v_row}")
            except Exception as e:
                log(f"      차트{i} 계열{si} 범위 설정 실패: {e}")


# ─────────────────────────────────────────────────────────────────────────────
# 시트 처리
# ─────────────────────────────────────────────────────────────────────────────
def process_sheet(wb, xl, cfg, target_year, target_month):
    log(f"  [{cfg['name']}] 월별 카운트 수식/서식 갱신:")
    ws = wb.Worksheets(cfg["name"])
    L = cfg["left"]
    fill_block(ws, L["header_row"], L["data_row"], cfg["left_start_col"],
               None, cfg["extra_filter"], None, target_year, target_month, "전사")
    for b in cfg["right"]:
        fill_block(ws, b["header_row"], b["data_row"], cfg["right_start_col"],
                   b["division"], cfg["extra_filter"], cfg["year_header_row"],
                   target_year, target_month, b["division"])
    try:
        xl.CutCopyMode = False
    except Exception:
        pass
    log(f"  [{cfg['name']}] 차트 범위 갱신:")
    update_charts_for_sheet(ws, cfg, target_year, target_month)


# ─────────────────────────────────────────────────────────────────────────────
# 메인
# ─────────────────────────────────────────────────────────────────────────────
def parse_args():
    here = os.path.dirname(os.path.abspath(__file__))
    p = argparse.ArgumentParser(description="장애 주간보고 엑셀 데이터 업데이트")
    p.add_argument("--src", default=os.path.join(here, "data", appconfig.EXCEL_NAME),
                   help="원본 엑셀 경로")
    p.add_argument("--out", default=None, help="결과 저장 경로(미지정 시 output 폴더에 자동 생성)")
    p.add_argument("--date", default=None, help="기준일 (YYYY-MM-DD, 기본: 오늘)")
    p.add_argument("--outdir", default=os.path.join(here, "output"),
                   help="결과 폴더(--out 미지정 시)")
    p.add_argument("--visible", action="store_true", help="Excel 창을 보이게 실행(디버그)")
    return p.parse_args()


def main():
    args = parse_args()

    if args.date:
        base_date = dt.datetime.strptime(args.date, "%Y-%m-%d")
    else:
        t = dt.date.today()
        base_date = dt.datetime(t.year, t.month, t.day)
    target_year, target_month = base_date.year, base_date.month

    src = os.path.abspath(args.src)
    if not os.path.exists(src):
        log(f"[오류] 원본 파일을 찾을 수 없습니다: {src}")
        sys.exit(1)

    # 인플레이스 갱신(기본): 원본을 '기존파일명_날짜.xlsx'로 백업하고, 갱신 결과로 원본을 덮어쓴다.
    stamp = base_date.strftime("%Y%m%d")
    if args.out:
        target = os.path.abspath(args.out)   # 명시 지정 시 해당 경로에 저장(백업 없음)
        backup = None
    else:
        target = src
        base_name = os.path.splitext(os.path.basename(src))[0]
        backup = os.path.join(os.path.dirname(src), f"{base_name}_{stamp}.xlsx")

    log("=" * 70)
    log("장애 주간보고 - 엑셀 데이터 업데이트")
    log("=" * 70)
    log(f"  원본       : {src}")
    if backup:
        log(f"  백업 예정  : {backup}")
    log(f"  저장 대상  : {target}" + ("  (원본 덮어쓰기)" if os.path.abspath(target) == src else ""))
    log(f"  기준일     : {base_date.strftime('%Y-%m-%d')}  (대상월 {target_year}-{target_month:02d})")
    log("-" * 70)

    # 대상 파일이 다른 프로그램(Excel 등)에서 열려 있으면 저장 불가 → 미리 중단
    if os.path.exists(target):
        try:
            with open(target, "r+b"):
                pass
        except OSError:
            log("[오류] 대상 엑셀이 다른 프로그램(예: Excel)에서 열려 있어 저장할 수 없습니다.")
            log(f"       Excel에서 해당 파일을 닫고 다시 실행하세요:")
            log(f"       {target}")
            sys.exit(1)

    if backup:
        shutil.copyfile(src, backup)          # 원본 백업
        log("  백업 완료")
    if os.path.abspath(target) != src:        # 다른 경로 저장 시 원본을 복사 후 갱신
        os.makedirs(os.path.dirname(target), exist_ok=True)
        shutil.copyfile(src, target)

    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = bool(args.visible)
    xl.DisplayAlerts = False
    xl.AskToUpdateLinks = False
    wb = None
    try:
        wb = xl.Workbooks.Open(target, UpdateLinks=0)

        # 읽기전용으로 열렸으면(다른 곳에서 사용 중) 저장이 안 되므로 명확히 중단
        if getattr(wb, "ReadOnly", False):
            log("[오류] 파일이 읽기전용으로 열렸습니다(다른 프로그램에서 사용 중).")
            log(f"       Excel에서 파일을 닫고 다시 실행하세요: {target}")
            sys.exit(1)

        # (1) 기준일 설정 (일련번호로 넣어 시간대 밀림 방지)
        ws_sum = wb.Worksheets(SUMMARY_SHEET)
        ws_sum.Range("B1").Value2 = excel_serial(base_date)
        log(f"  [요약] B1(기준일) = {base_date.strftime('%Y-%m-%d')} 설정")

        # (2) 사업부별 / 사업부별_계열사별 월별 카운트 + 서식 + 차트
        for cfg in SHEETS:
            process_sheet(wb, xl, cfg, target_year, target_month)

        # (3) 피벗 새로고침 + 전체 재계산
        log("  피벗 테이블 새로고침(RefreshAll) 및 전체 재계산...")
        wb.RefreshAll()
        try:
            xl.CalculateUntilAsyncQueriesDone()
        except Exception:
            pass
        xl.CalculateFullRebuild()
        try:
            xl.CalculateUntilAsyncQueriesDone()
        except Exception:
            pass

        # 결과 확인
        try:
            b2_serial = ws_sum.Range("B2").Value2
            b2 = (dt.date(1899, 12, 30) + dt.timedelta(days=int(b2_serial))).strftime("%Y-%m-%d")
            log("-" * 70)
            log(f"  결과 확인 -> 요약 B2(기준일)={b2}, B4(올해 총건수)={ws_sum.Range('B4').Value}")
        except Exception:
            pass

        wb.Save()
        log(f"  저장 완료: {target}")
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        try:
            xl.Quit()
        except Exception:
            pass

    log("=" * 70)
    log("완료.")
    log("=" * 70)


if __name__ == "__main__":
    main()
