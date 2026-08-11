# -*- coding: utf-8 -*-
"""
회사 고유 설정 로더
====================
사업부명·내부 파일명 등 회사 관련 값을 소스에서 분리해 외부 JSON으로 관리한다.
- report_config.json         : 실제 값 (git 미포함 / .gitignore)
- report_config.sample.json  : 예시 템플릿 (git 포함, 플레이스홀더)

실제 파일(report_config.json)이 있으면 그것을, 없으면 sample을 읽는다.
"""

import json
import os
import sys


def _cfg_dir():
    if getattr(sys, "frozen", False):
        return getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _load():
    d = _cfg_dir()
    for name in ("report_config.json", "report_config.sample.json"):
        p = os.path.join(d, name)
        if os.path.exists(p):
            with open(p, encoding="utf-8") as f:
                return json.load(f)
    raise FileNotFoundError(
        "report_config.json (또는 report_config.sample.json) 을 찾을 수 없습니다. "
        "report_config.sample.json 을 복사해 report_config.json 을 만들고 값을 채우세요."
    )


_cfg = _load()

EXCEL_NAME = _cfg["excel_filename"]                 # 원본 엑셀 파일명
TEMPLATE_NAME = _cfg["template_filename"]           # PPT 템플릿(지난주 PPT) 파일명
DIVISIONS_MAIN = dict(_cfg["divisions_main"])       # {사업부명: 데이터행}  (사업부별 시트)
DIVISIONS_AFFILIATE = dict(_cfg["divisions_affiliate"])  # {사업부명: 데이터행} (사업부별_계열사별 시트)
FONT_EA = _cfg.get("font_ea", "맑은 고딕")           # 제목/부제목 한글 서체
FONT_LATIN = _cfg.get("font_latin", "Arial Narrow")  # 숫자/영문 서체
