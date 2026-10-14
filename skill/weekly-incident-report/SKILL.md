---
name: weekly-incident-report
description: 기존 WeeklyIncidentReport 코드와 회사 보고서 양식을 사용해 주간 장애보고 Excel을 갱신하고 지난주 PowerPoint에서 이번 주 PPT를 생성한다. 장애공유 보고서 전체 생성, Excel 집계 갱신, 갱신된 Excel의 PPT 변환 요청에 사용한다.
---

# 주간 장애보고서 생성

기존 `update_excel.py`와 `make_ppt.py`를 실행한다. 보고서 집계·비교·차트 처리 로직을 새로 작성하지 않는다. 이 스킬은 저장소 코드를 사용하는 지침이며, 단독 실행 패키지가 아니다.

## 코드와 실행 환경

- 현재 작업 디렉터리에서 `update_excel.py`, `make_ppt.py`, `appconfig.py`, `requirements.txt`가 있는 저장소 루트를 찾는다. 이 파일이 원래 위치인 `skill/weekly-incident-report/`에 있다면 두 단계 위가 루트다. 설치 위치를 저장소 루트로 가정하지 않는다. 저장소를 찾을 수 없으면 사용자에게 경로를 요청한다.
- 저장소 루트에서 기존 Python 환경을 사용한다. 전체 실행과 Excel 갱신에는 Windows, Microsoft Excel, `pywin32`, `openpyxl`, `python-pptx`가 필요하다. PPT만 생성할 때는 `openpyxl`, `python-pptx`를 사용하며 Excel COM은 호출하지 않는다. PowerPoint 앱은 생성에 필수가 아니며 시각적 결과 확인에 사용할 수 있다.
- 의존성이 없으면 해당 Python 환경에서 `python -m pip install -r requirements.txt`로 설치한다. Excel COM이 없는 환경에서는 피벗 갱신을 다른 라이브러리로 대체하지 말고 필요한 환경을 안내한다.
- `appconfig.py`는 루트의 `report_config.json`, 없으면 `report_config.sample.json`을 읽는다. 실제 부서명과 행 번호가 입력 Excel에 맞는지 확인한다. 예시 값으로 실제 보고서를 실행하지 않는다. 설정이 없으면 sample을 복사하고 사용자 제공 값으로 채우며, 부서명·행 번호를 추측하지 않는다. 실제 설정과 장애 데이터를 스킬 배포 파일에 포함하지 않는다.

## 입력과 실행 전 확인

1. 요청 범위를 전체 실행, Excel만 갱신, PPT만 생성 중에서 판단한다. 보고서 전체 생성 요청은 Excel 갱신 후 PPT 생성으로 처리한다.
2. 기준일(`YYYY-MM-DD`), 원본 Excel 경로, PPT 생성 시 지난주 PPT 경로를 확보한다. 사용자가 오늘을 지정하면 로컬 날짜를 사용한다. 기존 대화나 명확한 파일 지정에서 알 수 없는 필수 입력만 질문하며, 여러 후보 중 임의로 고르지 않는다.
3. 입력 파일의 존재와 `.xlsx` / `.pptx` 확장자를 확인한다. Excel은 `요약`, `데이터`, `사업부별`, `사업부별_계열사별` 시트와 기존 코드의 셀·열 구조를 사용한다. PPT는 기존 회사 양식의 슬라이드·차트·표 구조를 사용한다. 임의의 Excel이나 PPT 양식을 자동 변환하는 기능이 아니다. 구조가 다르면 실행을 멈추고 차이를 안내한다. 필요할 때 저장소의 `README_엑셀업데이트.md`, `README_PPT생성.md`와 해당 Python 코드를 읽는다.
4. 기본 Excel 갱신은 원본을 `<원본이름>_YYYYMMDD.xlsx`로 백업한 뒤 원본에 저장한다. 원본 보존을 요청했다면 다른 경로를 `--out`으로 지정한다. 이 경우 스크립트는 별도 백업을 만들지 않으며, PPT 입력도 그 출력 Excel로 바꾼다.
5. 백업이나 출력 파일이 이미 있으면 기존 결과가 지워지지 않게 충돌 없는 경로를 선택한다. 사용자가 해당 파일의 덮어쓰기를 이미 지정했다면 그대로 처리한다. 기본 백업명이 충돌하면 원본을 다른 이름으로 먼저 백업한 뒤 Excel 명령에 `--out`으로 원본 경로를 명시한다. `--out`과 원본 경로가 같으면 자동 백업이 생기지 않는 점에 주의한다. 출력 PPT 경로는 지난주 PPT나 별도 템플릿 경로와 같게 지정하지 않는다.
6. 실행할 단계, 기준일, 입력과 저장·백업 경로를 간단히 알리고 실행한다. Excel 파일 잠금 오류가 나면 사용자가 파일을 닫도록 안내한다. 다른 Excel 프로세스를 강제 종료하지 않는다.

## 실행

아래 명령의 날짜와 경로를 확보한 입력으로 바꾼다. 경로는 절대 경로로 지정하고, 공백과 한글이 포함된 경로는 셸에 맞게 인용한다. 사용자가 제공한 문자열을 실행 코드로 평가하지 않는다. 각 명령을 따로 실행해 종료 코드를 확인한다.

### Excel 갱신

```powershell
python update_excel.py --src "C:\Reports\장애데이터.xlsx" --date 2026-10-14
```

원본을 유지하고 복사본을 갱신하려면:

```powershell
python update_excel.py --src "C:\Reports\장애데이터.xlsx" --date 2026-10-14 --out "C:\Reports\장애데이터_갱신.xlsx"
```

`--outdir`는 현재 구현에서 출력 대상 결정에 사용되지 않는다. 다른 경로에 저장할 때는 `--out`을 사용한다.

### PPT 생성

```powershell
python make_ppt.py --xlsx "C:\Reports\장애데이터.xlsx" --prev "C:\Reports\20261007_장애공유.pptx" --date 2026-10-14 --out "C:\Reports\20261014_장애공유.pptx"
```

- 항상 `--xlsx`로 이번에 갱신한 파일을 명시한다. PPT만 요청했다면 사용자가 지정한 갱신된 Excel을 사용한다.
- `--prev`는 비교용 지난주 PPT이며 필수다. 기본적으로 템플릿 역할도 한다. 사용자가 별도 템플릿을 지정했을 때만 `--template`을 추가한다.
- 기준일 기본 정책은 Excel 우선(`--date-source auto`)이다. `요약!B2`, 없으면 `B1`의 날짜를 사용하고, 읽지 못하면 입력일 또는 오늘로 대체한다. `--date`만으로 Excel 날짜를 강제로 바꿀 수 없다.
- Excel 날짜를 읽어 요청일과 일치하는지 확인한다. 전체 실행 후에도 날짜가 다르거나 읽을 수 없으면 PPT를 생성하지 말고 원인을 확인한다. PPT만 생성하는 요청에서도 불일치를 알리고 사용자의 날짜 선택을 확인한다. 사용자가 제목 날짜 강제를 명시했을 때만 `--date-source input`을 사용하며, 집계 데이터의 날짜는 그대로임을 알린다. `--date-source excel`도 날짜 읽기 실패를 오류로 처리하지 않으므로 검증을 대신할 수 없다.
- 장애 내용 요약은 코드의 규칙 기반 처리다. 사용자가 원문을 요구할 때만 `--raw-content`를 추가한다. 장애 사실이나 건수를 모델이 임의로 수정하지 않는다.

## 실패 처리와 결과 확인

- 전체 실행에서는 Excel 명령이 성공하고 출력 파일과 기준일을 확인한 뒤에만 PPT 명령을 실행한다. 실패하면 단계·종료 코드·핵심 오류와 이미 생성된 백업/출력 경로를 안내한다. 이전 파일로 조용히 이어서 실행하거나 원본을 자동 복원하지 않는다.
- Excel 출력은 `openpyxl.load_workbook(..., data_only=True, read_only=True)`로 열어 필요한 시트와 `요약!B2`(없으면 `B1`) 기준일을 확인하고 닫는다. 검증 과정에서 Excel을 `openpyxl`로 다시 저장하지 않는다. 피벗·차트 보존과 수식 재계산은 기존 Excel COM 코드가 담당한다.
- PPT 출력 파일이 존재하고 비어 있지 않은지 확인하고, `pptx.Presentation`으로 열어 슬라이드가 있는지 확인한다. 렌더링이나 PowerPoint 앱으로 확인할 수 있으면 제목 날짜·차트·상세 표의 잘림을 점검한다. 열기 검사만 했으면 시각적 검증까지 완료했다고 말하지 않는다.
- 최종 응답에는 적용 기준일, 갱신된 Excel, 백업(생성한 경우), 생성한 PPT의 경로와 실제 수행한 검증을 간결하게 전달한다. 내부 장애 상세 내용은 요청 없이 답변이나 외부 서비스에 전송하지 않는다.

## Codex와 Claude Code에서 사용

이 폴더는 요청에 따라 저장소의 `skill/` 아래에 두었다. 이 경로만으로 자동 검색된다고 가정하지 않는다. 설치 전에는 이 `SKILL.md` 경로를 직접 지정해 사용하도록 요청할 수 있다.

- Codex 프로젝트 설치: 이 폴더를 저장소의 `.agents/skills/weekly-incident-report/`로 복사한다.
- Claude Code 프로젝트 설치: 이 폴더를 저장소의 `.claude/skills/weekly-incident-report/`로 복사한다.
- 설치 후 Codex에서는 `$weekly-incident-report`, Claude Code에서는 `/weekly-incident-report`로 호출할 수 있다. 두 설치본을 사용하는 경우 이 파일을 원본으로 유지해 함께 갱신한다.
- 설치해도 실행 코드는 저장소 루트에 남는다. 다른 프로젝트나 PC에서 사용하려면 WeeklyIncidentReport 저장소와 설정, 입력 파일, 실행 환경이 함께 필요하다.
