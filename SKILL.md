---
name: hwp-batch-convert
description: Batch-convert 한컴 한글 문서(HWP/HWPX) on Windows into PDF and other export formats with a Korean-friendly automation workflow. Use when the user asks for 이 폴더 hwp 전부 pdf로 바꿔줘, hwp/hwpx/doc/pdf 일괄 변환, 한글문서 일괄 처리, 폴더 단위 변환, 여러 한글 파일을 다른 형식으로 내보내기, or wants plan-only / mock / real conversion runs with machine-readable reports. Supports HWP/HWPX to PDF, HWPX, DOCX, ODT, HTML, RTF, TXT, PNG, JPG, BMP, and GIF, plus optional automatic approval of known 한글 보안 확인 팝업. Prefer this skill for Windows environments with Hancom HWP installed; do not use it for non-HWP document families unless the task is explicitly about HWP/HWPX conversion.
---

# Hwp Batch Convert

Use this skill for **Windows 기반 한글(HWP/HWPX) 문서 일괄 변환**.

Current scope:
- 폴더 단위 일괄 변환
- 파일 여러 개 일괄 변환
- HWP/HWPX → PDF 기본 변환
- HWP/HWPX → HWPX/DOCX/ODT/HTML/RTF/TXT/PNG/JPG/BMP/GIF 변환
- 동일 형식 자동 건너뜀
- 출력 파일명 충돌 시 자동 번호 부여
- 지원하지 않는 단일 파일 조기 에러 처리
- 작업 계획만 확인하는 `--plan-only`
- OpenClaw 상위 레이어 연동용 `--json`, `--report-json`
- 한글 보안 확인 팝업 자동 허용용 `--auto-allow-dialogs`
- `--startup-timeout-seconds`, `--file-timeout-seconds` timeout 제어
- `--kill-owned-hwp-on-timeout` 자동화로 띄운 HWP 정리 시도
- `--fail-fast`, `--allow-partial-success`, `--allow-empty`
- `--preserve-source-root` 로 여러 입력 source 결과 구분
- 로컬 UI 검증용 `--self-test-dialog-handler`
- 테스트용 `--mode mock`

## Source basis

This skill reuses the design of the local/source repo:
- `tmp/HwpMate`
- upstream: `https://github.com/twbeatles/HwpMate`

Main logic origin:
- `hwpmate/services/hwp_converter.py`
- `hwpmate/services/task_planner.py`
- `hwpmate/constants.py`
- `hwpmate/path_utils.py`

If you need the mapping details or reuse rationale, read:
- `references/hwpmate-reuse-notes.md`

If you need the popup whitelist / safety details, read:
- `references/auto-allow-dialogs.md`

## Quick start

같은 폴더에 PDF 출력:

```bash
python skills/hwp-batch-convert/scripts/hwp_batch_convert.py "C:\docs\contracts" --format PDF --same-location
```

별도 출력 폴더로 변환:

```bash
python skills/hwp-batch-convert/scripts/hwp_batch_convert.py "C:\docs\hwp" --format PDF --output-dir "C:\docs\pdf" --auto-allow-dialogs
```

여러 파일 한 번에 변환:

```bash
python skills/hwp-batch-convert/scripts/hwp_batch_convert.py "C:\docs\a.hwp" "C:\docs\b.hwpx" --format DOCX --output-dir "C:\docs\docx"
```

실제 변환 없이 계획만 확인:

```bash
python skills/hwp-batch-convert/scripts/hwp_batch_convert.py "C:\docs\hwp" --format PDF --output-dir "C:\docs\pdf" --plan-only --json
```

테스트용 모의 변환:

```bash
python skills/hwp-batch-convert/scripts/hwp_batch_convert.py "C:\docs\sample" --format PDF --output-dir "C:\docs\out" --mode mock --json
```

## Main script

### `scripts/hwp_batch_convert.py`

Parameters:
- `sources...`: 입력 파일/폴더 경로 하나 이상
- `--format`: 출력 형식 (`PDF`, `HWPX`, `DOCX`, `ODT`, `HTML`, `RTF`, `TXT`, `PNG`, `JPG`, `BMP`, `GIF`, `HWP`)
- `--same-location`: 원본과 같은 폴더에 출력
- `--output-dir`: 출력 루트 폴더
- `--include-sub`: 하위 폴더 포함(기본값)
- `--no-include-sub`: 하위 폴더 제외
- `--overwrite`: 같은 이름 출력 허용
- `--plan-only`: 실제 변환 없이 작업 계획만 생성
- `--mode real|mock`: 실변환 또는 모의 변환
- `--auto-allow-dialogs`: 제목 `한글`, 본문에 `접근하려는 시도`, 버튼 `모두 허용`/`허용` 조건을 모두 만족하고 현재 실행이 띄운 HWP 프로세스 범위에 속한 보안 팝업만 자동 클릭
- `--startup-timeout-seconds`: real 모드 초기화 timeout
- `--file-timeout-seconds`: real 모드 파일별 timeout
- `--kill-owned-hwp-on-timeout`: timeout 시 owned HWP 정리 시도
- `--fail-fast`: 한 파일 실패 시 남은 작업 중단
- `--allow-partial-success`: 일부 실패가 있어도 종료 코드 `0`
- `--allow-empty`: 변환 대상이 없어도 빈 결과 허용
- `--preserve-source-root`: 여러 source를 output-dir 아래에서 source 이름별로 분리
- `--json`: stdout JSON 출력
- `--report-json`: 결과/에러 JSON 파일 저장
- `--self-test-dialog-handler`: 로컬 테스트용 샘플 `한글` 창을 띄워 자동 클릭 로직만 검증

## Recommended workflow

1. 사용자 요청이 폴더/여러 파일 기반 HWP/HWPX 변환인지 확인한다.
2. 출력 형식이 명시되지 않았으면 보통 `PDF`를 기본 제안으로 사용한다.
3. 먼저 `--plan-only --json` 으로 대상/건너뜀/출력 경로를 확인한다.
4. 여러 입력 source를 함께 넣으면 `--preserve-source-root` 필요 여부를 먼저 판단한다.
5. 환경 검증이 먼저 필요하면 `--mode mock` 으로 경로/출력 구조만 검증한다.
6. 환경이 Windows + 한글 설치 + pywin32 가능하면 `--mode real` 과 timeout 옵션으로 실행한다.
7. 보안 팝업 개입 가능성이 있으면 `--auto-allow-dialogs` 를 함께 검토한다.
8. 필요하면 `--report-json` 으로 결과 파일을 남긴다.

## Operational notes

- 이 스킬은 사실상 **Windows 전용**이다.
- 실변환(`--mode real`)은 **한컴 한글 설치**와 **pywin32**가 필요하다.
- `--auto-allow-dialogs` 는 화이트리스트 + owned PID 제한 방식이다. 제목이 `한글` 이고, 본문에 `접근하려는 시도` 가 있으며, 버튼이 `모두 허용` 또는 `허용` 인 경우에만 클릭한다.
- 위 조건에 맞지 않는 다른 팝업은 자동으로 건드리지 않는다. 감지되더라도 클릭 없이 이벤트만 남기거나 무시한다.
- 자동 허용 기록은 stdout JSON/`--report-json` 의 `auto_dialog_*` 필드와 `auto_dialog_events` 배열에서 확인한다.
- delayed button 같은 UI 지연 상황을 고려해 watcher가 같은 창을 다시 스캔할 수 있다.
- 한글 COM 자동화가 실패하면 `--mode mock` 으로 스킬/경로 로직만 먼저 검증하고, 이후 환경 문제를 분리 진단한다.
- 여러 폴더를 동시에 입력하면서 `--output-dir` 를 쓰면, `--preserve-source-root` 없이 평탄화될 수 있으니 결과 경로를 확인한다.
- 기본 종료 코드는 실패가 1건이라도 있으면 `1` 이고, 부분 성공을 허용하려면 `--allow-partial-success` 를 명시한다.
- 입력이 비어 있거나 지원하지 않는 파일이면 기본적으로 에러 처리되고, `--allow-empty` 일 때만 빈 결과를 허용한다.
- 기본 입력 확장자는 `.hwp`, `.hwpx` 만 대상으로 한다.
- 사용자 요청이 DOCX/PDF 일반 문서 처리 중심이고 HWP가 핵심이 아니면 다른 문서 스킬/도구를 우선 고려한다.

## Verification

- 기본 자동 검증: `pytest tests/test_hwp_batch_convert.py -q`
- 로컬 UI 클릭 검증: `python skills/hwp-batch-convert/scripts/hwp_batch_convert.py --self-test-dialog-handler`
