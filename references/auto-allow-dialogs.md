# 한글 보안 팝업 자동 허용 메모

## 목적

한글(HWP) COM 자동화 중 뜨는 보안 확인 팝업 때문에 `Open`/`SaveAs` 흐름이 멈추는 경우를 줄이기 위해, 매우 제한된 조건의 팝업만 자동 클릭한다.

현재 구현은 단순 화이트리스트를 넘어서, **이번 실행이 띄운 HWP 프로세스(PID) 범위 안에서만** 감시하도록 제한한다.

## 화이트리스트 조건

다음 조건을 모두 만족할 때만 클릭한다.

1. 최상위 창 제목이 정확히 `한글`
2. 자식 컨트롤 텍스트를 합친 본문에 `접근하려는 시도` 포함
3. 버튼 텍스트가 `모두 허용` 우선, 없으면 `허용`

## 안전장치

- 제목만 `한글` 인 다른 창은 본문/버튼 조건까지 함께 확인하기 전에는 클릭하지 않는다.
- 버튼 텍스트가 없거나 다른 경우 클릭하지 않는다.
- 화이트리스트와 맞지 않는 창은 `text-mismatch`, `button-mismatch` 등 이유를 이벤트에 남길 수 있다.
- 클릭이 끝난 창만 처리 완료로 간주하고, 버튼이 늦게 나타나는 경우를 위해 같은 창을 다시 볼 수 있다.
- 현재 실행이 띄운 HWP PID에 속하지 않는 창은 검사 대상에서 제외한다.
- 외부 UI 자동화 패키지 없이 Win32 API(`EnumWindows`, `EnumChildWindows`, `SendMessageW(BM_CLICK)`)만 사용한다.

## 기록 필드

자동 허용 관련 진단은 아래 필드에서 확인한다.

- `summary.auto_dialog_enabled`
- `summary.auto_dialog_detected_count`
- `summary.auto_dialog_clicked_count`
- `auto_dialog_events[]`
- `auto_dialog_events[].process_id`

## 테스트 방법

### UI 클릭 자체 테스트

```powershell
python skills/hwp-batch-convert/scripts/hwp_batch_convert.py --self-test-dialog-handler
```

테스트는 PowerShell WinForms로 아래와 유사한 샘플 창을 띄우고, 자동 클릭 로직이 `모두 허용` 버튼을 실제로 누르는지 확인한다.

- 제목: `한글`
- 본문: `한글 문서에 접근하려는 시도를 허용하시겠습니까?`
- 버튼: `모두 허용`

### 변환 경로 테스트

```powershell
python skills/hwp-batch-convert/scripts/hwp_batch_convert.py <입력> --format PDF --output-dir <출력> --mode mock --auto-allow-dialogs --json
```

mock 모드에서는 실제 한글 창이 없으므로 `auto_dialog_detected_count = 0` 이 정상이다.

### real 모드 권장 예시

```powershell
python skills/hwp-batch-convert/scripts/hwp_batch_convert.py <입력> --format PDF --output-dir <출력> --mode real --auto-allow-dialogs --startup-timeout-seconds 20 --file-timeout-seconds 120 --kill-owned-hwp-on-timeout --json --report-json <report.json>
```
