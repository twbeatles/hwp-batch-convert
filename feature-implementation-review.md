# Feature Implementation Review

## Status

초기 리뷰에서 지적한 주요 구현 이슈는 현재 코드에 반영되었다.

반영된 핵심 항목:

- `real` 모드 startup/file timeout 추가
- timeout 시 owned HWP PID 정리 시도 옵션 추가
- `--auto-allow-dialogs` 의 PID 범위 제한
- delayed button 상황을 고려한 watcher 재스캔
- unsupported 단일 입력의 조기 에러 처리
- `--same-location` / `--output-dir` 상호배타 처리
- `--preserve-source-root` 추가
- 실패 시 non-zero exit code 기본화
- `--allow-partial-success`, `--fail-fast`, `--allow-empty` 정책 추가
- 실패 시 `--report-json` 에 최소 에러 payload 저장
- `pytest` 자동 검증 추가

## 남아 있는 성격의 리스크

아래는 코드 결함이라기보다 운영상 한계에 가깝다.

- Hancom HWP COM 동작 자체는 Windows 환경과 설치 상태에 의존한다.
- timeout/정리 로직이 있어도 COM 내부 hang를 항상 완벽하게 복구한다고 보장할 수는 없다.
- 새 보안 팝업 패턴이 등장하면 `--auto-allow-dialogs` 가 자동 처리하지 않을 수 있다.

## Related docs

- `README.md`
- `SKILL.md`
- `references/auto-allow-dialogs.md`
- `references/hwpmate-reuse-notes.md`
