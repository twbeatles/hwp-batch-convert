from __future__ import annotations

import argparse
import json
import subprocess
import sys
import time
from pathlib import Path
from typing import Sequence

from hwp_batch_core import (
    DEFAULT_FILE_TIMEOUT_SECONDS,
    DEFAULT_STARTUP_TIMEOUT_SECONDS,
    FAIL_FAST_DETAIL,
    FORMAT_TYPES,
    STATUS_FAILED,
    STATUS_PLANNED,
    STATUS_SKIPPED,
    STATUS_SUCCESS,
    ConversionSummary,
    TaskPlanner,
    dedupe_strings,
    determine_exit_code,
    render_human,
    write_json_file,
)
from hwp_batch_dialogs import AutoAllowDialogWatcher
from hwp_batch_core import MockConverter
from hwp_batch_real import run_internal_real_worker, run_real_worker_task


def choose_converter(mode: str):
    if mode == "mock":
        return MockConverter()
    raise ValueError(f"지원하지 않는 모드입니다: {mode}")


def run_conversion(args: argparse.Namespace) -> ConversionSummary:
    planner = TaskPlanner()
    plan = planner.build_tasks(
        sources=args.sources,
        format_type=args.format,
        include_sub=args.include_sub,
        same_location=args.same_location,
        output_path=args.output_dir or "",
        allow_empty=args.allow_empty,
        preserve_source_root=args.preserve_source_root,
    )
    plan.conflict_renamed_count = planner.resolve_output_conflicts(plan.tasks, overwrite=args.overwrite)
    if plan.conflict_renamed_count:
        plan.warnings.append(f"출력 파일 충돌 {plan.conflict_renamed_count}건은 자동으로 새 이름을 부여했습니다.")

    all_tasks = list(plan.skipped_tasks)
    auto_dialog_events = []
    warnings = list(plan.warnings)
    progids: list[str] = []
    start = time.time()

    if args.plan_only:
        for task in plan.tasks:
            task.status = STATUS_PLANNED
        all_tasks.extend(plan.tasks)
        return ConversionSummary(
            format_type=args.format,
            tasks=all_tasks,
            warnings=dedupe_strings(warnings),
            elapsed_seconds=round(time.time() - start, 3),
            mode="plan",
            auto_dialog_enabled=args.auto_allow_dialogs,
        )

    if not plan.tasks:
        return ConversionSummary(
            format_type=args.format,
            tasks=all_tasks,
            warnings=dedupe_strings(warnings),
            elapsed_seconds=round(time.time() - start, 3),
            mode=args.mode,
            auto_dialog_enabled=args.auto_allow_dialogs,
        )

    if args.mode == "real":
        script_path = Path(__file__).resolve()
        for index, task in enumerate(plan.tasks):
            result = run_real_worker_task(task, args, script_path)
            task.status = STATUS_SUCCESS if result.ok else STATUS_FAILED
            task.error = result.error
            all_tasks.append(task)
            warnings.extend(result.warnings)
            auto_dialog_events.extend(result.auto_dialog_events)
            if result.progid_used:
                progids.append(result.progid_used)
            if not result.ok and args.fail_fast:
                warnings.append("--fail-fast에 따라 남은 작업을 건너뛰었습니다.")
                for remaining_task in plan.tasks[index + 1 :]:
                    remaining_task.status = STATUS_SKIPPED
                    remaining_task.error = FAIL_FAST_DETAIL
                    all_tasks.append(remaining_task)
                break
    else:
        converter = choose_converter(args.mode)
        converter.initialize()
        try:
            for index, task in enumerate(plan.tasks):
                ok, error = converter.convert_file(task.input_file, task.output_file, args.format)
                task.status = STATUS_SUCCESS if ok else STATUS_FAILED
                task.error = error
                all_tasks.append(task)
                if not ok and args.fail_fast:
                    warnings.append("--fail-fast에 따라 남은 작업을 건너뛰었습니다.")
                    for remaining_task in plan.tasks[index + 1 :]:
                        remaining_task.status = STATUS_SKIPPED
                        remaining_task.error = FAIL_FAST_DETAIL
                        all_tasks.append(remaining_task)
                    break
        finally:
            converter.cleanup()
        if args.auto_allow_dialogs:
            warnings.append("--auto-allow-dialogs 는 real 모드에서만 실제 동작합니다.")
        progid = getattr(converter, "progid_used", None)
        if progid:
            progids.append(progid)

    if auto_dialog_events:
        clicked = len([event for event in auto_dialog_events if event.clicked])
        warnings.append(f"보안 팝업 자동 허용 기록: 감지 {len(auto_dialog_events)}건, 클릭 {clicked}건")

    return ConversionSummary(
        format_type=args.format,
        tasks=all_tasks,
        warnings=dedupe_strings(warnings),
        elapsed_seconds=round(time.time() - start, 3),
        progid_used=", ".join(dedupe_strings(progids)) or None,
        mode=args.mode,
        auto_dialog_enabled=args.auto_allow_dialogs,
        auto_dialog_events=auto_dialog_events,
    )


def run_dialog_self_test(timeout_seconds: float = 8.0) -> dict[str, object]:
    dialog_script = r"""
Add-Type -AssemblyName System.Windows.Forms
$form = New-Object System.Windows.Forms.Form
$form.Text = '한글'
$form.Width = 420
$form.Height = 170
$form.StartPosition = 'CenterScreen'
$label = New-Object System.Windows.Forms.Label
$label.AutoSize = $true
$label.Left = 20
$label.Top = 20
$label.Text = '한글 문서에 접근하려는 시도를 허용하시겠습니까?'
$form.Controls.Add($label)
$button = New-Object System.Windows.Forms.Button
$button.Text = '모두 허용'
$button.Left = 20
$button.Top = 70
$button.Width = 100
$button.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $button
$form.Controls.Add($button)
[void]$form.ShowDialog()
""".strip()
    proc = subprocess.Popen(["powershell", "-NoProfile", "-STA", "-Command", dialog_script])
    watcher = AutoAllowDialogWatcher(enabled=True, allowed_pids={proc.pid}, poll_interval=0.2)
    try:
        clicked = watcher.click_once_for_test(timeout_seconds=timeout_seconds)
        proc.wait(timeout=timeout_seconds)
        events = watcher.snapshot_events()
        return {
            "clicked": clicked,
            "returncode": proc.returncode,
            "events": [event.to_record() for event in events],
        }
    finally:
        if proc.poll() is None:
            proc.kill()


def build_error_payload(exc: Exception) -> dict[str, object]:
    return {
        "error": str(exc),
        "summary": {
            "failed_count": 1,
        },
    }


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="한글(HWP/HWPX) 문서를 PDF/DOCX/HWPX 등으로 일괄 변환합니다.")
    parser.add_argument("sources", nargs="*", help="입력 파일 또는 폴더 경로(여러 개 가능)")
    parser.add_argument("--format", default="PDF", choices=sorted(FORMAT_TYPES.keys()), help="출력 형식")
    parser.add_argument("--include-sub", action="store_true", default=True, help="하위 폴더 포함(기본값: 켜짐)")
    parser.add_argument("--no-include-sub", dest="include_sub", action="store_false", help="하위 폴더 미포함")
    output_group = parser.add_mutually_exclusive_group()
    output_group.add_argument("--same-location", action="store_true", default=False, help="원본과 같은 폴더에 출력")
    output_group.add_argument("--output-dir", help="출력 루트 폴더")
    parser.add_argument("--overwrite", action="store_true", help="같은 이름 출력 파일 덮어쓰기 허용")
    parser.add_argument("--plan-only", action="store_true", help="실제 변환 없이 작업 계획만 출력")
    parser.add_argument("--mode", choices=["real", "mock"], default="real", help="real=한글 COM 실변환, mock=테스트용 가짜 변환")
    parser.add_argument("--auto-allow-dialogs", action="store_true", help="한글 보안 확인 팝업을 소유 HWP 프로세스 범위에서만 자동 클릭")
    parser.add_argument("--json", action="store_true", help="JSON 출력")
    parser.add_argument("--report-json", help="결과 JSON 파일 저장 경로")
    parser.add_argument("--startup-timeout-seconds", type=float, default=DEFAULT_STARTUP_TIMEOUT_SECONDS, help="real 모드 초기화 timeout(초)")
    parser.add_argument("--file-timeout-seconds", type=float, default=DEFAULT_FILE_TIMEOUT_SECONDS, help="real 모드 파일별 변환 timeout(초)")
    parser.add_argument("--fail-fast", action="store_true", help="작업 하나가 실패하면 남은 작업을 건너뜀")
    parser.add_argument("--allow-partial-success", action="store_true", help="일부 파일 실패가 있어도 전체 종료 코드를 성공으로 유지")
    parser.add_argument("--allow-empty", action="store_true", help="변환 대상이 없어도 오류로 처리하지 않음")
    parser.add_argument("--preserve-source-root", action="store_true", help="여러 입력 source를 output-dir 아래에서 source 이름별로 구분")
    parser.add_argument("--kill-owned-hwp-on-timeout", action="store_true", help="timeout 발생 시 이번 실행이 띄운 HWP 프로세스를 정리 시도")
    parser.add_argument("--self-test-dialog-handler", action="store_true", help="보안 팝업 자동 클릭 로직의 로컬 UI 테스트를 실행")
    parser.add_argument("--internal-worker-real-convert", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--worker-input", help=argparse.SUPPRESS)
    parser.add_argument("--worker-output", help=argparse.SUPPRESS)
    parser.add_argument("--worker-format", help=argparse.SUPPRESS)
    parser.add_argument("--worker-state-json", help=argparse.SUPPRESS)
    parser.add_argument("--worker-auto-allow-dialogs", action="store_true", help=argparse.SUPPRESS)

    args = parser.parse_args(argv)
    if args.internal_worker_real_convert:
        missing = [name for name in ("worker_input", "worker_output", "worker_format", "worker_state_json") if not getattr(args, name)]
        if missing:
            parser.error(f"worker 인자가 누락되었습니다: {', '.join(missing)}")
        return args
    if args.self_test_dialog_handler:
        return args
    if not args.sources:
        parser.error("입력 파일 또는 폴더를 하나 이상 지정해주세요.")
    if not args.same_location and not args.output_dir:
        parser.error("--same-location 또는 --output-dir 중 하나를 지정해주세요.")
    if args.startup_timeout_seconds <= 0:
        parser.error("--startup-timeout-seconds 는 0보다 커야 합니다.")
    if args.file_timeout_seconds <= 0:
        parser.error("--file-timeout-seconds 는 0보다 커야 합니다.")
    return args


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    if args.internal_worker_real_convert:
        return run_internal_real_worker(args)

    try:
        if args.self_test_dialog_handler:
            payload = run_dialog_self_test()
            print(json.dumps(payload, ensure_ascii=False, indent=2))
            return 0 if payload["clicked"] and payload["returncode"] == 0 else 1
        summary = run_conversion(args)
    except Exception as exc:
        payload = build_error_payload(exc)
        if getattr(args, "report_json", None):
            write_json_file(Path(args.report_json), payload)
        if getattr(args, "json", False):
            print(json.dumps(payload, ensure_ascii=False, indent=2))
        else:
            print(f"오류: {exc}", file=sys.stderr)
        return 1

    payload = summary.to_json_dict()
    if args.report_json:
        write_json_file(Path(args.report_json), payload)
    if args.json:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    else:
        print(render_human(summary))
    return determine_exit_code(summary, allow_partial_success=args.allow_partial_success)


if __name__ == "__main__":
    raise SystemExit(main())
