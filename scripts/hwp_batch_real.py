from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import time
from pathlib import Path

from hwp_batch_core import (
    FORMAT_TYPES,
    HWP_PROGIDS,
    RealWorkerResult,
    AutoDialogEvent,
    dedupe_strings,
    kill_processes,
    parse_json_text,
    read_json_file,
    safe_unlink,
    write_json_file,
)
from hwp_batch_dialogs import AutoAllowDialogWatcher

DOCUMENT_LOAD_DELAY = 1.0
HWP_PROCESS_NAMES = {"hwp.exe", "hwpctrl.exe"}
WORKER_POLL_INTERVAL_SECONDS = 0.2


def snapshot_hwp_pids() -> set[int]:
    try:
        result = subprocess.run(
            ["tasklist", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
            check=False,
        )
        if result.returncode != 0:
            return set()
        import csv
        import io

        reader = csv.reader(io.StringIO(result.stdout))
        pids: set[int] = set()
        for row in reader:
            if len(row) < 2:
                continue
            image_name = row[0].strip().lower()
            if image_name not in HWP_PROCESS_NAMES:
                continue
            try:
                pids.add(int(row[1]))
            except ValueError:
                pass
        return pids
    except Exception:
        return set()


class RealHwpConverter:
    def __init__(self) -> None:
        self.hwp = None
        self.progid_used: str | None = None
        self.is_initialized = False
        self.owned_pids: set[int] = set()
        self.pythoncom = None

    def initialize(self) -> bool:
        if self.is_initialized:
            return True

        try:
            import pythoncom
            from win32com import client as win32_client
        except ImportError as exc:
            raise RuntimeError("pywin32가 필요합니다. `pip install pywin32` 후 다시 실행해주세요.") from exc

        self.pythoncom = pythoncom
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass

        dispatch_factory = getattr(win32_client, "DispatchEx", win32_client.Dispatch)
        errors: list[str] = []

        for progid in HWP_PROGIDS:
            before_pids = snapshot_hwp_pids()
            try:
                self.hwp = dispatch_factory(progid)
                self.progid_used = progid
                try:
                    self.hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")
                except Exception:
                    pass
                self.hwp.SetMessageBoxMode(0x00000001)
                time.sleep(0.2)
                self.owned_pids = snapshot_hwp_pids() - before_pids
                self.is_initialized = True
                return True
            except Exception as exc:
                errors.append(f"{progid}: {exc}")

        raise RuntimeError("한글 COM 객체 생성에 실패했습니다.\n" + "\n".join(errors))

    def convert_file(self, input_path: Path, output_path: Path, format_type: str = "PDF") -> tuple[bool, str | None]:
        if not self.is_initialized or self.hwp is None:
            return False, "한글 COM 객체가 초기화되지 않았습니다."
        try:
            output_path.parent.mkdir(parents=True, exist_ok=True)
            self.hwp.Open(str(input_path), "", "forceopen:true")
            time.sleep(DOCUMENT_LOAD_DELAY)
            save_format = FORMAT_TYPES[format_type]["save_format"]
            try:
                self.hwp.SaveAs(str(output_path), save_format)
            except Exception:
                self.hwp.SaveAs(str(output_path), save_format, "")
            self.hwp.Clear(option=1)
            return True, None
        except Exception as exc:
            try:
                self.hwp.Clear(option=1)
            except Exception:
                pass
            return False, str(exc)

    def cleanup(self) -> None:
        if self.hwp is not None and self.is_initialized:
            try:
                self.hwp.Clear(3)
            except Exception:
                pass
            try:
                self.hwp.Quit()
            except Exception:
                pass
            self.hwp = None
            self.is_initialized = False

        if self.pythoncom is not None:
            try:
                self.pythoncom.CoUninitialize()
            except Exception:
                pass


def _make_worker_state_path() -> Path:
    handle = tempfile.NamedTemporaryFile(prefix="hwp-convert-state-", suffix=".json", delete=False)
    handle.close()
    return Path(handle.name)


def _worker_command(
    *,
    script_path: Path,
    input_path: Path,
    output_path: Path,
    format_type: str,
    auto_allow_dialogs: bool,
    state_path: Path,
) -> list[str]:
    command = [
        sys.executable,
        str(script_path),
        "--internal-worker-real-convert",
        "--worker-input",
        str(input_path),
        "--worker-output",
        str(output_path),
        "--worker-format",
        format_type,
        "--worker-state-json",
        str(state_path),
    ]
    if auto_allow_dialogs:
        command.append("--worker-auto-allow-dialogs")
    return command


def run_real_worker_task(task, args, script_path: Path) -> RealWorkerResult:
    state_path = _make_worker_state_path()
    before_pids = snapshot_hwp_pids()
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    proc = subprocess.Popen(
        _worker_command(
            script_path=script_path,
            input_path=task.input_file,
            output_path=task.output_file,
            format_type=args.format,
            auto_allow_dialogs=args.auto_allow_dialogs,
            state_path=state_path,
        ),
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=env,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )

    started_at = time.monotonic()
    initialized_at: float | None = None
    state_payload: dict[str, object] = {}
    timeout_stage: str | None = None

    try:
        while proc.poll() is None:
            latest_state = read_json_file(state_path)
            if latest_state:
                state_payload = latest_state
                if state_payload.get("initialized") and initialized_at is None:
                    initialized_at = time.monotonic()

            now = time.monotonic()
            if initialized_at is None:
                if now - started_at > args.startup_timeout_seconds:
                    timeout_stage = "startup"
                    break
            elif now - initialized_at > args.file_timeout_seconds:
                timeout_stage = "file"
                break
            time.sleep(WORKER_POLL_INTERVAL_SECONDS)

        if timeout_stage is not None:
            proc.kill()
            try:
                proc.communicate(timeout=5)
            except Exception:
                pass
            latest_state = read_json_file(state_path)
            if latest_state:
                state_payload = latest_state

            warnings = list(state_payload.get("warnings", []))
            owned_pids = {int(pid) for pid in state_payload.get("owned_pids", [])}
            if args.kill_owned_hwp_on_timeout and not owned_pids:
                owned_pids = snapshot_hwp_pids() - before_pids

            if args.kill_owned_hwp_on_timeout and owned_pids:
                killed_pids = kill_processes(owned_pids)
                if killed_pids:
                    warnings.append(f"timeout 후 정리한 HWP PID: {', '.join(str(pid) for pid in killed_pids)}")
            elif args.kill_owned_hwp_on_timeout:
                warnings.append("timeout 후 정리할 HWP PID를 찾지 못했습니다.")

            if timeout_stage == "startup":
                error = f"초기화 시간 제한 {args.startup_timeout_seconds:.1f}초를 초과했습니다."
            else:
                error = f"파일 변환 시간 제한 {args.file_timeout_seconds:.1f}초를 초과했습니다."
            return RealWorkerResult(
                ok=False,
                error=error,
                warnings=dedupe_strings(warnings),
                progid_used=state_payload.get("progid_used"),
            )

        stdout, stderr = proc.communicate()
        latest_state = read_json_file(state_path)
        if latest_state:
            state_payload = latest_state

        payload = parse_json_text(stdout)
        warnings = list(state_payload.get("warnings", []))
        if payload and isinstance(payload.get("warnings"), list):
            warnings.extend(str(item) for item in payload["warnings"])

        if payload is None:
            error = stderr.strip() or stdout.strip() or "real worker 결과를 파싱하지 못했습니다."
            return RealWorkerResult(
                ok=False,
                error=error,
                warnings=dedupe_strings(warnings),
                progid_used=state_payload.get("progid_used"),
            )

        events = [
            AutoDialogEvent.from_record(record)
            for record in payload.get("auto_dialog_events", [])
            if isinstance(record, dict)
        ]
        ok = bool(payload.get("ok", False)) and proc.returncode == 0
        error = None if ok else str(payload.get("error") or stderr.strip() or "real worker 실행에 실패했습니다.")
        return RealWorkerResult(
            ok=ok,
            error=error,
            warnings=dedupe_strings(warnings),
            progid_used=str(payload.get("progid_used") or state_payload.get("progid_used") or "") or None,
            auto_dialog_events=events,
        )
    finally:
        safe_unlink(state_path)


def run_internal_real_worker(args) -> int:
    state_path = Path(args.worker_state_json)
    write_json_file(state_path, {"initialized": False, "owned_pids": [], "warnings": []})

    converter = RealHwpConverter()
    warnings: list[str] = []
    watcher: AutoAllowDialogWatcher | None = None
    payload: dict[str, object]

    try:
        converter.initialize()
        allowed_pids = set(converter.owned_pids)
        if args.worker_auto_allow_dialogs and not allowed_pids:
            warnings.append("자동 허용을 요청했지만 소유 HWP PID를 확인하지 못해 watcher를 비활성화했습니다.")

        watcher = AutoAllowDialogWatcher(
            enabled=args.worker_auto_allow_dialogs and bool(allowed_pids),
            allowed_pids=allowed_pids if allowed_pids else set(),
        )
        write_json_file(
            state_path,
            {
                "initialized": True,
                "owned_pids": sorted(converter.owned_pids),
                "progid_used": converter.progid_used,
                "warnings": warnings,
            },
        )
        watcher.start()
        ok, error = converter.convert_file(Path(args.worker_input), Path(args.worker_output), args.worker_format)
        events = watcher.snapshot_events()
        payload = {
            "ok": ok,
            "error": error,
            "progid_used": converter.progid_used,
            "owned_pids": sorted(converter.owned_pids),
            "warnings": warnings,
            "auto_dialog_events": [event.to_record() for event in events],
        }
    except Exception as exc:
        events = watcher.snapshot_events() if watcher is not None else []
        payload = {
            "ok": False,
            "error": str(exc),
            "progid_used": converter.progid_used,
            "owned_pids": sorted(converter.owned_pids),
            "warnings": warnings,
            "auto_dialog_events": [event.to_record() for event in events],
        }
        write_json_file(
            state_path,
            {
                "initialized": converter.is_initialized,
                "owned_pids": sorted(converter.owned_pids),
                "progid_used": converter.progid_used,
                "warnings": warnings,
                "error": str(exc),
            },
        )
    finally:
        if watcher is not None:
            watcher.stop()
        converter.cleanup()

    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0 if payload["ok"] else 1
