from __future__ import annotations

import json
import os
import subprocess
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable, Optional

SUPPORTED_EXTENSIONS = (".hwp", ".hwpx")
MAX_FILENAME_COUNTER = 1000
DEFAULT_STARTUP_TIMEOUT_SECONDS = 20.0
DEFAULT_FILE_TIMEOUT_SECONDS = 120.0

HWP_PROGIDS = [
    "HWPControl.HwpCtrl.1",
    "HwpObject.HwpObject",
    "HWPFrame.HwpObject",
]
FORMAT_TYPES: dict[str, dict[str, str]] = {
    "HWP": {"ext": ".hwp", "save_format": "HWP"},
    "HWPX": {"ext": ".hwpx", "save_format": "HWPX"},
    "PDF": {"ext": ".pdf", "save_format": "PDF"},
    "DOCX": {"ext": ".docx", "save_format": "OOXML"},
    "ODT": {"ext": ".odt", "save_format": "ODT"},
    "HTML": {"ext": ".html", "save_format": "HTML"},
    "RTF": {"ext": ".rtf", "save_format": "RTF"},
    "TXT": {"ext": ".txt", "save_format": "TEXT"},
    "PNG": {"ext": ".png", "save_format": "PNG"},
    "JPG": {"ext": ".jpg", "save_format": "JPG"},
    "BMP": {"ext": ".bmp", "save_format": "BMP"},
    "GIF": {"ext": ".gif", "save_format": "GIF"},
}

STATUS_PENDING = "대기"
STATUS_PLANNED = "계획됨"
STATUS_SUCCESS = "성공"
STATUS_FAILED = "실패"
STATUS_SKIPPED = "건너뜀"
FAIL_FAST_DETAIL = "이전 작업 실패 후 --fail-fast로 중단했습니다."


def canonicalize_path(path: str | Path) -> str:
    return os.path.abspath(os.path.normpath(str(path)))


def iter_supported_files(
    root_path: Path,
    include_sub: bool = True,
    allowed_exts: Optional[Iterable[str]] = None,
) -> Iterable[Path]:
    allowed = {ext.lower() for ext in (allowed_exts or SUPPORTED_EXTENSIONS)}
    if root_path.is_file():
        if root_path.suffix.lower() in allowed:
            yield root_path
        return
    if not root_path.is_dir():
        return
    if include_sub:
        for dirpath, _, filenames in os.walk(root_path):
            for filename in filenames:
                _, ext = os.path.splitext(filename)
                if ext.lower() in allowed:
                    yield Path(dirpath) / filename
        return
    with os.scandir(root_path) as entries:
        for entry in entries:
            if not entry.is_file():
                continue
            _, ext = os.path.splitext(entry.name)
            if ext.lower() in allowed:
                yield Path(entry.path)


def dedupe_strings(values: Iterable[str]) -> list[str]:
    seen: set[str] = set()
    deduped: list[str] = []
    for value in values:
        if not value or value in seen:
            continue
        seen.add(value)
        deduped.append(value)
    return deduped


def write_json_file(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def read_json_file(path: Path) -> dict[str, Any] | None:
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None


def parse_json_text(text: str) -> dict[str, Any] | None:
    text = text.strip()
    if not text:
        return None
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        return None


def safe_unlink(path: Path) -> None:
    try:
        path.unlink()
    except FileNotFoundError:
        return


def kill_processes(pids: Iterable[int]) -> list[int]:
    killed: list[int] = []
    for pid in sorted({int(pid) for pid in pids if int(pid) > 0}):
        try:
            result = subprocess.run(
                ["taskkill", "/PID", str(pid), "/T", "/F"],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="ignore",
                check=False,
            )
        except Exception:
            continue
        if result.returncode == 0:
            killed.append(pid)
    return killed


@dataclass
class AutoDialogEvent:
    window_title: str
    window_text: str
    button_text: str
    clicked: bool
    reason: str
    process_id: int | None = None
    timestamp: float = field(default_factory=time.time)

    def to_record(self) -> dict[str, Any]:
        return {
            "timestamp": round(self.timestamp, 3),
            "window_title": self.window_title,
            "window_text": self.window_text,
            "button_text": self.button_text,
            "clicked": self.clicked,
            "reason": self.reason,
            "process_id": self.process_id,
        }

    @classmethod
    def from_record(cls, record: dict[str, Any]) -> "AutoDialogEvent":
        return cls(
            window_title=str(record.get("window_title", "")),
            window_text=str(record.get("window_text", "")),
            button_text=str(record.get("button_text", "")),
            clicked=bool(record.get("clicked", False)),
            reason=str(record.get("reason", "")),
            process_id=record.get("process_id"),
            timestamp=float(record.get("timestamp", time.time())),
        )


@dataclass
class ConversionTask:
    input_file: Path
    output_file: Path
    source_root: Path | None = None
    status: str = STATUS_PENDING
    error: str | None = None

    def to_record(self) -> dict[str, Any]:
        return {
            "input_file": str(self.input_file),
            "output_file": str(self.output_file),
            "source_root": str(self.source_root) if self.source_root else None,
            "status": self.status,
            "detail": self.error or "",
        }


@dataclass
class PlannedConversion:
    format_type: str
    same_location: bool
    output_path: str
    tasks: list[ConversionTask] = field(default_factory=list)
    skipped_tasks: list[ConversionTask] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    conflict_renamed_count: int = 0


@dataclass
class ConversionSummary:
    format_type: str
    tasks: list[ConversionTask] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    elapsed_seconds: float | None = None
    progid_used: str | None = None
    mode: str = "real"
    auto_dialog_enabled: bool = False
    auto_dialog_events: list[AutoDialogEvent] = field(default_factory=list)

    def to_json_dict(self) -> dict[str, Any]:
        success = len([task for task in self.tasks if task.status == STATUS_SUCCESS])
        failed = len([task for task in self.tasks if task.status == STATUS_FAILED])
        skipped = len([task for task in self.tasks if task.status == STATUS_SKIPPED])
        clicked = len([event for event in self.auto_dialog_events if event.clicked])
        detected = len(self.auto_dialog_events)
        return {
            "summary": {
                "format_type": self.format_type,
                "mode": self.mode,
                "total_requested": len(self.tasks),
                "success_count": success,
                "failed_count": failed,
                "skipped_count": skipped,
                "elapsed_seconds": self.elapsed_seconds,
                "progid_used": self.progid_used,
                "warnings": dedupe_strings(self.warnings),
                "auto_dialog_enabled": self.auto_dialog_enabled,
                "auto_dialog_detected_count": detected,
                "auto_dialog_clicked_count": clicked,
            },
            "tasks": [task.to_record() for task in sorted(self.tasks, key=lambda item: str(item.input_file).lower())],
            "auto_dialog_events": [event.to_record() for event in self.auto_dialog_events],
        }


@dataclass
class RealWorkerResult:
    ok: bool
    error: str | None
    warnings: list[str] = field(default_factory=list)
    progid_used: str | None = None
    auto_dialog_events: list[AutoDialogEvent] = field(default_factory=list)


class TaskPlanner:
    def build_tasks(
        self,
        *,
        sources: list[str],
        format_type: str,
        include_sub: bool,
        same_location: bool,
        output_path: str,
        allow_empty: bool,
        preserve_source_root: bool,
    ) -> PlannedConversion:
        tasks: list[ConversionTask] = []
        skipped: list[ConversionTask] = []
        warnings: list[str] = []
        out_ext = FORMAT_TYPES[format_type]["ext"]

        if not sources:
            raise ValueError("파일 또는 폴더를 하나 이상 지정해주세요.")

        normalized_sources = [Path(canonicalize_path(src)) for src in sources]
        multiple_sources = len(normalized_sources) > 1
        explicit_output_root = Path(canonicalize_path(output_path)) if output_path else None

        if multiple_sources and explicit_output_root and not same_location and not preserve_source_root:
            warnings.append("여러 입력 소스를 함께 변환할 때 결과 추적이 중요하면 --preserve-source-root 사용을 권장합니다.")

        for source in normalized_sources:
            if not source.exists():
                raise ValueError(f"입력 경로가 존재하지 않습니다: {source}")

            source_root = source if source.is_dir() else source.parent
            source_files: list[Path]

            if source.is_file():
                if source.suffix.lower() not in SUPPORTED_EXTENSIONS:
                    if allow_empty:
                        warnings.append(f"지원하지 않는 입력 파일을 건너뜀: {source}")
                        continue
                    raise ValueError(f"지원하지 않는 입력 파일입니다: {source}")
                source_files = [source]
            else:
                source_files = sorted(iter_supported_files(source, include_sub=include_sub), key=lambda path: str(path).lower())
                if not source_files:
                    warnings.append(f"지원 파일이 없는 폴더를 건너뜀: {source}")

            for input_file in source_files:
                if input_file.suffix.lower() == out_ext.lower():
                    skipped.append(
                        ConversionTask(
                            input_file=input_file,
                            output_file=input_file,
                            source_root=source_root,
                            status=STATUS_SKIPPED,
                            error=f"이미 {format_type} 형식입니다.",
                        )
                    )
                    continue

                if same_location:
                    output_file = input_file.parent / f"{input_file.stem}{out_ext}"
                else:
                    if explicit_output_root is None:
                        raise ValueError("--output-dir 또는 --same-location 중 하나가 필요합니다.")
                    output_file = self._build_output_file(
                        source=source,
                        input_file=input_file,
                        output_root=explicit_output_root,
                        output_ext=out_ext,
                        multiple_sources=multiple_sources,
                        preserve_source_root=preserve_source_root,
                    )

                tasks.append(
                    ConversionTask(
                        input_file=input_file,
                        output_file=output_file,
                        source_root=source_root,
                    )
                )

        if skipped:
            warnings.append(f"동일 형식 {len(skipped)}개는 자동으로 건너뜁니다.")

        if not tasks and not skipped:
            if allow_empty:
                warnings.append("변환 대상이 없습니다.")
            else:
                raise ValueError("변환할 지원 파일이 없습니다.")

        return PlannedConversion(
            format_type=format_type,
            same_location=same_location,
            output_path=output_path,
            tasks=tasks,
            skipped_tasks=skipped,
            warnings=warnings,
        )

    def _build_output_file(
        self,
        *,
        source: Path,
        input_file: Path,
        output_root: Path,
        output_ext: str,
        multiple_sources: bool,
        preserve_source_root: bool,
    ) -> Path:
        filename = f"{input_file.stem}{output_ext}"
        if preserve_source_root:
            prefix = self._source_prefix(source)
            if source.is_file():
                return output_root / prefix / filename
            relative_path = input_file.relative_to(source)
            return output_root / prefix / relative_path.parent / filename
        if source.is_file() or multiple_sources:
            return output_root / filename
        relative_path = input_file.relative_to(source)
        return output_root / relative_path.parent / filename

    def _source_prefix(self, source: Path) -> Path:
        if source.is_dir():
            return Path(source.name or "source")
        parent_name = source.parent.name or "files"
        return Path(parent_name)

    def resolve_output_conflicts(self, tasks: list[ConversionTask], overwrite: bool) -> int:
        if overwrite:
            return 0
        used_paths: set[Path] = set()
        renamed_count = 0
        for task in tasks:
            original_path = task.output_file
            if task.output_file.exists() or task.output_file in used_paths:
                counter = 1
                stem = original_path.stem
                ext = original_path.suffix
                parent = original_path.parent
                while counter <= MAX_FILENAME_COUNTER:
                    candidate = parent / f"{stem} ({counter}){ext}"
                    if (not candidate.exists()) and (candidate not in used_paths):
                        task.output_file = candidate
                        break
                    counter += 1
                if task.output_file == original_path:
                    task.output_file = parent / f"{stem}_{int(time.time())}{ext}"
                if task.output_file != original_path:
                    renamed_count += 1
            used_paths.add(task.output_file)
        return renamed_count


class MockConverter:
    def __init__(self) -> None:
        self.progid_used = "mock"

    def initialize(self) -> bool:
        return True

    def convert_file(self, input_path: Path, output_path: Path, format_type: str = "PDF") -> tuple[bool, str | None]:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        payload = f"mock-converted:{Path(input_path).name}->{format_type}\n"
        output_path.write_text(payload, encoding="utf-8")
        return True, None

    def cleanup(self) -> None:
        return None


def render_human(summary: ConversionSummary) -> str:
    data = summary.to_json_dict()["summary"]
    lines = [
        f"형식: {data['format_type']} ({data['mode']})",
        f"총 {data['total_requested']}건 | 성공 {data['success_count']} | 실패 {data['failed_count']} | 건너뜀 {data['skipped_count']}",
    ]
    if data["warnings"]:
        lines.append("경고: " + " / ".join(data["warnings"]))
    if summary.auto_dialog_enabled:
        lines.append(
            f"보안 팝업 자동 허용: 감지 {data['auto_dialog_detected_count']} | 클릭 {data['auto_dialog_clicked_count']}"
        )
    failed = [task for task in summary.tasks if task.status == STATUS_FAILED]
    if failed:
        lines.append("실패 목록:")
        for task in failed[:10]:
            lines.append(f"- {task.input_file.name}: {task.error}")
    return "\n".join(lines)


def determine_exit_code(summary: ConversionSummary, allow_partial_success: bool) -> int:
    failed = len([task for task in summary.tasks if task.status == STATUS_FAILED])
    if failed and not allow_partial_success:
        return 1
    return 0
