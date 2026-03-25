from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPTS_DIR = REPO_ROOT / "scripts"

if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

import hwp_batch_convert as cli
from hwp_batch_dialogs import AutoAllowDialogWatcher


def make_sample_tree(tmp_path: Path) -> Path:
    src = tmp_path / "src"
    (src / "sub").mkdir(parents=True)
    (src / "a.hwp").write_text("a", encoding="utf-8")
    (src / "b.hwpx").write_text("b", encoding="utf-8")
    (src / "sub" / "c.hwp").write_text("c", encoding="utf-8")
    return src


def test_preserve_source_root_and_source_root_field(tmp_path: Path) -> None:
    src1 = tmp_path / "src1"
    src2 = tmp_path / "src2"
    (src1 / "nested").mkdir(parents=True)
    (src2 / "deep").mkdir(parents=True)
    (src1 / "nested" / "a.hwp").write_text("a", encoding="utf-8")
    (src2 / "deep" / "b.hwp").write_text("b", encoding="utf-8")
    out = tmp_path / "out"

    args = cli.parse_args(
        [
            str(src1),
            str(src2),
            "--format",
            "PDF",
            "--output-dir",
            str(out),
            "--preserve-source-root",
            "--plan-only",
        ]
    )
    summary = cli.run_conversion(args)
    payload = summary.to_json_dict()

    outputs = {Path(task["output_file"]).relative_to(out).as_posix() for task in payload["tasks"]}
    source_roots = {Path(task["source_root"]).name for task in payload["tasks"]}

    assert outputs == {"src1/nested/a.pdf", "src2/deep/b.pdf"}
    assert source_roots == {"src1", "src2"}


def test_unsupported_single_file_is_error(tmp_path: Path) -> None:
    note = tmp_path / "note.txt"
    note.write_text("x", encoding="utf-8")

    with pytest.raises(ValueError, match="지원하지 않는 입력 파일"):
        cli.run_conversion(
            cli.parse_args(
                [
                    str(note),
                    "--format",
                    "PDF",
                    "--output-dir",
                    str(tmp_path / "out"),
                ]
            )
        )


def test_allow_empty_returns_summary(tmp_path: Path) -> None:
    empty = tmp_path / "empty"
    empty.mkdir()

    summary = cli.run_conversion(
        cli.parse_args(
            [
                str(empty),
                "--format",
                "PDF",
                "--output-dir",
                str(tmp_path / "out"),
                "--allow-empty",
                "--json",
            ]
        )
    )

    assert summary.mode == "real"
    assert summary.tasks == []
    assert "변환 대상이 없습니다." in summary.warnings


def test_same_location_and_output_dir_are_mutually_exclusive(tmp_path: Path) -> None:
    src = make_sample_tree(tmp_path)
    with pytest.raises(SystemExit):
        cli.parse_args(
            [
                str(src),
                "--format",
                "PDF",
                "--same-location",
                "--output-dir",
                str(tmp_path / "out"),
            ]
        )


def test_failed_tasks_return_nonzero_without_partial_success(tmp_path: Path, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]) -> None:
    src = make_sample_tree(tmp_path)

    def fake_convert(self, input_path, output_path, format_type="PDF"):
        if Path(input_path).name == "a.hwp":
            return False, "boom"
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text("ok", encoding="utf-8")
        return True, None

    monkeypatch.setattr(cli.MockConverter, "convert_file", fake_convert)

    exit_code = cli.main(
        [
            str(src),
            "--format",
            "PDF",
            "--output-dir",
            str(tmp_path / "out"),
            "--mode",
            "mock",
            "--json",
        ]
    )
    capsys.readouterr()

    assert exit_code == 1


def test_allow_partial_success_keeps_zero_exit(tmp_path: Path, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]) -> None:
    src = make_sample_tree(tmp_path)

    def fake_convert(self, input_path, output_path, format_type="PDF"):
        if Path(input_path).name == "a.hwp":
            return False, "boom"
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text("ok", encoding="utf-8")
        return True, None

    monkeypatch.setattr(cli.MockConverter, "convert_file", fake_convert)

    exit_code = cli.main(
        [
            str(src),
            "--format",
            "PDF",
            "--output-dir",
            str(tmp_path / "out"),
            "--mode",
            "mock",
            "--json",
            "--allow-partial-success",
        ]
    )
    capsys.readouterr()

    assert exit_code == 0


def test_report_json_written_on_exception(tmp_path: Path, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]) -> None:
    src = make_sample_tree(tmp_path)
    report = tmp_path / "report.json"

    def boom(_args):
        raise RuntimeError("planned failure")

    monkeypatch.setattr(cli, "run_conversion", boom)

    exit_code = cli.main(
        [
            str(src),
            "--format",
            "PDF",
            "--output-dir",
            str(tmp_path / "out"),
            "--json",
            "--report-json",
            str(report),
        ]
    )
    capsys.readouterr()

    payload = json.loads(report.read_text(encoding="utf-8"))
    assert exit_code == 1
    assert payload["error"] == "planned failure"


def _launch_dialog_process(*, delayed_button: bool) -> subprocess.Popen[str]:
    if delayed_button:
        script = r"""
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
$buttonTimer = New-Object System.Windows.Forms.Timer
$buttonTimer.Interval = 700
$buttonTimer.Add_Tick({
    $button = New-Object System.Windows.Forms.Button
    $button.Text = '모두 허용'
    $button.Left = 20
    $button.Top = 70
    $button.Width = 100
    $button.Add_Click({ $form.Close() })
    $form.Controls.Add($button)
    $buttonTimer.Stop()
})
$closeTimer = New-Object System.Windows.Forms.Timer
$closeTimer.Interval = 4000
$closeTimer.Add_Tick({
    $closeTimer.Stop()
    $form.Close()
})
$buttonTimer.Start()
$closeTimer.Start()
[void]$form.ShowDialog()
"""
    else:
        script = r"""
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
$button.Add_Click({ $form.Close() })
$form.Controls.Add($button)
$closeTimer = New-Object System.Windows.Forms.Timer
$closeTimer.Interval = 2000
$closeTimer.Add_Tick({
    $closeTimer.Stop()
    $form.Close()
})
$closeTimer.Start()
[void]$form.ShowDialog()
"""
    return subprocess.Popen(["powershell", "-NoProfile", "-STA", "-Command", script], text=True)


def test_dialog_watcher_retries_until_button_appears() -> None:
    proc = _launch_dialog_process(delayed_button=True)
    watcher = AutoAllowDialogWatcher(enabled=True, allowed_pids={proc.pid}, poll_interval=0.1)
    try:
        assert watcher.click_once_for_test(timeout_seconds=3.5) is True
        proc.wait(timeout=3)
        events = watcher.snapshot_events()
        assert any(event.reason == "button-mismatch" for event in events)
        assert any(event.clicked for event in events)
    finally:
        if proc.poll() is None:
            proc.kill()


def test_dialog_watcher_respects_allowed_pid() -> None:
    proc = _launch_dialog_process(delayed_button=False)
    watcher = AutoAllowDialogWatcher(enabled=True, allowed_pids={proc.pid + 100000}, poll_interval=0.1)
    try:
        assert watcher.click_once_for_test(timeout_seconds=1.5) is False
        proc.wait(timeout=3)
    finally:
        if proc.poll() is None:
            proc.kill()
