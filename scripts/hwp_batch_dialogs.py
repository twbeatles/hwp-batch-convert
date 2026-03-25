from __future__ import annotations

import ctypes
import threading
import time
from ctypes import wintypes

from hwp_batch_core import AutoDialogEvent

DIALOG_TITLE_WHITELIST = {"한글"}
DIALOG_TEXT_KEYWORDS = ("접근하려는 시도",)
DIALOG_ALLOW_BUTTONS = ("모두 허용", "허용")
POLL_INTERVAL_SECONDS = 0.35
WINDOW_SCAN_TIMEOUT_SECONDS = 3.0
BM_CLICK = 0x00F5
WM_GETTEXT = 0x000D
WM_GETTEXTLENGTH = 0x000E

USER32 = ctypes.WinDLL("user32", use_last_error=True)
USER32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
USER32.GetWindowThreadProcessId.restype = wintypes.DWORD


def _get_window_text(hwnd: int) -> str:
    length = USER32.SendMessageW(hwnd, WM_GETTEXTLENGTH, 0, 0)
    if length <= 0:
        return ""
    buffer = ctypes.create_unicode_buffer(length + 1)
    USER32.SendMessageW(hwnd, WM_GETTEXT, length + 1, ctypes.byref(buffer))
    return buffer.value.strip()


def _get_class_name(hwnd: int) -> str:
    buffer = ctypes.create_unicode_buffer(256)
    USER32.GetClassNameW(hwnd, buffer, 256)
    return buffer.value


def _get_window_pid(hwnd: int) -> int:
    pid = wintypes.DWORD()
    USER32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    return int(pid.value)


class AutoAllowDialogWatcher:
    def __init__(
        self,
        *,
        enabled: bool = False,
        allowed_pids: set[int] | None = None,
        poll_interval: float = POLL_INTERVAL_SECONDS,
    ) -> None:
        self.enabled = enabled
        self.allowed_pids = None if allowed_pids is None else set(allowed_pids)
        self.poll_interval = poll_interval
        self.events: list[AutoDialogEvent] = []
        self._lock = threading.Lock()
        self._stop_event = threading.Event()
        self._thread: threading.Thread | None = None
        self._handled_hwnds: set[int] = set()
        self._recorded_signatures: dict[int, tuple[str, str, str, str]] = {}

    def start(self) -> None:
        if not self.enabled:
            return
        self._thread = threading.Thread(target=self._run, name="hwp-auto-allow-dialogs", daemon=True)
        self._thread.start()

    def stop(self) -> None:
        self._stop_event.set()
        if self._thread is not None:
            self._thread.join(timeout=2.0)

    def snapshot_events(self) -> list[AutoDialogEvent]:
        with self._lock:
            return list(self.events)

    def click_once_for_test(self, timeout_seconds: float = WINDOW_SCAN_TIMEOUT_SECONDS) -> bool:
        deadline = time.time() + timeout_seconds
        while time.time() < deadline:
            if self._scan_once():
                return True
            time.sleep(self.poll_interval)
        return False

    def _run(self) -> None:
        while not self._stop_event.wait(self.poll_interval):
            self._scan_once()

    def _scan_once(self) -> bool:
        matched = False
        hwnds: list[int] = []
        enum_proc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)(lambda hwnd, _: hwnds.append(hwnd) or True)
        USER32.EnumWindows(enum_proc, 0)
        active_hwnds = set(hwnds)

        for hwnd in hwnds:
            if hwnd in self._handled_hwnds or not USER32.IsWindowVisible(hwnd):
                continue

            pid = _get_window_pid(hwnd)
            if self.allowed_pids is not None and pid not in self.allowed_pids:
                continue

            title_length = USER32.GetWindowTextLengthW(hwnd)
            if title_length <= 0:
                continue

            title_buffer = ctypes.create_unicode_buffer(title_length + 1)
            USER32.GetWindowTextW(hwnd, title_buffer, title_length + 1)
            title_text = title_buffer.value.strip()
            if title_text not in DIALOG_TITLE_WHITELIST:
                continue

            text_parts, allow_button_hwnd, allow_button_text = self._inspect_dialog(hwnd)
            window_text = " ".join(part for part in text_parts if part).strip()
            reason = self._classify_candidate(title_text, window_text, allow_button_text)
            signature = (title_text, window_text, allow_button_text, reason)

            if reason != "match":
                previous = self._recorded_signatures.get(hwnd)
                if window_text and previous != signature:
                    self._record_event(
                        AutoDialogEvent(
                            window_title=title_text,
                            window_text=window_text,
                            button_text=allow_button_text,
                            clicked=False,
                            reason=reason,
                            process_id=pid,
                        )
                    )
                    self._recorded_signatures[hwnd] = signature
                continue

            clicked = False
            if allow_button_hwnd:
                USER32.SendMessageW(allow_button_hwnd, BM_CLICK, 0, 0)
                clicked = True

            self._record_event(
                AutoDialogEvent(
                    window_title=title_text,
                    window_text=window_text,
                    button_text=allow_button_text,
                    clicked=clicked,
                    reason="clicked" if clicked else "allow-button-not-found",
                    process_id=pid,
                )
            )
            self._handled_hwnds.add(hwnd)
            self._recorded_signatures.pop(hwnd, None)
            matched = matched or clicked

        self._handled_hwnds.intersection_update(active_hwnds)
        self._recorded_signatures = {
            hwnd: signature
            for hwnd, signature in self._recorded_signatures.items()
            if hwnd in active_hwnds
        }
        return matched

    def _inspect_dialog(self, hwnd: int) -> tuple[list[str], int | None, str]:
        parts: list[str] = []
        allow_button_hwnd: int | None = None
        allow_button_text = ""
        child_hwnds: list[int] = []
        enum_proc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)(lambda child, _: child_hwnds.append(child) or True)
        USER32.EnumChildWindows(hwnd, enum_proc, 0)

        for child in child_hwnds:
            text = _get_window_text(child)
            class_name = _get_class_name(child)
            if text:
                parts.append(text)
            if "BUTTON" in class_name.upper() and text in DIALOG_ALLOW_BUTTONS and allow_button_hwnd is None:
                allow_button_hwnd = child
                allow_button_text = text
        return parts, allow_button_hwnd, allow_button_text

    def _classify_candidate(self, title: str, window_text: str, allow_button_text: str) -> str:
        if title not in DIALOG_TITLE_WHITELIST:
            return "title-mismatch"
        if not all(keyword in window_text for keyword in DIALOG_TEXT_KEYWORDS):
            return "text-mismatch"
        if not allow_button_text:
            return "button-mismatch"
        return "match"

    def _record_event(self, event: AutoDialogEvent) -> None:
        with self._lock:
            self.events.append(event)
