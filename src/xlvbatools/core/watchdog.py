"""
Excel Dialog Watchdog
======================
Background thread that monitors for and automatically dismisses Excel pop-up
dialogs during headless COM automation. Captures dialog text for diagnostics.

This is the SINGLE most important component for reliable headless VBA iteration.
Without it, any compile error, MsgBox, or runtime error dialog will hang the
COM call indefinitely, forcing a taskkill and losing all diagnostic context.

Supported dialog types:
    - Compile errors ("Compile error: ...")
    - Runtime errors ("Run-time error '1004': ...")
    - MsgBox calls from VBA code
    - File dialogs (GetSaveAsFilename, GetOpenFilename)
    - Application alerts (save warnings, overwrite confirmations)
    - VBE error dialogs

Usage (standalone):
    from xlvbatools.core.watchdog import DialogWatchdog

    watchdog = DialogWatchdog(target_pid=excel_pid)
    watchdog.start()
    # ... do COM stuff that might trigger dialogs ...
    events = watchdog.stop()
    for event in events:
        print(f"[{event.dialog_type}] {event.title}: {event.text}")

Usage (integrated -- via ExcelSession):
    The ExcelSession context manager starts/stops the watchdog automatically.
"""

import ctypes
import ctypes.wintypes
import logging
import os
import re
import threading
import time
from dataclasses import dataclass, field
from typing import List, Optional, Callable

from xlvbatools._compat import IS_WINDOWS

logger = logging.getLogger(__name__)


# ===================================================================
#  Win32 Constants
# ===================================================================

WM_GETTEXT = 0x000D
WM_GETTEXTLENGTH = 0x000E
WM_CLOSE = 0x0010
WM_CANCELMODE = 0x001F
BM_CLICK = 0x00F5
DIALOG_CLASS = "#32770"

# Button labels we'll click to dismiss dialogs (case-insensitive matching)
DISMISS_BUTTONS = {"ok", "end", "close", "abort", "cancel", "no", "yes"}


def _normalize_button_text(text: str) -> str:
    """Remove Win32 mnemonic markers and punctuation from a button label."""
    return (text or "").replace("&", "").strip().lower().rstrip(".")


# ===================================================================
#  Win32 API Wrappers (ctypes -- no pywin32 dependency for the watchdog)
# ===================================================================

if IS_WINDOWS:
    user32 = ctypes.windll.user32

    # Type aliases
    HWND = ctypes.wintypes.HWND
    BOOL = ctypes.wintypes.BOOL
    LPARAM = ctypes.wintypes.LPARAM
    EnumWindowsProc = ctypes.WINFUNCTYPE(BOOL, HWND, LPARAM)
    EnumChildProc = ctypes.WINFUNCTYPE(BOOL, HWND, LPARAM)

    # ctypes otherwise assumes 32-bit integer arguments, which can truncate
    # HWNDs and text-buffer pointers in 64-bit Python.
    user32.GetWindowTextLengthW.argtypes = [HWND]
    user32.GetWindowTextLengthW.restype = ctypes.c_int
    user32.GetWindowTextW.argtypes = [HWND, ctypes.c_wchar_p, ctypes.c_int]
    user32.GetWindowTextW.restype = ctypes.c_int
    user32.GetClassNameW.argtypes = [HWND, ctypes.c_wchar_p, ctypes.c_int]
    user32.GetClassNameW.restype = ctypes.c_int
    user32.SendMessageTimeoutW.argtypes = [
        HWND, ctypes.wintypes.UINT, ctypes.c_size_t,
        ctypes.c_ssize_t, ctypes.wintypes.UINT, ctypes.wintypes.UINT,
        ctypes.POINTER(ctypes.c_size_t),
    ]
    user32.SendMessageTimeoutW.restype = ctypes.wintypes.LPARAM
else:
    user32 = None
    HWND = BOOL = LPARAM = None
    EnumWindowsProc = EnumChildProc = None


def _get_window_text(hwnd: int) -> str:
    """Get the text/title of a window by handle."""
    length = user32.GetWindowTextLengthW(hwnd)
    if length == 0:
        return ""
    buf = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, buf, length + 1)
    return buf.value


def _get_window_class(hwnd: int) -> str:
    """Get the window class name."""
    buf = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, buf, 256)
    return buf.value


def _get_control_text(hwnd: int) -> str:
    """Get control text without allowing a hung window to block the watchdog."""
    result_len = ctypes.c_size_t(0)
    res = user32.SendMessageTimeoutW(
        hwnd,
        WM_GETTEXTLENGTH,
        0,
        0,
        0x0002,  # SMTO_ABORTIFHUNG
        250,     # 250ms timeout
        ctypes.byref(result_len),
    )
    text = ""
    if res != 0 and result_len.value:
        length = result_len.value
        buf = ctypes.create_unicode_buffer(length + 1)
        result_text = ctypes.c_size_t(0)
        res = user32.SendMessageTimeoutW(
            hwnd, WM_GETTEXT, length + 1, ctypes.addressof(buf),
            0x0002, 250, ctypes.byref(result_text),
        )
        if res != 0:
            text = buf.value

    # Some VBA dialog controls expose text only through GetWindowTextW.
    if not text:
        text = _get_window_text(hwnd)
    return _normalize_text(text)


def _normalize_text(text: str) -> str:
    """Normalize mixed newlines while retaining the original multiline layout."""
    return (text or "").replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")


def _click_button(hwnd: int):
    """Click a button control by sending BM_CLICK with timeout protection."""
    import ctypes
    result = ctypes.c_size_t(0)
    user32.SendMessageTimeoutW(
        hwnd,
        BM_CLICK,
        0,
        0,
        0x0002,  # SMTO_ABORTIFHUNG
        250,     # 250ms timeout
        ctypes.byref(result),
    )


def _is_window_visible(hwnd: int) -> bool:
    """Check if a window is visible."""
    return bool(user32.IsWindowVisible(hwnd))


def _close_window(hwnd: int):
    """Send WM_CLOSE to a window."""
    return bool(user32.PostMessageW(hwnd, WM_CLOSE, 0, 0))


def _get_window_pid(hwnd: int) -> tuple:
    """Get the thread ID and process ID for a window handle."""
    pid = ctypes.wintypes.DWORD()
    tid = user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    return tid, pid.value


def _get_code_pane_selection(pane) -> tuple:
    """Read a VBE selection across tuple-return and by-reference COM variants."""
    try:
        selection = pane.GetSelection()
        if selection is not None and len(selection) == 4:
            return tuple(int(value) for value in selection)
    except TypeError:
        pass

    import pythoncom
    from win32com.client import VARIANT
    refs = [VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0) for _ in range(4)]
    pane.GetSelection(*refs)
    return tuple(int(ref.value) for ref in refs)


def _populate_compile_location(vbe, wb, result: dict) -> None:
    """Populate the selected compile-error location without showing the VBE."""
    panes = []
    try:
        if vbe.ActiveCodePane is not None:
            panes.append(vbe.ActiveCodePane)
    except Exception:
        pass
    try:
        selected = vbe.SelectedVBComponent
        if selected is not None and selected.CodeModule.CodePane is not None:
            panes.append(selected.CodeModule.CodePane)
    except Exception:
        pass
    if not panes:
        try:
            for component in wb.VBProject.VBComponents:
                try:
                    pane = component.CodeModule.CodePane
                    if pane is not None:
                        panes.append(pane)
                except Exception:
                    continue
        except Exception:
            pass

    best = None
    for pane in panes:
        try:
            selection = _get_code_pane_selection(pane)
            # A compile error normally selects a token beyond the default
            # 1,1 caret. Prefer that pane while retaining a final fallback.
            candidate = (selection != (1, 1, 1, 1), pane, selection)
            if best is None or candidate[0] > best[0]:
                best = candidate
        except Exception:
            continue
    if best is None or not best[0]:
        return

    _, pane, selection = best
    module = pane.CodeModule
    error_line, error_column = selection[0], selection[1]
    result["error_module"] = module.Name
    result["error_line"] = error_line
    result["error_column"] = error_column
    start = max(1, error_line - 5)
    end = min(module.CountOfLines, error_line + 5)
    result["error_context"] = [
        f"{'>>> ' if line == error_line else '    '}{line:4d}: {module.Lines(line, 1)}"
        for line in range(start, end + 1)
    ]


def _populate_static_compile_location(wb, result: dict) -> None:
    """Fallback to UV001 analysis when hidden VBE exposes no selection."""
    from xlvbatools.analysis.project_context import (
        VBAModuleSource,
        build_project_index,
    )
    from xlvbatools.analysis.rules import run_all_rules
    from xlvbatools.vba.manifest import get_type_info

    modules = []
    sources = []
    for component in wb.VBProject.VBComponents:
        try:
            module = component.CodeModule
            count = module.CountOfLines
            code = module.Lines(1, count) if count else ""
            lines = code.splitlines(keepends=True)
            modules.append((component.Name, module, count, lines))
            type_info = get_type_info(component.Type)
            sources.append(
                VBAModuleSource.create(
                    name=component.Name,
                    rel_path=f"{type_info['dir']}/{component.Name}{type_info['ext']}",
                    module_kind=type_info["name"],
                    lines=lines,
                )
            )
        except Exception:
            continue

    project_index = build_project_index(sources)
    for component_name, module, count, lines in modules:
        try:
            if not count:
                continue
            source = next(
                (
                    item for item in sources
                    if item.name.casefold() == component_name.casefold()
                ),
                None,
            )
            rel_path = source.rel_path if source is not None else component_name
            issues = run_all_rules(
                rel_path,
                lines,
                project_index=project_index,
            )
            issue = next(
                (item for item in issues if item.rule_id == "UV001" and item.severity == "ERROR"),
                None,
            )
            if issue is None:
                continue
            line_text = module.Lines(issue.line_num, 1)
            match = re.search(r"Undeclared variable '([^']+)'", issue.message)
            column = 1
            if match:
                position = line_text.lower().find(match.group(1).lower())
                if position >= 0:
                    column = position + 1
            result["error_module"] = component_name
            result["error_line"] = issue.line_num
            result["error_column"] = column
            start = max(1, issue.line_num - 5)
            end = min(count, issue.line_num + 5)
            result["error_context"] = [
                f"{'>>> ' if line == issue.line_num else '    '}{line:4d}: {module.Lines(line, 1)}"
                for line in range(start, end + 1)
            ]
            return
        except Exception:
            continue


def _hide_vbe_ui(excel) -> None:
    """Cancel transient VBE command menus and keep the editor headless."""
    try:
        window = excel.VBE.MainWindow
        result = ctypes.c_size_t(0)
        user32.SendMessageTimeoutW(
            int(window.HWnd), WM_CANCELMODE, 0, 0,
            0x0002, 250, ctypes.byref(result),
        )
        window.Visible = False
        del window
    except Exception:
        pass


# ===================================================================
#  Dialog Event
# ===================================================================

@dataclass
class DialogControl:
    hwnd: int
    class_name: str
    text: str

    def to_dict(self) -> dict:
        return {"hwnd": self.hwnd, "class_name": self.class_name, "text": self.text}


@dataclass
class DialogEvent:
    """A captured dialog event with full diagnostics."""

    timestamp: float
    hwnd: int
    title: str
    dialog_type: str  # "compile_error", "runtime_error", "msgbox", "file_dialog", "alert", "unknown"
    texts: List[str] = field(default_factory=list)
    controls: List[DialogControl] = field(default_factory=list)
    sequence: int = 0
    button_clicked: str = ""
    dismissed: bool = False

    @property
    def text(self) -> str:
        """Concatenated dialog text for display."""
        return " | ".join(t for t in self.texts if t)

    def to_dict(self) -> dict:
        return {
            "timestamp": self.timestamp,
            "hwnd": self.hwnd,
            "title": self.title,
            "type": self.dialog_type,
            "text": self.text,
            "texts": self.texts,
            "controls": [control.to_dict() for control in self.controls],
            "sequence": self.sequence,
            "button_clicked": self.button_clicked,
            "dismissed": self.dismissed,
        }

    def __str__(self) -> str:
        status = "DISMISSED" if self.dismissed else "CAPTURED"
        return f"[{status}] [{self.dialog_type}] {self.title}: {self.text}"


# ===================================================================
#  Dialog Classifier
# ===================================================================

_COMPILE_RE = re.compile(r"compile\s+error", re.IGNORECASE)
_RUNTIME_RE = re.compile(r"run.?time\s+error", re.IGNORECASE)
_VB_TITLE_RE = re.compile(r"microsoft\s+visual\s+basic", re.IGNORECASE)
_EXCEL_TITLE_RE = re.compile(r"microsoft\s+excel", re.IGNORECASE)


def _classify_dialog(title: str, texts: List[str]) -> str:
    """Classify a dialog based on its title and content."""
    all_text = " ".join([title] + texts).lower()

    if _COMPILE_RE.search(all_text):
        return "compile_error"
    if _RUNTIME_RE.search(all_text):
        return "runtime_error"
    if _VB_TITLE_RE.search(title):
        return "vb_error"
    if "save" in all_text and ("file" in all_text or "changes" in all_text):
        return "save_dialog"
    if "open" in all_text and "file" in all_text:
        return "file_dialog"
    if _EXCEL_TITLE_RE.search(title):
        return "excel_alert"
    return "unknown"


# ===================================================================
#  Dialog Watchdog Thread
# ===================================================================

class DialogWatchdog:
    """
    Background thread that polls for Win32 dialog windows (#32770 class),
    captures their text content, and automatically dismisses them.

    This prevents COM automation hangs caused by modal dialogs that block
    the Excel thread and wait for user input that never comes in headless mode.

    Parameters
    ----------
    poll_interval : float
        Seconds between each poll cycle (default 0.25s -- 4 polls/sec).
    timeout : float
        Maximum lifetime of the watchdog thread (default 300s = 5 min).
    auto_dismiss : bool
        Whether to automatically dismiss found dialogs (default True).
    dismiss_strategy : str
        Which button to prefer: "ok" (click OK/Close), "cancel" (click Cancel/No),
        or "close" (send WM_CLOSE). Default is "ok".
    on_dialog : callable, optional
        Optional callback fired when a dialog is detected, receives DialogEvent.
    """

    def __init__(
        self,
        poll_interval: float = 0.25,
        timeout: float = 300.0,
        auto_dismiss: bool = True,
        dismiss_strategy: str = "ok",
        on_dialog: Optional[Callable] = None,
        target_pid: Optional[int] = None,
        capture_attempts: int = 3,
        capture_retry_delay: float = 0.075,
    ):
        from xlvbatools._compat import require_windows
        require_windows("DialogWatchdog")

        self.poll_interval = poll_interval
        self.timeout = timeout
        self.auto_dismiss = auto_dismiss
        self.dismiss_strategy = dismiss_strategy
        self.on_dialog = on_dialog
        self.target_pid = target_pid
        if self.auto_dismiss and self.target_pid is None:
            raise ValueError("An auto-dismissing watchdog requires target_pid.")
        self.capture_attempts = max(1, capture_attempts)
        self.capture_retry_delay = max(0.0, capture_retry_delay)

        self._events: List[DialogEvent] = []
        self._lock = threading.Lock()
        self._stop_event = threading.Event()
        self._thread: Optional[threading.Thread] = None
        self._seen_hwnds: set = set()
        self._next_sequence = 1

    @property
    def events(self) -> List[DialogEvent]:
        """Thread-safe access to captured dialog events."""
        with self._lock:
            return list(self._events)

    @property
    def had_dialogs(self) -> bool:
        """True if any dialog was captured during this watchdog session."""
        with self._lock:
            return len(self._events) > 0

    @property
    def had_errors(self) -> bool:
        """True if any compile or runtime error dialog was captured."""
        with self._lock:
            return any(
                e.dialog_type in ("compile_error", "runtime_error", "vb_error")
                for e in self._events
            )

    @property
    def error_summary(self) -> str:
        """Human-readable summary of all error dialogs, or empty string."""
        with self._lock:
            errors = [
                e for e in self._events
                if e.dialog_type in ("compile_error", "runtime_error", "vb_error")
            ]
        if not errors:
            return ""
        lines = []
        for e in errors:
            lines.append(f"  [{e.dialog_type}] {e.title}")
            if e.text:
                lines.append(f"    Text: {e.text}")
        return "\n".join(lines)

    def start(self):
        """Start the watchdog background thread."""
        if self._thread is not None and self._thread.is_alive():
            return  # Already running

        self._stop_event.clear()
        self._seen_hwnds.clear()
        with self._lock:
            self._events.clear()
        self._next_sequence = 1

        self._thread = threading.Thread(
            target=self._poll_loop,
            name="DialogWatchdog",
            daemon=True,
        )
        self._thread.start()
        logger.debug("Dialog watchdog started")

    def stop(self) -> List[DialogEvent]:
        """Stop the watchdog and return all captured events."""
        self._stop_event.set()
        if self._thread is not None:
            self._thread.join(timeout=2.0)
            self._thread = None
        logger.debug(f"Dialog watchdog stopped ({len(self._events)} events)")
        return self.events

    def _poll_loop(self):
        """Main polling loop -- runs on background thread."""
        start_time = time.time()

        while not self._stop_event.is_set():
            # Timeout guard
            if time.time() - start_time > self.timeout:
                logger.warning(f"Dialog watchdog timed out after {self.timeout}s")
                break

            try:
                self._scan_for_dialogs()
            except Exception as e:
                logger.debug(f"Watchdog scan error (non-fatal): {e}")

            self._stop_event.wait(self.poll_interval)

    def _scan_for_dialogs(self):
        """Scan all top-level windows for dialog class (#32770)."""
        dialog_hwnds = []

        def _enum_callback(hwnd, _lparam):
            try:
                if not _is_window_visible(hwnd):
                    return True
                cls = _get_window_class(hwnd)
                if cls == DIALOG_CLASS and hwnd not in self._seen_hwnds:
                    # Filter by process if target_pid is set
                    if self.target_pid is not None:
                        _, pid = _get_window_pid(hwnd)
                        if pid != self.target_pid:
                            return True
                    title = _get_window_text(hwnd)
                    # Filter for Excel/VB related dialogs
                    if self._is_excel_dialog(title):
                        dialog_hwnds.append(hwnd)
            except Exception:
                pass
            return True

        callback = EnumWindowsProc(_enum_callback)
        user32.EnumWindows(callback, 0)

        for hwnd in dialog_hwnds:
            self._handle_dialog(hwnd)

    def _is_excel_dialog(self, title: str) -> bool:
        """Check if a dialog title indicates it belongs to Excel or VBE."""
        if not title:
            return True  # Untitled dialogs from Excel are common
        title_lower = title.lower()
        keywords = ["microsoft", "excel", "visual basic", "vba", "compile",
                     "error", "save", "open", "warning"]
        return any(k in title_lower for k in keywords)

    def _handle_dialog(self, hwnd: int):
        """Capture dialog content and optionally dismiss it."""
        title = _get_window_text(hwnd)
        texts, buttons, controls = [], [], []
        seen_texts, seen_controls = set(), set()

        for attempt in range(self.capture_attempts):
            def _enum_children(child_hwnd, _lparam):
                try:
                    child_class = _get_window_class(child_hwnd)
                    child_text = _get_control_text(child_hwnd)
                    key = (int(child_hwnd), child_class, child_text)
                    if key not in seen_controls:
                        controls.append(DialogControl(int(child_hwnd), child_class, child_text))
                        seen_controls.add(key)
                    if child_text:
                        if child_class == "Button":
                            button = (int(child_hwnd), child_text)
                            if button not in buttons:
                                buttons.append(button)
                        elif child_text not in seen_texts:
                            texts.append(child_text)
                            seen_texts.add(child_text)
                except Exception as error:
                    logger.debug(f"Error capturing child control {child_hwnd}: {error}")
                return True

            try:
                callback = EnumChildProc(_enum_children)
                user32.EnumChildWindows(hwnd, callback, 0)
            except Exception as e:
                logger.debug(f"Error enumerating dialog children: {e}")
            if texts or attempt == self.capture_attempts - 1:
                break
            self._stop_event.wait(self.capture_retry_delay)

        # Mark seen only after all bounded capture attempts complete.
        self._seen_hwnds.add(hwnd)

        # Classify
        dialog_type = _classify_dialog(title, texts)

        # Create event
        event = DialogEvent(
            timestamp=time.time(),
            hwnd=hwnd,
            title=title,
            dialog_type=dialog_type,
            texts=texts,
            controls=controls,
            sequence=self._next_sequence,
        )
        self._next_sequence += 1

        # Log immediately
        logger.warning(f"Dialog detected: {event}")

        # Dismiss if requested
        if self.auto_dismiss:
            event.dismissed = self._dismiss_dialog(hwnd, buttons, event)

        # Store event
        with self._lock:
            self._events.append(event)

        # Fire callback
        if self.on_dialog:
            try:
                self.on_dialog(event)
            except Exception as e:
                logger.debug(f"on_dialog callback error: {e}")

    def _dismiss_dialog(self, hwnd: int, buttons: list, event: DialogEvent) -> bool:
        """Attempt to dismiss a dialog window. Returns True if successful."""
        # Strategy 1: Click a known dismiss button
        preferred_order = self._get_button_priority()
        for preferred in preferred_order:
            for btn_hwnd, btn_text in buttons:
                if _normalize_button_text(btn_text) in preferred:
                    try:
                        _click_button(btn_hwnd)
                        event.button_clicked = btn_text
                        logger.info(f"Clicked '{btn_text}' on dialog: {event.title}")
                        return True
                    except Exception as e:
                        logger.debug(f"Failed to click '{btn_text}': {e}")

        # Strategy 2: Click ANY button we recognize
        for btn_hwnd, btn_text in buttons:
            if _normalize_button_text(btn_text) in DISMISS_BUTTONS:
                try:
                    _click_button(btn_hwnd)
                    event.button_clicked = btn_text
                    logger.info(f"Clicked fallback '{btn_text}' on dialog: {event.title}")
                    return True
                except Exception as e:
                    logger.debug(f"Failed to click fallback '{btn_text}': {e}")

        # Strategy 3: Send WM_CLOSE
        try:
            if _close_window(hwnd):
                event.button_clicked = "[WM_CLOSE]"
                logger.info(f"Sent WM_CLOSE to dialog: {event.title}")
                return True
            logger.debug(f"PostMessageW rejected WM_CLOSE for dialog: {event.title}")
            return False
        except Exception as e:
            logger.debug(f"Failed to send WM_CLOSE: {e}")
            return False

    def _get_button_priority(self) -> list:
        """Get button preference order based on dismiss strategy."""
        if self.dismiss_strategy == "cancel":
            return [{"cancel", "no", "abort"}, {"close", "end"}, {"ok", "yes"}]
        elif self.dismiss_strategy == "close":
            return [{"close", "end"}, {"ok"}, {"cancel"}]
        else:  # "ok" (default)
            return [{"ok"}, {"close", "end", "yes"}, {"cancel", "no"}]


# ===================================================================
#  VBE Compile Test with Dialog Protection
# ===================================================================

def compile_test_with_watchdog(excel, wb, watchdog: Optional[DialogWatchdog] = None) -> dict:
    """
    Trigger VBE compilation and capture any compile error dialogs.

    Inspects VBE compile control ID 578, then forces project compilation by
    running a temporary no-op probe procedure. This avoids the visible menu
    flash caused by executing VBE CommandBar controls. If a compile error is
    found, the function reads VBE selection or static Option Explicit evidence
    to identify the exact module and line.

    Parameters
    ----------
    excel : COM Excel.Application
        Active Excel application object.
    wb : COM Workbook
        Open workbook to compile.
    watchdog : DialogWatchdog, optional
        If None, creates a temporary one.

    Returns
    -------
    dict
        Result with keys:
        - success: bool
        - errors: list of error dicts
        - error_module: str (module name if compile error found)
        - error_line: int (line number if compile error found)
        - error_context: list of str (surrounding code lines)
    """
    own_watchdog = watchdog is None
    if own_watchdog:
        _, target_pid = _get_window_pid(excel.Hwnd)
        watchdog = DialogWatchdog(poll_interval=0.2, timeout=30.0, target_pid=target_pid)
        watchdog.start()

    result = {
        "success": True,
        "already_compiled": False,
        "errors": [],
        "error_module": "",
        "error_line": 0,
        "error_column": 0,
        "error_context": [],
        "target_project": "",
        "active_project": "",
        "target_project_file": "",
        "active_project_file": "",
        "compile_verified": True,
        "warnings": [],
    }
    probe_component = None

    try:
        # Make VBE available (required for compile button)
        vbe = excel.VBE

        # CommandBars control 578 compiles the active VB project, which is not
        # necessarily the workbook passed by the caller. Activate exactly one
        # component to target this project without walking every component or
        # making the VBE visible.
        wb.Activate()
        try:
            active_file = os.path.abspath(vbe.ActiveVBProject.FileName)
        except Exception:
            active_file = ""
        if os.path.normcase(active_file) != os.path.normcase(os.path.abspath(wb.FullName)):
            target_component = wb.VBProject.VBComponents.Item(1)
            target_component.Activate()
            del target_component
        result["target_project"] = wb.VBProject.Name
        result["target_project_file"] = wb.VBProject.FileName
        try:
            result["active_project"] = vbe.ActiveVBProject.Name
            result["active_project_file"] = vbe.ActiveVBProject.FileName
        except Exception:
            result["active_project"] = ""

        # Inspect control 578 for compatibility, but do not execute it: VBE's
        # CommandBar API can flash a File menu even while MainWindow is hidden.
        compile_btn = vbe.CommandBars.FindControl(Id=578)
        if compile_btn is None:
            result["errors"].append({
                "type": "warning",
                "message": "VBE compile button (ID=578) not found. Skipping compile test."
            })
            return result

        if not compile_btn.Enabled:
            result["already_compiled"] = True
            return result

        pre_sequence = max((event.sequence for event in watchdog.events), default=0)

        # Adding and running a no-op procedure forces VBA to compile the active
        # project without invoking any VBE command-bar UI. The component is
        # removed before returning and is never intentionally persisted.
        suffix = str(int(time.time() * 1000000))[-8:]
        module_name = f"modXlvbaProbe{suffix}"
        procedure_name = f"XlvbaCompileProbe{suffix}"
        probe_component = wb.VBProject.VBComponents.Add(1)
        probe_component.Name = module_name
        probe_component.CodeModule.AddFromString(
            f"Option Explicit\r\nPublic Sub {procedure_name}()\r\nEnd Sub\r\n"
        )
        _hide_vbe_ui(excel)
        probe_error = None
        try:
            escaped_name = wb.Name.replace("'", "''")
            excel.Run(f"'{escaped_name}'!{procedure_name}")
        except Exception as error:
            probe_error = error
        _hide_vbe_ui(excel)
        compile_incomplete = bool(compile_btn.Enabled)
        deadline = time.time() + 1.0
        new_error_events = []
        while time.time() < deadline:
            _hide_vbe_ui(excel)
            new_error_events = [
                event for event in watchdog.events
                if event.sequence > pre_sequence and event.dialog_type in
                ("compile_error", "runtime_error", "vb_error")
            ]
            if new_error_events:
                break
            time.sleep(0.05)

        try:
            wb.VBProject.VBComponents.Remove(probe_component)
            probe_component = None
        except Exception as error:
            result["warnings"].append(f"Could not remove temporary compile probe: {error}")

        # A successful compile disables control 578. If it remains enabled,
        # VBE stopped at an error even when no modal dialog was raised (a
        # common behavior while the VBE window is hidden).
        if new_error_events or compile_incomplete:
            try:
                _populate_compile_location(vbe, wb, result)
                if not result["error_line"]:
                    _populate_static_compile_location(wb, result)
            except Exception as e:
                logger.debug(f"Could not read VBE error location: {e}")

        if new_error_events or result["error_line"]:
            result["success"] = False
            result["errors"].extend(event.to_dict() for event in new_error_events)
            if not new_error_events:
                result["errors"].append({
                    "type": "compile_error",
                    "message": "VBE compilation did not complete; compile control remained enabled.",
                })
        elif compile_incomplete:
            result["compile_verified"] = False
            result["warnings"].append(
                "VBE compile control remained enabled, but no compile error dialog or static Option Explicit failure was found."
            )
        elif probe_error is not None:
            result["success"] = False
            result["errors"].append({
                "type": "exception",
                "message": f"Compile probe failed without a captured compile error: {probe_error}",
            })

    except Exception as e:
        result["success"] = False
        result["errors"].append({
            "type": "exception",
            "message": f"Compile test raised exception: {e}",
        })
    finally:
        _hide_vbe_ui(excel)
        if probe_component is not None:
            try:
                wb.VBProject.VBComponents.Remove(probe_component)
            except Exception:
                pass
        if own_watchdog:
            watchdog.stop()

    return result
