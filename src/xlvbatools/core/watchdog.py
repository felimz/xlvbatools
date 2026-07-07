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

    watchdog = DialogWatchdog()
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
BM_CLICK = 0x00F5
DIALOG_CLASS = "#32770"

# Button labels we'll click to dismiss dialogs (case-insensitive matching)
DISMISS_BUTTONS = {"ok", "end", "close", "abort", "cancel", "no", "yes"}


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
    """Get text from a control via WM_GETTEXT (works for static labels etc.)."""
    length = user32.SendMessageW(hwnd, WM_GETTEXTLENGTH, 0, 0)
    if length == 0:
        return ""
    buf = ctypes.create_unicode_buffer(length + 1)
    user32.SendMessageW(hwnd, WM_GETTEXT, length + 1, buf)
    return buf.value


def _click_button(hwnd: int):
    """Click a button control by sending BM_CLICK."""
    user32.SendMessageW(hwnd, BM_CLICK, 0, 0)


def _is_window_visible(hwnd: int) -> bool:
    """Check if a window is visible."""
    return bool(user32.IsWindowVisible(hwnd))


def _close_window(hwnd: int):
    """Send WM_CLOSE to a window."""
    user32.PostMessageW(hwnd, WM_CLOSE, 0, 0)


# ===================================================================
#  Dialog Event
# ===================================================================

@dataclass
class DialogEvent:
    """A captured dialog event with full diagnostics."""

    timestamp: float
    hwnd: int
    title: str
    dialog_type: str  # "compile_error", "runtime_error", "msgbox", "file_dialog", "alert", "unknown"
    texts: List[str] = field(default_factory=list)
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
    ):
        from xlvbatools._compat import require_windows
        require_windows("DialogWatchdog")

        self.poll_interval = poll_interval
        self.timeout = timeout
        self.auto_dismiss = auto_dismiss
        self.dismiss_strategy = dismiss_strategy
        self.on_dialog = on_dialog

        self._events: List[DialogEvent] = []
        self._lock = threading.Lock()
        self._stop_event = threading.Event()
        self._thread: Optional[threading.Thread] = None
        self._seen_hwnds: set = set()

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
        self._seen_hwnds.add(hwnd)

        title = _get_window_text(hwnd)
        texts = []
        buttons = []

        # Enumerate all child controls
        def _enum_children(child_hwnd, _lparam):
            try:
                child_class = _get_window_class(child_hwnd)
                child_text = _get_control_text(child_hwnd)

                if child_text:
                    if child_class == "Button":
                        buttons.append((child_hwnd, child_text))
                    else:
                        texts.append(child_text)
            except Exception:
                pass
            return True

        try:
            callback = EnumChildProc(_enum_children)
            user32.EnumChildWindows(hwnd, callback, 0)
        except Exception as e:
            logger.debug(f"Error enumerating dialog children: {e}")

        # Classify
        dialog_type = _classify_dialog(title, texts)

        # Create event
        event = DialogEvent(
            timestamp=time.time(),
            hwnd=hwnd,
            title=title,
            dialog_type=dialog_type,
            texts=texts,
        )

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
                if btn_text.strip().lower().rstrip(".") in preferred:
                    try:
                        _click_button(btn_hwnd)
                        event.button_clicked = btn_text
                        logger.info(f"Clicked '{btn_text}' on dialog: {event.title}")
                        return True
                    except Exception as e:
                        logger.debug(f"Failed to click '{btn_text}': {e}")

        # Strategy 2: Click ANY button we recognize
        for btn_hwnd, btn_text in buttons:
            if btn_text.strip().lower().rstrip(".") in DISMISS_BUTTONS:
                try:
                    _click_button(btn_hwnd)
                    event.button_clicked = btn_text
                    logger.info(f"Clicked fallback '{btn_text}' on dialog: {event.title}")
                    return True
                except Exception as e:
                    logger.debug(f"Failed to click fallback '{btn_text}': {e}")

        # Strategy 3: Send WM_CLOSE
        try:
            _close_window(hwnd)
            event.button_clicked = "[WM_CLOSE]"
            logger.info(f"Sent WM_CLOSE to dialog: {event.title}")
            return True
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

    Uses the VBE CommandBars compile button (ID=578) to force compilation,
    then checks the watchdog for any error dialogs that appeared. If a compile
    error is found, it also reads the VBE active code pane to identify the
    exact module and line of the error.

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
        watchdog = DialogWatchdog(poll_interval=0.2, timeout=30.0)
        watchdog.start()

    result = {
        "success": True,
        "errors": [],
        "error_module": "",
        "error_line": 0,
        "error_context": [],
    }

    try:
        # Make VBE available (required for compile button)
        vbe = excel.VBE

        # Activate all components so VBE loads them
        for comp in wb.VBProject.VBComponents:
            try:
                comp.Activate()
            except Exception:
                pass

        time.sleep(0.5)

        # Find and execute the compile button
        compile_btn = vbe.CommandBars.FindControl(Id=578)
        if compile_btn is None:
            result["errors"].append({
                "type": "warning",
                "message": "VBE compile button (ID=578) not found. Skipping compile test."
            })
            return result

        # Execute compile -- if there's an error, this will trigger a dialog
        compile_btn.Execute()
        time.sleep(2.0)  # Give dialog time to appear

        # Check watchdog for errors
        if watchdog.had_errors:
            result["success"] = False
            for event in watchdog.events:
                if event.dialog_type in ("compile_error", "runtime_error", "vb_error"):
                    result["errors"].append(event.to_dict())

            # Try to read the error location from the active code pane
            try:
                pane = vbe.ActiveCodePane
                if pane is not None:
                    module = pane.CodeModule
                    result["error_module"] = module.Name
                    sel = pane.GetSelection()
                    error_line = sel[0]  # sel is (startLine, startCol, endLine, endCol)
                    result["error_line"] = error_line

                    # Capture surrounding context (+/-5 lines)
                    start = max(1, error_line - 5)
                    end = min(module.CountOfLines, error_line + 5)
                    context = []
                    for i in range(start, end + 1):
                        prefix = ">>> " if i == error_line else "    "
                        line_text = module.Lines(i, 1)
                        context.append(f"{prefix}{i:4d}: {line_text}")
                    result["error_context"] = context
            except Exception as e:
                logger.debug(f"Could not read VBE error location: {e}")

    except Exception as e:
        result["success"] = False
        result["errors"].append({
            "type": "exception",
            "message": f"Compile test raised exception: {e}",
        })
    finally:
        if own_watchdog:
            watchdog.stop()

    return result
