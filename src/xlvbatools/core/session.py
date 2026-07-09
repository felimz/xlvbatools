"""
Excel COM Session Context Manager
===================================
Provides a safe, reusable context manager for all Excel COM operations.
Handles process cleanup, workbook opening, dialog protection, and graceful shutdown.

The session integrates the DialogWatchdog to automatically capture and dismiss
any pop-up dialogs (compile errors, MsgBox calls, runtime errors, file dialogs)
that would otherwise hang COM calls indefinitely.

Usage:
    from xlvbatools import ExcelSession

    with ExcelSession("path/to/workbook.xlsm") as session:
        # Access Excel application and workbook
        excel, wb = session.excel, session.wb

        # Run a macro safely (watchdog catches any dialogs)
        result = session.run_macro("MyMacro")

        # Check if any dialogs were captured
        if session.had_errors:
            print(session.error_summary)

        # Compile test with error location
        result = session.compile_test()
        if not result["success"]:
            for line in result["error_context"]:
                print(line)
"""

import gc
import os
import time
import logging
from typing import Any, Optional

from xlvbatools._compat import require_windows
from xlvbatools.core.process import (
    kill_excel,
    is_excel_running,
    close_excel_gracefully,
    is_process_running,
    kill_process_by_pid,
)

logger = logging.getLogger(__name__)


class ExcelSession:
    workbook_path: str
    visible: bool
    save_on_exit: bool
    kill_on_enter: bool
    init_delay: float
    enable_watchdog: bool
    watchdog_poll_interval: float
    excel: Any
    wb: Any
    watchdog: Optional["DialogWatchdog"]
    excel_pid: Optional[int]
    """
    Context manager for safe Excel COM automation sessions with dialog protection.

    The watchdog thread runs in the background, polling for Win32 dialog windows
    every 250ms. When it finds one, it captures the dialog text (for diagnostics),
    then automatically dismisses it by clicking OK/Close/Cancel -- preventing the
    COM call from hanging forever.

    Parameters
    ----------
    workbook_path : str
        Path to the .xlsm workbook (relative or absolute).
    visible : bool
        Whether to make Excel visible (default False for headless).
    save_on_exit : bool
        Whether to save the workbook on context exit (default True).
    kill_on_enter : bool
        Whether to kill stale Excel processes on entry (default True).
    init_delay : float
        Seconds to wait after opening for VBProject initialization (default 1.5).
    enable_watchdog : bool
        Whether to start the dialog watchdog (default True). Disable only for
        debugging when you want to see dialogs manually.
    watchdog_poll_interval : float
        How often (seconds) the watchdog polls for dialogs (default 0.25).
    """

    def __init__(
        self,
        workbook_path: str,
        visible: bool = False,
        save_on_exit: bool = True,
        kill_on_enter: bool = True,
        init_delay: float = 1.5,
        enable_watchdog: bool = True,
        watchdog_poll_interval: float = 0.25,
    ):
        require_windows("ExcelSession")

        self.workbook_path = os.path.abspath(workbook_path)
        self.visible = visible
        self.save_on_exit = save_on_exit
        self.kill_on_enter = kill_on_enter
        self.init_delay = init_delay
        self.enable_watchdog = enable_watchdog
        self.watchdog_poll_interval = watchdog_poll_interval

        self.excel = None
        self.wb = None
        self.watchdog = None
        self.excel_pid = None

    def __enter__(self):
        import win32com.client

        if self.kill_on_enter:
            if is_excel_running():
                # Attempt graceful closure of the target workbook first
                success = close_excel_gracefully(target_path=self.workbook_path)
                # If graceful closure failed, fall back to force-kill
                if not success:
                    logger.info("Graceful closure of target workbook failed. Force-killing EXCEL.EXE...")
                    kill_excel()

        if not os.path.exists(self.workbook_path):
            raise FileNotFoundError(f"Workbook not found: {self.workbook_path}")

        # Start the dialog watchdog BEFORE opening Excel
        if self.enable_watchdog:
            from xlvbatools.core.watchdog import DialogWatchdog
            self.watchdog = DialogWatchdog(
                poll_interval=self.watchdog_poll_interval,
                # 600s (10 min) instead of watchdog default (300s) because
                # sessions with macro execution may run longer than 5 minutes.
                timeout=600.0,
                auto_dismiss=True,
                dismiss_strategy="ok",
            )
            self.watchdog.start()

        logger.info(f"Opening Excel session: {self.workbook_path}")
        # Use DispatchEx to ensure a new, isolated Excel instance is spawned
        self.excel = win32com.client.DispatchEx("Excel.Application")
        self.excel.Visible = self.visible
        self.excel.DisplayAlerts = False
        self.excel.AutomationSecurity = 1  # Enable macros

        # Retrieve the PID of our spawned Excel instance
        try:
            import win32process
            _, self.excel_pid = win32process.GetWindowThreadProcessId(self.excel.Hwnd)
            logger.info(f"Spawned Excel process with PID: {self.excel_pid}")
            # Wire PID into watchdog so it only monitors our Excel instance
            if self.watchdog is not None:
                self.watchdog.target_pid = self.excel_pid
        except Exception as e:
            self.excel_pid = None
            logger.warning(f"Could not retrieve Excel PID: {e}")

        self.wb = self.excel.Workbooks.Open(self.workbook_path, UpdateLinks=0)
        time.sleep(self.init_delay)

        # Check if any dialogs appeared during open
        if self.watchdog and self.watchdog.had_errors:
            logger.warning(f"Dialogs during workbook open:\n{self.watchdog.error_summary}")

        logger.info(f"Workbook opened: {os.path.basename(self.workbook_path)}")
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        # Stop watchdog first
        if self.watchdog is not None:
            events = self.watchdog.stop()
            if events:
                logger.info(f"Watchdog captured {len(events)} dialog(s) during session:")
                for event in events:
                    logger.info(f"  {event}")

        try:
            if self.wb is not None:
                if self.save_on_exit and exc_type is None:
                    self.wb.Save()
                    logger.info("Workbook saved")
                self.wb.Close(False)
                pid_suffix = f" (PID: {self.excel_pid})" if self.excel_pid else ""
                logger.info(f"Workbook closed{pid_suffix}")
        except Exception as e:
            logger.warning(f"Error closing workbook: {e}")
        finally:
            self.wb = None
            gc.collect()

        try:
            if self.excel is not None:
                self.excel.Quit()
                pid_suffix = f" (PID: {self.excel_pid})" if self.excel_pid else ""
                logger.info(f"Excel quit{pid_suffix}")
        except Exception as e:
            logger.warning(f"Error quitting Excel: {e}")
        finally:
            self.excel = None
            gc.collect()

        # Final safety: wait a moment for Excel to exit. We only force-kill if there was an exception
        # (which could indicate a hang) or if the process remains active after a grace period.
        # Note: If the caller holds references to COM objects (e.g. from local variables),
        # Excel cannot exit until those references are cleared by GC. Force-killing too early
        # will trigger RPC server errors (0x800706ba) during GC.
        if self.excel_pid is not None:
            grace_period = 2.0 if exc_type is None else 0.5
            exited = False
            for _ in range(int(grace_period / 0.1)):
                if not is_process_running(self.excel_pid):
                    exited = True
                    break
                time.sleep(0.1)

            if not exited and (exc_type is not None or not self.kill_on_enter):
                logger.info(f"Excel process PID {self.excel_pid} did not exit cleanly. Force-killing spawned process...")
                kill_process_by_pid(self.excel_pid)

        return False  # Don't suppress exceptions

    # -- Convenience Properties --

    @property
    def had_dialogs(self) -> bool:
        """True if any dialog was captured during this session."""
        return self.watchdog is not None and self.watchdog.had_dialogs

    @property
    def had_errors(self) -> bool:
        """True if any compile/runtime error dialog was captured."""
        return self.watchdog is not None and self.watchdog.had_errors

    @property
    def error_summary(self) -> str:
        """Human-readable summary of error dialogs, or empty string."""
        if self.watchdog is None:
            return ""
        return self.watchdog.error_summary

    @property
    def dialog_events(self) -> list:
        """All captured dialog events."""
        if self.watchdog is None:
            return []
        return self.watchdog.events

    @property
    def vb_project(self):
        """
        Get the workbook's VBProject object with programmatic access trust checks.
        """
        if self.wb is None:
            raise ValueError("Workbook is not open")
        try:
            return self.wb.VBProject
        except Exception as e:
            err_str = str(e).lower()
            if "programmatic access" in err_str or "not trusted" in err_str or "0x800a03ec" in err_str:
                raise RuntimeError(
                    "Programmatic access to Visual Basic Project is not trusted in Excel.\n"
                    "To enable this, go to: Excel Options -> Trust Center -> Trust Center Settings "
                    "-> Macro Settings -> check 'Trust access to the VBA project object model'."
                ) from e
            raise

    # -- Macro Execution with Dialog Protection --

    def run_macro(self, macro_name: str, timeout: float = 120.0) -> dict:
        """
        Run a VBA macro with full dialog protection.

        If the macro triggers a dialog (compile error, MsgBox, runtime error),
        the watchdog dismisses it and the method returns with error details
        instead of hanging forever.

        Note: The ``timeout`` parameter is reserved for future use. Excel's
        ``Application.Run`` is a blocking COM call that cannot be externally
        interrupted. The dialog watchdog handles the most common hang scenario
        (modal dialogs blocking the COM thread).

        Returns
        -------
        dict
            Keys: success, elapsed_seconds, error, dialog_events
        """
        pre_count = 0
        if self.watchdog:
            pre_count = len(self.watchdog.events)

        t0 = time.time()
        try:
            self.excel.Run(macro_name)
            elapsed = time.time() - t0

            # Check if watchdog caught anything during the run
            new_events = []
            if self.watchdog:
                all_events = self.watchdog.events
                new_events = all_events[pre_count:]

            result = {
                "success": not any(
                    e.dialog_type in ("compile_error", "runtime_error", "vb_error")
                    for e in new_events
                ),
                "elapsed_seconds": elapsed,
                "dialog_events": [e.to_dict() for e in new_events],
            }

            if new_events:
                logger.warning(
                    f"{macro_name} completed but {len(new_events)} dialog(s) were dismissed"
                )
                for e in new_events:
                    logger.warning(f"  {e}")

            return result

        except Exception as e:
            elapsed = time.time() - t0
            error_msg = str(e)

            # Enrich error message with watchdog data
            new_events = []
            if self.watchdog:
                all_events = self.watchdog.events
                new_events = all_events[pre_count:]

            if new_events:
                dialog_text = "; ".join(
                    f"[{ev.dialog_type}] {ev.text}" for ev in new_events
                )
                error_msg = f"{error_msg} | Dialogs captured: {dialog_text}"

            logger.error(f"{macro_name} failed after {elapsed:.2f}s: {error_msg}")
            return {
                "success": False,
                "elapsed_seconds": elapsed,
                "error": error_msg,
                "dialog_events": [e.to_dict() for e in new_events],
            }

    # -- Compile Test --

    def compile_test(self) -> dict:
        """
        Trigger VBE compilation and capture any compile error details.

        Returns a dict with:
        - success: bool
        - errors: list of error details
        - error_module: str (module name where error occurred)
        - error_line: int (line number of error)
        - error_context: list[str] (surrounding code with >>> marker)
        """
        from xlvbatools.core.watchdog import compile_test_with_watchdog
        return compile_test_with_watchdog(self.excel, self.wb, self.watchdog)

    # -- Named Range Helpers --

    def set_named_range(self, name: str, value) -> bool:
        """Set a named range value. Returns True on success."""
        try:
            rng = self.wb.Names(name).RefersToRange
            rng.Value = value
            logger.info(f"Set {name} = {value}")
            return True
        except Exception as e:
            logger.warning(f"Could not set named range '{name}': {e}")
            return False

    def get_named_range(self, name: str, default=None):
        """Get a named range value. Returns default if not found."""
        try:
            return self.wb.Names(name).RefersToRange.Value
        except Exception:
            return default
