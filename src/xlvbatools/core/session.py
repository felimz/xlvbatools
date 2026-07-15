"""
Excel COM Session Context Manager
===================================
Provides a safe, reusable context manager for all Excel COM operations.
Handles process cleanup, workbook opening, dialog protection, and graceful shutdown.

The session integrates the DialogWatchdog to automatically capture and dismiss
any pop-up dialogs (compile errors, MsgBox calls, runtime errors, file dialogs)
that would otherwise hang COM calls indefinitely.

Internal worker usage:
    from xlvbatools.core.session import ExcelSession

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
import ctypes
import os
import time
import logging
import threading
import uuid
import sys
from typing import Any, Callable, Optional

from xlvbatools._compat import require_windows
from xlvbatools.core.process import (
    is_excel_running,
    close_excel_gracefully,
    is_process_running,
    kill_process_by_pid,
)
from xlvbatools.errors import TrustCenterError

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
        kill_on_enter: bool = False,
        init_delay: float = 1.5,
        enable_watchdog: bool = True,
        watchdog_poll_interval: float = 0.25,
        exit_grace_period: float = 20.0,
        terminate_owned_process: bool = True,
        on_excel_started: Optional[Callable[[int], None]] = None,
        read_only: bool = False,
        disable_macros: bool = False,
    ):
        require_windows("ExcelSession")

        self.workbook_path = os.path.abspath(workbook_path)
        self.visible = visible
        self.save_on_exit = save_on_exit
        self.kill_on_enter = kill_on_enter
        self.init_delay = init_delay
        self.enable_watchdog = enable_watchdog
        self.watchdog_poll_interval = watchdog_poll_interval
        self.exit_grace_period = max(0.0, exit_grace_period)
        self.terminate_owned_process = terminate_owned_process
        self.on_excel_started = on_excel_started
        self.read_only = read_only
        self.disable_macros = disable_macros

        self.excel = None
        self.wb = None
        self.watchdog = None
        self.excel_pid = None
        self.cleanup_result = {
            "pid": None, "quit_requested": False, "exited_gracefully": False,
            "force_terminated": False, "still_running": False,
        }
        self.phase = "session_start"
        self._com_initialized = False
        self._com_thread_id = None

    @staticmethod
    def _thread_has_com_apartment() -> bool:
        """Return whether COM is already initialized on the current thread."""
        try:
            apartment_type = ctypes.c_int()
            qualifier = ctypes.c_int()
            co_get_apartment_type = ctypes.windll.ole32.CoGetApartmentType
            co_get_apartment_type.argtypes = [
                ctypes.POINTER(ctypes.c_int), ctypes.POINTER(ctypes.c_int)
            ]
            co_get_apartment_type.restype = ctypes.c_long
            return co_get_apartment_type(
                ctypes.byref(apartment_type), ctypes.byref(qualifier)
            ) >= 0
        except Exception:
            return False

    def _initialize_com(self) -> None:
        """Enter a balanced STA apartment for this session."""
        if self._thread_has_com_apartment():
            logger.debug(
                f"Using caller-owned COM apartment on thread {threading.get_ident()}"
            )
            return
        import pythoncom

        pythoncom.CoInitialize()
        self._com_initialized = True
        self._com_thread_id = threading.get_ident()
        logger.debug(f"COM initialized on thread {self._com_thread_id}")

    def _uninitialize_com(self) -> None:
        """Leave the session's COM apartment on its owning thread."""
        if not self._com_initialized:
            return
        current_thread = threading.get_ident()
        if current_thread != self._com_thread_id:
            logger.error(
                f"COM teardown thread mismatch: initialized on {self._com_thread_id}, "
                f"exiting on {current_thread}"
            )
            return
        try:
            import pythoncom
            try:
                pythoncom.CoFreeUnusedLibraries()
            finally:
                pythoncom.CoUninitialize()
            logger.debug(f"COM uninitialized on thread {current_thread}")
        finally:
            self._com_initialized = False
            self._com_thread_id = None

    def __enter__(self):
        import win32com.client

        if not os.path.exists(self.workbook_path):
            raise FileNotFoundError(f"Workbook not found: {self.workbook_path}")

        self._initialize_com()

        if self.kill_on_enter:
            if is_excel_running():
                # Attempt graceful closure of the target workbook first
                success = close_excel_gracefully(target_path=self.workbook_path)
                # Never terminate unrelated Excel processes as a fallback.
                if not success:
                    logger.warning("Could not close the stale target workbook; continuing without global termination")

        logger.info(f"Opening Excel session: {self.workbook_path}")
        # Use DispatchEx to ensure a new, isolated Excel instance is spawned
        try:
            self.excel = win32com.client.DispatchEx("Excel.Application")
        except Exception:
            self._uninitialize_com()
            raise

        self.phase = "workbook_open"
        try:
            import win32process
            _, self.excel_pid = win32process.GetWindowThreadProcessId(self.excel.Hwnd)
            if not self.excel_pid:
                raise RuntimeError("Excel returned an invalid process ID")
            logger.info(f"Spawned Excel process with PID: {self.excel_pid}")
        except Exception as e:
            try:
                self.excel.Quit()
            finally:
                self.excel = None
                self._uninitialize_com()
            raise RuntimeError("Could not determine the spawned Excel PID; refusing to start an unscoped watchdog") from e

        if self.enable_watchdog:
            from xlvbatools.core.watchdog import DialogWatchdog
            self.watchdog = DialogWatchdog(
                poll_interval=self.watchdog_poll_interval, timeout=600.0,
                auto_dismiss=True, dismiss_strategy="ok", target_pid=self.excel_pid,
            )
            self.watchdog.start()

        if self.on_excel_started is not None:
            self.on_excel_started(self.excel_pid)

        try:
            self.excel.Visible = self.visible
            self.excel.DisplayAlerts = False
            self.excel.EnableEvents = not self.disable_macros
            self.excel.AskToUpdateLinks = False
            self.excel.AutomationSecurity = 3 if self.disable_macros else 1
            self.wb = self.excel.Workbooks.Open(
                self.workbook_path,
                UpdateLinks=0,
                ReadOnly=self.read_only,
                AddToMru=False,
                IgnoreReadOnlyRecommended=True,
            )
            time.sleep(self.init_delay)
        except Exception:
            self.__exit__(*sys.exc_info())
            raise

        # Check if any dialogs appeared during open
        if self.watchdog and self.watchdog.had_errors:
            logger.warning(f"Dialogs during workbook open:\n{self.watchdog.error_summary}")

        logger.info(f"Workbook opened: {os.path.basename(self.workbook_path)}")
        self.phase = "ready"
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        # Stop watchdog first
        if self.watchdog is not None:
            events = self.watchdog.stop()
            if events:
                logger.info(f"Watchdog captured {len(events)} dialog(s) during session:")
                for event in events:
                    logger.info(f"  {event}")

        # pywin32 wrappers can participate in cycles (especially VBE,
        # Names/RefersToRange, and pytest assertion temporaries). Finalize
        # those proxies while their Excel RPC server is still alive. Waiting
        # until after Workbook.Close/Application.Quit produces misleading
        # 0x800706ba finalizer diagnostics even when cleanup succeeds.
        gc.collect()
        gc.collect()

        close_error = None
        save_error = None
        try:
            if self.wb is not None:
                if self.save_on_exit and exc_type is None:
                    try:
                        self.wb.Save()
                        logger.info("Workbook saved")
                    except Exception as e:
                        save_error = e
                        logger.error(f"Error saving workbook: {e}")
                self.wb.Close(False)
                pid_suffix = f" (PID: {self.excel_pid})" if self.excel_pid else ""
                logger.info(f"Workbook closed{pid_suffix}")
        except Exception as e:
            close_error = e
            logger.warning(f"Error closing workbook: {e}")
        finally:
            self.wb = None
            gc.collect()

        quit_requested = False
        try:
            if self.excel is not None:
                quit_requested = True
                self.excel.Quit()
                pid_suffix = f" (PID: {self.excel_pid})" if self.excel_pid else ""
                logger.info(f"Excel quit{pid_suffix}")
        except Exception as e:
            logger.warning(f"Error quitting Excel: {e}")
        finally:
            self.excel = None
            gc.collect()

        self.cleanup_result.update({"pid": self.excel_pid, "quit_requested": quit_requested})

        # Release this session's COM apartment before waiting for the out-of-
        # process Excel server. When the caller owns the apartment, retain its
        # ownership but still ask pywin32 to release unused COM libraries.
        # Waiting first can leave Excel alive until the grace deadline even
        # though Workbook.Close and Application.Quit both succeeded.
        gc.collect()
        gc.collect()
        if self._com_initialized:
            self._uninitialize_com()
        else:
            try:
                import pythoncom
                pythoncom.CoFreeUnusedLibraries()
            except Exception:
                pass

        if self.excel_pid is not None:
            deadline = time.time() + self.exit_grace_period
            while time.time() < deadline:
                if not is_process_running(self.excel_pid):
                    self.cleanup_result["exited_gracefully"] = True
                    break
                time.sleep(0.1)
            if is_process_running(self.excel_pid) and self.terminate_owned_process:
                logger.warning(f"Excel PID {self.excel_pid} did not exit; terminating the owned process")
                self.cleanup_result["force_terminated"] = kill_process_by_pid(self.excel_pid)
                deadline = time.time() + max(1.0, self.exit_grace_period)
                while time.time() < deadline and is_process_running(self.excel_pid):
                    time.sleep(0.1)
            self.cleanup_result["still_running"] = is_process_running(self.excel_pid)
            if self.cleanup_result["still_running"]:
                logger.error(f"Cleanup failed: owned Excel PID {self.excel_pid} is still running")
        if close_error:
            self.cleanup_result["workbook_close_error"] = str(close_error)
        if save_error:
            self.cleanup_result["workbook_save_error"] = str(save_error)
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
                raise TrustCenterError(
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
        pre_sequence = 0
        if self.watchdog:
            pre_sequence = max((event.sequence for event in self.watchdog.events), default=0)

        t0 = time.time()
        try:
            self.excel.Run(macro_name)
            elapsed = time.time() - t0

            # Check if watchdog caught anything during the run
            new_events = []
            if self.watchdog:
                all_events = self.watchdog.events
                new_events = [event for event in all_events if event.sequence > pre_sequence]

            result = {
                "success": not any(
                    e.dialog_type in ("compile_error", "runtime_error", "vb_error")
                    for e in new_events
                ),
                "elapsed_seconds": elapsed,
                "run_id": str(uuid.uuid4()),
                "macro": macro_name,
                "phase": "macro_execution",
                "dialog_events": [e.to_dict() for e in new_events],
                "cleanup": self.cleanup_result,
            }
            if not result["success"]:
                result["primary_error"] = self._primary_error(new_events, None)

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
                new_events = [event for event in all_events if event.sequence > pre_sequence]

            if new_events:
                dialog_text = "; ".join(
                    f"[{ev.dialog_type}] {ev.text}" for ev in new_events
                )
                error_msg = f"{error_msg} | Dialogs captured: {dialog_text}"

            logger.error(f"{macro_name} failed after {elapsed:.2f}s: {error_msg}")
            return {
                "success": False,
                "elapsed_seconds": elapsed,
                "run_id": str(uuid.uuid4()),
                "macro": macro_name,
                "phase": "macro_execution",
                "error": error_msg,
                "com_error": self._com_error(e),
                "primary_error": self._primary_error(new_events, e),
                "dialog_events": [e.to_dict() for e in new_events],
                "cleanup": self.cleanup_result,
            }

    @staticmethod
    def _com_error(error: Exception) -> dict:
        hresult = getattr(error, "hresult", None)
        message = getattr(error, "strerror", None) or str(error)
        return {"hresult": hresult, "message": message}

    @staticmethod
    def _primary_error(events: list, error: Optional[Exception]) -> str:
        for event in events:
            if event.dialog_type in ("runtime_error", "compile_error", "vb_error") and event.text:
                return event.text
        if error is not None:
            excepinfo = getattr(error, "excepinfo", None)
            if excepinfo and len(excepinfo) > 2 and excepinfo[2]:
                return str(excepinfo[2])
            return str(error)
        return ""

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

    def set_named_range(self, name: str, value, strict: bool = False) -> bool:
        """Set a named range value. Returns True on success."""
        try:
            rng = self.wb.Names(name).RefersToRange
            rng.Value = value
            logger.info(f"Set {name} = {value}")
            return True
        except Exception as e:
            if strict:
                raise KeyError(f"Could not set named range '{name}': {e}") from e
            logger.warning(f"Could not set named range '{name}': {e}")
            return False

    def get_named_range(self, name: str, default=None):
        """Get a named range value. Returns default if not found."""
        try:
            return self.wb.Names(name).RefersToRange.Value
        except Exception:
            return default
