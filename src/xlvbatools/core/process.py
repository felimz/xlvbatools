"""
Excel Process Management
=========================
Utilities for managing Excel.exe processes: kill, detect, gracefully close.

These are standalone functions with no state -- safe to call from anywhere.
"""

import subprocess
import time
import logging

from xlvbatools._compat import IS_WINDOWS

logger = logging.getLogger(__name__)


def kill_excel(timeout: float = 2.0) -> bool:
    """
    Force-kill all running Excel processes.

    Parameters
    ----------
    timeout : float
        Seconds to wait after killing for process cleanup.

    Returns
    -------
    bool
        True if any processes were killed.
    """
    if not IS_WINDOWS:
        logger.debug("kill_excel: skipped (not Windows)")
        return False
    result = subprocess.run(
        ["taskkill", "/f", "/im", "EXCEL.EXE"],
        capture_output=True, text=True
    )
    killed = result.returncode == 0
    if killed:
        logger.info("Killed stale Excel processes")
        time.sleep(timeout)
    return killed


def is_excel_running() -> bool:
    """Check if any Excel process is currently running."""
    if not IS_WINDOWS:
        return False
    try:
        output = subprocess.check_output(
            ["tasklist", "/fi", "IMAGENAME eq EXCEL.EXE"],
            text=True
        )
        return "EXCEL.EXE" in output
    except Exception:
        return False


def elegant_close_excel() -> bool:
    """
    Attempt to elegantly save and close all open workbooks and quit
    running Excel instances via COM ROT (Running Object Table).

    Returns True if all running instances were successfully closed.
    """
    if not IS_WINDOWS:
        return True
    import win32com.client

    logger.info("Attempting elegant closure of existing Excel instances...")
    try:
        for _attempt in range(20):  # Safety limit to prevent infinite loop
            try:
                excel = win32com.client.GetActiveObject("Excel.Application")
            except Exception:
                # No active Excel instances registered in the ROT
                break

            logger.info("Found running Excel instance in ROT. Saving and closing...")
            wbs = excel.Workbooks
            for i in range(wbs.Count, 0, -1):
                try:
                    wb = wbs.Item(i)
                    logger.info(f"Saving and closing workbook: {wb.Name}")
                    wb.Close(SaveChanges=True)
                except Exception as e:
                    logger.warning(f"Failed to close workbook: {e}")
                    return False

            try:
                excel.Quit()
                logger.info("Quit Excel instance elegantly.")
            except Exception as e:
                logger.warning(f"Failed to quit Excel: {e}")
                return False

            time.sleep(0.5)
        return True
    except Exception as e:
        logger.warning(f"Error during elegant Excel closure: {e}")
        return False


def is_process_running(pid: int) -> bool:
    """Check if a process with the given PID is running."""
    if not IS_WINDOWS:
        return False
    try:
        output = subprocess.check_output(
            ["tasklist", "/fi", f"PID eq {pid}"],
            text=True
        )
        return str(pid) in output
    except Exception:
        return False


def kill_process_by_pid(pid: int):
    """Forcefully kill a process by PID."""
    if not IS_WINDOWS:
        return
    subprocess.run(
        ["taskkill", "/f", "/pid", str(pid)],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL
    )
