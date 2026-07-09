"""
Excel Process Management
=========================
Utilities for managing Excel.exe processes: kill, detect, gracefully close.

These are standalone functions with no state -- safe to call from anywhere.
"""

import os
import subprocess
import time
import logging
from typing import Optional

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


def close_excel_gracefully(target_path: Optional[str] = None) -> bool:
    """
    Attempt to gracefully save and close open workbooks and quit
    running Excel instances via COM ROT (Running Object Table).

    If target_path is specified, only that specific workbook is closed.
    If the parent Excel instance has no other open workbooks after closing
    the target, that instance is quit. Unrelated workbooks are left untouched.

    Returns True if successful.
    """
    if not IS_WINDOWS:
        return True
    import win32com.client

    if target_path:
        target_path = os.path.abspath(target_path).lower()

    logger.info("Attempting graceful closure of existing Excel instances...")
    try:
        quit_hwnds = set()
        keep_hwnds = set()
        for _attempt in range(20):  # Safety limit to prevent infinite loop
            try:
                excel = win32com.client.GetActiveObject("Excel.Application")
            except Exception:
                # No active Excel instances registered in the ROT
                break

            try:
                hwnd = excel.Hwnd
            except Exception:
                hwnd = None

            if hwnd is not None:
                if hwnd in quit_hwnds:
                    # We already instructed this instance to quit. Wait for it to shut down.
                    time.sleep(0.2)
                    continue
                if hwnd in keep_hwnds:
                    # We already processed this instance and decided to keep it running.
                    # Since GetActiveObject will continue returning the same active object, we can stop.
                    break

            wbs = excel.Workbooks
            target_found = False
            unrelated_open = False

            for i in range(wbs.Count, 0, -1):
                try:
                    wb = wbs.Item(i)
                    wb_path = os.path.abspath(wb.FullName).lower()
                    if target_path and wb_path == target_path:
                        logger.info(f"Saving and closing target workbook: {wb.Name}")
                        wb.Close(SaveChanges=True)
                        target_found = True
                    elif target_path:
                        unrelated_open = True
                    else:
                        logger.info(f"Saving and closing workbook: {wb.Name}")
                        wb.Close(SaveChanges=True)
                except Exception as e:
                    logger.warning(f"Failed to close workbook: {e}")
                    return False

            if not target_path or (target_found and not unrelated_open):
                try:
                    excel.Quit()
                    logger.info("Quit Excel instance gracefully.")
                    if hwnd is not None:
                        quit_hwnds.add(hwnd)
                except Exception as e:
                    logger.warning(f"Failed to quit Excel: {e}")
                    return False
            else:
                if target_found and unrelated_open:
                    logger.info("Closed target workbook gracefully. Leaving Excel instance running for unrelated workbooks.")
                if hwnd is not None:
                    keep_hwnds.add(hwnd)

            time.sleep(0.3)
        return True
    except Exception as e:
        logger.warning(f"Error during graceful Excel closure: {e}")
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
