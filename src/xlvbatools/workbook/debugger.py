"""
Workbook Debugger
==================
Opens Excel and VBE visibly for interactive debugging. Disables the
dialog watchdog so the user can see native VBA dialogs.

Usage:
    from xlvbatools.workbook import launch_debug_session

    launch_debug_session("workbook.xlsm")
"""

import logging

from xlvbatools.core.session import ExcelSession

logger = logging.getLogger(__name__)


def launch_debug_session(
    workbook_path: str,
    open_vbe: bool = True,
):
    """
    Open Excel and VBE visibly for interactive debugging.

    This function blocks until the user presses Enter in the console.
    The watchdog is disabled so native VBA dialogs appear normally.

    Parameters
    ----------
    workbook_path : str
        Path to the .xlsm workbook.
    open_vbe : bool
        Whether to automatically open the VBE window (default True).
    """
    print("Opening Excel visibly via ExcelSession...")
    with ExcelSession(
        workbook_path,
        visible=True,
        enable_watchdog=False,
        save_on_exit=True,
    ) as session:
        excel = session.excel
        excel.DisplayAlerts = True

        if open_vbe:
            print("Activating VBE window...")
            try:
                excel.VBE.MainWindow.Visible = True
                print("VBE active. Set breakpoints, add watches, or step through code.")
            except Exception:
                print("\n[NOTE] Could not open VBE automatically.")
                print("  Check: Excel Options -> Trust Center -> Trust Center Settings")
                print("         -> Macro Settings -> 'Trust access to the VBA project object model'")
                print("\n  You can still open VBE manually with Alt+F11.\n")

        print("Excel is running. Press Enter to save and close...")
        input()

    print("Excel session closed.")
