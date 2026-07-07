"""
Macro Runner
==============
Executes VBA macros with full dialog protection through ExcelSession.

Usage:
    from xlvbatools.macro import run_macro

    result = run_macro("workbook.xlsm", "MyMacro")
    result = run_macro("workbook.xlsm", "MyMacro", named_ranges={"FilePath": "C:/data.txt"})
"""

import logging
import os
import time

from xlvbatools.core.session import ExcelSession

logger = logging.getLogger(__name__)


def run_macro(
    workbook_path: str,
    macro_name: str,
    named_ranges: dict | None = None,
    timeout: float = 120.0,
    visible: bool = False,
    save_on_exit: bool = True,
) -> dict:
    """
    Run a VBA macro with full dialog protection.

    Parameters
    ----------
    workbook_path : str
        Path to the .xlsm workbook.
    macro_name : str
        Name of the macro to execute (e.g. "MyMacro" or "Sheet1.MyMacro").
    named_ranges : dict, optional
        Named ranges to set before running (name -> value pairs).
    timeout : float
        Maximum seconds to wait for the macro (default 120).
    visible : bool
        Whether Excel should be visible during execution.
    save_on_exit : bool
        Whether to save the workbook after execution.

    Returns
    -------
    dict
        Result with keys: success, elapsed_seconds, error (if any), dialog_events.
    """
    wb_path = os.path.abspath(workbook_path)

    with ExcelSession(wb_path, visible=visible, save_on_exit=save_on_exit) as session:
        # Set any named ranges before running
        if named_ranges:
            for name, value in named_ranges.items():
                if not session.set_named_range(name, value):
                    logger.warning(f"Could not set named range '{name}' = {value}")

        # Run the macro
        result = session.run_macro(macro_name, timeout=timeout)

        return result
