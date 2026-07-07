"""
Workbook Modifier
==================
Programmatically modifies cell values, formulas, and named ranges.

Usage:
    from xlvbatools.workbook import modify_cell

    modify_cell("workbook.xlsm", sheet="Sheet1", cell="A1", value=42)
    modify_cell("workbook.xlsm", sheet="Sheet1", cell="B2", formula="=A1*2")
"""

import logging
import os

from xlvbatools.core.session import ExcelSession

logger = logging.getLogger(__name__)


def modify_cell(
    workbook_path: str,
    sheet: str = "Sheet1",
    cell: str | None = None,
    value=None,
    formula: str | None = None,
    name: str | None = None,
    refers_to: str | None = None,
    delete_name: bool = False,
) -> bool:
    """
    Modify a cell value/formula or manage a named range.

    Parameters
    ----------
    workbook_path : str
        Path to the .xlsm workbook.
    sheet : str
        Worksheet name for cell modifications.
    cell : str, optional
        Cell coordinate (e.g. "C30").
    value : any, optional
        Value to write.
    formula : str, optional
        Formula to write (e.g. "=A1+B1").
    name : str, optional
        Named range to create, delete, or target.
    refers_to : str, optional
        Reference for creating a named range.
    delete_name : bool
        If True, deletes the named range specified by `name`.

    Returns
    -------
    bool
        True if all modifications succeeded.
    """
    wb_path = os.path.abspath(workbook_path)
    if not os.path.exists(wb_path):
        logger.error(f"Workbook not found: {wb_path}")
        return False

    success = True
    with ExcelSession(wb_path, visible=False, save_on_exit=True) as session:
        excel, wb = session.excel, session.wb

        try:
            # Named range deletion
            if name and delete_name:
                try:
                    wb.Names(name).Delete()
                    logger.info(f"Deleted named range '{name}'")
                except Exception as e:
                    logger.error(f"Failed to delete named range '{name}': {e}")
                    success = False

            # Named range creation
            elif name and refers_to:
                try:
                    wb.Names.Add(Name=name, RefersTo=refers_to)
                    logger.info(f"Created named range '{name}' -> '{refers_to}'")
                except Exception as e:
                    logger.error(f"Failed to create named range '{name}': {e}")
                    success = False

            # Cell modification
            target_cell = None
            if cell:
                try:
                    ws = wb.Worksheets(sheet)
                    target_cell = ws.Range(cell)
                except Exception as e:
                    logger.error(f"Failed to resolve '{cell}' on '{sheet}': {e}")
                    success = False
            elif name and not refers_to and not delete_name:
                try:
                    target_cell = wb.Names(name).RefersToRange
                except Exception as e:
                    logger.error(f"Failed to resolve named range '{name}': {e}")
                    success = False

            if target_cell is not None:
                if formula is not None:
                    target_cell.Formula = formula
                    logger.info(f"Set formula on '{target_cell.Address}': {formula}")
                if value is not None:
                    target_cell.Value = value
                    logger.info(f"Set value on '{target_cell.Address}': {value}")

            excel.Calculate()

        except Exception as e:
            logger.error(f"Excel COM error: {e}")
            success = False

    return success
