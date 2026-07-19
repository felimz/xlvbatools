"""
Workbook Modifier
==================
Programmatically modifies cell values, formulas, and named ranges.

Usage:
    from xlvbatools import Project

    result = Project.open("book.xlsm").modify(cell="A1", value=42)

    modify_cell("workbook.xlsm", sheet="Sheet1", cell="A1", value=42)
    modify_cell("workbook.xlsm", sheet="Sheet1", cell="B2", formula="=A1*2")
"""

import logging
import os
from contextlib import nullcontext
from typing import Any, Mapping

from xlvbatools.core.session import ExcelSession

logger = logging.getLogger(__name__)


def _write_ranges_in_session(
    session: ExcelSession,
    *,
    sheet: str,
    values: Mapping[str, Any],
    calculate: bool = False,
) -> dict[str, Any]:
    """Apply validated range values through an already-owned Excel session."""
    worksheet = None
    target = None
    writes: list[dict[str, Any]] = []
    try:
        worksheet = session.wb.Worksheets(sheet)
        for address, raw_value in values.items():
            target = worksheet.Range(address)
            rows = int(target.Rows.Count)
            columns = int(target.Columns.Count)
            if isinstance(raw_value, (list, tuple)):
                matrix = tuple(tuple(row) for row in raw_value)
                value_rows = len(matrix)
                value_columns = len(matrix[0]) if matrix else 0
                if value_rows != rows or value_columns != columns:
                    raise ValueError(
                        f"Values for {sheet}!{address} have shape "
                        f"{value_rows}x{value_columns}; range is {rows}x{columns}"
                    )
                target.Value = matrix
            else:
                if rows != 1 or columns != 1:
                    raise ValueError(
                        f"Scalar value requires a single cell, not "
                        f"{sheet}!{address} ({rows}x{columns})"
                    )
                target.Value = raw_value
            writes.append({
                "sheet": sheet,
                "range": address,
                "rows": rows,
                "columns": columns,
            })
            target = None
        if calculate:
            session.excel.Calculate()
        return {
            "applied": True,
            "writes": writes,
            "calculated": calculate,
        }
    finally:
        target = None
        worksheet = None


def modify_cell(
    workbook_path: str,
    sheet: str = "Sheet1",
    cell: str | None = None,
    value=None,
    formula: str | None = None,
    name: str | None = None,
    refers_to: str | None = None,
    delete_name: bool = False,
    *,
    _session=None,
    _raise_on_error: bool = False,
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
        message = f"Workbook not found: {wb_path}"
        logger.error(message)
        if _raise_on_error:
            raise FileNotFoundError(message)
        return False

    success = True
    errors = []
    session_context = (
        nullcontext(_session) if _session is not None
        else ExcelSession(wb_path, visible=False, save_on_exit=True)
    )
    with session_context as session:
        excel, wb = session.excel, session.wb

        try:
            # Named range deletion
            if name and delete_name:
                try:
                    wb.Names(name).Delete()
                    logger.info(f"Deleted named range '{name}'")
                except Exception as e:
                    logger.error(f"Failed to delete named range '{name}': {e}")
                    errors.append(f"Failed to delete named range '{name}': {e}")
                    success = False

            # Named range creation
            elif name and refers_to:
                try:
                    wb.Names.Add(Name=name, RefersTo=refers_to)
                    logger.info(f"Created named range '{name}' -> '{refers_to}'")
                except Exception as e:
                    logger.error(f"Failed to create named range '{name}': {e}")
                    errors.append(f"Failed to create named range '{name}': {e}")
                    success = False

            # Cell modification
            target_cell = None
            if cell:
                try:
                    ws = wb.Worksheets(sheet)
                    target_cell = ws.Range(cell)
                except Exception as e:
                    logger.error(f"Failed to resolve '{cell}' on '{sheet}': {e}")
                    errors.append(f"Failed to resolve '{cell}' on '{sheet}': {e}")
                    success = False
            elif name and not refers_to and not delete_name:
                try:
                    target_cell = wb.Names(name).RefersToRange
                except Exception as e:
                    logger.error(f"Failed to resolve named range '{name}': {e}")
                    errors.append(f"Failed to resolve named range '{name}': {e}")
                    success = False

            if target_cell is not None:
                if formula is not None:
                    target_cell.Formula = formula
                    logger.info(f"Set formula on '{target_cell.Address}': {formula}")
                if value is not None:
                    target_cell.Value = value
                    logger.info(f"Set value on '{target_cell.Address}': {value}")

            if success:
                excel.Calculate()

        except Exception as e:
            logger.error(f"Excel COM error: {e}")
            errors.append(f"Excel COM error: {e}")
            success = False

    if not success and _raise_on_error:
        raise RuntimeError("; ".join(errors) or "Workbook modification failed")
    return success
