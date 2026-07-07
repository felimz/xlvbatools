"""
Workbook Dumper
================
Exports worksheet screenshots with cell/column headers and dumps
worksheet contents (values, formatted texts, formulas) to JSON and Markdown.

Usage:
    from xlvbatools.workbook import dump_sheet_data, export_screenshots

    export_screenshots("workbook.xlsm", ["Sheet1"], "screenshots/")
    dump_sheet_data("workbook.xlsm", ["Sheet1"], output_json="dump.json")
"""

import json
import logging
import os
import time

from xlvbatools.core.session import ExcelSession

logger = logging.getLogger(__name__)

# Excel cell error value mapping
EXCEL_ERROR_MAP = {
    -2146826281: "#DIV/0!",
    -2146826273: "#VALUE!",
    -2146826265: "#REF!",
    -2146826259: "#NAME?",
    -2146826252: "#NUM!",
    -2146826246: "#N/A",
    -2146826288: "#NULL!",
}


def get_column_letter(col_idx: int) -> str:
    """Convert a 1-based column index to Excel column letter(s)."""
    result = ""
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(65 + remainder) + result
    return result


def dump_named_ranges(wb) -> dict:
    """Extract all named ranges from a workbook COM object."""
    named_ranges = {}
    try:
        for name in wb.Names:
            name_str = name.Name
            try:
                refers_to = name.RefersTo
                val = None
                try:
                    rng = name.RefersToRange
                    if rng is not None:
                        r_count = rng.Rows.Count
                        c_count = rng.Columns.Count
                        if r_count > 1 or c_count > 1:
                            raw_val = rng.Value
                            val = [list(row) if isinstance(row, tuple) else [row]
                                   for row in raw_val] if isinstance(raw_val, tuple) else raw_val
                        else:
                            val = rng.Value
                except Exception:
                    try:
                        val = wb.Application.Evaluate(refers_to)
                    except Exception:
                        val = None

                error_type = None
                if isinstance(val, int) and val in EXCEL_ERROR_MAP:
                    error_type = EXCEL_ERROR_MAP[val]
                    val = error_type

                named_ranges[name_str] = {
                    "refers_to": refers_to,
                    "value": val,
                    "error": error_type,
                }
            except Exception as e:
                named_ranges[name_str] = {"refers_to": None, "value": None, "error": str(e)}
    except Exception as e:
        logger.warning(f"Failed to query Named Ranges: {e}")
    return named_ranges


def dump_sheet_shapes(sheet) -> list:
    """Extract interactive shapes (buttons, form controls) from a worksheet."""
    shapes_info = []
    try:
        for shape in sheet.Shapes:
            on_action, text, linked_cell = "", "", ""
            try:
                on_action = shape.OnAction
            except Exception:
                pass
            try:
                text = shape.TextFrame.Characters.Text
            except Exception:
                pass
            try:
                lc = shape.ControlFormat.LinkedCell
                linked_cell = str(lc) if not callable(lc) else lc()
            except Exception:
                pass

            if on_action or text or linked_cell:
                shapes_info.append({
                    "name": shape.Name,
                    "type": shape.Type,
                    "text": text.strip() if text else "",
                    "on_action": on_action,
                    "linked_cell": linked_cell,
                })
    except Exception as e:
        logger.warning(f"Failed to extract shapes: {e}")
    return shapes_info


def export_screenshots(
    workbook_path: str,
    sheets: list[str],
    output_dir: str,
    custom_range: str | None = None,
) -> dict[str, str]:
    """
    Export worksheet screenshots with column/row headers.

    Returns dict mapping sheet name -> output file path (or error message).
    """
    os.makedirs(output_dir, exist_ok=True)
    results = {}

    with ExcelSession(workbook_path, visible=True, save_on_exit=False) as session:
        excel, wb = session.excel, session.wb

        for sheet_name in sheets:
            try:
                try:
                    sheet = wb.Worksheets(sheet_name)
                except Exception:
                    results[sheet_name] = "Not found"
                    continue

                ur = sheet.Range(custom_range) if custom_range else sheet.UsedRange
                r_count, c_count = ur.Rows.Count, ur.Columns.Count

                if r_count == 0 or c_count == 0:
                    results[sheet_name] = "Empty"
                    continue

                # Copy to temp sheet, freeze values, export as image
                sheet.Copy()
                wb_temp = excel.ActiveWorkbook
                ws_temp = wb_temp.ActiveSheet

                try:
                    ws_temp.UsedRange.Copy()
                    ws_temp.UsedRange.PasteSpecial(Paste=-4163)  # xlPasteValues
                    excel.CutCopyMode = False
                except Exception:
                    pass

                export_range = ws_temp.UsedRange
                temp_img = os.path.join(output_dir, f"{sheet_name}_raw.png")
                export_range.CopyPicture(Appearance=1, Format=2)  # xlScreen, xlBitmap

                # Use a temp chart to export
                chart = ws_temp.ChartObjects.Add(0, 0, export_range.Width, export_range.Height)
                chart.Chart.Paste()
                chart.Chart.Export(temp_img, "PNG")
                chart.Delete()

                wb_temp.Close(False)

                # Add headers with PIL if available
                output_path = os.path.join(output_dir, f"{sheet_name}.png")
                try:
                    from PIL import Image
                    img = Image.open(temp_img)
                    img.save(output_path)
                    os.remove(temp_img)
                except ImportError:
                    os.rename(temp_img, output_path)

                results[sheet_name] = output_path
                logger.info(f"Screenshot: {sheet_name} -> {output_path}")

            except Exception as e:
                results[sheet_name] = f"Error: {e}"
                logger.error(f"Screenshot failed for {sheet_name}: {e}")

    return results


def dump_sheet_data(
    workbook_path: str,
    sheets: list[str],
    output_json: str | None = None,
    output_md: str | None = None,
    custom_range: str | None = None,
    dump_names: bool = True,
    max_md_rows: int = 500,
) -> dict:
    """
    Dump worksheet contents to JSON and/or Markdown.

    Returns the complete dump data dict.
    """
    wb_path = os.path.abspath(workbook_path)
    dump_data = {"workbook": os.path.basename(wb_path), "sheets": {}}

    with ExcelSession(wb_path, visible=False, save_on_exit=False) as session:
        excel, wb = session.excel, session.wb

        for sheet_name in sheets:
            try:
                sheet = wb.Worksheets(sheet_name)
            except Exception:
                dump_data["sheets"][sheet_name] = {"error": "Not found"}
                continue

            ur = sheet.Range(custom_range) if custom_range else sheet.UsedRange
            r_start, c_start = ur.Row, ur.Column
            r_count, c_count = ur.Rows.Count, ur.Columns.Count

            # Read values and formulas
            values = _read_range_data(ur, r_count, c_count)
            formulas = _read_range_formulas(ur, r_count, c_count)

            sheet_data = {
                "used_range": ur.Address,
                "rows": r_count,
                "cols": c_count,
                "values": values,
                "formulas": formulas,
                "shapes": dump_sheet_shapes(sheet),
            }
            dump_data["sheets"][sheet_name] = sheet_data

        if dump_names:
            dump_data["named_ranges"] = dump_named_ranges(wb)

    # Write outputs
    if output_json:
        os.makedirs(os.path.dirname(os.path.abspath(output_json)), exist_ok=True)
        with open(output_json, "w", encoding="utf-8") as f:
            json.dump(dump_data, f, indent=2, default=str)
        logger.info(f"JSON dump: {output_json}")

    if output_md:
        os.makedirs(os.path.dirname(os.path.abspath(output_md)), exist_ok=True)
        md_content = _render_markdown(dump_data, max_md_rows)
        with open(output_md, "w", encoding="utf-8") as f:
            f.write(md_content)
        logger.info(f"Markdown dump: {output_md}")

    return dump_data


def _read_range_data(ur, r_count: int, c_count: int) -> list:
    """Read values from a range into a 2D list."""
    if r_count == 1 and c_count == 1:
        val = ur.Value
        if isinstance(val, int) and val in EXCEL_ERROR_MAP:
            val = EXCEL_ERROR_MAP[val]
        return [[val]]
    raw = ur.Value
    if isinstance(raw, tuple):
        result = []
        for row in raw:
            row_data = []
            for cell in (row if isinstance(row, tuple) else (row,)):
                if isinstance(cell, int) and cell in EXCEL_ERROR_MAP:
                    cell = EXCEL_ERROR_MAP[cell]
                row_data.append(cell)
            result.append(row_data)
        return result
    return [[raw]]


def _read_range_formulas(ur, r_count: int, c_count: int) -> list:
    """Read formulas from a range into a 2D list."""
    try:
        if r_count == 1 and c_count == 1:
            return [[ur.Formula]]
        raw = ur.Formula
        if isinstance(raw, tuple):
            return [list(row) if isinstance(row, tuple) else [row] for row in raw]
        return [[raw]]
    except Exception:
        return []


def _render_markdown(data: dict, max_rows: int) -> str:
    """Render dump data as a Markdown document."""
    lines = [f"# Workbook Dump: {data['workbook']}\n"]

    for sheet_name, sheet_data in data.get("sheets", {}).items():
        lines.append(f"\n## {sheet_name}\n")

        if "error" in sheet_data:
            lines.append(f"**Error**: {sheet_data['error']}\n")
            continue

        lines.append(f"- Range: `{sheet_data.get('used_range', '?')}`")
        lines.append(f"- Rows: {sheet_data.get('rows', 0)}, Cols: {sheet_data.get('cols', 0)}\n")

        values = sheet_data.get("values", [])
        if values and len(values) <= max_rows:
            lines.append("### Values\n")
            lines.append(_table_to_md(values))

    nr = data.get("named_ranges", {})
    if nr:
        lines.append("\n## Named Ranges\n")
        lines.append("| Name | Value | RefersTo |")
        lines.append("|:---|:---|:---|")
        for name, info in sorted(nr.items()):
            val = str(info.get("value", ""))[:50]
            ref = str(info.get("refers_to", ""))
            lines.append(f"| {name} | {val} | {ref} |")

    return "\n".join(lines)


def _table_to_md(data: list) -> str:
    """Convert a 2D list to a Markdown table."""
    if not data:
        return ""
    cols = max(len(row) for row in data)
    header = "| " + " | ".join(get_column_letter(c + 1) for c in range(cols)) + " |"
    sep = "|" + "|".join(":---" for _ in range(cols)) + "|"
    rows = []
    for row in data:
        cells = [str(row[c] if c < len(row) and row[c] is not None else "")[:40]
                 for c in range(cols)]
        rows.append("| " + " | ".join(cells) + " |")
    return "\n".join([header, sep] + rows)
