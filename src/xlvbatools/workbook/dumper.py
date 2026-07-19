"""
Workbook Dumper
================
Exports worksheet screenshots with cell/column headers and dumps
worksheet contents (values, formatted texts, formulas) to JSON and Markdown.

Features:
- Cell-level data model with address, value, text, formula, and error info
- Formatted text (.Text property) alongside raw values
- Error detection with per-sheet error summaries
- Named range extraction with scope (workbook vs. sheet-level)
- Interactive shapes/controls inventory
- Rich Markdown report with row/column headers, formula annotations
- Screenshot export with Pillow-composited headers and gridlines

Usage:
    from xlvbatools import Project

    result = Project.open("book.xlsm").inspect(["Sheet1"])

    export_screenshots("workbook.xlsm", ["Sheet1"], "screenshots/")
    dump_sheet_data("workbook.xlsm", ["Sheet1"], output_json="dump.json")
"""

import json
import gc
import logging
import os
import time
from contextlib import nullcontext
from typing import Any

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


def sanitize_markdown_cell(text) -> str:
    """Sanitize text for insertion into a markdown table cell."""
    if text is None:
        return ""
    text_str = str(text)
    # Escape vertical bars
    text_str = text_str.replace("|", "\\|")
    # Replace newlines with <br> to keep table rows on a single line
    text_str = text_str.replace("\r\n", "<br>").replace("\n", "<br>").replace("\r", "<br>")
    return text_str


def normalize_2d(raw_data, r_count: int, c_count: int) -> list:
    """Normalize raw data from Excel Value/Formula properties into a list of lists."""
    if r_count == 1 and c_count == 1:
        return [[raw_data]]
    if isinstance(raw_data, tuple):
        return [list(row) if isinstance(row, tuple) else [row] for row in raw_data]
    return [[raw_data]]


def dump_named_ranges(wb) -> dict:
    """
    Extract all named ranges from a workbook COM object.

    Returns a dict keyed by name, with each entry containing:
    refers_to, scope, value, and error (if any).
    """
    named_ranges = {}
    names = None
    name = None
    rng = None
    try:
        names = wb.Names
        for index in range(1, int(names.Count) + 1):
            name = names.Item(index)
            name_str = name.Name
            try:
                refers_to = name.RefersTo

                # Determine scope (Workbook vs. specific sheet)
                scope = "Workbook"
                if refers_to and "!" in refers_to:
                    parts = refers_to.split("!")
                    if len(parts) > 0:
                        scope = parts[0].lstrip("=")

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
                    # Fallback to Evaluate if it refers to a formula/constant
                    try:
                        val = wb.Application.Evaluate(refers_to)
                    except Exception:
                        val = None
                finally:
                    rng = None

                # Format error values
                error_type = None
                if isinstance(val, int) and val in EXCEL_ERROR_MAP:
                    error_type = EXCEL_ERROR_MAP[val]
                    val = error_type

                named_ranges[name_str] = {
                    "refers_to": refers_to,
                    "scope": scope,
                    "value": val,
                    "error": error_type,
                }
            except Exception as e:
                named_ranges[name_str] = {
                    "refers_to": None,
                    "scope": "Unknown",
                    "value": None,
                    "error": str(e),
                }
            finally:
                name = None
    except Exception as e:
        logger.warning(f"Failed to query Named Ranges: {e}")
    finally:
        rng = None
        name = None
        names = None
    return named_ranges


def _collection_item(collection_factory, name):
    """Resolve a named item from Excel's callable or Item-based collections."""
    collection = collection_factory()
    try:
        return collection.Item(name)
    except Exception:
        return collection_factory(name)


def _shape_display_text(sheet, shape, shape_type: int) -> tuple[str, str]:
    """Return visible shape/control text and the COM accessor that supplied it."""
    candidates = [
        ("TextFrame.Characters.Text", lambda: shape.TextFrame.Characters.Text),
        ("TextFrame2.TextRange.Text", lambda: shape.TextFrame2.TextRange.Text),
    ]
    if shape_type == 8:  # msoFormControl
        candidates.append((
            "Buttons.Caption",
            lambda: _collection_item(sheet.Buttons, shape.Name).Caption,
        ))
    elif shape_type == 12:  # msoOLEControlObject (ActiveX)
        candidates.append((
            "OLEObjects.Object.Caption",
            lambda: _collection_item(sheet.OLEObjects, shape.Name).Object.Caption,
        ))

    for accessor, getter in candidates:
        try:
            value = getter()
            if value is not None and str(value).strip():
                return str(value).strip(), accessor
        except Exception:
            continue
    return "", ""


def dump_sheet_shapes(sheet) -> list:
    """Extract shapes and controls, including Forms and ActiveX captions."""
    shapes_info = []
    shapes = None
    shape = None
    shape_iterator = None
    try:
        shapes = sheet.Shapes
        shape_iterator = (
            (shapes.Item(index) for index in range(1, int(shapes.Count) + 1))
            if hasattr(shapes, "Count") and hasattr(shapes, "Item")
            else iter(shapes)
        )
        for shape in shape_iterator:
            on_action, linked_cell = "", ""
            try:
                shape_type = int(shape.Type)
            except Exception:
                shape_type = -1
            try:
                on_action = shape.OnAction
            except Exception:
                pass
            text, text_accessor = _shape_display_text(sheet, shape, shape_type)
            try:
                lc = shape.ControlFormat.LinkedCell
                linked_cell = str(lc) if not callable(lc) else lc()
            except Exception:
                pass

            if on_action or text or linked_cell:
                shapes_info.append({
                    "name": shape.Name,
                    "type": shape_type,
                    "control_type": {
                        8: "forms_control",
                        12: "activex_control",
                    }.get(shape_type, "shape"),
                    "text": text,
                    "text_accessor": text_accessor,
                    "on_action": on_action,
                    "linked_cell": linked_cell,
                })
            shape = None
    except Exception as e:
        logger.warning(f"Failed to extract shapes for sheet {sheet.Name}: {e}")
    finally:
        shape = None
        shape_iterator = None
        shapes = None
    return shapes_info


def _scaled_boundaries(sizes: list[float], pixel_extent: int) -> list[int]:
    """Map Excel point sizes to stable pixel boundaries without rounding drift."""
    if pixel_extent <= 0 or not sizes:
        raise ValueError("Overlay geometry requires a positive extent and sizes")
    normalized = []
    for size in sizes:
        try:
            normalized.append(max(0.0, float(size)))
        except (TypeError, ValueError):
            normalized.append(0.0)
    total = sum(normalized)
    if total <= 0:
        raise ValueError("Overlay geometry contains no visible rows or columns")
    boundaries = [0]
    cumulative = 0.0
    for size in normalized:
        cumulative += size
        boundaries.append(round(pixel_extent * cumulative / total))
    boundaries[-1] = pixel_extent
    return [max(0, min(pixel_extent, value)) for value in boundaries]


def _worksheet_is_renderable(sheet, include_hidden_sheets: bool) -> bool:
    """Return whether a worksheet may be screenshotted under the policy."""
    return include_hidden_sheets or int(sheet.Visible) == -1


def _apply_headers_and_grid_overlay(
    image_path: str,
    column_widths: list[float],
    row_heights: list[float],
    first_column: int,
    first_row: int,
    include_headers: bool = True,
    include_grid_overlay: bool = True,
) -> None:
    """Add geometry-accurate headers/grid around an Excel-rendered PNG."""
    if not include_headers and not include_grid_overlay:
        return
    from PIL import Image, ImageDraw, ImageFont

    with Image.open(image_path) as source:
        source.load()
        image = source.convert("RGBA")
    x_bounds = _scaled_boundaries(column_widths, image.width)
    y_bounds = _scaled_boundaries(row_heights, image.height)

    try:
        font = ImageFont.truetype("arial.ttf", 12)
    except OSError:
        font = ImageFont.load_default()
    measure = ImageDraw.Draw(Image.new("RGB", (1, 1)))

    def text_size(value: str) -> tuple[int, int]:
        box = measure.textbbox((0, 0), value, font=font)
        return box[2] - box[0], box[3] - box[1]

    max_row_label = str(first_row + len(row_heights) - 1)
    row_label_width, font_height = text_size(max_row_label)
    pad_left = max(32, row_label_width + 14) if include_headers else 0
    pad_top = max(20, font_height + 8) if include_headers else 0
    canvas = Image.new(
        "RGBA", (image.width + pad_left, image.height + pad_top),
        (255, 255, 255, 255),
    )
    canvas.paste(image, (pad_left, pad_top))
    draw = ImageDraw.Draw(canvas)
    header_fill = (240, 240, 240, 255)
    header_line = (180, 180, 180, 255)
    grid_line = (210, 210, 210, 255)

    if include_headers:
        draw.rectangle((0, 0, canvas.width - 1, pad_top - 1), fill=header_fill)
        draw.rectangle((0, 0, pad_left - 1, canvas.height - 1), fill=header_fill)
        draw.line((pad_left, 0, pad_left, canvas.height - 1), fill=header_line)
        draw.line((0, pad_top, canvas.width - 1, pad_top), fill=header_line)

    for index in range(len(column_widths)):
        left, right = x_bounds[index], x_bounds[index + 1]
        x = pad_left + right
        if include_grid_overlay and 0 < right < image.width:
            draw.line((x, pad_top, x, canvas.height - 1), fill=grid_line)
        if include_headers and right > left:
            label = get_column_letter(first_column + index)
            width, height = text_size(label)
            if width + 4 <= right - left:
                draw.text(
                    (pad_left + left + (right - left - width) // 2,
                     max(0, (pad_top - height) // 2)),
                    label, fill=(50, 50, 50, 255), font=font,
                )

    for index in range(len(row_heights)):
        top, bottom = y_bounds[index], y_bounds[index + 1]
        y = pad_top + bottom
        if include_grid_overlay and 0 < bottom < image.height:
            draw.line((pad_left, y, canvas.width - 1, y), fill=grid_line)
        if include_headers and bottom > top:
            label = str(first_row + index)
            width, height = text_size(label)
            if height + 2 <= bottom - top:
                draw.text(
                    (max(0, (pad_left - width) // 2),
                     pad_top + top + (bottom - top - height) // 2),
                    label, fill=(50, 50, 50, 255), font=font,
                )

    draw.rectangle(
        (0, 0, canvas.width - 1, canvas.height - 1),
        outline=header_line, width=1,
    )
    temporary_path = image_path + ".overlay.tmp"
    try:
        canvas.convert("RGB").save(temporary_path, "PNG")
        os.replace(temporary_path, image_path)
    finally:
        if os.path.exists(temporary_path):
            os.remove(temporary_path)
        image.close()
        canvas.close()


def _show_owned_excel_offscreen(excel) -> None:
    """Make only the owned Excel HWND renderable without activating onscreen."""
    import ctypes

    hwnd = int(excel.Hwnd)
    user32 = ctypes.windll.user32
    swp_no_zorder = 0x0004
    swp_noactivate = 0x0010
    # Position the hidden HWND first, because setting Application.Visible can
    # otherwise display it at its last normal desktop coordinates for a frame.
    user32.SetWindowPos(
        hwnd, 0, -30000, -30000, 1600, 1000,
        swp_no_zorder | swp_noactivate,
    )
    excel.Visible = True
    user32.SetWindowPos(
        hwnd, 0, -30000, -30000, 1600, 1000,
        swp_no_zorder | swp_noactivate,
    )


def export_screenshots(
    workbook_path: str,
    sheets: list[str],
    output_dir: str,
    custom_range: str | None = None,
    render_mode: str = "excel_native",
    include_headers: bool = True,
    include_grid_overlay: bool = True,
    include_hidden_sheets: bool = False,
    dpi: int = 144,
    _session: ExcelSession | None = None,
) -> dict[str, str]:
    """
    Export worksheet screenshots with column/row headers composited via Pillow.

    Includes a retry loop for CopyPicture (common COM flake), proper Pillow
    header composition with column letters and row numbers, gridline overlay,
    and graceful custom_range fallback.

    Returns dict mapping sheet name -> output file path (or error message).
    """
    os.makedirs(output_dir, exist_ok=True)
    if render_mode != "excel_native":
        raise NotImplementedError(
            f"Render mode '{render_mode}' is not implemented; use 'excel_native'"
        )
    if dpi <= 0:
        raise ValueError("dpi must be positive")
    results = {}

    session_context = nullcontext(_session) if _session is not None else ExcelSession(
        workbook_path, visible=False, save_on_exit=False, kill_on_enter=False,
        read_only=True, disable_macros=True,
    )
    with session_context as session:
        excel, wb = session.excel, session.wb

        for sheet_name in sheets:
            try:
                try:
                    sheet = wb.Worksheets(sheet_name)
                except Exception:
                    logger.warning(f"Worksheet '{sheet_name}' not found in workbook")
                    results[sheet_name] = "Not found"
                    continue
                if not _worksheet_is_renderable(sheet, include_hidden_sheets):
                    logger.info(f"Skipping hidden worksheet '{sheet_name}'")
                    results[sheet_name] = "Hidden (skipped)"
                    continue

                # Resolve range with fallback (D8)
                if custom_range:
                    try:
                        ur = sheet.Range(custom_range)
                    except Exception as e:
                        logger.warning(
                            f"Custom range '{custom_range}' is invalid on sheet "
                            f"'{sheet_name}'. Falling back to UsedRange. Error: {e}"
                        )
                        ur = sheet.UsedRange
                else:
                    ur = sheet.UsedRange

                r_start = ur.Row
                c_start = ur.Column
                r_count = ur.Rows.Count
                c_count = ur.Columns.Count

                if r_count == 0 or c_count == 0 or (
                    r_count == 1 and c_count == 1 and ur.Value is None
                ):
                    logger.warning(f"Worksheet '{sheet_name}' is empty")
                    results[sheet_name] = "Empty"
                    continue

                r_end = r_start + r_count - 1
                c_end = c_start + c_count - 1

                logger.info(f"Rendering original range '{sheet_name}'!{ur.Address}")

                range_width = float(ur.Width)
                range_height = float(ur.Height)
                # Check the aggregate COM properties before enumerating every
                # row/column; a 70k-row range must fail in O(1), not after 70k
                # cross-process Height calls.
                pixel_scale = dpi / 72.0
                estimated_width = round(range_width * pixel_scale)
                estimated_height = round(range_height * pixel_scale)
                estimated_pixels = estimated_width * estimated_height
                max_native_axis_pixels = 32767
                max_native_pixels = 80_000_000
                if (estimated_width > max_native_axis_pixels or
                        estimated_height > max_native_axis_pixels or
                        estimated_pixels > max_native_pixels):
                    raise ValueError(
                        f"Range {ur.Address} is too large for excel_native "
                        f"rendering (estimated {estimated_width} x "
                        f"{estimated_height} pixels at {dpi} DPI); "
                        "request a smaller range"
                    )

                # Get column widths and row heights in points (for Pillow scaling)
                widths = [sheet.Cells(r_start, c).Width for c in range(c_start, c_end + 1)]
                heights = [sheet.Cells(r, c_start).Height for r in range(r_start, r_end + 1)]

                # Enable gridlines
                try:
                    excel.ActiveWindow.DisplayGridlines = True
                except Exception:
                    pass

                export_range = sheet.Range(
                    sheet.Cells(r_start, c_start),
                    sheet.Cells(r_end, c_end),
                )

                # Build output filename (D7 — custom range in filename)
                if custom_range:
                    range_clean = custom_range.replace(":", "_").replace("$", "")
                    filename = f"{sheet_name.lower().replace(' ', '_')}_zoom_{range_clean}.png"
                else:
                    filename = f"{sheet_name.lower().replace(' ', '_')}_screenshot.png"
                out_path = os.path.abspath(os.path.join(output_dir, filename))

                max_retries = 5
                wb_temp = None
                chart_obj = None
                chart = None
                try:
                    for attempt in range(max_retries):
                        try:
                            sheet.Activate()
                            try:
                                export_range.Select()
                            except Exception:
                                # Large/non-contiguous used ranges can reject
                                # Select while still supporting CopyPicture.
                                pass
                            export_range.CopyPicture(Appearance=1, Format=2)
                            break
                        except Exception:
                            if attempt == max_retries - 1:
                                raise
                            if attempt == 0:
                                # Some Excel builds require a visible normal
                                # window for CopyPicture. Only the owned
                                # instance is exposed, far outside the desktop.
                                _show_owned_excel_offscreen(excel)
                            time.sleep(0.5)

                    # Never create the bitmap-only chart workbook while the
                    # owned application is visible, even off-screen.
                    excel.Visible = False
                    # The clean workbook receives only the clipboard bitmap;
                    # worksheet code and formulas never cross this boundary.
                    wb_temp = excel.Workbooks.Add()
                    ws_temp = wb_temp.Worksheets(1)
                    chart_obj = ws_temp.ChartObjects().Add(
                        0, 0, range_width, range_height
                    )
                    chart = chart_obj.Chart

                    try:
                        chart_obj.Select()
                    except Exception:
                        pass

                # Paste with retry loop to handle transient clipboard lock issues
                    for attempt in range(max_retries):
                        try:
                            chart.Paste()
                            break
                        except Exception:
                            if attempt == max_retries - 1:
                                raise
                            time.sleep(0.5)

                # Size the pasted shape to match the chart area
                    if chart.Shapes.Count > 0:
                        shape = chart.Shapes.Item(1)
                        shape.Left = 0
                        shape.Top = 0
                        shape.Width = range_width
                        shape.Height = range_height
                    time.sleep(1.0)
                    if not chart.Export(out_path, "PNG"):
                        raise RuntimeError("Excel chart export returned false")
                finally:
                    shape = None
                    chart = None
                    if chart_obj is not None:
                        try:
                            chart_obj.Delete()
                        except Exception:
                            pass
                    chart_obj = None
                    if wb_temp is not None:
                        try:
                            wb_temp.Close(False)
                        except Exception:
                            pass
                    wb_temp = None
                    excel.CutCopyMode = False

                try:
                    _apply_headers_and_grid_overlay(
                        out_path, widths, heights, c_start, r_start,
                        include_headers=include_headers,
                        include_grid_overlay=include_grid_overlay,
                    )
                    logger.info(f"Applied headers/grid overlay to screenshot: {out_path}")

                except ImportError:
                    logger.info("Pillow not available, screenshot saved without headers")
                except Exception as e:
                    logger.warning(f"Failed to apply Pillow headers: {e}")

                logger.info(f"Saved sheet '{sheet_name}' screenshot to {out_path}")
                results[sheet_name] = out_path

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
    _session: ExcelSession | None = None,
) -> dict:
    """
    Dump worksheet contents to JSON and/or Markdown.

    Uses a cell-level data model: each non-empty cell is stored as
    ``{address: {row, col, value, text, formula, is_error, error_type}}``.

    Returns the complete dump data dict.
    """
    wb_path = os.path.abspath(workbook_path)
    dump_data: dict[str, Any] = {
        "metadata": {
            "workbook": os.path.basename(wb_path),
            "timestamp": time.strftime("%Y-%m-%dT%H:%M:%S"),
            "sheets_processed": sheets,
        },
        "sheets": {},
    }

    session_context = nullcontext(_session) if _session is not None else ExcelSession(
        wb_path, visible=False, save_on_exit=False, kill_on_enter=False,
        read_only=True, disable_macros=True,
    )
    with session_context as session:
        wb = session.wb

        if dump_names:
            dump_data["named_ranges"] = dump_named_ranges(wb)

        for sheet_name in sheets:
            sheet = None
            ur = None
            cell = None
            try:
                try:
                    sheet = wb.Worksheets(sheet_name)
                except Exception:
                    logger.warning(f"Worksheet '{sheet_name}' not found for data dump")
                    dump_data["sheets"][sheet_name] = {"error": "Not found"}
                    continue

                # Resolve range with graceful fallback (D8)
                if custom_range:
                    try:
                        ur = sheet.Range(custom_range)
                    except Exception as e:
                        logger.warning(
                            f"Custom range '{custom_range}' is invalid on sheet "
                            f"'{sheet_name}'. Falling back to UsedRange. Error: {e}"
                        )
                        ur = sheet.UsedRange
                else:
                    ur = sheet.UsedRange

                r_start = ur.Row
                c_start = ur.Column
                r_count = ur.Rows.Count
                c_count = ur.Columns.Count

                if r_count == 0 or c_count == 0 or (
                    r_count == 1 and c_count == 1 and ur.Value is None
                ):
                    logger.info(f"Worksheet '{sheet_name}' is empty")
                    dump_data["sheets"][sheet_name] = {
                        "bounds": {
                            "r_start": 0, "c_start": 0,
                            "r_count": 0, "c_count": 0,
                        },
                        "cells": {},
                    }
                    continue

                logger.info(f"Extracting data from sheet '{sheet_name}' (UsedRange: {ur.Address})")

                # Bulk retrieve values and formulas
                raw_values = ur.Value
                raw_formulas = ur.Formula

                values = normalize_2d(raw_values, r_count, c_count)
                formulas = normalize_2d(raw_formulas, r_count, c_count)

                # Build cell-level data model (D1, D2, D3)
                sheet_cells = {}
                errors_found = []

                for r in range(r_count):
                    row_vals = values[r]
                    row_forms = formulas[r]

                    for c in range(c_count):
                        val = row_vals[c]
                        form = row_forms[c]

                        is_formula = isinstance(form, str) and form.startswith("=")

                        # Only extract cells that have content
                        if val is not None or is_formula:
                            row_num = r_start + r
                            col_num = c_start + c
                            addr = f"{get_column_letter(col_num)}{row_num}"

                            # Get formatted text (D2)
                            try:
                                cell = ur.Cells(r + 1, c + 1)
                                text = cell.Text
                            except Exception:
                                text = str(val) if val is not None else ""
                            finally:
                                cell = None

                            # Detect errors (D3)
                            is_error = False
                            error_type = None
                            if isinstance(val, int) and val in EXCEL_ERROR_MAP:
                                is_error = True
                                error_type = EXCEL_ERROR_MAP[val]
                                val = error_type
                                text = error_type
                                errors_found.append(addr)

                            sheet_cells[addr] = {
                                "row": row_num,
                                "col": col_num,
                                "value": val,
                                "text": text,
                                "formula": form if is_formula else None,
                                "is_error": is_error,
                                "error_type": error_type,
                            }

                sheet_shapes = dump_sheet_shapes(sheet)

                dump_data["sheets"][sheet_name] = {
                    "bounds": {
                        "r_start": r_start,
                        "c_start": c_start,
                        "r_count": r_count,
                        "c_count": c_count,
                        "address": ur.Address,
                    },
                    "errors_found": errors_found,
                    "shapes": sheet_shapes,
                    "cells": sheet_cells,
                }

            except Exception as e:
                logger.error(f"Error dumping worksheet '{sheet_name}': {e}")
                dump_data["sheets"][sheet_name] = {"error": str(e)}
            finally:
                cell = None
                ur = None
                sheet = None
                gc.collect()

        wb = None
        gc.collect()

    # Write JSON output
    if output_json:
        parent = os.path.dirname(os.path.abspath(output_json))
        if parent:
            os.makedirs(parent, exist_ok=True)
        with open(output_json, "w", encoding="utf-8") as f:
            json.dump(dump_data, f, indent=2, default=str)
        logger.info(f"JSON dump: {output_json}")

    # Write Markdown output
    if output_md:
        parent = os.path.dirname(os.path.abspath(output_md))
        if parent:
            os.makedirs(parent, exist_ok=True)
        md_content = _render_markdown(dump_data, max_md_rows)
        with open(output_md, "w", encoding="utf-8") as f:
            f.write(md_content)
        logger.info(f"Markdown dump: {output_md}")

    return dump_data


def _inspect_workbook_in_process(
    workbook_path: str,
    sheets: list[str],
    output_dir: str = "screenshots",
    custom_range: str | None = None,
    include_data: bool = True,
    include_screenshots: bool = True,
    output_json: str | None = None,
    output_md: str | None = None,
    continue_on_render_error: bool = False,
    include_hidden_sheets: bool = False,
    on_excel_started=None,
) -> dict:
    """Inspect data and render ranges in one isolated, read-only Excel session."""
    result = {
        "success": True,
        "phase": "session_start",
        "screenshots": {},
        "data": None,
        "primary_error": None,
        "dialog_events": [],
        "cleanup": {},
    }
    session = ExcelSession(
        workbook_path,
        visible=False,
        save_on_exit=False,
        kill_on_enter=False,
        read_only=True,
        disable_macros=True,
        on_excel_started=on_excel_started,
    )
    try:
        with session:
            inspection = _inspect_existing_session(
                session,
                workbook_path=workbook_path,
                sheets=sheets,
                output_dir=output_dir,
                custom_range=custom_range,
                include_data=include_data,
                include_screenshots=include_screenshots,
                output_json=output_json,
                output_md=output_md,
                continue_on_render_error=continue_on_render_error,
                include_hidden_sheets=include_hidden_sheets,
            )
            result["screenshots"] = inspection["screenshots"]
            result["data"] = inspection["workbook_data"]
        result["phase"] = "complete"
    except Exception as error:
        result["success"] = False
        result["primary_error"] = str(error)
    finally:
        result["dialog_events"] = (
            [event.to_dict() for event in session.watchdog.events]
            if session.watchdog is not None else []
        )
        result["cleanup"] = dict(session.cleanup_result)
    return result


def _inspect_existing_session(
    session: ExcelSession,
    *,
    workbook_path: str,
    sheets: list[str],
    output_dir: str = "screenshots",
    custom_range: str | None = None,
    include_data: bool = True,
    include_screenshots: bool = True,
    output_json: str | None = None,
    output_md: str | None = None,
    continue_on_render_error: bool = False,
    include_hidden_sheets: bool = False,
) -> dict[str, Any]:
    """Inspect workbook state through an existing workflow-owned session."""
    screenshots: dict[str, str] = {}
    workbook_data = None
    if include_screenshots:
        screenshots = export_screenshots(
            workbook_path,
            sheets,
            output_dir,
            custom_range=custom_range,
            _session=session,
            include_hidden_sheets=include_hidden_sheets,
        )
        render_errors = {
            name: value for name, value in screenshots.items()
            if value in ("Not found", "Empty") or str(value).startswith("Error:")
        }
        if render_errors and not continue_on_render_error:
            raise RuntimeError(f"Screenshot rendering failed: {render_errors}")
    if include_data:
        workbook_data = dump_sheet_data(
            workbook_path,
            sheets,
            output_json=output_json,
            output_md=output_md,
            custom_range=custom_range,
            _session=session,
        )
    return {
        "workbook_data": workbook_data,
        "screenshots": screenshots,
    }


def inspect_workbook(
    workbook_path: str,
    sheets: list[str],
    output_dir: str = "screenshots",
    custom_range: str | None = None,
    include_data: bool = True,
    include_screenshots: bool = True,
    output_json: str | None = None,
    output_md: str | None = None,
    continue_on_render_error: bool = False,
    include_hidden_sheets: bool = False,
    timeout_seconds: float = 60,
) -> dict:
    """Run combined workbook inspection in a timeout-controlled worker."""
    from xlvbatools.core.worker import execute_worker_request

    request = {
        "workbook_path": os.path.abspath(workbook_path), "sheets": sheets,
        "output_dir": os.path.abspath(output_dir), "custom_range": custom_range,
        "include_data": include_data, "include_screenshots": include_screenshots,
        "output_json": os.path.abspath(output_json) if output_json else None,
        "output_md": os.path.abspath(output_md) if output_md else None,
        "continue_on_render_error": continue_on_render_error,
        "include_hidden_sheets": include_hidden_sheets,
    }
    return execute_worker_request("inspect", request, timeout=timeout_seconds)


# ── Markdown Report Generation (D6) ──


def _render_markdown(dump_data: dict, max_rows: int) -> str:
    """
    Render dump data as a rich, agent-readable Markdown document.

    Includes metadata, named ranges with scope, per-sheet grids with
    row/column headers, error highlights, formula annotations, shapes tables,
    and a separate formula listing.
    """
    lines = []
    metadata = dump_data.get("metadata", {})
    lines.append(f"# Workbook Data Dump: {metadata.get('workbook', '?')}")
    lines.append(f"- **Generated At**: {metadata.get('timestamp', '?')}")
    lines.append("- **Sheets Processed**: " + ", ".join(metadata.get("sheets_processed", [])))
    lines.append("\n---\n")

    # Named Ranges (D5 — with scope)
    nr = dump_data.get("named_ranges", {})
    if nr:
        lines.append("## Named Ranges")
        lines.append("| Name | Scope | Refers To | Evaluated Value |")
        lines.append("|---|---|---|---|")
        for name, info in sorted(nr.items()):
            ref = info.get("refers_to") or ""
            scope = info.get("scope") or "Workbook"
            val = str(info.get("value")) if info.get("value") is not None else ""
            if info.get("error"):
                val = f"**❌ {info['error']}**"
            lines.append(f"| {name} | {scope} | `{ref}` | {val} |")
        lines.append("\n---\n")

    # Per-sheet data
    for sheet_name, sheet_data in dump_data.get("sheets", {}).items():
        lines.append(f"## Sheet: {sheet_name}")

        if "error" in sheet_data:
            lines.append(f"**Error**: {sheet_data['error']}\n")
            continue

        bounds = sheet_data.get("bounds", {})
        cells = sheet_data.get("cells", {})

        if bounds.get("r_count", 0) == 0:
            lines.append("*(This sheet is empty)*\n")
            continue

        lines.append(f"- **Used Range**: {bounds.get('address', '?')} "
                      f"({bounds.get('r_count', 0)} rows, {bounds.get('c_count', 0)} columns)")
        lines.append(f"- **Total Non-Empty Cells**: {len(cells)}")

        # Error summary (D3)
        errors = sheet_data.get("errors_found", [])
        if errors:
            lines.append(f"- **🚨 Errors Detected**: **{len(errors)} cells contain errors**")
            lines.append("")
            lines.append("### 🚨 Cell Errors Detected")
            lines.append("| Cell | Value | Formula |")
            lines.append("|---|---|---|")
            for addr in sorted(errors):
                info = cells.get(addr, {})
                form_str = f"`{info.get('formula', '')}`" if info.get("formula") else ""
                lines.append(f"| {addr} | **❌ {info.get('error_type', '?')}** | {form_str} |")
            lines.append("")
        else:
            lines.append("")

        # Shapes / interactive controls
        shapes = sheet_data.get("shapes", [])
        if shapes:
            lines.append("### 🎛️ Interactive Controls & Buttons")
            lines.append("| Shape Name | Type | Label / Text | Text Accessor | Linked Macro (OnAction) | Linked Cell |")
            lines.append("|---|---|---|---|---|---|")
            for shape in shapes:
                lines.append(
                    f"| {shape['name']} | {shape['type']} | {shape['text']} "
                    f"| `{shape.get('text_accessor', '')}` "
                    f"| `{shape['on_action']}` | `{shape['linked_cell']}` |"
                )
            lines.append("")

        # Grid table with row/column headers
        r_start = bounds.get("r_start", 1)
        c_start = bounds.get("c_start", 1)
        r_count = bounds.get("r_count", 0)
        c_count = bounds.get("c_count", 0)

        if r_count <= max_rows:
            grid = [["" for _ in range(c_count)] for _ in range(r_count)]
            formulas_in_grid = {}

            for addr, cell in cells.items():
                r_idx = cell["row"] - r_start
                c_idx = cell["col"] - c_start

                text = cell.get("text", "")
                if cell.get("is_error"):
                    grid[r_idx][c_idx] = f"**❌ {cell['error_type']}**"
                elif cell.get("formula"):
                    formulas_in_grid[addr] = cell["formula"]
                    # Mark cell with asterisk to indicate formula
                    grid[r_idx][c_idx] = f"*{sanitize_markdown_cell(text)}*"
                else:
                    grid[r_idx][c_idx] = sanitize_markdown_cell(text)

            # Render Markdown table with row/column headers
            headers = [""] + [get_column_letter(c) for c in range(c_start, c_start + c_count)]
            lines.append("| " + " | ".join(headers) + " |")
            lines.append("| " + " | ".join(["---"] * len(headers)) + " |")

            for r in range(r_count):
                row_label = str(r_start + r)
                row_cells = [row_label] + grid[r]
                lines.append("| " + " | ".join(row_cells) + " |")

            lines.append("\n*Note: Italicized cells contain formulas.*\n")
        else:
            lines.append(
                f"*(Grid table omitted because row count {r_count} "
                f"exceeds maximum limit of {max_rows})*\n"
            )

        # Separate formula listing
        formula_cells = {addr: info for addr, info in cells.items() if info.get("formula")}
        if formula_cells:
            lines.append(f"### Formulas in '{sheet_name}'")
            lines.append("| Cell | Value | Formula |")
            lines.append("|---|---|---|")
            for addr in sorted(formula_cells.keys()):
                info = formula_cells[addr]
                val_str = sanitize_markdown_cell(info.get("text", ""))
                form_str = f"`{info['formula']}`"
                lines.append(f"| {addr} | {val_str} | {form_str} |")
            lines.append("")
        else:
            lines.append("*(No formulas found in this sheet)*\n")

        lines.append("\n---\n")

    return "\n".join(lines)
