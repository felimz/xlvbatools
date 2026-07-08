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
    try:
        for name in wb.Names:
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
        logger.warning(f"Failed to extract shapes for sheet {sheet.Name}: {e}")
    return shapes_info


def export_screenshots(
    workbook_path: str,
    sheets: list[str],
    output_dir: str,
    custom_range: str | None = None,
) -> dict[str, str]:
    """
    Export worksheet screenshots with column/row headers composited via Pillow.

    Includes a retry loop for CopyPicture (common COM flake), proper Pillow
    header composition with column letters and row numbers, gridline overlay,
    and graceful custom_range fallback.

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
                    logger.warning(f"Worksheet '{sheet_name}' not found in workbook")
                    results[sheet_name] = "Not found"
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

                logger.info(f"Copying sheet '{sheet_name}' for screenshot (UsedRange: {ur.Address})")

                # Copy sheet to a new temporary workbook to freeze values
                sheet.Copy()
                wb_temp = excel.ActiveWorkbook
                ws_temp = wb_temp.ActiveSheet

                # Freeze values to prevent external reference / formula errors
                try:
                    ws_temp.UsedRange.Copy()
                    ws_temp.UsedRange.PasteSpecial(Paste=-4163)  # xlPasteValues
                    excel.CutCopyMode = False
                except Exception as e:
                    logger.warning(f"Failed to freeze values on temp sheet: {e}")

                # Get column widths and row heights in points (for Pillow scaling)
                widths = [ws_temp.Cells(r_start, c).Width for c in range(c_start, c_end + 1)]
                heights = [ws_temp.Cells(r, c_start).Height for r in range(r_start, r_end + 1)]

                # Enable gridlines
                try:
                    excel.ActiveWindow.DisplayGridlines = True
                except Exception:
                    pass

                export_range = ws_temp.Range(
                    ws_temp.Cells(r_start, c_start),
                    ws_temp.Cells(r_end, c_end),
                )

                # Build output filename (D7 — custom range in filename)
                if custom_range:
                    range_clean = custom_range.replace(":", "_").replace("$", "")
                    filename = f"{sheet_name.lower().replace(' ', '_')}_zoom_{range_clean}.png"
                else:
                    filename = f"{sheet_name.lower().replace(' ', '_')}_screenshot.png"
                out_path = os.path.abspath(os.path.join(output_dir, filename))

                # CopyPicture with retry loop (D7)
                max_retries = 5
                for attempt in range(max_retries):
                    try:
                        ws_temp.Activate()
                        time.sleep(0.5)
                        try:
                            export_range.Select()
                        except Exception:
                            pass
                        export_range.CopyPicture(Appearance=1, Format=2)  # xlScreen, xlBitmap
                        break
                    except Exception as e:
                        if attempt == max_retries - 1:
                            logger.error(f"CopyPicture failed after {max_retries} attempts: {e}")
                            raise
                        logger.warning(
                            f"CopyPicture attempt {attempt + 1} failed, retrying in 0.5s... Error: {e}"
                        )
                        time.sleep(0.5)

                # Paste to a temporary chart object
                chart_obj = ws_temp.ChartObjects().Add(
                    0, 0, export_range.Width, export_range.Height
                )
                chart = chart_obj.Chart

                try:
                    chart_obj.Select()
                except Exception:
                    pass

                chart.Paste()

                # Size the pasted shape to match the chart area
                try:
                    if chart.Shapes.Count > 0:
                        shape = chart.Shapes.Item(1)
                        shape.Left = 0
                        shape.Top = 0
                        shape.Width = export_range.Width
                        shape.Height = export_range.Height
                except Exception as e:
                    logger.warning(f"Could not format pasted shape: {e}")

                time.sleep(1.0)

                chart.Export(out_path, "PNG")
                chart_obj.Delete()
                wb_temp.Close(False)

                # Apply column/row headers and gridline overlay with Pillow (D7)
                try:
                    from PIL import Image, ImageDraw

                    if os.path.exists(out_path):
                        img = Image.open(out_path).convert("RGBA")

                        scale_x = img.width / sum(widths) if sum(widths) > 0 else 1
                        scale_y = img.height / sum(heights) if sum(heights) > 0 else 1

                        pad_left = 45   # Width of row header column
                        pad_top = 22    # Height of column header row

                        final_w = img.width + pad_left
                        final_h = img.height + pad_top
                        final_img = Image.new("RGBA", (final_w, final_h), (255, 255, 255, 255))

                        # Paste raw Excel screenshot
                        final_img.paste(img, (pad_left, pad_top))

                        draw = ImageDraw.Draw(final_img)

                        # Load font
                        try:
                            from PIL import ImageFont
                            font = ImageFont.truetype("arial.ttf", 12)
                        except Exception:
                            from PIL import ImageFont
                            font = ImageFont.load_default()

                        # Draw header backgrounds (gray)
                        draw.rectangle([(0, 0), (final_w, pad_top)], fill=(240, 240, 240, 255))
                        draw.rectangle([(0, 0), (pad_left, final_h)], fill=(240, 240, 240, 255))

                        # Draw dividing lines
                        draw.line([(pad_left, 0), (pad_left, final_h)], fill=(180, 180, 180, 255), width=1)
                        draw.line([(0, pad_top), (final_w, pad_top)], fill=(180, 180, 180, 255), width=1)

                        # Draw column letter headers
                        curr_x = pad_left
                        for c_idx in range(c_start, c_end + 1):
                            w_px = int(widths[c_idx - c_start] * scale_x)
                            col_letter = get_column_letter(c_idx)

                            try:
                                bbox = draw.textbbox((0, 0), col_letter, font=font)
                                text_w = bbox[2] - bbox[0]
                                text_h = bbox[3] - bbox[1]
                            except AttributeError:
                                text_w, text_h = draw.textsize(col_letter, font=font)

                            tx = curr_x + (w_px - text_w) // 2
                            ty = (pad_top - text_h) // 2
                            draw.text((tx, ty), col_letter, fill=(50, 50, 50, 255), font=font)

                            curr_x += w_px
                            draw.line([(curr_x, 0), (curr_x, final_h)], fill=(220, 220, 220, 255), width=1)

                        # Draw row number headers
                        curr_y = pad_top
                        for r_idx in range(r_start, r_end + 1):
                            h_px = int(heights[r_idx - r_start] * scale_y)
                            row_str = str(r_idx)

                            try:
                                bbox = draw.textbbox((0, 0), row_str, font=font)
                                text_w = bbox[2] - bbox[0]
                                text_h = bbox[3] - bbox[1]
                            except AttributeError:
                                text_w, text_h = draw.textsize(row_str, font=font)

                            tx = (pad_left - text_w) // 2
                            ty = curr_y + (h_px - text_h) // 2
                            draw.text((tx, ty), row_str, fill=(50, 50, 50, 255), font=font)

                            curr_y += h_px
                            draw.line([(0, curr_y), (final_w, curr_y)], fill=(220, 220, 220, 255), width=1)

                        # Draw outer border
                        draw.rectangle(
                            [(0, 0), (final_w - 1, final_h - 1)],
                            outline=(180, 180, 180, 255), width=1,
                        )

                        # Save composite image
                        final_img.convert("RGB").save(out_path, "PNG")
                        logger.info(f"Applied Pillow headers and grid overlay to screenshot: {out_path}")

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
) -> dict:
    """
    Dump worksheet contents to JSON and/or Markdown.

    Uses a cell-level data model: each non-empty cell is stored as
    ``{address: {row, col, value, text, formula, is_error, error_type}}``.

    Returns the complete dump data dict.
    """
    wb_path = os.path.abspath(workbook_path)
    dump_data = {
        "metadata": {
            "workbook": os.path.basename(wb_path),
            "timestamp": time.strftime("%Y-%m-%dT%H:%M:%S"),
            "sheets_processed": sheets,
        },
        "sheets": {},
    }

    with ExcelSession(wb_path, visible=False, save_on_exit=False) as session:
        excel, wb = session.excel, session.wb

        if dump_names:
            dump_data["named_ranges"] = dump_named_ranges(wb)

        for sheet_name in sheets:
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
                                text = ur.Cells(r + 1, c + 1).Text
                            except Exception:
                                text = str(val) if val is not None else ""

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
            lines.append("| Shape Name | Type | Label / Text | Linked Macro (OnAction) | Linked Cell |")
            lines.append("|---|---|---|---|---|")
            for shape in shapes:
                lines.append(
                    f"| {shape['name']} | {shape['type']} | {shape['text']} "
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
