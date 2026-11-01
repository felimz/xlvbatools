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
- Native Excel screenshot pixels with optional custom gridlines

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
from contextlib import contextmanager, nullcontext
from typing import Any, Mapping

from xlvbatools.core.session import ExcelSession
from xlvbatools.rendering import GridlineOptions, _coerce_gridlines

logger = logging.getLogger(__name__)

RICH_TEXT_MAX_CHARACTERS = 4096
RICH_TEXT_MAX_RUNS = 256
_RICH_TEXT_FONT_PROPERTIES = (
    "Name", "Size", "Bold", "Italic", "Underline", "Strikethrough",
    "Subscript", "Superscript", "Color", "ColorIndex",
)


class ScreenshotRenderError(RuntimeError):
    """Excel could not produce a trustworthy native range capture."""

    def __init__(
        self,
        message: str,
        *,
        code: str = "render_content_mismatch",
        details: Mapping[str, Any],
    ) -> None:
        super().__init__(message)
        self.code = code
        self.phase = "screenshot_capture"
        self.details = dict(details)

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


def _worksheet_is_renderable(sheet, include_hidden_sheets: bool) -> bool:
    """Return whether a worksheet may be screenshotted under the policy."""
    return include_hidden_sheets or int(sheet.Visible) == -1


def _capture_window_state(excel, wb, sheet, export_range) -> dict[str, Any]:
    """Collect bounded, JSON-safe evidence for a CopyPicture attempt."""
    state: dict[str, Any] = {}
    probes = {
        "excel_visible": lambda: bool(excel.Visible),
        "screen_updating": lambda: bool(excel.ScreenUpdating),
        "display_gridlines": lambda: bool(excel.ActiveWindow.DisplayGridlines),
        "hwnd": lambda: int(excel.Hwnd),
        "window_state": lambda: int(excel.ActiveWindow.WindowState),
        "scroll_row": lambda: int(excel.ActiveWindow.ScrollRow),
        "scroll_column": lambda: int(excel.ActiveWindow.ScrollColumn),
        "active_workbook": lambda: str(excel.ActiveWorkbook.Name),
        "target_workbook": lambda: str(wb.Name),
        "active_sheet": lambda: str(excel.ActiveSheet.Name),
        "target_sheet": lambda: str(sheet.Name),
        "selection": lambda: str(excel.Selection.Address),
        "target_range": lambda: str(export_range.Address),
    }
    for name, probe in probes.items():
        try:
            state[name] = probe()
        except Exception as error:
            state[f"{name}_error"] = f"{type(error).__name__}: {error}"
    try:
        import ctypes
        from ctypes import wintypes

        rect = wintypes.RECT()
        if ctypes.windll.user32.GetWindowRect(int(excel.Hwnd), ctypes.byref(rect)):
            state["window_rect"] = [rect.left, rect.top, rect.right, rect.bottom]
    except Exception as error:
        state["window_rect_error"] = f"{type(error).__name__}: {error}"
    return state


@contextmanager
def _screenshot_gridlines(excel, wb, sheet, options: GridlineOptions):
    """Override only the target sheet's native grid, then restore its view."""
    try:
        active_book = str(excel.ActiveWorkbook.Name) == str(wb.Name)
        active_sheet = active_book and str(excel.ActiveSheet.Name) == str(sheet.Name)
    except AttributeError:
        active_book = active_sheet = False
    if not active_book:
        wb.Activate()
    if not active_sheet:
        sheet.Activate()
    window = excel.ActiveWindow
    original = bool(window.DisplayGridlines)
    try:
        if options.mode != "native":
            window.DisplayGridlines = False
        yield original
    finally:
        try:
            if options.mode != "native":
                window.DisplayGridlines = original
        finally:
            window = None  # A retained exception traceback must not pin this COM proxy.


def _native_image_metrics(image_path: str) -> dict[str, Any]:
    """Measure source pixels before headers or grid overlays can mask a blank export."""
    from PIL import Image

    with Image.open(image_path) as image:
        rgb = image.convert("RGB")
        histogram = rgb.getcolors(maxcolors=4096)
        if histogram is None:
            # RGB has at most 16.7m colors; quantization keeps this bounded for
            # photographic worksheet backgrounds without changing the file.
            histogram = rgb.quantize(colors=256).convert("RGB").getcolors(256)
        meaningful = 0
        darkest = 255
        for count, color in histogram or []:
            red, green, blue = color
            darkest = min(darkest, red, green, blue)
            if min(color) < 175 or max(color) - min(color) > 35:
                meaningful += count
        total = rgb.width * rgb.height
        return {
            "width": rgb.width,
            "height": rgb.height,
            "file_size": os.path.getsize(image_path),
            "meaningful_pixel_count": meaningful,
            "meaningful_pixel_ratio": meaningful / total if total else 0.0,
            "darkest_channel": darkest,
        }


def _image_is_implausibly_blank(
    metrics: Mapping[str, Any], expected_visible_items: int,
) -> bool:
    """Return true only when visible workbook evidence has no bitmap counterpart."""
    if expected_visible_items <= 0:
        return False
    minimum_pixels = max(8, min(64, expected_visible_items * 2))
    return int(metrics.get("meaningful_pixel_count", 0)) < minimum_pixels


def _range_visible_content_count(export_range, sheet) -> int:
    """Count cheap, render-relevant evidence in a bounded screenshot range."""
    count = 0
    rows = int(export_range.Rows.Count)
    columns = int(export_range.Columns.Count)
    values = normalize_2d(export_range.Value2, rows, columns)
    for row in values:
        for value in row:
            if value is not None and str(value) != "":
                count += 1

    first_row = int(export_range.Row)
    first_column = int(export_range.Column)
    last_row = first_row + rows - 1
    last_column = first_column + columns - 1
    try:
        for shape in sheet.Shapes:
            if int(getattr(shape, "Visible", -1)) == 0:
                continue
            top_left = shape.TopLeftCell
            bottom_right = shape.BottomRightCell
            intersects = not (
                int(bottom_right.Row) < first_row
                or int(top_left.Row) > last_row
                or int(bottom_right.Column) < first_column
                or int(top_left.Column) > last_column
            )
            if intersects:
                count += 1
    except Exception:
        logger.debug("Could not count shapes in screenshot range", exc_info=True)
    return count


def _structured_visible_content_counts(workbook_data: Mapping[str, Any]) -> dict[str, int]:
    """Derive per-sheet render evidence from the inspection's cell model."""
    counts: dict[str, int] = {}
    for sheet_name, sheet_data in (workbook_data.get("sheets") or {}).items():
        if not isinstance(sheet_data, Mapping) or sheet_data.get("error"):
            continue
        count = 0
        for cell in (sheet_data.get("cells") or {}).values():
            if not isinstance(cell, Mapping):
                continue
            text = cell.get("text")
            value = cell.get("value")
            if (text is not None and str(text) != "") or (
                value is not None and str(value) != ""
            ):
                count += 1
        counts[str(sheet_name)] = count
    return counts


def _capture_native_range(
    excel, wb, sheet, export_range, out_path: str, *,
    range_width: float, range_height: float, expected_visible_items: int,
    gridlines: GridlineOptions, include_headers: bool,
    max_render_attempts: int = 2,
) -> dict[str, Any]:
    """Validate real worksheet pixels before adding a requested grid overlay."""
    from xlvbatools.workbook._native_capture import capture_native_sheet

    attempts: list[dict[str, Any]] = []
    captured = None
    for attempt in range(1, max_render_attempts + 1):
        candidate = f"{out_path}.capture-{attempt}.png"
        try:
            captured = capture_native_sheet(
                excel, wb, sheet, export_range,
                preserve_borders=gridlines.mode in ("forced", "custom"),
            )
            captured.body.save(candidate, "PNG")
            metrics = _native_image_metrics(candidate)
            window = _capture_window_state(excel, wb, sheet, export_range)
            details = {
                "attempt": attempt, "metrics": metrics, "window": window,
                **captured.details,
            }
            attempts.append(details)
            if _image_is_implausibly_blank(metrics, expected_visible_items):
                continue
            with captured.render(gridlines, include_headers) as rendered:
                if rendered.width * rendered.height > 80_000_000:
                    raise ValueError("Screenshot with headings exceeds the output pixel limit")
                rendered.save(candidate, "PNG")
            os.replace(candidate, out_path)
            return {
                "metrics": metrics, "attempts": attempts,
                **captured.details,
            }
        except Exception as error:
            attempts.append({
                "attempt": attempt,
                "capture_error": f"{type(error).__name__}: {error}",
            })
            # Discard backend tracebacks before the session closes. They can
            # otherwise retain live COM proxies after an unsuccessful capture.
            error.__traceback__ = None
            error.__context__ = None
            error.__cause__ = None
        finally:
            if captured is not None:
                captured.close()
            captured = None
            if os.path.exists(candidate):
                os.remove(candidate)
        if attempt < max_render_attempts:
            time.sleep(0.2 * attempt)
    has_metrics = any("metrics" in item for item in attempts)
    sheet_name, range_address = str(sheet.Name), str(export_range.Address)
    excel = wb = sheet = export_range = None
    gc.collect()
    reason = next(
        (item["capture_error"] for item in reversed(attempts) if "capture_error" in item),
        "the rendered body was implausibly blank",
    )
    raise ScreenshotRenderError(
        f"Excel could not produce a faithful native worksheet capture: {reason}",
        code="render_content_mismatch" if has_metrics else "screenshot_capture_failed",
        details={
            "sheet": sheet_name, "range": range_address,
            "expected_visible_items": expected_visible_items, "attempts": attempts,
        },
    )


def export_screenshots(
    workbook_path: str,
    sheets: list[str],
    output_dir: str,
    custom_range: str | None = None,
    render_mode: str = "excel_native",
    include_headers: bool = True,
    gridlines: GridlineOptions | str = "native",
    include_hidden_sheets: bool = False,
    dpi: int = 144,
    expected_visible_content: Mapping[str, int] | None = None,
    continue_on_render_error: bool = False,
    diagnostics: dict[str, Any] | None = None,
    _session: ExcelSession | None = None,
) -> dict[str, str]:
    """
    Export native Excel pixels and optional native row/column headings.

    Captures use bounded attempts and exact native cell boundaries; custom
    gridlines are composited without resizing the worksheet or its headings.

    Returns dict mapping sheet name -> output file path (or error message).
    """
    options = _coerce_gridlines(gridlines)
    if not isinstance(include_headers, bool):
        raise TypeError("include_headers must be boolean")
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
        read_only=True,
        allow_workbook_events=False,
        allow_macro_execution=False,
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

                range_visible_items = _range_visible_content_count(
                    export_range, sheet,
                )
                expected_items = max(
                    range_visible_items,
                    int(expected_visible_content.get(sheet_name, 0))
                    if expected_visible_content is not None else 0,
                )
                with _screenshot_gridlines(excel, wb, sheet, options) as original_gridlines:
                    capture_diagnostics = _capture_native_range(
                        excel,
                        wb,
                        sheet,
                        export_range,
                        out_path,
                        range_width=range_width,
                        range_height=range_height,
                        expected_visible_items=expected_items,
                        gridlines=options, include_headers=include_headers,
                    )
                capture_diagnostics["gridlines"] = options.to_dict()
                capture_diagnostics["native_gridlines_original"] = original_gridlines
                capture_diagnostics["include_headers"] = include_headers
                if diagnostics is not None:
                    diagnostics[sheet_name] = capture_diagnostics

                logger.info(f"Saved sheet '{sheet_name}' screenshot to {out_path}")
                results[sheet_name] = out_path

            except ScreenshotRenderError as error:
                if diagnostics is not None:
                    diagnostics[sheet_name] = {
                        "success": False,
                        "code": error.code,
                        "message": str(error),
                        "details": error.details,
                    }
                if not continue_on_render_error:
                    excel = wb = sheet = ur = export_range = None
                    gc.collect()
                    raise
                results[sheet_name] = f"Error: {error}"
                logger.error(f"Screenshot failed for {sheet_name}: {error}")
            except Exception as e:
                results[sheet_name] = f"Error: {e}"
                logger.error(f"Screenshot failed for {sheet_name}: {e}")

    return results


def _dump_cell_rich_text(
    cell: Any,
    text: str,
    *,
    max_characters: int = RICH_TEXT_MAX_CHARACTERS,
    max_runs: int = RICH_TEXT_MAX_RUNS,
) -> dict[str, Any]:
    """Return bounded, 1-based partial font runs for one Excel cell."""
    text = str(text)
    text_length = len(text)
    if text_length == 0:
        return {
            "status": "complete",
            "text_length": 0,
            "characters_inspected": 0,
            "truncated": False,
            "runs": [],
        }

    inspected = min(text_length, max_characters)
    runs: list[dict[str, Any]] = []
    start = 1
    try:
        while start <= inspected and len(runs) < max_runs:
            first_signature = _rich_text_font_signature(cell, start, 1)
            remaining = inspected - start + 1
            length = _largest_matching_font_span(
                cell, start, remaining, first_signature,
            )
            runs.append({
                "start": start,
                "length": length,
                "text": text[start - 1:start - 1 + length],
                "font": dict(first_signature),
            })
            start += length
    except Exception as error:
        return {
            "status": "unsupported",
            "text_length": text_length,
            "characters_inspected": max(0, start - 1),
            "truncated": False,
            "runs": runs,
            "error": f"{type(error).__name__}: {error}",
        }

    truncated = inspected < text_length or start <= inspected
    return {
        "status": "truncated" if truncated else "complete",
        "text_length": text_length,
        "characters_inspected": min(inspected, start - 1),
        "truncated": truncated,
        "runs": runs,
    }


def _largest_matching_font_span(
    cell: Any,
    start: int,
    maximum: int,
    signature: tuple[tuple[str, Any], ...],
) -> int:
    """Find a maximal uniform run using logarithmic COM span probes."""
    if maximum <= 1:
        return 1
    low = 1
    high = maximum
    while low < high:
        candidate = (low + high + 1) // 2
        if _rich_text_font_signature(cell, start, candidate) == signature:
            low = candidate
        else:
            high = candidate - 1
    return low


def _rich_text_font_signature(
    cell: Any, start: int, length: int,
) -> tuple[tuple[str, Any], ...]:
    characters = None
    font = None
    try:
        characters = _cell_characters(cell, start, length)
        font = characters.Font
        return tuple(
            (name.casefold(), _json_safe_com_scalar(getattr(font, name)))
            for name in _RICH_TEXT_FONT_PROPERTIES
        )
    finally:
        font = None
        characters = None


def _cell_characters(cell: Any, start: int, length: int) -> Any:
    """Acquire Excel Characters across pywin32 dispatch variants."""
    getter = getattr(cell, "GetCharacters", None)
    if callable(getter):
        try:
            return getter(Start=start, Length=length)
        except TypeError:
            return getter(start, length)

    characters = getattr(cell, "Characters")
    if callable(characters):
        try:
            return characters(Start=start, Length=length)
        except TypeError:
            return characters(start, length)
    raise TypeError("Excel cell does not expose a callable Characters accessor")


def _json_safe_com_scalar(value: Any) -> Any:
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    return str(value)


def dump_sheet_data(
    workbook_path: str,
    sheets: list[str],
    output_json: str | None = None,
    output_md: str | None = None,
    custom_range: str | None = None,
    dump_names: bool = True,
    max_md_rows: int = 500,
    include_rich_text: bool = False,
    _session: ExcelSession | None = None,
) -> dict:
    """
    Dump worksheet contents to JSON and/or Markdown.

    Uses a cell-level data model: each non-empty cell is stored as
    ``{address: {row, col, value, text, formula, is_error, error_type}}``.
    When requested, ``rich_text`` adds bounded partial font runs without
    making per-character COM calls.

    Returns the complete dump data dict.
    """
    wb_path = os.path.abspath(workbook_path)
    dump_data: dict[str, Any] = {
        "metadata": {
            "workbook": os.path.basename(wb_path),
            "timestamp": time.strftime("%Y-%m-%dT%H:%M:%S"),
            "sheets_processed": sheets,
            "rich_text": {
                "included": include_rich_text,
                "max_characters_per_cell": RICH_TEXT_MAX_CHARACTERS,
                "max_runs_per_cell": RICH_TEXT_MAX_RUNS,
            },
        },
        "sheets": {},
    }

    session_context = nullcontext(_session) if _session is not None else ExcelSession(
        wb_path, visible=False, save_on_exit=False, kill_on_enter=False,
        read_only=True,
        allow_workbook_events=False,
        allow_macro_execution=False,
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
                            rich_text = None
                            try:
                                cell = ur.Cells(r + 1, c + 1)
                                text = cell.Text
                                if include_rich_text:
                                    rich_text = _dump_cell_rich_text(
                                        cell,
                                        val if isinstance(val, str) else str(text or ""),
                                    )
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

                            cell_data = {
                                "row": row_num,
                                "col": col_num,
                                "value": val,
                                "text": text,
                                "formula": form if is_formula else None,
                                "is_error": is_error,
                                "error_type": error_type,
                            }
                            if include_rich_text:
                                cell_data["rich_text"] = rich_text or {
                                    "status": "unsupported",
                                    "text_length": len(str(text or "")),
                                    "characters_inspected": 0,
                                    "truncated": False,
                                    "runs": [],
                                    "error": "Characters accessor unavailable",
                                }
                            sheet_cells[addr] = cell_data

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
    include_rich_text: bool = False,
    gridlines: GridlineOptions | str = "native",
    include_headers: bool = True,
    on_excel_started=None,
    on_phase=None,
) -> dict:
    """Inspect data and render ranges in one isolated, read-only Excel session."""
    gridline_options = _coerce_gridlines(gridlines)
    if not isinstance(include_headers, bool):
        raise TypeError("include_headers must be boolean")
    result = {
        "success": True,
        "phase": "session_start",
        "screenshots": {},
        "screenshot_diagnostics": {},
        "data": None,
        "primary_error": None,
        "error": None,
        "dialog_events": [],
        "cleanup": {},
    }
    session = ExcelSession(
        workbook_path,
        visible=False,
        save_on_exit=False,
        kill_on_enter=False,
        read_only=True,
        allow_workbook_events=False,
        allow_macro_execution=False,
        on_excel_started=on_excel_started,
    )

    def publish_phase(phase: str) -> None:
        result["phase"] = phase
        if on_phase is not None:
            on_phase(phase)

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
                include_rich_text=include_rich_text,
                gridlines=gridline_options,
                include_headers=include_headers,
                on_phase=publish_phase,
            )
            result["screenshots"] = inspection["screenshots"]
            result["screenshot_diagnostics"] = inspection[
                "screenshot_diagnostics"
            ]
            result["data"] = inspection["workbook_data"]
        result["phase"] = "complete"
    except Exception as error:
        result["success"] = False
        result["primary_error"] = str(error)
        result["error"] = {
            "code": getattr(error, "code", "inspection_failed"),
            "message": str(error),
            "type": type(error).__name__,
            "details": dict(getattr(error, "details", {}) or {}),
        }
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
    include_rich_text: bool = False,
    gridlines: GridlineOptions | str = "native",
    include_headers: bool = True,
    on_phase=None,
) -> dict[str, Any]:
    """Inspect workbook state through an existing workflow-owned session."""
    gridline_options = _coerce_gridlines(gridlines)
    if not isinstance(include_headers, bool):
        raise TypeError("include_headers must be boolean")
    screenshots: dict[str, str] = {}
    screenshot_diagnostics: dict[str, Any] = {}
    workbook_data = None
    if include_data:
        if on_phase is not None:
            on_phase("inspect_data")
        workbook_data = dump_sheet_data(
            workbook_path,
            sheets,
            output_json=output_json,
            output_md=output_md,
            custom_range=custom_range,
            include_rich_text=include_rich_text,
            _session=session,
        )
    if include_screenshots:
        if on_phase is not None:
            on_phase("screenshot_capture")
        expected_visible_content = (
            _structured_visible_content_counts(workbook_data)
            if workbook_data is not None else None
        )
        screenshots = export_screenshots(
            workbook_path,
            sheets,
            output_dir,
            custom_range=custom_range,
            _session=session,
            include_hidden_sheets=include_hidden_sheets,
            expected_visible_content=expected_visible_content,
            continue_on_render_error=continue_on_render_error,
            diagnostics=screenshot_diagnostics,
            gridlines=gridline_options,
            include_headers=include_headers,
        )
        render_errors = {
            name: value for name, value in screenshots.items()
            if value in ("Not found", "Empty") or str(value).startswith("Error:")
        }
        if render_errors and not continue_on_render_error:
            raise RuntimeError(f"Screenshot rendering failed: {render_errors}")
    return {
        "workbook_data": workbook_data,
        "screenshots": screenshots,
        "screenshot_diagnostics": screenshot_diagnostics,
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
    include_rich_text: bool = False,
    timeout_seconds: float = 60,
    gridlines: GridlineOptions | str = "native",
    include_headers: bool = True,
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
        "include_rich_text": include_rich_text,
        "gridlines": _coerce_gridlines(gridlines).to_dict(),
        "include_headers": include_headers,
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
