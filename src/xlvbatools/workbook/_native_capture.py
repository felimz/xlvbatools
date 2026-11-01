"""Capture Excel's own grid and headings without chart scaling or font recreation."""

from contextlib import contextmanager
from collections import Counter
import ctypes
from ctypes import wintypes
from dataclasses import dataclass
import gc
import time
from typing import Any

from PIL import Image

from xlvbatools.rendering import GridlineOptions, _paint_gridlines


class NativeCaptureError(RuntimeError):
    """Exact native pixels or their cell geometry could not be established."""


@dataclass
class NativePixels:
    body: Any
    columns: Any
    rows: Any
    corner: Any
    x_bounds: list[int]
    y_bounds: list[int]
    grid_color: str
    grid_width: int
    details: dict[str, Any]
    protected_edges: tuple[tuple[str, int, int, int], ...] = ()

    def render(self, options: GridlineOptions, include_headers: bool):
        body = self.body.copy()
        _paint_gridlines(
            body, self.x_bounds, self.y_bounds, options, self.grid_color, self.grid_width,
            self.protected_edges,
        )
        if not include_headers:
            return body
        left, top = self.rows.width, self.columns.height
        image = Image.new("RGB", (body.width + left, body.height + top))
        image.paste(body, (left, top))
        image.paste(self.columns, (left, 0))
        image.paste(self.rows, (0, top))
        image.paste(self.corner, (0, 0))
        body.close()
        return image

    def close(self):
        for item in (self.body, self.columns, self.rows, self.corner):
            item.close()


@contextmanager
def _physical_pixels():
    user = ctypes.WinDLL("user32", use_last_error=True)
    user.SetThreadDpiAwarenessContext.argtypes = [wintypes.HANDLE]
    user.SetThreadDpiAwarenessContext.restype = wintypes.HANDLE
    old = user.SetThreadDpiAwarenessContext(wintypes.HANDLE(-4))
    if not old:
        raise NativeCaptureError("Per-monitor DPI-aware native capture is unavailable")
    try:
        yield
    finally:
        user.SetThreadDpiAwarenessContext(old)


def _window_pixels(hwnd: int):
    import win32gui  # type: ignore[import-untyped]

    class BitmapInfo(ctypes.Structure):
        _fields_ = [
            ("size", wintypes.DWORD), ("width", wintypes.LONG), ("height", wintypes.LONG),
            ("planes", wintypes.WORD), ("bits", wintypes.WORD),
            ("compression", wintypes.DWORD), ("image_size", wintypes.DWORD),
            ("xppm", wintypes.LONG), ("yppm", wintypes.LONG),
            ("colors", wintypes.DWORD), ("important", wintypes.DWORD),
        ]

    user = ctypes.WinDLL("user32", use_last_error=True)
    gdi = ctypes.WinDLL("gdi32", use_last_error=True)
    user.GetDC.argtypes, user.GetDC.restype = [wintypes.HWND], wintypes.HDC
    user.ReleaseDC.argtypes = [wintypes.HWND, wintypes.HDC]
    user.PrintWindow.argtypes = [wintypes.HWND, wintypes.HDC, wintypes.UINT]
    gdi.CreateCompatibleDC.argtypes, gdi.CreateCompatibleDC.restype = [wintypes.HDC], wintypes.HDC
    gdi.CreateDIBSection.argtypes = [
        wintypes.HDC, ctypes.c_void_p, wintypes.UINT,
        ctypes.POINTER(ctypes.c_void_p), wintypes.HANDLE, wintypes.DWORD,
    ]
    gdi.CreateDIBSection.restype = wintypes.HBITMAP
    gdi.SelectObject.argtypes = [wintypes.HDC, wintypes.HANDLE]
    gdi.SelectObject.restype = wintypes.HANDLE
    gdi.DeleteObject.argtypes, gdi.DeleteDC.argtypes = [wintypes.HANDLE], [wintypes.HDC]
    _, _, width, height = win32gui.GetClientRect(hwnd)
    if width <= 0 or height <= 0 or width * height > 80_000_000:
        raise NativeCaptureError("Native Excel viewport exceeds the capture pixel limit")
    dc = memory = bitmap = old = None
    try:
        dc = user.GetDC(hwnd)
        memory = gdi.CreateCompatibleDC(dc)
        bits = ctypes.c_void_p()
        info = BitmapInfo(
            ctypes.sizeof(BitmapInfo), width, -height, 1, 32, 0,
            width * height * 4, 0, 0, 0, 0,
        )
        bitmap = gdi.CreateDIBSection(dc, ctypes.byref(info), 0, ctypes.byref(bits), None, 0)
        if not dc or not memory or not bitmap or not bits.value:
            raise NativeCaptureError("Could not allocate the native capture surface")
        old = gdi.SelectObject(memory, bitmap)
        if not user.PrintWindow(hwnd, memory, 3):
            raise NativeCaptureError("Excel did not render its native window")
        return Image.frombytes(
            "RGB", (width, height), ctypes.string_at(bits, width * height * 4), "raw", "BGRX",
        )
    finally:
        if old:
            gdi.SelectObject(memory, old)
        if bitmap:
            gdi.DeleteObject(bitmap)
        if memory:
            gdi.DeleteDC(memory)
        if dc:
            user.ReleaseDC(hwnd, dc)


def _frame(window):
    import win32gui

    return win32gui.GetAncestor(int(window.Hwnd), 2)


def _canvas(window):
    import win32gui

    candidates = []
    win32gui.EnumChildWindows(
        _frame(window),
        lambda hwnd, _: candidates.append(hwnd)
        if win32gui.GetClassName(hwnd) == "EXCEL7" and win32gui.IsWindowVisible(hwnd) else None,
        None,
    )
    if not candidates:
        raise NativeCaptureError("The owned Excel worksheet window is unavailable")
    return max(candidates, key=lambda hwnd: win32gui.GetClientRect(hwnd)[2])


def _paint(hwnd):
    import win32gui
    from xlvbatools.core.session import ExcelSession

    win32gui.RedrawWindow(hwnd, None, None, 0x0001 | 0x0004 | 0x0080 | 0x0100)
    win32gui.UpdateWindow(hwnd)
    ExcelSession._pump_shutdown_messages()
    time.sleep(0.2)


def _grid_origin(window, origin, size):
    width, height = size
    try:
        left = int(window.PointsToScreenPixelsX(0)) - origin[0]
        top = int(window.PointsToScreenPixelsY(0)) - origin[1]
        if not (0 < left < width and 0 < top < height):
            raise NativeCaptureError("Could not locate native worksheet pixels")
        return left, top
    finally:
        window = None


def _axis_bounds(pixels, start, sizes, grid_color, grid_width):
    """Read cell edges from the native unformatted ruler, including hidden cells."""
    positions = []
    cursor = 0
    while cursor < len(pixels):
        if pixels[cursor] == grid_color:
            end = cursor + 1
            while end < len(pixels) and pixels[end] == grid_color:
                end += 1
            if (end - cursor) % grid_width:
                raise NativeCaptureError("Native grid stroke does not match its measured width")
            positions.extend(range(cursor + grid_width, end + 1, grid_width))
            cursor = end
        else:
            cursor += 1
    bounds, position = [0], 0
    for offset, point_size in enumerate(sizes):
        if not point_size:
            bounds.append(bounds[-1])
            continue
        if position >= len(positions):
            if not position:
                raise NativeCaptureError("A cell does not fit the native capture viewport")
            return start + offset - 1, bounds
        bounds.append(positions[position])
        position += 1
    return start + len(sizes) - 1, bounds


def _grid_style(pixels, left, top):
    background = pixels.getpixel((left + 2, top + 2))
    rulers = [
        [pixels.getpixel((x, top + 2)) for x in range(left, pixels.width - 128)],
        [pixels.getpixel((left + 2, y)) for y in range(top, pixels.height - 128)],
    ]
    counts = Counter(color for ruler in rulers for color in ruler if color != background)
    if not counts:
        raise NativeCaptureError("Excel's automatic grid style could not be sampled")
    color = counts.most_common(1)[0][0]
    strokes = []
    for ruler in rulers:
        run = 0
        for pixel in ruler + [background]:
            if pixel == color:
                run += 1
            elif run:
                strokes.append(run)
                run = 0
    return color, min(strokes)


def _heading_origin(pixels, signature, maximum_width):
    """Locate the native select-all corner's right edge without document offsets."""
    # The outer top border changes by one pixel when workbook activation
    # changes. Match the native interior marker; copy all output pixels intact.
    with signature.crop((0, 1, signature.width, signature.height)) as marker:
        expected = marker.tobytes()
    for left in range(signature.width, min(maximum_width, pixels.width) + 1):
        with pixels.crop((left - signature.width, 1, left, signature.height)) as candidate:
            if candidate.tobytes() == expected:
                return left, signature.height
    raise NativeCaptureError("Native worksheet heading pixels did not match their reference")


def _prepare(excel, book, row, column, outside, viewport, signature=None, maximum_header_width=256):
    import win32gui

    book.Activate()
    sheet = book.Worksheets(1)
    window = excel.ActiveWindow
    try:
        sheet.Cells(*outside).Select()
        window.ScrollRow, window.ScrollColumn = row, column
        window.DisplayHeadings = True
        window.WindowState = -4143
        window.Visible = True
        hwnd = _frame(window)
        win32gui.SetWindowPos(hwnd, 0, -30000, -30000, *viewport, 0x0004 | 0x0010)
        excel.Visible = True
        book.Activate()
        sheet.Activate()
        window = excel.ActiveWindow
        window.ScrollRow, window.ScrollColumn = row, column
        hwnd = _frame(window)
        win32gui.SetWindowPos(hwnd, 0, -30000, -30000, *viewport, 0x0004 | 0x0010)
        canvas = _canvas(window)
        _paint(canvas)
        origin = win32gui.ClientToScreen(canvas, (0, 0))
        _, _, width, height = win32gui.GetClientRect(canvas)
        if signature is None:
            left, top = _grid_origin(window, origin, (width, height))
        else:
            with _window_pixels(canvas) as pixels:
                left, top = _heading_origin(pixels, signature, maximum_header_width)
        return canvas, origin, left, top, width, height
    finally:
        sheet = window = book = excel = None


def _border_segments(sheet, first_row, first_column, x_bounds, y_bounds):
    """Retain displayed cell borders while adding gridlines through fills."""
    edges = []

    def inspect(axis, logical, start, end, pixel, bounds):
        region = merge = None
        try:
            if axis == "x":
                region = sheet.Range(sheet.Cells(start, logical), sheet.Cells(end, logical))
                border_index = 10  # xlEdgeRight
            else:
                region = sheet.Range(sheet.Cells(logical, start), sheet.Cells(logical, end))
                border_index = 9  # xlEdgeBottom
            style = region.DisplayFormat.Borders(border_index).LineStyle
            if style in (-4142, 0):
                return
            merged = region.MergeCells
            if start < end and (style is None or merged is not False):
                middle = (start + end) // 2
                inspect(axis, logical, start, middle, pixel, bounds)
                inspect(axis, logical, middle + 1, end, pixel, bounds)
                return
            if merged:
                merge = region.MergeArea
                merge_end = (
                    int(merge.Column + merge.Columns.Count - 1) if axis == "x" else
                    int(merge.Row + merge.Rows.Count - 1)
                )
                if logical != merge_end:
                    return
            offset = first_row if axis == "x" else first_column
            edge_start, edge_end = bounds[start - offset], bounds[end - offset + 1]
            if edge_start < edge_end:
                edges.append((axis, pixel, edge_start, edge_end))
        finally:
            region = merge = None

    try:
        last_row, last_column = first_row + len(y_bounds) - 2, first_column + len(x_bounds) - 2
        for index, boundary in enumerate(x_bounds[1:-1], 1):
            if boundary > 0:
                inspect("x", first_column + index - 1, first_row, last_row, boundary - 1, y_bounds)
        for index, boundary in enumerate(y_bounds[1:-1], 1):
            if boundary > 0:
                inspect("y", first_row + index - 1, first_column, last_column, boundary - 1, x_bounds)
        return tuple(edges)
    finally:
        sheet = None


def _match_column_width(column, point_width, initial_width):
    """Match native point width exactly, independent of a book's font metrics."""
    try:
        column.ColumnWidth = initial_width
        column.Hidden = point_width == 0
        if not point_width:
            return
        for _ in range(6):
            actual = float(column.Width)
            if abs(actual - point_width) < 0.0001:
                return
            character_width = float(column.ColumnWidth)
            column.ColumnWidth = min(255, max(1 / 256, character_width * point_width / actual))
        low, high = 1, 255 * 256
        while low <= high:
            middle = (low + high) // 2
            column.ColumnWidth = middle / 256
            actual = float(column.Width)
            if abs(actual - point_width) < 0.0001:
                return
            if actual < point_width:
                low = middle + 1
            else:
                high = middle - 1
        raise NativeCaptureError(f"Could not match the native column width of {point_width} points")
    finally:
        column = None


def capture_native_sheet(excel, book, sheet, target, *, preserve_borders=True) -> NativePixels:
    """Tile real worksheet pixels using an exact, unmerged geometry reference."""
    import win32gui

    reference = source_window = reference_window = reference_sheet = font = signature = None
    source_signature = None
    tiles = []
    source_window = excel.ActiveWindow
    if str(source_window.ActiveSheet.Parent.Name) != str(book.Name):
        source_window = book.Windows(1)
    saved = {name: getattr(source_window, name) for name in (
        "ScrollRow", "ScrollColumn", "DisplayHeadings", "View", "FreezePanes",
        "SplitRow", "SplitColumn", "WindowState",
    )}
    pane_scrolls = [
        (int(source_window.Panes(index).ScrollRow), int(source_window.Panes(index).ScrollColumn))
        for index in range(1, int(source_window.Panes.Count) + 1)
    ]
    active_pane = int(source_window.ActivePane.Index)
    original_sheet = str(source_window.ActiveSheet.Name)
    visible, updating = bool(excel.Visible), bool(excel.ScreenUpdating)
    try:
        selected = str(source_window.RangeSelection.Address)
    except Exception:
        selected = None
    with _physical_pixels():
        hwnd = _frame(source_window)
        rectangle = win32gui.GetWindowRect(hwnd)
        try:
            book.Activate()
            sheet.Activate()
            if bool(source_window.DisplayRightToLeft):
                raise NativeCaptureError("Native screenshots currently require a left-to-right view")
            excel.ScreenUpdating = True
            source_window.FreezePanes = False
            source_window.SplitRow = source_window.SplitColumn = 0
            source_window.View = 1
            source_window.WindowState = -4143
            zoom = float(source_window.Zoom)
            first_row, first_column = int(target.Row), int(target.Column)
            last_row = first_row + int(target.Rows.Count) - 1
            last_column = first_column + int(target.Columns.Count) - 1
            source_window.DisplayHeadings = True
            source_window.Visible = True
            sheet.Cells(min(last_row + 1, 1048576), min(last_column + 1, 16384)).Select()
            source_window.ScrollRow = source_window.ScrollColumn = 1
            win32gui.SetWindowPos(hwnd, 0, -30000, -30000, 1600, 1000, 0x0004 | 0x0010)
            excel.Visible = True
            canvas = _canvas(source_window)
            _paint(canvas)
            origin = win32gui.ClientToScreen(canvas, (0, 0))
            _, _, cw, ch = win32gui.GetClientRect(canvas)
            source_left, source_top = _grid_origin(source_window, origin, (cw, ch))
            with _window_pixels(canvas) as calibration:
                source_signature = calibration.crop((max(0, source_left - 8), 0, source_left, source_top))
            widths = [float(sheet.Columns(i).Width) for i in range(first_column, last_column + 1)]
            heights = [float(sheet.Rows(i).Height) for i in range(first_row, last_row + 1)]
            font = book.Styles("Normal").Font
            font_values = {name: getattr(font, name) for name in ("Name", "Size", "Bold", "Italic")}
            font = None
            user = ctypes.WinDLL("user32")
            user.GetDpiForWindow.argtypes = [wintypes.HWND]
            dpi = int(user.GetDpiForWindow(hwnd))
            scale = dpi / 72 * zoom / 100
            header_w = round(max(11, float(font_values["Size"])) * dpi / 72 * 8 + 48)
            header_h = round(max(11, float(font_values["Size"])) * dpi / 72 * 2 + 24)
            viewport = (
                max(1600, round(max(widths) * scale) + header_w + 250),
                max(1000, round(max(heights) * scale) + header_h + 600),
            )
            excel.Visible = False
            reference = excel.Workbooks.Add()
            reference_sheet = reference.Worksheets(1)
            reference_window = excel.ActiveWindow
            reference_window.Zoom = zoom
            reference_window.DisplayGridlines = True
            reference_window.GridlineColorIndex = -4105
            reference_window.View = 1
            ref_canvas, _, left, top, _, _ = _prepare(
                excel, reference, 1, 1, (20, 20), viewport,
            )
            try:
                with _window_pixels(ref_canvas) as style_pixels:
                    native_color, grid_width = _grid_style(style_pixels, left, top)
                grid_color = "#%02X%02X%02X" % native_color
            finally:
                gc.collect()
            font = reference.Styles("Normal").Font
            for name, value in font_values.items():
                setattr(font, name, value)
            font = None
            for index, width in enumerate(widths, first_column):
                _match_column_width(
                    reference_sheet.Columns(index), width, sheet.Columns(index).ColumnWidth,
                )
            for index, height in enumerate(heights, first_row):
                if height:
                    reference_sheet.Rows(index).RowHeight = height
                reference_sheet.Rows(index).Hidden = height == 0
                if abs(float(reference_sheet.Rows(index).Height) - height) > 0.0001:
                    raise NativeCaptureError("Reference row height differs from the source")
            ref_canvas, _, left, top, _, _ = _prepare(
                excel, reference, 1, 1, (20, 20), viewport,
            )
            with _window_pixels(ref_canvas) as calibration:
                signature = calibration.crop((max(0, left - 8), 0, left, top))
            row, output_height, output_width = first_row, 0, 0
            all_x: list[int] = []
            all_y = [0]
            while row <= last_row:
                column, x_offset, row_end = first_column, 0, None
                row_bounds = None
                row_x: list[int] = []
                while column <= last_column:
                    outside = (min(last_row + 1, 1048576), min(last_column + 1, 16384))
                    if outside[0] == last_row:
                        outside = (max(1, first_row - 1), outside[1])
                    if outside[1] == last_column:
                        outside = (outside[0], max(1, first_column - 1))
                    ref_canvas, ref_origin, left, top, width, height = _prepare(
                        excel, reference, row, column, outside, viewport,
                        signature, header_w,
                    )
                    try:
                        budget_x, budget_y = width - header_w - 128, height - header_h - 128
                        with _window_pixels(ref_canvas) as reference_pixels:
                            col_end, x_bounds = _axis_bounds(
                                [reference_pixels.getpixel((x, top + 2))
                                 for x in range(left, left + budget_x)],
                                column, widths[column - first_column:], native_color, grid_width,
                            )
                            measured_end, y_bounds = _axis_bounds(
                                [reference_pixels.getpixel((left + 2, y))
                                 for y in range(top, top + budget_y)],
                                row, heights[row - first_row:], native_color, grid_width,
                            )
                        if row_end is None:
                            row_end, row_bounds = measured_end, y_bounds
                        elif measured_end != row_end or y_bounds != row_bounds:
                            raise NativeCaptureError("Native row geometry changed between tiles")
                        tile_width, tile_height = x_bounds[-1], y_bounds[-1]
                    finally:
                        gc.collect()
                    book.Activate()
                    sheet.Activate()
                    # Use the actual source worksheet, not Worksheets(1).
                    sheet.Cells(*outside).Select()
                    source_window.ScrollRow, source_window.ScrollColumn = row, column
                    source_window.DisplayHeadings = True
                    win32gui.SetWindowPos(hwnd, 0, -30000, -30000, *viewport, 0x0004 | 0x0010)
                    excel.Visible = True
                    canvas = _canvas(source_window)
                    _paint(canvas)
                    _, _, sw, sh = win32gui.GetClientRect(canvas)
                    with _window_pixels(canvas) as pixels:
                        source_left, source_top = _heading_origin(pixels, source_signature, header_w)
                        if source_left + tile_width > sw - 96 or source_top + tile_height > sh - 96:
                            raise NativeCaptureError("Source tile does not fit the native viewport")
                        tiles.append({
                            "x": x_offset, "y": output_height,
                            "body": pixels.crop((source_left, source_top,
                                                 source_left + tile_width, source_top + tile_height)),
                            "columns": pixels.crop((source_left, 0, source_left + tile_width, source_top)),
                            "rows": pixels.crop((0, source_top, source_left, source_top + tile_height)),
                            "corner": pixels.crop((0, 0, source_left, source_top)),
                            "source_hwnd": canvas,
                            "source_rectangle": [source_left, source_top,
                                                 source_left + tile_width, source_top + tile_height],
                            "logical_bounds": [row, column, row_end, col_end],
                        })
                    row_x.extend(x_offset + edge for edge in x_bounds[:-1])
                    x_offset += tile_width
                    column = col_end + 1
                    if x_offset > 32767 or (output_height + tile_height) * x_offset > 80_000_000:
                        raise NativeCaptureError("Native screenshot exceeds the output pixel limit")
                row_x.append(x_offset)
                if not all_x:
                    all_x, output_width = row_x, x_offset
                elif row_x != all_x:
                    raise NativeCaptureError("Native column geometry changed between tiles")
                assert row_bounds is not None and row_end is not None
                all_y.extend(output_height + edge for edge in row_bounds[1:])
                output_height += row_bounds[-1]
                row = row_end + 1
                if output_height > 32767:
                    raise NativeCaptureError("Native screenshot exceeds the output axis limit")
            if not grid_color or not grid_width:
                raise NativeCaptureError("Could not sample Excel's native grid style")
            left = max(tile["rows"].width for tile in tiles)
            top = tiles[0]["columns"].height
            body = Image.new("RGB", (output_width, output_height))
            columns = Image.new("RGB", (output_width, top))
            rows = Image.new("RGB", (left, output_height))
            corner = max(tiles, key=lambda tile: tile["corner"].width)["corner"].copy()
            for tile in tiles:
                if tile["columns"].height != top:
                    raise NativeCaptureError("Native heading height changed between tiles")
                body.paste(tile["body"], (tile["x"], tile["y"]))
                if tile["y"] == 0:
                    columns.paste(tile["columns"], (tile["x"], 0))
                if tile["x"] == 0:
                    strip = tile["rows"]
                    extra = left - strip.width
                    if extra:
                        with strip.crop((0, 0, 1, strip.height)) as edge:
                            rows.paste(edge.resize((extra, strip.height)), (0, tile["y"]))
                    rows.paste(strip, (extra, tile["y"]))
            return NativePixels(
                body, columns, rows, corner, all_x, all_y, grid_color, grid_width,
                {"capture_backend": "excel_window", "dpi": dpi, "zoom": zoom,
                 "tiles": len(tiles), "header_source": "native_pixels",
                 "x_boundaries": all_x, "y_boundaries": all_y,
                 "native_grid_color": grid_color, "native_grid_width": grid_width,
                 "original_scroll": [saved["ScrollRow"], saved["ScrollColumn"]],
                 "tile_geometry": [
                     {name: tile[name] for name in (
                         "x", "y", "source_hwnd", "source_rectangle", "logical_bounds",
                     )} for tile in tiles
                 ]},
                _border_segments(sheet, first_row, first_column, all_x, all_y)
                if preserve_borders else (),
            )
        finally:
            if signature is not None:
                signature.close()
            if source_signature is not None:
                source_signature.close()
            for tile in tiles:
                for name in ("body", "columns", "rows", "corner"):
                    tile[name].close()
            font = reference_sheet = reference_window = None
            if reference is not None:
                reference.Close(False)
            reference = None
            try:
                book.Activate()
                book.Sheets(original_sheet).Activate()
                source_window.FreezePanes = False
                source_window.View = saved["View"]
                source_window.SplitRow = source_window.SplitColumn = 0
                source_window.ScrollRow, source_window.ScrollColumn = pane_scrolls[0]
                source_window.SplitRow = saved["SplitRow"]
                source_window.SplitColumn = saved["SplitColumn"]
                source_window.FreezePanes = saved["FreezePanes"]
                source_window.DisplayHeadings = saved["DisplayHeadings"]
                if selected is not None:
                    book.Sheets(original_sheet).Range(selected).Select()
                win32gui.SetWindowPos(
                    hwnd, 0, rectangle[0], rectangle[1],
                    rectangle[2] - rectangle[0], rectangle[3] - rectangle[1], 0x0004 | 0x0010,
                )
                source_window.WindowState = saved["WindowState"]
                excel.Visible, excel.ScreenUpdating = visible, updating
                source_window.Panes(active_pane).Activate()
                for index, (pane_row, pane_column) in enumerate(pane_scrolls, 1):
                    source_window.Panes(index).ScrollRow = pane_row
                    source_window.Panes(index).ScrollColumn = pane_column
                source_window.ScrollRow = saved["ScrollRow"]
                source_window.ScrollColumn = saved["ScrollColumn"]
            finally:
                source_window = None
                book = sheet = target = excel = None
                gc.collect()
