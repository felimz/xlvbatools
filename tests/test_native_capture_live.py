"""Byte-exact native body/headings and overlay geometry at Excel zoom levels."""

import gc
import sys

from PIL import ImageChops
import pytest

from xlvbatools import CleanupReport, GridlineOptions
from xlvbatools.core.session import ExcelSession
from xlvbatools.workbook import _native_capture as backend
from xlvbatools.workbook import dumper


pytestmark = [
    pytest.mark.excel,
    pytest.mark.skipif(sys.platform != "win32", reason="Windows Excel required"),
]


@pytest.mark.parametrize("zoom", [75, 100, 125, 150])
def test_native_pixels_and_overlay_edges_are_exact_for_cropped_hidden_ranges(
    minimal_workbook, monkeypatch, zoom,
):
    frames = []
    original_capture = backend._window_pixels

    def record(hwnd):
        image = original_capture(hwnd)
        frames.append((hwnd, image.copy()))
        return image

    monkeypatch.setattr(backend, "_window_pixels", record)
    captured = sheet = target = window = border = font = None
    try:
        with ExcelSession(minimal_workbook, save_on_exit=False) as session:
            try:
                sheet = session.wb.Worksheets(1)
                font = session.wb.Styles("Normal").Font
                font.Name, font.Size = "Arial", 12
                font = None
                for column, width in enumerate((9.25, 17.3, 10.1, 7.7, 14.2), 2):
                    sheet.Columns(column).ColumnWidth = width
                for row, height in enumerate((18, 31.5, 15, 27, 19.5, 36, 24, 15.75, 21), 91):
                    sheet.Rows(row).RowHeight = height
                sheet.Columns(3).Hidden = True
                sheet.Rows(94).Hidden = True
                sheet.Range("B91").Value = "Pixel fidelity"
                sheet.Range("B92:D92").Merge()
                sheet.Range("B93:E96").Interior.Color = 40 + 80 * 256 + 120 * 65536
                border = sheet.Range("B95:F95").Borders(9)
                border.LineStyle, border.Color = 1, 255
                border = None
                edge_colors = {2: (0, 128, 0), 4: (128, 0, 128),
                               5: (0, 0, 255), 6: (255, 128, 0)}
                for column, (red, green, blue) in edge_colors.items():
                    border = sheet.Range(sheet.Cells(91, column), sheet.Cells(99, column)).Borders(10)
                    border.LineStyle = 1
                    border.Color = red + green * 256 + blue * 65536
                    border = None
                window = session.excel.ActiveWindow
                window.Zoom = zoom
                window.DisplayGridlines = True
                window.DisplayHeadings = False
                window.SplitRow, window.SplitColumn = 2, 1
                window.FreezePanes = True
                sheet.Range("G103").Select()
                window.ScrollRow, window.ScrollColumn = 88, 2
                saved = {name: getattr(window, name) for name in (
                    "Zoom", "DisplayHeadings", "DisplayGridlines", "FreezePanes",
                    "SplitRow", "SplitColumn", "ScrollRow", "ScrollColumn",
                )}
                target = sheet.Range("B91:F99")
                captured = backend.capture_native_sheet(session.excel, session.wb, sheet, target)
                assert session.excel.Workbooks.Count == 1
                assert {name: getattr(window, name) for name in saved} == saved, captured.details
                assert captured.details["zoom"] == zoom
                assert captured.x_bounds[1] == captured.x_bounds[2]  # Hidden C.
                assert captured.y_bounds[3] == captured.y_bounds[4]  # Hidden 94.
                assert captured.grid_width == 1
                sample_y = (captured.y_bounds[5] + captured.y_bounds[6]) // 2
                for column, color in edge_colors.items():
                    edge_x = captured.x_bounds[column - 1] - 1
                    assert captured.body.getpixel((edge_x, sample_y)) == color
                tiles = captured.details["tile_geometry"]
                source_hwnd = tiles[0]["source_hwnd"]
                originals = [image for hwnd, image in frames if hwnd == source_hwnd]
                assert len(originals) >= len(tiles)
                originals = originals[-len(tiles):]
                with captured.render(GridlineOptions(), True) as image:
                    left, top = captured.rows.width, captured.columns.height
                    for tile, original in zip(tiles, originals):
                        x, y = tile["x"], tile["y"]
                        sx, sy, right, bottom = tile["source_rectangle"]
                        width, height = right - sx, bottom - sy
                        # Every body pixel and every heading glyph/box comes
                        # unchanged from the actual Excel window; no tolerance.
                        assert image.crop((left+x, top+y, left+x+width, top+y+height)).tobytes() == \
                            original.crop((sx, sy, right, bottom)).tobytes()
                        if y == 0:
                            assert image.crop((left+x, 0, left+x+width, top)).tobytes() == \
                                original.crop((sx, 0, right, sy)).tobytes()
                        if x == 0:
                            assert image.crop((left-sx, top+y, left, top+y+height)).tobytes() == \
                                original.crop((0, sy, sx, bottom)).tobytes()
                with captured.render(GridlineOptions(mode="forced"), False) as forced:
                    # Difference pixels can occur only on measured native
                    # gridline strokes; a one-pixel offset fails this check.
                    difference = ImageChops.difference(captured.body, forced)
                    xs = {value - 1 for value in captured.x_bounds[1:-1] if value > 0}
                    ys = {value - 1 for value in captured.y_bounds[1:-1] if value > 0}
                    for y in range(difference.height):
                        for x in range(difference.width):
                            if difference.getpixel((x, y)) != (0, 0, 0):
                                assert x in xs or y in ys, (x, y)
                    assert difference.getbbox() is not None
                    # An explicit red cell border retains its exact native row.
                    border_y = captured.y_bounds[5] - 1
                    border_x = captured.body.width - 4
                    assert forced.getpixel((border_x, border_y)) == (255, 0, 0)
            finally:
                if captured is not None:
                    captured.close()
                captured = font = border = target = window = sheet = None
                gc.collect()
        assert CleanupReport.from_mapping(session.cleanup_result).is_clean
    finally:
        for _, image in frames:
            image.close()


def test_failed_native_capture_closes_reference_and_cleans_up_owned_excel(
    minimal_workbook, monkeypatch, tmp_path,
):
    calls = []

    def fail(hwnd):
        calls.append(hwnd)
        raise backend.NativeCaptureError("Injected native rendering failure")

    monkeypatch.setattr(backend, "_window_pixels", fail)
    sheet = None
    with pytest.raises(dumper.ScreenshotRenderError) as caught:
        with ExcelSession(minimal_workbook, save_on_exit=False) as session:
            sheet = session.wb.Worksheets(1)
            sheet.Range("A1").Value = "Visible content"
            sheet = None
            dumper.export_screenshots(
                minimal_workbook, ["Sheet1"], str(tmp_path), custom_range="A1:B2",
                _session=session,
            )
    assert len(calls) == 2
    assert caught.value.code == "screenshot_capture_failed"
    assert CleanupReport.from_mapping(session.cleanup_result).is_clean, session.cleanup_result
