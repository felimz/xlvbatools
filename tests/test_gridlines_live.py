"""Native Excel acceptance for workbook and overlay gridline modes."""

import gc
from pathlib import Path
import sys
import xml.etree.ElementTree as ET
import zipfile

from PIL import Image
import pytest

from xlvbatools import CleanupReport, GridlineOptions, InspectStep, Project
from xlvbatools.core.session import ExcelSession


pytestmark = [
    pytest.mark.excel,
    pytest.mark.skipif(sys.platform != "win32", reason="Windows Excel required"),
]


def _saved_gridlines(workbook):
    namespace = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(workbook) as archive:
        sheet = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
    view = sheet.find("x:sheetViews/x:sheetView", namespace)
    return view.get("showGridLines", "1") != "0"


def _boundary_pixels(image, row):
    # Uniform columns/rows let the test probe a cell boundary without depending
    # on the host's DPI, font metrics, or Excel's bitmap rounding.
    x, y = round(image.width / 4), round((row - 0.5) * image.height / 6)
    return [image.getpixel((x + offset, y)) for offset in range(-2, 3)]


@pytest.mark.parametrize("original_visible", [True, False])
def test_live_gridlines_respect_formatting_and_restore_the_saved_view(
    minimal_workbook, tmp_path, original_visible,
):
    with ExcelSession(minimal_workbook, save_on_exit=True) as setup:
        sheet = setup.wb.Worksheets("Sheet1")
        sheet.Range("A1:D6").RowHeight = 36
        sheet.Range("A1:D6").ColumnWidth = 12
        sheet.Range("A1:D6").Interior.ColorIndex = -4142  # No fill.
        sheet.Range("A1").Value = "Content"
        sheet.Range("A2:B2").Merge()
        sheet.Range("A3:B4").Interior.Color = 40 + 80 * 256 + 120 * 65536
        border = sheet.Range("A5:D5").Borders(9)  # xlEdgeBottom.
        border.LineStyle = 1
        border.Color = 255  # Red cell border must remain visible in off mode.
        border.Weight = 2
        setup.excel.ActiveWindow.DisplayGridlines = original_visible
        border = None
        sheet = None
        gc.collect()
    assert CleanupReport.from_mapping(setup.cleanup_result).is_clean
    assert _saved_gridlines(minimal_workbook) is original_visible

    modes = ["forced", "off", "custom", "native"]
    steps = [
        InspectStep(
            mode, ("Sheet1",), output_dir=str(tmp_path / mode),
            cell_range="A1:D6", include_data=False, include_screenshots=True,
            include_headers=False,
            gridlines=(
                GridlineOptions(mode="custom", color="#FF0000", opacity=0.5)
                if mode == "custom" else mode
            ),
        )
        for mode in modes
    ]
    result = Project.open(minimal_workbook).workflow(steps, save=True, timeout=120)
    workflow = result.require_success()
    result.require_clean_shutdown()
    assert _saved_gridlines(minimal_workbook) is original_visible
    images = {}
    try:
        for mode in modes:
            inspection = workflow.step(mode).data
            with Image.open(Path(inspection.screenshots["Sheet1"])) as captured:
                images[mode] = captured.convert("RGBA").convert("RGB")
            diagnostics = inspection.screenshot_diagnostics["Sheet1"]
            assert diagnostics["gridlines"]["mode"] == mode
            assert diagnostics["native_gridlines_original"] is original_visible
            capture = diagnostics["attempts"][-1]["window"]
            assert capture["display_gridlines"] is (original_visible and mode == "native")
        assert len({image.size for image in images.values()}) == 1
        # Fills and merged cells suppress only the native grid. Forced/custom
        # overlays intentionally cross those boundaries.
        assert _boundary_pixels(images["native"], 4) == [(40, 80, 120)] * 5
        assert _boundary_pixels(images["off"], 4) == [(40, 80, 120)] * 5
        style = workflow.step("forced").data.screenshot_diagnostics["Sheet1"]
        color = style["native_grid_color"]
        native_gray = tuple(int(color[i:i + 2], 16) for i in (1, 3, 5))
        assert (native_gray != (255, 255, 255)) and style["native_grid_width"] >= 1
        assert native_gray in _boundary_pixels(images["forced"], 4)
        assert (148, 40, 60) in _boundary_pixels(images["custom"], 4)
        assert _boundary_pixels(images["native"], 2) == [(255, 255, 255)] * 5
        assert native_gray in _boundary_pixels(images["forced"], 2)
        native_empty_boundary = _boundary_pixels(images["native"], 6)
        assert any(pixel != (255, 255, 255) for pixel in native_empty_boundary) is original_visible
        assert _boundary_pixels(images["off"], 6) == [(255, 255, 255)] * 5
        # Off hides gridlines while retaining explicitly formatted cell borders.
        for image in images.values():
            x = round(image.width * 0.625)
            y = style["y_boundaries"][5] - 1
            assert image.getpixel((x, y)) == (255, 0, 0)
    finally:
        for image in images.values():
            image.close()
