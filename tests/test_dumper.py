"""Unit coverage for isolated workbook inspection orchestration."""

import gc
import os
from types import SimpleNamespace

import pytest


class _RichFont:
    def __init__(self, styles):
        defaults = {
            "Name": "Aptos", "Size": 11, "Bold": False, "Italic": False,
            "Underline": 0, "Strikethrough": False, "Subscript": False,
            "Superscript": False, "Color": 0, "ColorIndex": 1,
        }
        for name, default in defaults.items():
            values = {style.get(name, default) for style in styles}
            setattr(self, name, values.pop() if len(values) == 1 else None)


class _RichCharacters:
    def __init__(self, styles):
        self.Font = _RichFont(styles)


class _RichCell:
    def __init__(self, styles):
        self.styles = styles
        self.calls = []

    def GetCharacters(self, Start, Length):
        self.calls.append((Start, Length))
        return _RichCharacters(self.styles[Start - 1:Start - 1 + Length])


@pytest.mark.unit
def test_rich_text_dump_models_partial_font_runs_with_bounded_probes():
    from xlvbatools.workbook.dumper import _dump_cell_rich_text

    styles = [
        *({"Bold": True} for _ in range(5)),
        *({"Italic": True, "Color": 255} for _ in range(5)),
    ]
    cell = _RichCell(styles)

    result = _dump_cell_rich_text(cell, "HelloWorld")

    assert result["status"] == "complete"
    assert [(run["start"], run["length"], run["text"]) for run in result["runs"]] == [
        (1, 5, "Hello"),
        (6, 5, "World"),
    ]
    assert result["runs"][0]["font"]["bold"] is True
    assert result["runs"][1]["font"]["italic"] is True
    assert len(cell.calls) < len(styles)


@pytest.mark.unit
def test_rich_text_dump_reports_character_and_run_truncation():
    from xlvbatools.workbook.dumper import _dump_cell_rich_text

    styles = [{"Bold": index % 2 == 0} for index in range(12)]
    result = _dump_cell_rich_text(
        _RichCell(styles), "abcdefghijkl", max_characters=8, max_runs=3,
    )

    assert result["status"] == "truncated"
    assert result["truncated"] is True
    assert result["text_length"] == 12
    assert result["characters_inspected"] == 3
    assert len(result["runs"]) == 3


@pytest.mark.unit
def test_rich_text_dump_degrades_per_cell_when_characters_are_unsupported():
    from xlvbatools.workbook.dumper import _dump_cell_rich_text

    result = _dump_cell_rich_text(object(), "content")

    assert result["status"] == "unsupported"
    assert result["runs"] == []
    assert "Characters" in result["error"]


@pytest.mark.unit
def test_shared_worker_inspection_result_is_returned(monkeypatch, tmp_path):
    from xlvbatools.core import worker
    from xlvbatools.workbook import dumper

    expected = {"success": True, "phase": "complete", "cleanup": {"pid": 42}}

    calls = []

    def fake_run(operation, arguments, *, timeout):
        calls.append((operation, arguments, timeout))
        return expected

    monkeypatch.setattr(worker, "execute_worker_request", fake_run)
    result = dumper.inspect_workbook("book.xlsm", ["Input"], output_dir=str(tmp_path))

    assert result == expected
    assert calls[0][0] == "inspect"
    assert calls[0][1]["workbook_path"].endswith("book.xlsm")
    assert calls[0][2] == 60


@pytest.mark.unit
def test_excel_session_safe_defaults():
    from xlvbatools.core.session import ExcelSession

    session = ExcelSession("book.xlsm")
    assert session.kill_on_enter is False
    assert session.read_only is False
    assert session.allow_workbook_events is False
    assert session.allow_macro_execution is False
    assert session.allow_vbe_visible is False


@pytest.mark.unit
def test_renderer_does_not_copy_worksheet_source():
    import inspect
    from xlvbatools.workbook import dumper, _native_capture

    source = inspect.getsource(dumper.export_screenshots) + inspect.getsource(
        _native_capture.capture_native_sheet,
    )
    assert "sheet.Copy(" not in source
    assert "excel.Workbooks.Add()" in source
    assert "PrintWindow" in inspect.getsource(_native_capture._window_pixels)


@pytest.mark.unit
def test_native_capture_validates_body_before_headings_and_retains_attempts(monkeypatch, tmp_path):
    from PIL import Image, ImageDraw
    from xlvbatools import GridlineOptions
    from xlvbatools.workbook import dumper, _native_capture

    calls = []
    def capture(*args, **kwargs):
        body = Image.new("RGB", (240, 120), "white")
        if calls:
            ImageDraw.Draw(body).rectangle((20, 20, 80, 50), fill="black")
        calls.append(True)
        return _native_capture.NativePixels(
            body, Image.new("RGB", (240, 29), "black"),
            Image.new("RGB", (40, 120), "black"), Image.new("RGB", (40, 29), "black"),
            [0, 120, 240], [0, 60, 120], "#E0E0E0", 1,
            {"capture_backend": "excel_window"},
        )

    monkeypatch.setattr(_native_capture, "capture_native_sheet", capture)
    target = SimpleNamespace(Address="$B$2:$C$3")
    path = tmp_path / "native.png"
    result = dumper._capture_native_range(
        object(), SimpleNamespace(Name="book"), SimpleNamespace(Name="Input"),
        target, str(path), range_width=120, range_height=60,
        expected_visible_items=3, gridlines=GridlineOptions(), include_headers=True,
    )
    assert len(calls) == 2
    assert len(result["attempts"]) == 2
    assert result["attempts"][0]["metrics"]["meaningful_pixel_count"] == 0
    assert result["attempts"][1]["metrics"]["meaningful_pixel_count"] > 64
    assert result["capture_backend"] == "excel_window"
    assert path.is_file()
    assert list(tmp_path.glob("*.capture-*.png")) == []


@pytest.mark.unit
def test_screenshot_capture_error_has_public_phase_and_code():
    from xlvbatools.workbook.dumper import ScreenshotRenderError

    error = ScreenshotRenderError(
        "copy failed",
        code="screenshot_capture_failed",
        details={"attempts": []},
    )

    assert error.phase == "screenshot_capture"
    assert error.code == "screenshot_capture_failed"
    assert error.details == {"attempts": []}


@pytest.mark.unit
def test_native_bitmap_validation_precedes_overlay(tmp_path):
    from PIL import Image, ImageDraw
    from xlvbatools.workbook.dumper import (
        _image_is_implausibly_blank,
        _native_image_metrics,
    )

    blank = tmp_path / "blank.png"
    Image.new("RGB", (240, 120), (255, 255, 255)).save(blank)
    blank_metrics = _native_image_metrics(str(blank))
    assert _image_is_implausibly_blank(blank_metrics, 3) is True
    assert _image_is_implausibly_blank(blank_metrics, 0) is False

    content = tmp_path / "content.png"
    image = Image.new("RGB", (240, 120), (255, 255, 255))
    ImageDraw.Draw(image).rectangle((20, 20, 80, 50), fill=(0, 0, 0))
    image.save(content)
    content_metrics = _native_image_metrics(str(content))
    assert _image_is_implausibly_blank(content_metrics, 3) is False


@pytest.mark.unit
def test_structured_content_count_uses_visible_cell_values():
    from xlvbatools.workbook.dumper import _structured_visible_content_counts

    counts = _structured_visible_content_counts({
        "sheets": {
            "Input": {
                "cells": {
                    "A1": {"text": "heading", "value": "heading"},
                    "A2": {"text": "", "value": None},
                    "A3": {"text": "0", "value": 0},
                },
                "shapes": [{"name": "Run"}],
            },
        },
    })

    assert counts == {"Input": 2}


@pytest.mark.unit
def test_native_ruler_uses_exact_edges_for_unequal_and_hidden_cells():
    from xlvbatools.workbook._native_capture import _axis_bounds
    white, gray = (255, 255, 255), (224, 224, 224)
    pixels = [white] * 101
    for edge in (33, 100):
        pixels[edge] = gray
    end, boundaries = _axis_bounds(pixels, 27, [10, 0, 20], gray, 1)
    assert end == 29
    assert boundaries == [0, 34, 34, 101]


@pytest.mark.unit
def test_grid_only_does_not_resize_native_pixels():
    from PIL import Image
    from xlvbatools import GridlineOptions
    from xlvbatools.rendering import _paint_gridlines

    image = Image.new("RGB", (100, 50), "white")
    _paint_gridlines(image, [0, 50, 100], [0, 25, 50],
                     GridlineOptions(mode="forced"), "#E0E0E0", 1)
    assert image.size == (100, 50)
    assert image.getpixel((49, 10)) == (224, 224, 224)
    assert image.getpixel((50, 10)) == (255, 255, 255)


@pytest.mark.unit
def test_renderer_rejects_invalid_dpi_before_starting_excel(tmp_path):
    import pytest
    from xlvbatools.workbook.dumper import export_screenshots

    with pytest.raises(ValueError, match="dpi must be positive"):
        export_screenshots(
            "does-not-open.xlsm", ["Input"], str(tmp_path), dpi=0,
        )


@pytest.mark.unit
def test_hidden_worksheets_require_explicit_opt_in():
    from types import SimpleNamespace
    from xlvbatools.workbook.dumper import _worksheet_is_renderable

    assert _worksheet_is_renderable(SimpleNamespace(Visible=-1), False) is True
    assert _worksheet_is_renderable(SimpleNamespace(Visible=0), False) is False
    assert _worksheet_is_renderable(SimpleNamespace(Visible=2), False) is False
    assert _worksheet_is_renderable(SimpleNamespace(Visible=0), True) is True


@pytest.mark.unit
def test_shape_dump_reports_forms_activex_and_text_accessors():
    from xlvbatools.workbook.dumper import dump_sheet_shapes

    class NamedCollection:
        def __init__(self, items):
            self.items = items

        def __call__(self, name=None):
            return self if name is None else self.items[name]

        def Item(self, name):
            return self.items[name]

    ordinary = SimpleNamespace(
        Name="Title", Type=1, OnAction="",
        TextFrame=SimpleNamespace(
            Characters=SimpleNamespace(Text="Ordinary text"),
        ),
    )
    forms = SimpleNamespace(
        Name="RunButton", Type=8, OnAction="RunModel",
        TextFrame=SimpleNamespace(Characters=SimpleNamespace(Text="")),
        TextFrame2=SimpleNamespace(TextRange=SimpleNamespace(Text="")),
    )
    activex = SimpleNamespace(
        Name="ActiveButton", Type=12,
        TextFrame=SimpleNamespace(Characters=SimpleNamespace(Text="")),
        TextFrame2=SimpleNamespace(TextRange=SimpleNamespace(Text="")),
    )
    sheet = SimpleNamespace(
        Name="Input",
        Shapes=[ordinary, forms, activex],
        Buttons=NamedCollection({
            "RunButton": SimpleNamespace(Caption="Run model"),
        }),
        OLEObjects=NamedCollection({
            "ActiveButton": SimpleNamespace(
                Object=SimpleNamespace(Caption="Active action"),
            ),
        }),
    )

    result = {item["name"]: item for item in dump_sheet_shapes(sheet)}

    assert result["Title"]["text"] == "Ordinary text"
    assert result["Title"]["text_accessor"] == "TextFrame.Characters.Text"
    assert result["Title"]["control_type"] == "shape"
    assert result["RunButton"]["text"] == "Run model"
    assert result["RunButton"]["text_accessor"] == "Buttons.Caption"
    assert result["RunButton"]["control_type"] == "forms_control"
    assert result["ActiveButton"]["text"] == "Active action"
    assert result["ActiveButton"]["text_accessor"] == "OLEObjects.Object.Caption"
    assert result["ActiveButton"]["control_type"] == "activex_control"


@pytest.mark.unit
def test_shape_dump_uses_textframe2_when_legacy_textframe_is_empty():
    from xlvbatools.workbook.dumper import dump_sheet_shapes

    shape = SimpleNamespace(
        Name="ModernText", Type=1, OnAction="",
        TextFrame=SimpleNamespace(Characters=SimpleNamespace(Text="")),
        TextFrame2=SimpleNamespace(
            TextRange=SimpleNamespace(Text="TextFrame2 label"),
        ),
    )
    sheet = SimpleNamespace(Name="Input", Shapes=[shape])

    result = dump_sheet_shapes(sheet)

    assert result[0]["text"] == "TextFrame2 label"
    assert result[0]["text_accessor"] == "TextFrame2.TextRange.Text"


@pytest.mark.excel
def test_live_dump_reports_partial_rich_text_runs(minimal_workbook):
    from xlvbatools import Project
    from xlvbatools.core.session import ExcelSession

    sheet = None
    cell = None
    characters = None
    with ExcelSession(minimal_workbook, save_on_exit=True) as session:
        sheet = session.wb.Worksheets("Sheet1")
        cell = sheet.Range("A1")
        cell.Value = "BoldPlain"
        cell.Font.Bold = False
        cell.Font.Italic = False
        characters = cell.GetCharacters(1, 4)
        characters.Font.Bold = True
        characters = None
        characters = cell.GetCharacters(5, 5)
        characters.Font.Italic = True
        characters = None
        cell = None
        sheet = None

    result = Project.open(minimal_workbook).inspect(
        ["Sheet1"], include_data=True, include_screenshots=False,
        include_rich_text=True,
    )

    assert result.success is True, result.to_dict()
    rich_text = result.data.workbook_data["sheets"]["Sheet1"]["cells"]["A1"][
        "rich_text"
    ]
    assert rich_text["status"] == "complete"
    assert [(run["start"], run["length"]) for run in rich_text["runs"]] == [
        (1, 4), (5, 5),
    ]
    assert bool(rich_text["runs"][0]["font"]["bold"]) is True
    assert bool(rich_text["runs"][1]["font"]["italic"]) is True
    assert result.diagnostics.cleanup.still_running is False


@pytest.mark.excel
@pytest.mark.smoke
def test_combined_range_data_and_screenshot_share_one_clean_session(
    minimal_workbook, tmp_path,
):
    """Combined inspection honors its range and leaves no owned Excel process."""
    from xlvbatools.core.session import ExcelSession
    from xlvbatools.workbook.dumper import inspect_workbook

    sheet = None
    cells = None
    with ExcelSession(
        minimal_workbook, save_on_exit=True, kill_on_enter=False,
    ) as session:
        sheet = session.wb.Worksheets("Sheet1")
        cells = sheet.Range("A1:B2")
        cells.Value = (("alpha", 2), ("beta", 4))
        cells = None
        sheet = None
        gc.collect()

    output_dir = tmp_path / "screenshots"
    output_json = tmp_path / "dump.json"
    result = inspect_workbook(
        minimal_workbook,
        ["Sheet1"],
        output_dir=str(output_dir),
        custom_range="A1:B2",
        include_data=True,
        include_screenshots=True,
        output_json=str(output_json),
        timeout_seconds=60,
    )

    assert result["success"] is True, result
    assert result["phase"] == "complete"
    assert result["cleanup"]["pid"] is not None
    assert result["cleanup"]["still_running"] is False
    workbook_data = result["data"]["workbook_data"]
    assert workbook_data["sheets"]["Sheet1"]["bounds"]["address"] == "$A$1:$B$2"
    assert set(workbook_data["sheets"]["Sheet1"]["cells"]) == {
        "A1", "B1", "A2", "B2",
    }
    assert "screenshots" not in result
    screenshot = result["data"]["screenshots"]["Sheet1"]
    assert os.path.isfile(screenshot), result
    capture = result["data"]["screenshot_diagnostics"]["Sheet1"]
    assert capture["capture_backend"] == "excel_window"
    assert capture["attempts"][-1]["window"]["target_range"] == "$A$1:$B$2"
    assert capture["attempts"][-1]["metrics"]["meaningful_pixel_count"] > 64
    assert output_json.is_file()
