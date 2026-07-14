"""Unit coverage for isolated workbook inspection orchestration."""

import gc
import json
import os
import subprocess
from types import SimpleNamespace

import pytest


def test_inspection_worker_result_is_returned(monkeypatch, tmp_path):
    from xlvbatools.workbook import dumper

    expected = {"success": True, "phase": "complete", "cleanup": {"pid": 42}}

    def fake_run(command, **kwargs):
        result_path = command[-2]
        with open(result_path, "w", encoding="utf-8") as handle:
            json.dump(expected, handle)
        return subprocess.CompletedProcess(command, 0)

    monkeypatch.setattr(dumper.subprocess, "run", fake_run)
    result = dumper.inspect_workbook("book.xlsm", ["Input"], output_dir=str(tmp_path))

    assert result == expected


def test_excel_session_safe_defaults():
    from xlvbatools.core.session import ExcelSession

    session = ExcelSession("book.xlsm")
    assert session.kill_on_enter is False
    assert session.read_only is False
    assert session.disable_macros is False


def test_renderer_does_not_copy_worksheet_source():
    import inspect
    from xlvbatools.workbook import dumper

    source = inspect.getsource(dumper.export_screenshots)
    assert "sheet.Copy(" not in source
    assert "export_range.CopyPicture" in source
    assert "excel.Workbooks.Add()" in source


def test_scaled_boundaries_are_exact_and_support_hidden_cells():
    from xlvbatools.workbook.dumper import _scaled_boundaries

    boundaries = _scaled_boundaries([10, 0, 20], 101)

    assert boundaries == [0, 34, 34, 101]
    assert boundaries[-1] == 101


def test_overlay_geometry_and_atomic_output(tmp_path):
    from PIL import Image
    from xlvbatools.workbook.dumper import _apply_headers_and_grid_overlay

    path = tmp_path / "render.png"
    Image.new("RGB", (101, 61), (12, 34, 56)).save(path)

    _apply_headers_and_grid_overlay(
        str(path), [10, 0, 20], [5, 10], first_column=27, first_row=999,
    )

    with Image.open(path) as rendered:
        # Font metrics vary by host, but the body retains its exact native
        # dimensions and every boundary is relative to the computed header.
        pad_left = rendered.width - 101
        pad_top = rendered.height - 61
        assert pad_left >= 32
        assert pad_top >= 20
        assert rendered.getpixel((pad_left + 1, pad_top + 1)) == (12, 34, 56)
        assert rendered.getpixel((pad_left + 34, pad_top + 10)) == (210, 210, 210)
        assert rendered.getpixel((pad_left + 48, pad_top + 20)) == (210, 210, 210)
    assert not (tmp_path / "render.png.overlay.tmp").exists()


def test_grid_only_does_not_resize_native_image(tmp_path):
    from PIL import Image
    from xlvbatools.workbook.dumper import _apply_headers_and_grid_overlay

    path = tmp_path / "grid.png"
    Image.new("RGB", (100, 50), "white").save(path)
    _apply_headers_and_grid_overlay(
        str(path), [1, 1], [1, 1], 1, 1,
        include_headers=False, include_grid_overlay=True,
    )

    with Image.open(path) as rendered:
        assert rendered.size == (100, 50)
        assert rendered.getpixel((50, 10)) == (210, 210, 210)


def test_renderer_rejects_invalid_dpi_before_starting_excel(tmp_path):
    import pytest
    from xlvbatools.workbook.dumper import export_screenshots

    with pytest.raises(ValueError, match="dpi must be positive"):
        export_screenshots(
            "does-not-open.xlsm", ["Input"], str(tmp_path), dpi=0,
        )


def test_hidden_worksheets_require_explicit_opt_in():
    from types import SimpleNamespace
    from xlvbatools.workbook.dumper import _worksheet_is_renderable

    assert _worksheet_is_renderable(SimpleNamespace(Visible=-1), False) is True
    assert _worksheet_is_renderable(SimpleNamespace(Visible=0), False) is False
    assert _worksheet_is_renderable(SimpleNamespace(Visible=2), False) is False
    assert _worksheet_is_renderable(SimpleNamespace(Visible=0), True) is True


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


@pytest.mark.com
@pytest.mark.e2e
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
    assert result["data"]["sheets"]["Sheet1"]["bounds"]["address"] == "$A$1:$B$2"
    assert set(result["data"]["sheets"]["Sheet1"]["cells"]) == {
        "A1", "B1", "A2", "B2",
    }
    screenshot = result["screenshots"]["Sheet1"]
    assert os.path.isfile(screenshot), result
    assert output_json.is_file()
