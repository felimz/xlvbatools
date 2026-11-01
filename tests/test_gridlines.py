"""Screenshot appearance contracts and pixel-level gridline behavior."""

from dataclasses import FrozenInstanceError
import json
from types import SimpleNamespace
from unittest.mock import Mock, patch

from PIL import Image, ImageDraw
import pytest

from xlvbatools import GridlineOptions, InspectStep, OperationResult, Project
from xlvbatools.rendering import _coerce_gridlines, _paint_gridlines
from xlvbatools.workbook.dumper import _screenshot_gridlines
from xlvbatools.workbook._native_capture import NativePixels
from xlvbatools.workflow import _steps_from_payload, _steps_to_worker


pytestmark = pytest.mark.unit


@pytest.mark.parametrize("mode", ["native", "forced", "off", "custom"])
def test_api_and_workflow_send_the_same_gridline_configuration(tmp_path, mode):
    options = GridlineOptions(mode=mode)
    executor = Mock()
    executor.execute.return_value = OperationResult(
        operation="inspect", success=True, phase="complete", data={},
    )
    project = Project.open(tmp_path / "book.xlsm", executor=executor)
    project.inspect(["Input"], gridlines=mode, include_headers=False)
    request = executor.execute.call_args.args[0]
    assert request.arguments["gridlines"] == options.to_dict()
    assert request.arguments["include_headers"] is False

    step = InspectStep("capture", ("Input",), gridlines=mode, include_headers=False)
    payload = _steps_to_worker([step])
    assert payload[0]["gridlines"] == request.arguments["gridlines"]
    assert _steps_from_payload(json.loads(json.dumps(payload))) == (step,)


def test_custom_options_are_immutable_and_survive_workflow_json():
    options = GridlineOptions(
        mode="custom", color="#12ab34", opacity=0.25, width=3, style="dotted",
    )
    step = InspectStep("capture", ("Input",), gridlines=options)
    encoded = json.loads(json.dumps(_steps_to_worker([step])))
    assert encoded[0]["gridlines"]["color"] == "#12AB34"
    assert _steps_from_payload(encoded) == (step,)
    with pytest.raises(FrozenInstanceError):
        options.opacity = 0.9


@pytest.mark.parametrize("gridlines", [
    "default", None, {"mode": "off", "typo": True},
    {"mode": "custom", "opacity": float("nan")},
    {"mode": "custom", "opacity": 1.1},
    {"mode": "custom", "opacity": True},
    {"mode": "custom", "color": "#1234"},
    {"mode": "custom", "width": 0},
    {"mode": "custom", "width": True},
    {"mode": "custom", "style": "double"},
    {"mode": "native", "color": "#FF0000"},
])
def test_invalid_appearance_is_rejected_before_starting_a_worker(tmp_path, gridlines):
    executor = Mock()
    project = Project.open(tmp_path / "book.xlsm", executor=executor)
    with pytest.raises((TypeError, ValueError)):
        project.inspect(["Input"], gridlines=gridlines)
    executor.execute.assert_not_called()
    with pytest.raises((TypeError, ValueError)):
        _steps_from_payload([{
            "id": "capture", "kind": "inspect", "sheets": ["Input"],
            "gridlines": gridlines,
        }])


@pytest.mark.parametrize("mode", ["native", "off"])
def test_native_and_off_do_not_paint_over_fills_merges_or_borders(tmp_path, mode):
    path = tmp_path / "native.png"
    image = Image.new("RGB", (100, 60), (40, 80, 120))
    ImageDraw.Draw(image).line((0, 45, 99, 45), fill="red", width=3)
    image.save(path)
    before = image.tobytes()
    _paint_gridlines(image, [0, 50, 100], [0, 20, 40, 60],
                     GridlineOptions(mode=mode), "#E0E0E0", 1)
    assert image.tobytes() == before


@pytest.mark.parametrize("opacity,expected", [
    (0, (40, 80, 120)), (0.5, (148, 40, 60)), (1, (255, 0, 0)),
])
def test_custom_opacity_blends_over_the_actual_fill_once(tmp_path, opacity, expected):
    path = tmp_path / "alpha.png"
    Image.new("RGB", (100, 60), (40, 80, 120)).save(path)
    with Image.open(path) as image:
        _paint_gridlines(
            image, [0, 50, 100], [0, 30, 60],
            GridlineOptions(mode="custom", color="#FF0000", opacity=opacity),
            "#E0E0E0", 1,
        )
        assert image.getpixel((49, 10)) == expected
        assert image.getpixel((10, 29)) == expected
        assert image.getpixel((49, 29)) == expected
        assert image.getpixel((10, 10)) == (40, 80, 120)


@pytest.mark.parametrize("style,on,off", [
    ("solid", 4, None), ("dashed", 4, 14), ("dotted", 6, 4),
])
def test_custom_pattern_and_width_have_visible_pixels_and_gaps(tmp_path, style, on, off):
    path = tmp_path / "pattern.png"
    Image.new("RGB", (100, 80), "white").save(path)
    with Image.open(path) as image:
        _paint_gridlines(
            image, [0, 50, 100], [0, 80],
            GridlineOptions(mode="custom", color="#008000", width=2, style=style),
            "#E0E0E0", 1,
        )
        assert image.getpixel((49, on)) == (0, 128, 0)
        assert image.getpixel((50, on)) == (0, 128, 0)
        assert image.getpixel((51, on)) == (255, 255, 255)
        if off is not None:
            assert image.getpixel((49, off)) == (255, 255, 255)


def test_native_heading_pixels_are_copied_without_font_recreation():
    body = Image.new("RGB", (100, 60), (40, 80, 120))
    columns = Image.new("RGB", (100, 29), (11, 22, 33))
    rows = Image.new("RGB", (43, 60), (44, 55, 66))
    corner = Image.new("RGB", (43, 29), (77, 88, 99))
    ImageDraw.Draw(columns).line((13, 4, 60, 22), fill="white", width=2)
    capture = NativePixels(
        body, columns, rows, corner, [0, 50, 100], [0, 30, 60], "#E0E0E0", 1, {},
    )
    try:
        with capture.render(GridlineOptions(), True) as image:
            assert image.size == (143, 89)
            assert image.crop((43, 29, 143, 89)).tobytes() == body.tobytes()
            assert image.crop((43, 0, 143, 29)).tobytes() == columns.tobytes()
            assert image.crop((0, 29, 43, 89)).tobytes() == rows.tobytes()
            assert image.crop((0, 0, 43, 29)).tobytes() == corner.tobytes()
    finally:
        capture.close()


def test_grid_overlay_preserves_existing_border_pixels():
    image = Image.new("RGB", (100, 60), (40, 80, 120))
    ImageDraw.Draw(image).line((0, 29, 99, 29), fill="red")
    _paint_gridlines(
        image, [0, 50, 100], [0, 30, 60], GridlineOptions(mode="forced"),
        "#E0E0E0", 1, (("y", 29, 0, 100),),
    )
    assert image.getpixel((10, 29)) == (255, 0, 0)
    assert image.getpixel((49, 29)) == (255, 0, 0)
    assert image.getpixel((49, 10)) == (224, 224, 224)


def test_native_heading_anchor_excludes_activation_border_but_requires_exact_interior():
    from xlvbatools.workbook._native_capture import _heading_origin, NativeCaptureError

    signature = Image.new("RGB", (8, 29), "black")
    ImageDraw.Draw(signature).polygon(((1, 25), (6, 4), (6, 25)), fill=(80, 80, 80))
    signature.putpixel((7, 0), (50, 50, 50))
    frame = Image.new("RGB", (100, 60), "black")
    frame.paste(signature, (32, 0))
    frame.putpixel((39, 0), (0, 0, 0))
    assert _heading_origin(frame, signature, 70) == (40, 29)
    frame.putpixel((38, 5), (255, 0, 0))
    with pytest.raises(NativeCaptureError):
        _heading_origin(frame, signature, 70)


@pytest.mark.parametrize("mode", ["native", "forced", "off", "custom"])
@pytest.mark.parametrize("fail", [False, True])
def test_capture_restores_the_target_view_even_on_failure(mode, fail):
    window = SimpleNamespace(DisplayGridlines=True)
    other_window = SimpleNamespace(DisplayGridlines=False)
    excel = SimpleNamespace(ActiveWindow=other_window)
    wb = SimpleNamespace(Activate=lambda: None)
    sheet = SimpleNamespace(Activate=lambda: setattr(excel, "ActiveWindow", window))

    def capture():
        with _screenshot_gridlines(excel, wb, sheet, _coerce_gridlines(mode)) as original:
            assert original is True
            assert window.DisplayGridlines is (mode == "native")
            excel.ActiveWindow = other_window  # Temporary geometry workbook took focus.
            if fail:
                raise RuntimeError("capture failed")

    if fail:
        with pytest.raises(RuntimeError, match="capture failed"):
            capture()
    else:
        capture()
    assert window.DisplayGridlines is True
    assert other_window.DisplayGridlines is False


def test_cli_custom_gridlines_forward_typed_options_and_headers(capsys):
    from xlvbatools.cli.main import main

    with patch("xlvbatools.config.loader.load_config"), \
         patch("xlvbatools.logging.setup_logging"), \
         patch("xlvbatools.cli.commands._project") as project:
        project.return_value.inspect.return_value = OperationResult(
            operation="inspect", success=True, phase="complete", data={},
        )
        main([
            "dump", "--sheets", "Input", "--screenshot", "--gridlines", "custom",
            "--gridline-color", "#006699", "--gridline-opacity", "0.3",
            "--gridline-width", "2", "--gridline-style", "dashed", "--no-headers",
        ])
        kwargs = project.return_value.inspect.call_args.kwargs
        assert kwargs["gridlines"] == GridlineOptions(
            mode="custom", color="#006699", opacity=0.3, width=2, style="dashed",
        )
        assert kwargs["include_headers"] is False
    assert json.loads(capsys.readouterr().out)["success"] is True


@pytest.mark.parametrize("options", [
    ["--gridlines", "custom", "--gridline-opacity", "nan"],
    ["--gridlines", "custom", "--gridline-width", "0"],
    ["--gridline-color", "#FFFFFF"],
])
def test_cli_invalid_appearance_is_structured_and_does_not_open_excel(capsys, options):
    from xlvbatools.cli.main import main

    with patch("xlvbatools.cli.commands._project") as project:
        with pytest.raises(SystemExit) as error:
            main(["dump", "--sheets", "Input", "--screenshot", *options])
        assert error.value.code == 2
        project.assert_not_called()
    assert json.loads(capsys.readouterr().out)["error"]["code"] == "invalid_arguments"
