"""Screenshot appearance options shared by the API, CLI, and workers."""

from dataclasses import dataclass
import math
import re
from typing import Any, Literal, Mapping, cast


GridlineMode = Literal["native", "forced", "off", "custom"]
GridlineStyle = Literal["solid", "dashed", "dotted"]
GRIDLINE_MODES = ("native", "forced", "off", "custom")
GRIDLINE_STYLES = ("solid", "dashed", "dotted")


@dataclass(frozen=True)
class GridlineOptions:
    """Control screenshot gridlines without changing cell formatting.

    ``native`` preserves Excel's view; ``forced`` overlays every cell boundary;
    ``off`` hides native gridlines; ``custom`` styles a forced overlay. Custom
    width is measured in output pixels and opacity ranges from 0 to 1. A
    ``None`` color or width uses Excel's measured native automatic grid style.
    """

    mode: GridlineMode = "native"
    color: str | None = None
    opacity: float = 1.0
    width: int | None = None
    style: GridlineStyle = "solid"

    def __post_init__(self) -> None:
        if not isinstance(self.mode, str) or self.mode not in GRIDLINE_MODES:
            raise ValueError(f"gridlines mode must be one of {', '.join(GRIDLINE_MODES)}")
        if self.color is not None and (
            not isinstance(self.color, str) or not re.fullmatch(r"#[0-9a-fA-F]{6}", self.color)
        ):
            raise ValueError("gridline color must be a six-digit hex color such as #D2D2D2")
        if isinstance(self.opacity, bool) or not isinstance(self.opacity, (int, float)):
            raise TypeError("gridline opacity must be a number between 0 and 1")
        if not math.isfinite(self.opacity) or not 0 <= self.opacity <= 1:
            raise ValueError("gridline opacity must be a finite number between 0 and 1")
        if self.width is not None and (
            isinstance(self.width, bool) or not isinstance(self.width, int) or self.width < 1
        ):
            raise ValueError("gridline width must be a positive integer in output pixels")
        if not isinstance(self.style, str) or self.style not in GRIDLINE_STYLES:
            raise ValueError(f"gridline style must be one of {', '.join(GRIDLINE_STYLES)}")
        if self.mode != "custom" and (
            self.color is not None or self.opacity != 1 or
            self.width is not None or self.style != "solid"
        ):
            raise ValueError("custom gridline styling requires mode='custom'")
        if self.color is not None:
            object.__setattr__(self, "color", self.color.upper())
        object.__setattr__(self, "opacity", float(self.opacity))

    def to_dict(self) -> dict[str, Any]:
        """Return the JSON-safe screenshot configuration."""
        return {
            "mode": self.mode, "color": self.color, "opacity": self.opacity,
            "width": self.width, "style": self.style,
        }


def _coerce_gridlines(value: Any) -> GridlineOptions:
    """Validate typed API options or their strict worker/JSON representation."""
    if isinstance(value, GridlineOptions):
        return value
    if isinstance(value, str):
        return GridlineOptions(mode=cast(GridlineMode, value))
    if isinstance(value, Mapping):
        unknown = set(value) - {"mode", "color", "opacity", "width", "style"}
        if unknown:
            raise ValueError("unknown gridline options: " + ", ".join(sorted(map(str, unknown))))
        return GridlineOptions(**dict(value))
    raise TypeError("gridlines must be a mode name or GridlineOptions object")


def _paint_gridlines(
    image, x_bounds, y_bounds, options, native_color, native_width, protected_edges=(),
):
    """Paint at measured native cell edges, compositing intersections once."""
    if options.mode not in ("forced", "custom") or options.opacity == 0:
        return
    from PIL import Image, ImageDraw

    color = options.color or native_color
    width = options.width or native_width
    rgba = tuple(int(color[i:i + 2], 16) for i in (1, 3, 5)) + (round(options.opacity * 255),)
    with Image.new("RGBA", image.size, (0, 0, 0, 0)) as overlay:
        draw = ImageDraw.Draw(overlay)

        def stroke(x1, y1, x2, y2):
            if options.style == "solid":
                draw.line((x1, y1, x2, y2), fill=rgba, width=width)
                return
            dash = width * (6 if options.style == "dashed" else 1)
            gap = width * (4 if options.style == "dashed" else 2)
            horizontal = y1 == y2
            length = x2 - x1 if horizontal else y2 - y1
            for offset in range(0, length + 1, dash + gap):
                end = min(offset + dash - 1, length)
                coords = ((x1 + offset, y1, x1 + end, y2) if horizontal else
                          (x1, y1 + offset, x2, y1 + end))
                draw.line(coords, fill=rgba, width=width)

        for boundary in sorted(set(x_bounds[1:-1])):
            if boundary > 0:
                stroke(boundary - 1, 0, boundary - 1, image.height - 1)
        for boundary in sorted(set(y_bounds[1:-1])):
            if boundary > 0:
                stroke(0, boundary - 1, image.width - 1, boundary - 1)
        for axis, edge, start, end in protected_edges:
            if axis == "x":
                rectangle = (edge - width + 1, start, edge + width - 1, end - 1)
            else:
                rectangle = (start, edge - width + 1, end - 1, edge + width - 1)
            draw.rectangle(rectangle, fill=(0, 0, 0, 0))
        painted = Image.alpha_composite(image.convert("RGBA"), overlay)
        image.paste(painted.convert(image.mode))
        painted.close()
