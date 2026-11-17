"""Square, centered icon slots in the source deck's coordinate system.

The source drawing is left intact. These helpers describe a replacement frame
and its contained picture; the caller owns OOXML relationships and shape IDs.
Parent group scaling is compensated so the frame is square on the actual slide,
not merely square in a group's potentially stretched child coordinates.
"""
from __future__ import annotations

import math
import re
from functools import lru_cache
from pathlib import Path

from lxml import etree as E
from PIL import Image


P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"p": P, "a": A}
_IDENTITY = (1.0, 0.0, 0.0, 1.0)
_LENGTH = re.compile(r"^\s*([+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?)\s*([a-zA-Z]*)\s*$")
_SVG_UNITS = {"": 1.0, "px": 1.0, "in": 96.0, "cm": 96 / 2.54,
              "mm": 96 / 25.4, "q": 96 / 101.6, "pt": 96 / 72,
              "pc": 16.0}


def _element(shape):
    node = getattr(shape, "_element", shape)
    if not hasattr(node, "find"):
        raise ValueError("Ожидается исходная фигура PowerPoint")
    return node


def _transform(node):
    for path in (f"{{{P}}}spPr/{{{A}}}xfrm", f"{{{P}}}grpSpPr/{{{A}}}xfrm", f"{{{P}}}xfrm"):
        transform = node.find(path)
        if transform is not None:
            return transform
    return None


def _geometry(shape, inherited_shape):
    own = _transform(shape)
    inherited = _transform(_element(inherited_shape)) if inherited_shape is not None else None
    values = []
    used_inheritance = False
    for tag, keys in (("off", ("x", "y")), ("ext", ("cx", "cy"))):
        node = own.find(f"{{{A}}}{tag}") if own is not None else None
        if node is None and inherited is not None:
            node = inherited.find(f"{{{A}}}{tag}")
            used_inheritance = True
        if node is None:
            raise ValueError("У места для иконки не определена исходная геометрия")
        try:
            values.extend(int(node.get(key)) for key in keys)
        except (TypeError, ValueError):
            raise ValueError("Некорректная исходная геометрия места для иконки") from None
    if values[2] <= 0 or values[3] <= 0:
        raise ValueError("Место для иконки должно иметь положительные ширину и высоту")
    attrs = {}
    for key in ("rot", "flipH", "flipV"):
        if own is not None and key in own.attrib:
            attrs[key] = own.get(key)
        elif inherited is not None and key in inherited.attrib:
            attrs[key] = inherited.get(key)
    return values, attrs, used_inheritance


def _multiply(left, right):
    # Row-major 2x2 matrices, acting on column vectors.
    a, b, c, d = left
    e, f, g, h = right
    return a * e + b * g, a * f + b * h, c * e + d * g, c * f + d * h


def _rotation_flip(attrs):
    try:
        angle = math.radians(int(attrs.get("rot", 0)) / 60000)
    except (TypeError, ValueError):
        raise ValueError("Некорректный угол поворота места для иконки") from None
    cosine, sine = math.cos(angle), math.sin(angle)
    flip_x = -1 if attrs.get("flipH") in ("1", "true") else 1
    flip_y = -1 if attrs.get("flipV") in ("1", "true") else 1
    return cosine * flip_x, -sine * flip_y, sine * flip_x, cosine * flip_y


def _group_linear(group):
    transform = _transform(group)
    if transform is None:
        return _IDENTITY
    ext = transform.find(f"{{{A}}}ext")
    child_ext = transform.find(f"{{{A}}}chExt")
    if ext is None or child_ext is None:
        raise ValueError("У родительской группы иконки не определён масштаб")
    try:
        width, height = int(ext.get("cx")), int(ext.get("cy"))
        child_width, child_height = int(child_ext.get("cx")), int(child_ext.get("cy"))
        if min(width, height, child_width, child_height) <= 0:
            raise ValueError
        scale_x, scale_y = width / child_width, height / child_height
    except (TypeError, ValueError, ZeroDivisionError):
        raise ValueError("Некорректный масштаб родительской группы иконки") from None
    return _multiply(_rotation_flip(transform.attrib), (scale_x, 0, 0, scale_y))


def _parent_linear(shape):
    result = _IDENTITY
    for ancestor in shape.iterancestors(f"{{{P}}}grpSp"):
        result = _multiply(_group_linear(ancestor), result)
    return result


def _axis_scales(matrix):
    a, b, c, d = matrix
    return math.hypot(a, c), math.hypot(b, d)


def describe_icon_slot(shape, *, inherited_shape=None):
    """Return an inscribed square frame, retaining the source slot's center.

    All returned boxes use the source shape's parent coordinates. Frame width
    and height can differ locally when compensating nonuniform ancestor scales;
    after the parent's transforms the displayed frame is square. Rotation and
    flips are returned for the caller to copy to the new square group.
    """
    node = _element(shape)
    box, attrs, inherited = _geometry(node, inherited_shape)
    parent = _parent_linear(node)
    combined = _multiply(parent, _rotation_flip(attrs))
    scale_x, scale_y = _axis_scales(combined)
    if min(scale_x, scale_y) <= 0 or not all(math.isfinite(v) for v in (scale_x, scale_y)):
        raise ValueError("Вырожденный масштаб места для иконки")
    a, b, c, d = combined
    # Rotation inside a nonuniformly scaled group can produce shear. A box-only
    # transform cannot cancel it; reject future unsupported sources explicitly.
    if abs(a * b + c * d) > scale_x * scale_y * 1e-8:
        raise ValueError("Поворот иконки внутри растянутой группы создаёт сдвиг осей; требуется отдельная нормализация группы")
    x, y, width, height = box
    physical_side = min(width * scale_x, height * scale_y)
    frame_width = max(1, round(physical_side / scale_x))
    frame_height = max(1, round(physical_side / scale_y))
    frame_x = round(x + (width - frame_width) / 2)
    frame_y = round(y + (height - frame_height) / 2)
    return {
        "fit": "contain", "preserve_aspect_ratio": True, "alignment": "center",
        "slot_shape": "square", "square_side_policy": "min_original_extent",
        "coordinate_space": "parent", "source_box_emu": box,
        "slot_box_emu": [frame_x, frame_y, frame_width, frame_height],
        "effective_scale": [scale_x, scale_y], "parent_scale": list(_axis_scales(parent)),
        "transform_attributes": attrs, "geometry_inherited": inherited,
        "display_side_emu": physical_side,
    }


def _svg_length(value):
    match = _LENGTH.fullmatch(value or "")
    if match is None or match[2].lower() not in _SVG_UNITS:
        return None
    number = float(match[1]) * _SVG_UNITS[match[2].lower()]
    return number if math.isfinite(number) and number > 0 else None


@lru_cache(maxsize=512)
def _image_dimensions_cached(filename, modified_ns, file_size):
    path = Path(filename)
    if path.suffix.lower() == ".svg":
        root = E.fromstring(path.read_bytes(), E.XMLParser(resolve_entities=False, no_network=True))
        width, height = _svg_length(root.get("width")), _svg_length(root.get("height"))
        # The intrinsic canvas controls placement. A different viewBox is drawn
        # with SVG preserveAspectRatio inside that canvas, as in its PNG fallback.
        if width and height:
            return width, height
        try:
            viewbox = [float(v) for v in re.split(r"[\s,]+", root.get("viewBox", "").strip())]
            if len(viewbox) != 4 or not all(math.isfinite(v) for v in viewbox) or min(viewbox[2:]) <= 0:
                raise ValueError
        except ValueError:
            raise ValueError("У SVG иконки нет корректного размера холста или viewBox") from None
        if width:
            return width, width * viewbox[3] / viewbox[2]
        if height:
            return height * viewbox[2] / viewbox[3], height
        return viewbox[2], viewbox[3]
    with Image.open(path) as picture:
        width, height = picture.size
    if min(width, height) <= 0:
        raise ValueError("У иконки некорректный размер изображения")
    return float(width), float(height)


def image_dimensions(image_path):
    """Intrinsic canvas size, with original transparent padding preserved."""
    path = Path(image_path).resolve()
    stat = path.stat()
    return _image_dimensions_cached(str(path), stat.st_mtime_ns, stat.st_size)


def fitted_icon_box(shape, image_path, *, inherited_shape=None):
    """Return the picture box in square-group child coordinates (origin 0,0)."""
    result = describe_icon_slot(shape, inherited_shape=inherited_shape)
    asset_width, asset_height = image_dimensions(image_path)
    _, _, frame_width, frame_height = result["slot_box_emu"]
    scale_x, scale_y = result["effective_scale"]
    physical_width, physical_height = frame_width * scale_x, frame_height * scale_y
    uniform = min(physical_width / asset_width, physical_height / asset_height)
    width = min(frame_width, max(1, round(asset_width * uniform / scale_x)))
    height = min(frame_height, max(1, round(asset_height * uniform / scale_y)))
    result.update({
        "content_box_emu": [round((frame_width - width) / 2), round((frame_height - height) / 2), width, height],
        "asset_size": [asset_width, asset_height], "asset_aspect_ratio": asset_width / asset_height,
    })
    return result


def fit_icon_in_slot(shape, image_path, *, inherited_shape=None):
    """Alias for callers that use the placement rather than box terminology."""
    return fitted_icon_box(shape, image_path, inherited_shape=inherited_shape)
