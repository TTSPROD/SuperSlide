"""Centered cover crops for photographs in unchanged source-deck frames.

The crop is calculated for the replacement image's display dimensions, including
EXIF orientation, and the frame's physical aspect on the slide. The helpers only
describe placement; they do not change the source XML, mask, or image file.
"""
from __future__ import annotations

import math
from functools import lru_cache
from pathlib import Path

from PIL import Image

from .source_icons import (
    _axis_scales, _element, _geometry, _multiply, _parent_linear,
    _rotation_flip, image_dimensions,
)


def describe_photo_frame(shape, inherited_shape=None):
    """Describe a source frame without resizing or moving it.

    Boxes remain in the shape's parent coordinates. The effective scales include
    all ancestor groups and the shape's own rotation/flip, so the physical frame
    aspect remains correct inside nonuniformly scaled groups. A sheared frame
    cannot be compensated by rectangular cropping and is rejected explicitly.
    """
    node = _element(shape)
    try:
        box, attrs, inherited = _geometry(node, inherited_shape)
        parent = _parent_linear(node)
        combined = _multiply(parent, _rotation_flip(attrs))
    except ValueError as error:
        raise ValueError(str(error).replace("иконки", "фотографии")) from None
    scale_x, scale_y = _axis_scales(combined)
    if min(scale_x, scale_y) <= 0 or not all(math.isfinite(v) for v in (scale_x, scale_y)):
        raise ValueError("Вырожденный масштаб рамки фотографии")
    a, b, c, d = combined
    if abs(a * b + c * d) > scale_x * scale_y * 1e-8:
        raise ValueError("Поворот фотографии внутри растянутой группы создаёт сдвиг осей; требуется отдельная нормализация группы")
    physical_width, physical_height = box[2] * scale_x, box[3] * scale_y
    return {
        "fit": "cover", "preserve_aspect_ratio": True, "alignment": "center",
        "frame": "source", "preserve_source_mask": True,
        "coordinate_space": "parent", "source_box_emu": box,
        "frame_box_emu": list(box), "effective_scale": [scale_x, scale_y],
        "parent_scale": list(_axis_scales(parent)),
        "transform_attributes": attrs, "geometry_inherited": inherited,
        "physical_size_emu": [physical_width, physical_height],
        "physical_aspect_ratio": physical_width / physical_height,
    }


@lru_cache(maxsize=512)
def _image_info_cached(filename, modified_ns, file_size):
    path = Path(filename)
    width, height = image_dimensions(path)
    orientation = 1
    if path.suffix.lower() != ".svg":
        with Image.open(path) as picture:
            orientation = picture.getexif().get(274, 1)
        if orientation not in range(1, 9):
            orientation = 1
        if orientation in (5, 6, 7, 8):
            width, height = height, width
    return width, height, orientation


def photo_crop(shape, image_path, inherited_shape=None):
    """Return a centered OOXML ``srcRect`` crop for the replacement image.

    The crop uses 100000 units per full image dimension. Any existing source crop
    belongs to the old image and is intentionally ignored. Callers should embed
    raster bytes after EXIF transposition when ``exif_orientation`` is not 1;
    ``image_size`` already describes that displayed orientation.
    """
    result = describe_photo_frame(shape, inherited_shape=inherited_shape)
    path = Path(image_path).resolve()
    stat = path.stat()
    width, height, orientation = _image_info_cached(str(path), stat.st_mtime_ns, stat.st_size)
    image_aspect = width / height
    frame_aspect = result["physical_aspect_ratio"]
    crop = {"l": 0, "t": 0, "r": 0, "b": 0}
    if image_aspect > frame_aspect:
        edge = round((1 - frame_aspect / image_aspect) * 50000)
        crop["l"] = crop["r"] = edge
    else:
        edge = round((1 - image_aspect / frame_aspect) * 50000)
        crop["t"] = crop["b"] = edge
    # A zero-width or zero-height srcRect would be invalid. Avoid silently
    # stretching assets whose extreme ratio exceeds OOXML crop precision.
    if edge >= 50000:
        raise ValueError("Пропорции фотографии слишком сильно отличаются от рамки для точного кадрирования OOXML")
    result.update({
        "crop": crop, "image_size": [width, height],
        "image_aspect_ratio": image_aspect, "exif_orientation": orientation,
    })
    return result
