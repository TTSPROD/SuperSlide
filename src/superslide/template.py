"""Inspect and validate the local PPTX/manifest pair before writing any slides."""

from __future__ import annotations

import json
import math
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER

from .models import SUPPORTED_VARIANTS


REQUIRED_FIELDS = {
    "title_slide": {"title", "subtitle", "author", "date"},
    "bullet_list": {"title", "bullets"},
    "two_columns": {"title", "left_heading", "left_bullets", "right_heading", "right_bullets"},
    "quote": {"quote", "author"},
    "table": {"title", "table", "footnote"},
    "bar_chart": {"title", "chart", "unit", "takeaway"},
}

EXTENDED_FIELDS = {
    "section_divider": {"number", "title", "subtitle"},
    "photo_text": {"title", "image", "heading", "bullets", "caption"},
    "image_spotlight": {"image", "title", "caption"},
    "gallery": {"title", "image_1", "image_2", "image_3", "caption_1", "caption_2", "caption_3"},
    "kpi_cards": {"title", *[f"{name}_{i}" for name in ("value", "label", "detail") for i in range(1, 4)]},
    "timeline": {"title", *[f"step_{name}_{i}" for name in ("title", "text") for i in range(1, 5)]},
    "comparison": {"title", "left_heading", "left_bullets", "right_heading", "right_bullets"},
    "closing": {"title", "subtitle", "contact"},
}

PICTURE_FIELDS = {
    "title_slide": {"image"},
    "photo_text": {"image"}, "image_spotlight": {"image"},
    "gallery": {"image_1", "image_2", "image_3"},
}


class TemplateError(ValueError):
    pass


def iter_layout_designs(manifest: dict):
    """Yield semantic layout, design id and the complete PowerPoint binding."""
    for layout_id, info in manifest["layouts"].items():
        yield layout_id, "default", info
        for variant_id, variant in info.get("variants", {}).items():
            yield layout_id, variant_id, variant


def select_layout(manifest: dict, layout_id: str, variant_id: str = "default") -> dict:
    info = manifest["layouts"].get(layout_id)
    if info is None:
        raise TemplateError(f"Шаблон не содержит тип слайда {layout_id}")
    if variant_id == "default":
        return info
    variant = info.get("variants", {}).get(variant_id)
    if variant is None:
        raise TemplateError(f"Шаблон не содержит вариант {layout_id}/{variant_id}")
    return variant


def load_template(template_path: str | Path, manifest_path: str | Path):
    template_path, manifest_path = Path(template_path), Path(manifest_path)
    with manifest_path.open(encoding="utf-8-sig") as stream:
        manifest = json.load(stream)
    if manifest.get("template_version") not in ("1.0", "1.1", "1.2"):
        raise TemplateError("Не поддерживается версия manifest шаблона")
    if not isinstance(manifest.get("layouts"), dict):
        raise TemplateError("В manifest отсутствует объект layouts")
    prs = Presentation(str(template_path))
    if len(prs.slides):
        raise TemplateError(
            "Генераторный шаблон должен содержать 0 слайдов. "
            "Используйте templates/legacy/corporate_base.pptx, а не исходную библиотеку примеров."
        )
    layouts = {}
    for master in prs.slide_masters:
        for layout in master.slide_layouts:
            if layout.name in layouts:
                raise TemplateError(f"Повторяющееся имя макета: {layout.name}")
            layouts[layout.name] = layout
    issues = []
    required_fields = {**REQUIRED_FIELDS, **(EXTENDED_FIELDS if manifest["template_version"] != "1.0" else {})}
    for layout_id in required_fields:
        if layout_id not in manifest["layouts"]:
            issues.append(f"{layout_id}: отсутствует указанный макет")
    for layout_id, info in manifest["layouts"].items():
        if not isinstance(info, dict) or not isinstance(info.get("variants", {}), dict):
            raise TemplateError(f"{layout_id}: некорректный каталог variants")
    bound_names = set()
    for semantic_id, variant_id, info in iter_layout_designs(manifest):
        required = required_fields.get(semantic_id)
        if required is None:
            issues.append(f"Неизвестный тип слайда: {semantic_id}")
            continue
        layout_id = semantic_id if variant_id == "default" else f"{semantic_id}/{variant_id}"
        if variant_id not in SUPPORTED_VARIANTS[semantic_id]:
            issues.append(f"{layout_id}: неизвестный вариант оформления")
        if variant_id != "default" and manifest["template_version"] != "1.2":
            issues.append(f"{layout_id}: дополнительные варианты требуют manifest 1.2")
        if not isinstance(info, dict):
            issues.append(f"{layout_id}: вариант должен быть объектом")
            continue
        if info.get("layout_name") in bound_names:
            issues.append(f"{layout_id}: один макет назначен нескольким вариантам")
        bound_names.add(info.get("layout_name"))
        layout = layouts.get(info.get("layout_name"))
        if layout is None:
            issues.append(f"{layout_id}: отсутствует указанный макет")
            continue
        fields = info.get("fields", {})
        required = set(required)
        if semantic_id == "title_slide" and variant_id == "photo":
            required.add("image")
        if semantic_id == "bullet_list" and variant_id == "two_columns":
            required.add("bullets_right")
        placeholder_by_idx = {ph.placeholder_format.idx: ph for ph in layout.placeholders}
        if len(placeholder_by_idx) != len(layout.placeholders):
            issues.append(f"{layout_id}: повторяющиеся placeholder.idx")
        if not required.issubset(fields):
            issues.append(f"{layout_id}: отсутствуют поля {sorted(required - set(fields))}")
        used_indices = []
        for name, field in fields.items():
            if not isinstance(field, dict) or type(field.get("idx")) is not int:
                issues.append(f"{layout_id}.{name}: нужен целочисленный idx")
                continue
            idx = field["idx"]
            kind = field.get("type", "text")
            expected = ("picture" if name in PICTURE_FIELDS.get(semantic_id, set())
                        else "content" if name in ("table", "chart") else "text")
            if kind != expected:
                issues.append(f"{layout_id}.{name}: требуется тип поля {expected}")
            for axis in ("x", "y"):
                anchor = field.get(f"crop_anchor_{axis}")
                if anchor is not None and (kind != "picture" or type(anchor) not in (int, float)
                                           or not math.isfinite(anchor) or not 0 <= anchor <= 1):
                    issues.append(f"{layout_id}.{name}: crop_anchor_{axis} должен быть от 0 до 1 для picture")
            used_indices.append(idx)
            ph = placeholder_by_idx.get(idx)
            if ph is None:
                issues.append(f"{layout_id}.{name}: placeholder {idx} не найден")
                continue
            if ph.placeholder_format.type in (PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER):
                issues.append(f"{layout_id}.{name}: служебный placeholder не копируется add_slide; используйте обычное текстовое поле макета")
            if ph.width is None or ph.height is None or ph.width <= 0 or ph.height <= 0:
                issues.append(f"{layout_id}.{name}: некорректный размер placeholder")
            if ph.left < 0 or ph.top < 0 or ph.left + ph.width > prs.slide_width or ph.top + ph.height > prs.slide_height:
                issues.append(f"{layout_id}.{name}: placeholder выходит за границы слайда")
            for key, actual in (("x", ph.left), ("y", ph.top), ("width", ph.width), ("height", ph.height)):
                if key in field and abs(field[key] - actual) > 1:
                    issues.append(f"{layout_id}.{name}: геометрия {key} отличается от manifest")
            if field.get("type", "text") == "text" and not ph.has_text_frame:
                issues.append(f"{layout_id}.{name}: placeholder не содержит текст")
            if field.get("type") == "picture" and ph.placeholder_format.type != PP_PLACEHOLDER.PICTURE:
                issues.append(f"{layout_id}.{name}: нужен picture placeholder")
            if name in ("table", "chart") and field.get("type") != "content":
                issues.append(f"{layout_id}.{name}: нужен content placeholder")
        if len(used_indices) != len(set(used_indices)):
            issues.append(f"{layout_id}: один placeholder назначен нескольким полям")
    if issues:
        raise TemplateError("Некорректный шаблон:\n" + "\n".join(issues))
    return prs, manifest, layouts


def inspect_template(template_path: str | Path, manifest_path: str | Path) -> dict:
    prs, manifest, layouts = load_template(template_path, manifest_path)
    return {
        "template": str(Path(template_path).resolve()),
        "template_version": manifest["template_version"],
        "font_family": manifest.get("font_family", "Arial"),
        "slide_count": len(prs.slides),
        "master_count": len(prs.slide_masters),
        "design_count": sum(1 for _ in iter_layout_designs(manifest)),
        "size_inches": [round(prs.slide_width / 914400, 4), round(prs.slide_height / 914400, 4)],
        "layouts": {
            layout_id: {
                "layout_name": info["layout_name"],
                "fields": {name: {"idx": field["idx"], "type": field.get("type", "text")}
                           for name, field in info["fields"].items()},
                "description": info.get("description", ""),
                "variants": {variant_id: {
                    "layout_name": variant["layout_name"],
                    "description": variant.get("description", ""),
                    "source_reference": variant.get("source_reference", {}),
                    "fields": {name: {"idx": field["idx"], "type": field.get("type", "text")}
                               for name, field in variant["fields"].items()},
                } for variant_id, variant in info.get("variants", {}).items()},
            }
            for layout_id, info in manifest["layouts"].items()
        },
    }
