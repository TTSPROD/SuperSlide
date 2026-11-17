"""Deterministic native PowerPoint renderer with conservative text-fit checks."""

from __future__ import annotations

import json
import math
import os
import tempfile
from copy import deepcopy
from functools import lru_cache
from pathlib import Path

from PIL import ImageFont
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_TICK_MARK
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Pt

from .assets import AssetCatalog
from .models import ImageRef, PresentationSpec, iter_image_refs
from .template import load_template, select_layout


class LayoutOverflowError(ValueError):
    pass


def _rgb(hex_value: str) -> RGBColor:
    return RGBColor.from_string(hex_value.lstrip("#").upper())


@lru_cache(maxsize=100)
def _font(size_pt: float, bold: bool = False, font_family: str = "Arial"):
    # 144 dpi -> two pixels per point. Only use approved Arial font metrics.
    filename = "arialbd.ttf" if bold else "arial.ttf"
    candidates = [filename, str(Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts" / filename)]
    if font_family.casefold() != "arial":
        candidates = [font_family + (" Bold.ttf" if bold else ".ttf")]
    for name in candidates:
        try:
            return ImageFont.truetype(name, round(size_pt * 2)), True
        except OSError:
            pass
    return ImageFont.load_default(size=round(size_pt * 2)), False


def _wrapped_lines(text: str, font, max_width_px: float) -> int:
    lines = 0
    for hard_line in text.split("\n"):
        if not hard_line:
            lines += 1
            continue
        current = ""
        for word in hard_line.split():
            candidate = (current + " " + word).strip()
            if font.getlength(candidate) <= max_width_px:
                current = candidate
                continue
            if current:
                lines += 1
                current = ""
            if font.getlength(word) <= max_width_px:
                current = word
            else:
                # PowerPoint can break an overlong token; count such breaks conservatively.
                fragment = ""
                for char in word:
                    if fragment and font.getlength(fragment + char) > max_width_px:
                        lines += 1
                        fragment = char
                    else:
                        fragment += char
                current = fragment
        if current:
            lines += 1
    return max(lines, 1)


def _fit_font(texts, field, shape, font_family, bullets=False):
    nominal = float(field.get("font_size_pt", 20))
    minimum = float(field.get("min_font_size_pt", nominal))
    if minimum > nominal or minimum <= 0:
        raise ValueError("Некорректный диапазон кеглей в manifest")
    frame = shape.text_frame
    width_pt = (shape.width - frame.margin_left - frame.margin_right) / 12700
    height_pt = (shape.height - frame.margin_top - frame.margin_bottom) / 12700
    if bullets:
        width_pt -= 17
    width_pt *= 0.96  # Keep a small allowance for Office line-breaking differences.
    paragraph_space = float(field.get("space_after_pt", 7 if bullets else 0))
    spacing = float(field.get("line_spacing", 1.12))
    current = nominal
    while current >= minimum - 0.01:
        font, exact = _font(current, bool(field.get("bold", False)), font_family)
        if exact:
            if any(font.getlength(word) > width_pt * 2 for text in texts for word in text.split()):
                current = round(current - 0.5, 2)
                continue
            counts = [_wrapped_lines(text, font, max(width_pt * 2, 1)) for text in texts]
        else:
            # A deliberately conservative fallback, reported to the caller for visual QA.
            counts = [max(1, math.ceil(len(t) * current * 0.6 / max(width_pt, 1))) for t in texts]
        lines = sum(counts)
        # Office line boxes include ascenders/descenders beyond the nominal point size.
        required_height = lines * current * spacing * 1.2 + max(0, len(texts) - 1) * paragraph_space
        if required_height <= height_pt and lines <= int(field.get("max_lines", 1000)):
            return current, exact, lines
        current = round(current - 0.5, 2)
    raise LayoutOverflowError(
        f"Текст не помещается в поле даже при {minimum:g} pt. "
        "Сократите содержание или разделите его на несколько слайдов."
    )


def _bullet_style(paragraph, enabled):
    ppr = paragraph._p.get_or_add_pPr()
    for tag in ("a:buNone", "a:buChar", "a:buAutoNum", "a:buFont"):
        for child in ppr.findall(tag, ppr.nsmap):
            ppr.remove(child)
    if enabled:
        ppr.set("marL", "215900")
        ppr.set("indent", "-177800")
        font = OxmlElement("a:buFont")
        font.set("typeface", "Arial")
        ppr.append(font)
        bullet = OxmlElement("a:buChar")
        bullet.set("char", "•")
        ppr.append(bullet)
    else:
        ppr.append(OxmlElement("a:buNone"))


def _fill_text(slide, field, value, *, font_family, default_color, report, slide_number, field_name):
    # Font weight follows the semantic field rather than emphasis left in a
    # source template. Use the same weight for fitting and for the saved text.
    base, _, suffix = field_name.rpartition("_")
    heading = (field_name in {"title", "heading", "left_heading", "right_heading"}
               or suffix.isdecimal() and base in {"title", "heading", "label", "step_title"})
    field = {**field, "bold": heading}
    ph = slide.placeholders[field["idx"]]
    # Bboxes in the prepared template already include the intended padding.
    # Do not let Office's implicit default insets reduce the usable field height.
    for side in ("left", "right", "top", "bottom"):
        setattr(ph.text_frame, f"margin_{side}", Pt(float(field.get(f"margin_{side}_pt", 0))))
    texts = value if isinstance(value, list) else [value]
    is_bullets = isinstance(value, list)
    if not any(texts):
        ph.text_frame.clear()
        return
    try:
        size, measured, lines = _fit_font(texts, field, ph, font_family, is_bullets)
    except LayoutOverflowError as exc:
        raise LayoutOverflowError(f"Слайд {slide_number}, поле {field_name}: {exc}") from exc
    frame = ph.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.vertical_anchor = MSO_ANCHOR.TOP
    for index, text in enumerate(texts):
        p = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        p.text = text
        p.alignment = PP_ALIGN.LEFT
        p.level = 0
        p.line_spacing = float(field.get("line_spacing", 1.12))
        p.space_before = Pt(0)
        p.space_after = Pt(float(field.get("space_after_pt", 7 if is_bullets else 0)))
        p.font.bold = heading
        _bullet_style(p, is_bullets)
        for run in p.runs:
            run.font.name = font_family
            run.font.size = Pt(size)
            run.font.bold = bool(field.get("bold", False))
            run.font.color.rgb = _rgb(field.get("color", default_color))
    if size < float(field.get("font_size_pt", 20)):
        report["adjustments"].append({"slide": slide_number, "field": field_name, "font_size_pt": size})
    if not measured:
        warning = f"Метрики шрифта {font_family} недоступны: оценка вместимости приблизительная."
        if warning not in report["warnings"]:
            report["warnings"].append(warning)
    report["text_checks"].append({"slide": slide_number, "field": field_name, "estimated_lines": lines, "font_size_pt": size})


def _remove_placeholder(shape):
    shape._element.getparent().remove(shape._element)


def _picture(slide, field, image: ImageRef, path: Path):
    ph = slide.placeholders[field["idx"]]
    left, top, width, height = ph.left, ph.top, ph.width, ph.height
    if image.fit == "cover":
        # PowerPoint picture placeholders crop proportionally to fill the frame.
        shape = ph.insert_picture(str(path))
        for axis, before, after in (("x", "crop_left", "crop_right"), ("y", "crop_top", "crop_bottom")):
            anchor = field.get(f"crop_anchor_{axis}")
            if anchor is None:
                continue
            if type(anchor) not in (int, float) or not math.isfinite(anchor) or not 0 <= anchor <= 1:
                raise ValueError("Положение кадра crop_anchor должно быть числом от 0 до 1")
            cropped = getattr(shape, before) + getattr(shape, after)
            setattr(shape, before, cropped * anchor)
            setattr(shape, after, cropped * (1 - anchor))
    else:
        from PIL import Image
        with Image.open(path) as source:
            image_width, image_height = source.size
        scale = min(width / image_width, height / image_height)
        fitted_width, fitted_height = round(image_width * scale), round(image_height * scale)
        shape = slide.shapes.add_picture(str(path), left + (width - fitted_width) // 2,
                                         top + (height - fitted_height) // 2,
                                         width=fitted_width, height=fitted_height)
        _remove_placeholder(ph)
    shape.name = f"superslide_image_{image.asset_id}"
    shape._element.nvPicPr.cNvPr.set("descr", image.alt_text)
    # The photo must stay behind slide text, including on a full-photo cover.
    tree = slide.shapes._spTree
    tree.remove(shape._element)
    tree.insert(2, shape._element)
    return shape


def _foreground_art(slide, layout):
    """Keep the original vector logo and supergraphic above a full-slide photo."""
    tree = slide.shapes._spTree
    next_id = max((int(node.get("id")) for node in tree.xpath(".//p:cNvPr")), default=1) + 1
    insert_at = 2 + sum(shape.name.startswith("superslide_image_") for shape in slide.shapes)
    for shape in layout.shapes:
        if shape.is_placeholder:
            continue
        element = deepcopy(shape._element)
        for node in element.iter():
            for attribute, relation_id in list(node.attrib.items()):
                if attribute.startswith("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"):
                    relation = layout.part.rels[relation_id]
                    if relation.is_external:
                        raise ValueError("Графика макета содержит внешнюю ссылку")
                    node.set(attribute, slide.part.relate_to(relation.target_part, relation.reltype))
        for node in element.xpath(".//p:cNvPr"):
            node.set("id", str(next_id))
            next_id += 1
        tree.insert(insert_at, element)
        insert_at += 1


def _slide_content(item):
    content = item.model_dump()
    if item.layout_id == "title_slide" and getattr(item, "image", None) is not None:
        content["image"] = item.image
    if item.layout_id == "bullet_list" and getattr(item, "variant_id", "default") == "two_columns":
        midpoint = (len(item.bullets) + 1) // 2
        content["bullets"] = item.bullets[:midpoint]
        content["bullets_right"] = item.bullets[midpoint:]
    if item.layout_id in ("two_columns", "comparison"):
        content.update({"left_heading": item.left.heading, "left_bullets": item.left.bullets,
                        "right_heading": item.right.heading, "right_bullets": item.right.bullets})
    elif item.layout_id in ("photo_text", "image_spotlight"):
        content["image"] = item.image
    elif item.layout_id == "gallery":
        for i, image in enumerate(item.images, 1):
            content[f"image_{i}"] = image
            content[f"caption_{i}"] = image.caption
    elif item.layout_id == "kpi_cards":
        for i, card in enumerate(item.cards, 1):
            for name, value in card.model_dump().items():
                content[f"{name}_{i}"] = value
    elif item.layout_id == "timeline":
        for i, step in enumerate(item.steps, 1):
            content[f"step_title_{i}"] = step.title
            content[f"step_text_{i}"] = step.text
    return content


def _table(slide, field, spec, font_family, colors, number):
    ph = slide.placeholders[field["idx"]]
    shape = slide.shapes.add_table(len(spec.rows) + 1, len(spec.columns), ph.left, ph.top, ph.width, ph.height)
    _remove_placeholder(ph)
    shape.name = "table"
    table = shape.table
    row_count, col_count = len(spec.rows) + 1, len(spec.columns)
    target_width, target_height = shape.width, shape.height
    widths = [1] * col_count
    # Content proportions are bounded so one verbose column cannot consume the table.
    for c in range(col_count):
        widths[c] = min(3.0, max(1.0, max(len(row[c]) for row in [spec.columns, *spec.rows]) ** 0.5 / 3))
    total_width = sum(widths)
    assigned_width = 0
    for c, weight in enumerate(widths):
        width = target_width - assigned_width if c == col_count - 1 else int(target_width * weight / total_width)
        table.columns[c].width = width
        assigned_width += width
    table.first_row = True
    for r, values in enumerate([spec.columns, *spec.rows]):
        table.rows[r].height = int(target_height / row_count)
        for c, text in enumerate(values):
            cell = table.cell(r, c)
            cell.margin_left = Pt(7)
            cell.margin_right = Pt(7)
            cell.margin_top = Pt(5)
            cell.margin_bottom = Pt(5)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            cell.fill.fore_color.rgb = _rgb(colors["primary"] if r == 0 else ("F4F4F4" if r % 2 else "FFFFFF"))
            cell.text = text
            # Fit each cell rather than allowing silent overflow in an otherwise valid table.
            box = type("CellBox", (), {
                "width": table.columns[c].width, "height": table.rows[r].height,
                "text_frame": cell.text_frame,
            })()
            nominal = float(field.get("font_size_pt", 18))
            minimum = float(field.get("min_font_size_pt", 14))
            style = {"font_size_pt": nominal if r == 0 else max(minimum, nominal - 1), "min_font_size_pt": minimum,
                     "bold": r == 0, "line_spacing": 1.05}
            try:
                size, _, _ = _fit_font([text], style, box, font_family)
            except LayoutOverflowError as exc:
                raise LayoutOverflowError(f"Слайд {number}, таблица, строка {r + 1}, столбец {c + 1}: {exc}") from exc
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.LEFT
                p.space_after = Pt(0)
                p.line_spacing = 1.05
                _bullet_style(p, False)
                for run in p.runs:
                    run.font.name = font_family
                    run.font.size = Pt(size)
                    run.font.bold = r == 0
                    run.font.color.rgb = _rgb("FFFFFF" if r == 0 else colors.get("text", colors["primary"]))
    return shape


def _chart(slide, field, spec, font_family, colors):
    ph = slide.placeholders[field["idx"]]
    data = CategoryChartData()
    data.categories = spec.categories
    for series in spec.series:
        data.add_series(series.name, series.values)
    shape = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, ph.left, ph.top, ph.width, ph.height, data)
    _remove_placeholder(ph)
    shape.name = "chart"
    chart = shape.chart
    chart.has_title = False
    chart.font.name = font_family
    chart.font.size = Pt(12)
    chart.font.color.rgb = _rgb(colors.get("text", colors["primary"]))
    chart.has_legend = len(spec.series) > 1
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.name = font_family
        chart.legend.font.size = Pt(12)
    chart.category_axis.tick_labels.font.name = font_family
    chart.category_axis.tick_labels.font.size = Pt(12)
    chart.category_axis.major_tick_mark = XL_TICK_MARK.NONE
    chart.value_axis.tick_labels.font.name = font_family
    chart.value_axis.tick_labels.font.size = Pt(12)
    chart.value_axis.major_tick_mark = XL_TICK_MARK.NONE
    if all(value >= 0 for series in spec.series for value in series.values):
        chart.value_axis.minimum_scale = 0
    chart.value_axis.has_major_gridlines = True
    chart.value_axis.major_gridlines.format.line.color.rgb = _rgb(colors.get("grid", "DEDEDE"))
    chart.value_axis.tick_labels.number_format = "General"
    chart.value_axis.tick_labels.number_format_is_linked = False
    comparison_palette = ["BCBCBC", "0041A0", "00635B"]
    comparison_index = 0
    for output, series in zip(chart.series, spec.series):
        if series.role == "company":
            color = colors["primary"]
        elif series.role == "neutral":
            color = "909090"
        else:
            color = comparison_palette[comparison_index % len(comparison_palette)]
            comparison_index += 1
        output.format.fill.solid()
        output.format.fill.fore_color.rgb = _rgb(color)
        output.format.line.fill.background()
    chart.plots[0].gap_width = 70
    return shape


def build_presentation(spec: PresentationSpec, template_path, manifest_path, output_path=None, *, source_lookup=None,
                       report_path=None, asset_catalog: AssetCatalog | None = None):
    """Render the deck; output_path=None performs an in-memory dry run without saving a file."""
    if output_path is not None:
        output_path = Path(output_path)
        if output_path.suffix.casefold() != ".pptx":
            raise ValueError("Имя выходного файла должно иметь расширение .pptx")
        if output_path.resolve() == Path(template_path).resolve():
            raise ValueError("Выходной файл совпадает с шаблоном")
        if report_path and Path(report_path).resolve() in (output_path.resolve(), Path(template_path).resolve(), Path(manifest_path).resolve()):
            raise ValueError("Путь отчёта совпадает с презентацией или конфигурацией шаблона")
    prs, manifest, layouts = load_template(template_path, manifest_path)
    if spec.template_version not in manifest.get("compatible_versions", [manifest["template_version"]]):
        raise ValueError("Версия JSON не совпадает с версией шаблона")
    images = list(iter_image_refs(spec))
    if images and asset_catalog is None:
        raise ValueError("Слайды с изображениями требуют каталог изображений (--asset-catalog)")
    resolved_images = {image.asset_id: asset_catalog.resolve(image.asset_id) for image in images}
    protected_assets = ({asset_catalog.resolve(record["id"]) for record in asset_catalog.records}
                        if asset_catalog else set())
    if output_path is not None and (output_path.resolve() in protected_assets
                                    or report_path and Path(report_path).resolve() in protected_assets):
        raise ValueError("Путь результата совпадает с изображением из каталога")
    font_family = manifest.get("font_family", "Arial")
    colors = {"primary": "002F6C", "text": "002F6C", "grid": "DEDEDE", **manifest.get("colors", {})}
    report = {
        "status": "built", "title": spec.title, "template_version": spec.template_version,
        "slide_count": len(spec.slides), "warnings": [], "adjustments": [], "text_checks": [], "images": [], "designs": [],
        "verification": {
            "schema": "passed", "template_contract": "passed",
            "text_fit": "estimated; verify rendered slides in PowerPoint",
            "facts": "source references are traceability, not proof of factual correctness",
            "powerpoint": "not run by build",
        },
    }
    for number, item in enumerate(spec.slides, 1):
        variant_id = getattr(item, "variant_id", "default")
        info = select_layout(manifest, item.layout_id, variant_id)
        slide = prs.slides.add_slide(layouts[info["layout_name"]])
        slide.name = f"{number:02d}_{item.layout_id}_{variant_id}"
        report["designs"].append({"slide": number, "layout_id": item.layout_id, "variant_id": variant_id,
                                  "layout_name": info["layout_name"], "source_reference": info.get("source_reference", {})})
        fields = info["fields"]
        content = _slide_content(item)
        for field_name, field in fields.items():
            if field_name in ("table", "chart"):
                continue
            if field.get("type") == "picture":
                image = content[field_name]
                _picture(slide, field, image, resolved_images[image.asset_id])
                report["images"].append({"slide": number, "field": field_name,
                                          "asset_id": image.asset_id, "fit": image.fit})
                continue
            value = content.get(field_name, "")
            _fill_text(slide, field, value, font_family=font_family, default_color=colors["text"],
                       report=report, slide_number=number, field_name=field_name)
        if item.layout_id == "table":
            _table(slide, fields["table"], item, font_family, colors, number)
        elif item.layout_id == "bar_chart":
            _chart(slide, fields["chart"], item, font_family, colors)
        if info.get("foreground_art"):
            _foreground_art(slide, layouts[info["layout_name"]])
        # Explicit page number: DATE/FOOTER/SLIDE_NUMBER placeholders are not cloned by add_slide.
        if info.get("show_footer", item.layout_id != "title_slide"):
            page = slide.shapes.add_textbox(prs.slide_width - Pt(57), prs.slide_height - Pt(27), Pt(25), Pt(16))
            page.name = "superslide_page_number"
            page.text_frame.margin_left = 0
            page.text_frame.margin_right = 0
            page.text_frame.margin_top = 0
            page.text_frame.margin_bottom = 0
            page.text_frame.text = str(number)
            p = page.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.RIGHT
            p.runs[0].font.name = font_family
            p.runs[0].font.size = Pt(9)
            p.runs[0].font.color.rgb = _rgb("FFFFFF" if info.get("background_color") == "002F6C" else colors.get("muted", "909090"))
        notes = [item.notes] if item.notes else []
        if item.source_refs:
            notes.append("Источники:")
            for ref in item.source_refs:
                source = (source_lookup or {}).get(ref)
                if source:
                    notes.append(f"{ref}: {source['document']}; {source['location']}")
                else:
                    notes.append(ref)
                    report["warnings"].append(f"Слайд {number}: источник {ref} не раскрыт; передайте --sources.")
        elif source_lookup and item.layout_id != "title_slide":
            report["warnings"].append(f"Слайд {number}: отсутствуют source_refs. Проверьте факты вручную.")
        if notes:
            slide.notes_slide.notes_text_frame.text = "\n\n".join(notes)
    prs.core_properties.title = spec.title
    prs.core_properties.subject = "SuperSlide — template-driven corporate presentation"
    prs.core_properties.author = "SuperSlide"
    prs.core_properties.keywords = f"SuperSlide; template {manifest['template_version']}"
    if output_path is not None:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        # Replace only after a complete package has been written. No partial PPTX on validation failure.
        temp_name = None
        try:
            with tempfile.NamedTemporaryFile(dir=output_path.parent, suffix=".pptx", delete=False) as temp:
                temp_name = temp.name
            prs.save(temp_name)
            os.replace(temp_name, output_path)
        finally:
            if temp_name and Path(temp_name).exists():
                Path(temp_name).unlink()
        report["output"] = str(output_path.resolve())
    if report_path:
        report_path = Path(report_path)
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return report


def check_text_fit(spec: PresentationSpec, template_path, manifest_path, *, asset_catalog: AssetCatalog | None = None):
    """In-memory dry run of the renderer: raises LayoutOverflowError for text that cannot fit."""
    return build_presentation(spec, template_path, manifest_path, asset_catalog=asset_catalog)
