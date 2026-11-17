"""Source-native Severstal templates. Never reconstruct the source artwork.

The JSON describes the editable objects; the PPTX carries the original OOXML,
fonts, masters, layouts, charts, workbooks and pictures. Content updates and the
heading/body weight policy are applied to generated slides; each occurrence
gets its own slide/layout/master chain when an inherited field changes.
"""
from __future__ import annotations

import hashlib
import io
import json
import posixpath
import re
from copy import deepcopy
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Literal
from zipfile import ZipFile, ZIP_DEFLATED

from lxml import etree as E
from PIL import Image, ImageOps
from pydantic import Field, model_validator

from .models import StrictModel, SourceId, Notes, _xml_text
from .paths import PROJECT_ROOT, SOURCE_PHOTO_CATALOG, SOURCE_ICON_CATALOG
from .source_charts import replace_chart
from .source_slots import source_slot_metadata, verify_reviewed_slots
from .source_text import paragraph_roles as _paragraph_roles, source_heading_fields, source_title_metadata

ROOT = PROJECT_ROOT
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = {"p": P, "a": A, "r": R}
SOURCE_TYPES = {
    "title": ("Title", "Титульные слайды"),
    "contents": ("Contents", "Оглавление"),
    "separators": ("Separators", "Разделители"),
    "big_text": ("BigText", "Крупный текст"),
    "bullets": ("Bullits", "Буллиты"),
    "cards": ("Cards", "Карточки"),
    "tables": ("Tables", "Таблицы"),
    "diagram": ("Diagram", "Диаграммы и графики"),
    "schemes": ("Schemes", "Схемы"),
    "circle_schemes": ("CircleSchemes", "Круговые схемы"),
    "team": ("Team", "Команда"),
    "photo": ("Photo", "Фотографии"),
    "quotes": ("Quotes", "Цитаты"),
    "icons": ("Icons", "Иконки"),
    "last_slide": ("LastSlide", "Завершающие слайды"),
}
FOOTER_DATE_PLACEHOLDER = "##.##.####"
SourceLayout = Literal["title", "contents", "separators", "big_text", "bullets", "cards", "tables", "diagram", "schemes", "circle_schemes", "team", "photo", "quotes", "icons", "last_slide"]


class SourceChartSeries(StrictModel):
    name: str = Field(min_length=1, max_length=150)
    values: list[float] = Field(min_length=1, max_length=100)


class SourceChartData(StrictModel):
    categories: list[str] = Field(min_length=1, max_length=100)
    series: list[SourceChartSeries] = Field(min_length=1, max_length=30)


class SourceUpdate(StrictModel):
    field_id: str = Field(pattern=r"^(slide|layout|master)_(text|table|chart|picture|icon)_[0-9]+$")
    text: str | None = Field(default=None, max_length=12000)
    paragraphs: list[str] | None = Field(default=None, max_length=100)
    rows: list[list[str]] | None = Field(default=None, max_length=100)
    photo_id: str | None = None
    icon_id: str | None = None
    chart_data: SourceChartData | None = None

    @model_validator(mode="after")
    def one_operation(self):
        if sum(v is not None for v in (self.text, self.paragraphs, self.rows, self.photo_id, self.icon_id, self.chart_data)) != 1:
            raise ValueError("Укажите ровно одно изменение: text, paragraphs, rows, photo_id, icon_id или chart_data")
        for value in [self.text, *(self.paragraphs or []), *[c for row in self.rows or [] for c in row]]:
            if value is not None:
                # Native DrawingML soft breaks are represented as vertical tabs
                # in the source contract, then written as a:br, never XML text.
                _xml_text(value.replace("\v", ""))
        return self


class SourceSlide(StrictModel):
    layout_id: SourceLayout
    variant_id: str = Field(pattern=r"^slide_[0-9]{3}$")
    updates: list[SourceUpdate] = Field(default_factory=list, max_length=250)
    source_refs: list[SourceId] = Field(default_factory=list, max_length=30)
    notes: Notes = ""

    @model_validator(mode="after")
    def unique_fields(self):
        if len({u.field_id for u in self.updates}) != len(self.updates):
            raise ValueError("Одно поле нельзя менять несколько раз")
        if len(set(self.source_refs)) != len(self.source_refs):
            raise ValueError("Повторяющиеся ссылки на источники")
        return self


class SourcePresentationSpec(StrictModel):
    template_version: Literal["2.0"]
    title: str = Field(min_length=1, max_length=150)
    slides: list[SourceSlide] = Field(min_length=1, max_length=250)


def source_presentation_schema():
    return SourcePresentationSpec.model_json_schema()


def _xml(data):
    return E.fromstring(data, E.XMLParser(resolve_entities=False, no_network=True))


def _bytes(node):
    return E.tostring(node, xml_declaration=True, encoding="UTF-8", standalone=True)


def _relpath(part):
    return posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")


def _target(owner, target):
    return target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(posixpath.dirname(owner), target))


def _relations(package, owner):
    path = _relpath(owner)
    return _xml(package[path]) if path in package else E.Element(f"{{{REL}}}Relationships", nsmap={None: REL})


def _related(package, owner, kind):
    for r in _relations(package, owner):
        if r.get("Type", "").endswith("/" + kind):
            return _target(owner, r.get("Target"))
    return None


def load_source_manifest(path):
    manifest = json.loads(Path(path).read_text(encoding="utf-8-sig"))
    if manifest.get("template_version") != "2.0" or manifest.get("render_mode") != "source_native":
        raise ValueError("Нужен source_native manifest версии 2.0")
    if set(manifest.get("layouts", {})) != set(SOURCE_TYPES):
        raise ValueError("Шаблон должен содержать ровно 15 типов из исходной папки")
    return manifest


class SourceFieldError(ValueError):
    """A source binding failure with safe metadata, without generated content."""

    def __init__(self, slide_index, field_id, reason, message, **details):
        self.validation_details = {"slide_index": slide_index, "reason": reason}
        if field_id is not None:
            self.validation_details["field_id"] = field_id
        self.validation_details.update(details)
        self.validation_issues = [dict(self.validation_details)]
        location = f"Слайд {slide_index}"
        if field_id is not None:
            location += f", поле {field_id}"
        super().__init__(location + ": " + message)


def validate_source_presentation(payload, manifest, known_source_ids=None):
    spec = payload if isinstance(payload, SourcePresentationSpec) else SourcePresentationSpec.model_validate(payload)
    text_errors = []
    for index, slide in enumerate(spec.slides, 1):
        variants = manifest["layouts"][slide.layout_id]["variants"]
        if slide.variant_id not in variants:
            raise SourceFieldError(index, None, "unknown_variant", f"нет исходного варианта {slide.layout_id}/{slide.variant_id}")
        fields = variants[slide.variant_id]["fields"]
        if known_source_ids is not None and set(slide.source_refs) - known_source_ids:
            raise SourceFieldError(index, None, "unknown_source", "неизвестные source_refs")
        for u in slide.updates:
            if u.field_id not in fields:
                raise SourceFieldError(index, u.field_id, "unknown_field", "такого поля нет в выбранном исходном варианте")
            field = fields[u.field_id]
            if field["kind"] == "text":
                if u.text is None and u.paragraphs is None:
                    raise SourceFieldError(index, u.field_id, "text_operation", "ожидается text или paragraphs")
                paragraphs = u.text.split("\n") if u.text is not None else u.paragraphs
                if len(paragraphs) > len(field["paragraphs"]):
                    limit = len(field["paragraphs"])
                    text_errors.append(SourceFieldError(index, u.field_id, "paragraph_count", f"в оригинале {limit} абзацев; сохраните структуру или выберите другой вариант", limit=limit, actual=len(paragraphs), target=limit))
                # Keep the original font sizes and bounding boxes. No silent shrink.
                if sum(map(len, paragraphs)) > field["max_characters"]:
                    limit = field["max_characters"]
                    target = _target_characters(limit)
                    text_errors.append(SourceFieldError(index, u.field_id, "text_length", f"сократите текст до {limit} символов, целевой объём {target}", limit=limit, actual=sum(map(len, paragraphs)), target=target))
            elif field["kind"] == "table":
                if u.rows is None or len(u.rows) != field["row_count"] or any(len(r) != field["column_count"] for r in u.rows):
                    raise SourceFieldError(index, u.field_id, "table_dimensions", f"сохраните исходную сетку {field['row_count']} × {field['column_count']}", limit={"rows": field["row_count"], "columns": field["column_count"]})
                for row_index, (values, original) in enumerate(zip(u.rows, field.get("rows", [])), 1):
                    for column_index, (text, sample) in enumerate(zip(values, original), 1):
                        limit = len(sample.split("\n"))
                        actual = len(text.split("\n"))
                        if actual > limit:
                            raise SourceFieldError(index, u.field_id, "cell_paragraph_count", f"ячейка {row_index}:{column_index} содержит не больше {limit} абзацев", row_index=row_index, column_index=column_index, limit=limit, actual=actual)
            elif field["kind"] in ("picture", "icon"):
                if u.photo_id is None and u.icon_id is None:
                    raise SourceFieldError(index, u.field_id, "image_operation", "ожидается photo_id или icon_id")
            elif field["kind"] == "chart":
                meta = field.get("chart", {})
                if not meta.get("supported"):
                    raise SourceFieldError(index, u.field_id, "chart_unsupported", "эта диаграмма редактируется в PowerPoint; выберите другой исходный вариант для заполнения через JSON")
                if u.chart_data is None or len(u.chart_data.categories) != meta["category_count"] or len(u.chart_data.series) != meta["series_count"] or any(len(s.values) != meta["category_count"] for s in u.chart_data.series):
                    raise SourceFieldError(index, u.field_id, "chart_dimensions", f"сохраните {meta['series_count']} рядов × {meta['category_count']} категорий", limit={"series": meta["series_count"], "categories": meta["category_count"]})
    if text_errors:
        first = text_errors[0]
        first.validation_issues = [dict(error.validation_details) for error in text_errors]
        first.args = ("; ".join(str(error) for error in text_errors),)
        raise first
    return spec


def _prompt_sample(text, limit=80):
    # DrawingML soft line breaks are represented as vertical tabs internally;
    # those control characters must never be copied into a generated a:t value.
    return " ".join(text.replace("\v", " ").split())[:limit]


def _target_characters(maximum):
    """Leave an editorial reserve; the source's hard limit stays unchanged."""
    return max(0, maximum * 85 // 100)


def _text_budget(field, paragraphs):
    maximum = field["max_characters"]
    target = _target_characters(maximum)
    budget = {"max_characters": maximum, "target_characters": target,
              "max_paragraphs": len(paragraphs)}
    if len(paragraphs) > 1:
        # Proportions reflect the differently styled source paragraphs. These
        # remain recommendations, since character widths vary by text/font.
        capacities = [len(text) for text in paragraphs]
        total = max(sum(capacities), maximum, 1)
        budget["paragraph_targets"] = [min(capacity, target * capacity // total)
                                        for capacity in capacities]
        roles = _paragraph_roles(field, paragraphs)
        if roles:
            budget["paragraph_roles"] = roles
    return budget


def _compact_source_field(field):
    result = {"kind": field["kind"]}
    name = re.sub(r"\s+\d+$", "", field.get("name", "")).strip()
    if name and name.casefold() not in {"текст", "textbox", "graphic", "рисунок", "rectangle", "прямоугольник", "text placeholder"}:
        result["name"] = _prompt_sample(name, 50)
    if field["kind"] == "text":
        sample = field.get("sample_text", "")
        paragraphs = field.get("paragraphs", sample.split("\n"))
        result.update({"sample_text": _prompt_sample(sample), "max_characters": field["max_characters"],
                       "paragraph_count": len(paragraphs), "text_budget": _text_budget(field, paragraphs)})
        if len(paragraphs) > 1:
            result["paragraph_characters"] = [len(p) for p in paragraphs]
            result["paragraph_samples"] = [_prompt_sample(p, 24) for p in paragraphs[:3]]
    elif field["kind"] == "table":
        rows = field.get("rows", [])
        result.update({"row_count": field["row_count"], "column_count": field["column_count"],
                       "sample_rows": [[_prompt_sample(cell, 28) for cell in row] for row in rows[:2]],
                       "cell_paragraph_counts": [[len(cell.split("\n")) for cell in row] for row in rows]})
    elif field["kind"] == "chart":
        meta = field["chart"]
        result["chart"] = {key: meta[key] for key in ("supported", "chart_types", "series_count", "category_count", "requires_same_dimensions") if key in meta}
        result["chart"]["categories"] = [_prompt_sample(category, 45) for category in meta.get("categories", [])]
        result["chart"]["series"] = [{"name": _prompt_sample(series["name"], 45), "values": series["values"], "number_format": series.get("number_format", "General")} for series in meta.get("series", [])]
        result["operation"] = "chart_data"
    return result


def source_prompt_catalog(manifest):
    """Expose actual slide content, not inherited prompts or brand artwork.

    The source masters contain example title/body levels and the legal footer.
    Layout text consists of placeholder prompts and slide-number/date fields;
    editable slide instances already carry the visible content. These inherited
    fields remain in the full manifest for explicit editing but are not LLM slots.
    Paragraph character counts describe the original composition as guidance;
    only max_characters and paragraph_count are enforced limits.
    """
    if manifest.get("render_mode") == "source_native" and set(manifest.get("layouts", {})) == set(SOURCE_TYPES):
        verify_reviewed_slots(manifest)
    brand_hashes = {field["media_sha256"] for layout in manifest["layouts"].values()
                    for variant in layout["variants"].values() for field in variant["fields"].values()
                    if field.get("scope") == "master" and field.get("media_sha256")}
    catalog = {}
    for key, layout in manifest["layouts"].items():
        variants = {}
        for vid, variant in layout["variants"].items():
            fields = {}
            for fid, field in variant["fields"].items():
                if field.get("scope", fid.split("_", 1)[0]) != "slide":
                    continue
                if (field.get("automatic") or field.get("sample_text", "").strip() in {FOOTER_DATE_PLACEHOLDER, "‹#›", "<#>"}
                        or field.get("media_sha256") in brand_hashes):
                    continue
                if _prompt_sample(field.get("sample_text", "")).startswith("Любое несанкционированное копирование"):
                    continue
                if field["kind"] == "chart" and not field.get("chart", {}).get("supported"):
                    continue
                fields[fid] = _compact_source_field(field)
            variants[vid] = {"name": _prompt_sample(variant["name"], 120), "fields": fields}
            slots = source_slot_metadata(variant)
            if slots:
                variants[vid]["item_slots"] = slots
        catalog[key] = {"name": layout["name"], "text_budget_policy": {
            "counting_rule": "Characters include spaces and punctuation; paragraph separators do not count.",
            "target_rule": "Write at target_characters (85% of hard maximum); do not aim at the maximum.",
            "paragraphs": "paragraph_targets and paragraph_roles preserve the source composition as guidance, not extra hard limits.",
            "capacity_basis": "Source sample length, not a measured guarantee of visual fit for arbitrary text.",
        }, "variants": variants}
    return catalog


def _shape_id(node):
    found = node.xpath("./p:nvSpPr/p:cNvPr | ./p:nvPicPr/p:cNvPr | ./p:nvGraphicFramePr/p:cNvPr | ./p:nvGrpSpPr/p:cNvPr | ./p:nvCxnSpPr/p:cNvPr", namespaces=NS)
    return int(found[0].get("id")) if found else None


def _find_shape(root, shape_id):
    for node in root.xpath(".//p:sp | .//p:pic | .//p:graphicFrame | .//p:grpSp | .//p:cxnSp", namespaces=NS):
        if _shape_id(node) == shape_id:
            return node
    raise ValueError(f"В исходнике отсутствует shape {shape_id}")


def _paragraph_text(paragraph):
    return "".join("\v" if node.tag == f"{{{A}}}br" else node.text or "" for node in paragraph.iter() if node.tag in (f"{{{A}}}t", f"{{{A}}}br"))


def _write_paragraph(paragraph, text):
    """Keep every run property; distribute new text across the original runs."""
    if _paragraph_text(paragraph) == text:
        return
    runs = paragraph.findall(f"{{{A}}}r")
    if not runs:
        run = E.Element(f"{{{A}}}r")
        end = paragraph.find(f"{{{A}}}endParaRPr")
        if end is not None:
            paragraph.insert(list(paragraph).index(end), run)
            properties = deepcopy(end); properties.tag = f"{{{A}}}rPr"; run.append(properties)
        else:
            paragraph.append(run)
        E.SubElement(run, f"{{{A}}}t")
        runs = [run]
    for br in list(paragraph.findall(f"{{{A}}}br")):
        paragraph.remove(br)
    for fld in list(paragraph.findall(f"{{{A}}}fld")):
        paragraph.remove(fld)
    lengths = [len(r.findtext(f"{{{A}}}t", "")) for r in runs]
    weights = lengths if any(lengths) else [1] * len(runs)
    total = sum(weights)
    offset = 0
    for i, run in enumerate(runs):
        stop = len(text) if i == len(runs) - 1 else offset + round(len(text) * weights[i] / total)
        node = run.find(f"{{{A}}}t")
        if node is None:
            node = E.SubElement(run, f"{{{A}}}t")
        parts = text[offset:stop].split("\v")
        node.text = parts[0]
        previous = run
        for part in parts[1:]:
            # A line break inherits the surrounding run's font properties.
            # Insert both nodes before the next source run/endParaRPr.
            br = E.Element(f"{{{A}}}br")
            properties = run.find(f"{{{A}}}rPr")
            if properties is not None:
                br.append(deepcopy(properties))
            paragraph.insert(list(paragraph).index(previous) + 1, br)
            continuation = deepcopy(run)
            continuation.find(f"{{{A}}}t").text = part
            paragraph.insert(list(paragraph).index(br) + 1, continuation)
            previous = continuation
        offset = stop


def _write_body(body, paragraphs):
    original = body.findall(f"{{{A}}}p")
    for i, paragraph in enumerate(original):
        _write_paragraph(paragraph, paragraphs[i] if i < len(paragraphs) else "")
    # Empty trailing bullet slots still consume height in PowerPoint. Remove
    # only explicitly numbered/bulleted paragraphs that were not supplied;
    # internal empty separators and non-list layout paragraphs stay intact.
    for paragraph in reversed(original):
        if (not _paragraph_text(paragraph).strip() and len(body.findall(f"{{{A}}}p")) > 1
                and paragraph.xpath("./a:pPr/a:buChar | ./a:pPr/a:buAutoNum", namespaces=NS)):
            body.remove(paragraph)
        else:
            break


def _set_paragraph_bold(paragraph, bold):
    """Override run and paragraph inheritance without changing other styling."""
    value = "1" if bold else "0"
    properties = []
    ppr = paragraph.find(f"{{{A}}}pPr")
    if ppr is not None:
        default = ppr.find(f"{{{A}}}defRPr")
        if default is None:
            default = E.Element(f"{{{A}}}defRPr")
            extensions = ppr.find(f"{{{A}}}extLst")
            ppr.insert(list(ppr).index(extensions) if extensions is not None else len(ppr), default)
        properties.append(default)
    for owner in paragraph:
        if owner.tag not in {f"{{{A}}}r", f"{{{A}}}fld", f"{{{A}}}br"}:
            continue
        run_properties = owner.find(f"{{{A}}}rPr")
        if run_properties is None:
            run_properties = E.Element(f"{{{A}}}rPr")
            owner.insert(0, run_properties)
        properties.append(run_properties)
    end = paragraph.find(f"{{{A}}}endParaRPr")
    if end is None:
        end = E.SubElement(paragraph, f"{{{A}}}endParaRPr")
    properties.append(end)
    for node in properties:
        node.set("b", value)


def _normalize_text_weight(root, design, scope, title, heading_fields=(), field_ids=None):
    """Apply heading/body weight to output content using the original roles."""
    changed = False
    for field_id, field in design["fields"].items():
        if field.get("scope", "slide") != scope or field_ids is not None and field_id not in field_ids:
            continue
        sample = field.get("sample_text", "")
        if (field.get("automatic") or sample.strip() in {FOOTER_DATE_PLACEHOLDER, "‹#›", "<#>"}
                or " ".join(sample.split()).startswith("Любое несанкционированное копирование")):
            continue
        if field["kind"] not in {"text", "table"}:
            continue
        shape = _find_shape(root, field["shape_id"])
        if field["kind"] == "table":
            table = shape.find(f".//{{{A}}}tbl")
            props = table.find(f"{{{A}}}tblPr")
            first_row = props is not None and props.get("firstRow") in {"1", "true"}
            first_col = props is not None and props.get("firstCol") in {"1", "true"}
            for row_index, row in enumerate(table.findall(f"{{{A}}}tr")):
                for col_index, cell in enumerate(row.findall(f"{{{A}}}tc")):
                    for paragraph in cell.findall(f"{{{A}}}txBody/{{{A}}}p"):
                        _set_paragraph_bold(paragraph, first_row and row_index == 0 or first_col and col_index == 0)
            changed = True
            continue
        original_text = field.get("paragraphs", sample.split("\n"))
        roles = _paragraph_roles(field, original_text)
        # A separate label needs source semantics or a smaller body beneath it;
        # complete bold paragraphs can also be ordinary descriptions.
        standalone_heading = scope == "slide" and field_id in heading_fields
        main_title = scope == "slide" and field_id == title["title_field_id"]
        native_title = bool(shape.xpath("./p:nvSpPr/p:nvPr/p:ph[@type='title' or @type='ctrTitle']", namespaces=NS))
        for index, paragraph in enumerate(shape.findall(f"{{{P}}}txBody/{{{A}}}p")):
            bold = roles[index] == "heading" if roles and index < len(roles) else main_title or native_title or standalone_heading
            _set_paragraph_bold(paragraph, bold)
        changed = True
    return changed


def _current_date_text():
    """Use the user's Moscow calendar day, including on a host in another zone."""
    return datetime.now(timezone(timedelta(hours=3))).strftime("%d.%m.%Y")


def _is_footer_date(shape):
    return "".join(shape.xpath("./p:txBody//a:t/text()", namespaces=NS)).strip() == FOOTER_DATE_PLACEHOLDER


def _fill_footer_date(slide_root, layout_root, date_text):
    """Fill the original footer slot, whose empty slide body masks its prompt."""
    changed = False
    tree = slide_root.find(f"{{{P}}}cSld/{{{P}}}spTree")
    for source_shape in layout_root.xpath(".//p:sp", namespaces=NS):
        if not _is_footer_date(source_shape):
            continue
        source_ph = source_shape.find(f"{{{P}}}nvSpPr/{{{P}}}nvPr/{{{P}}}ph")
        if source_ph is None:
            continue
        idx = source_ph.get("idx", "0")
        target = next((shape for shape in slide_root.xpath(".//p:sp", namespaces=NS)
                       if (ph := shape.find(f"{{{P}}}nvSpPr/{{{P}}}nvPr/{{{P}}}ph")) is not None
                       and ph.get("idx", "0") == idx), None)
        if target is None:
            target = deepcopy(source_shape)
            properties = target.find(f"{{{P}}}nvSpPr/{{{P}}}cNvPr")
            next_id = 1 + max((int(n.get("id")) for n in slide_root.xpath(".//p:cNvPr", namespaces=NS)), default=0)
            properties.set("id", str(next_id))
            target.find(f"{{{P}}}nvSpPr/{{{P}}}nvPr/{{{P}}}ph").attrib.pop("hasCustomPrompt", None)
            tree.append(target)
        body = target.find(f"{{{P}}}txBody")
        if body is None:
            body = deepcopy(source_shape.find(f"{{{P}}}txBody"))
            target.append(body)
        _write_body(body, [date_text])
        changed = True
    # Handle a literal footer already present directly on a source slide.
    for shape in slide_root.xpath(".//p:sp", namespaces=NS):
        if _is_footer_date(shape):
            _write_body(shape.find(f"{{{P}}}txBody"), [date_text])
            changed = True
    return changed


def _fill_package_date_prompts(package, date_text):
    """Remove date prompts in the output's layout/master edit views as well."""
    for part in list(package):
        if part.endswith(".xml") and ("/slideLayouts/" in part or "/slideMasters/" in part or part.startswith(("ppt/generated/layout", "ppt/generated/master"))):
            root = _xml(package[part])
            changed = False
            for shape in root.xpath(".//p:sp", namespaces=NS):
                if _is_footer_date(shape):
                    _write_body(shape.find(f"{{{P}}}txBody"), [date_text])
                    changed = True
            if changed:
                package[part] = _bytes(root)


def _override(content_types, part, content_type):
    name = "/" + part
    for item in content_types:
        if item.get("PartName") == name:
            return
    E.SubElement(content_types, f"{{{CT}}}Override", PartName=name, ContentType=content_type)


def _append_rel(rels, kind, target):
    existing = {r.get("Id") for r in rels}
    index = 1
    while f"rId{index}" in existing:
        index += 1
    rid = f"rId{index}"
    E.SubElement(rels, f"{{{REL}}}Relationship", Id=rid, Type=R + "/" + kind, Target="/" + target)
    return rid


def _clone_part(package, content_types, old, new, replacements):
    package[new] = package[old]
    rels = _relations(package, old)
    for relation in rels:
        if relation.get("TargetMode") != "External":
            target = _target(old, relation.get("Target"))
            relation.set("Target", "/" + replacements.get(target, target))
    package[_relpath(new)] = _bytes(rels)
    source_type = next((r.get("ContentType") for r in content_types if r.get("PartName") == "/" + old), None)
    if source_type:
        _override(content_types, new, source_type)


def _asset(catalog_path, asset_id):
    path = Path(catalog_path)
    catalog = json.loads(path.read_text(encoding="utf-8-sig"))
    records = catalog if isinstance(catalog, list) else next((catalog[k] for k in ("assets", "photos", "icons", "images") if isinstance(catalog.get(k), list)), [])
    if not records and isinstance(catalog, dict):
        records = [icon for group in catalog.get("groups", []) for icon in group.get("icons", [])]
    record = next((r for r in records if r.get("id", r.get("asset_id")) == asset_id), None)
    if record is None:
        raise ValueError(f"Нет ресурса {asset_id} в {path}")
    filename = record.get("path", record.get("file", record.get("svg")))
    if not isinstance(filename, str):
        raise ValueError(f"У ресурса {asset_id} нет пути к файлу")
    candidates = [path.parent / filename, ROOT / filename]
    resolved = next((p.resolve() for p in candidates if p.is_file()), None)
    if resolved is None:
        raise ValueError(f"Нет файла ресурса {asset_id}")
    return resolved


def _replace_image(package, ct, owner, shape, image_path, *, icon_mode=False, inherited_shape=None):
    """Photos cover the source frame; icons contain within a square frame."""
    placement = None
    if not icon_mode:
        from .source_photos import photo_crop
        placement = photo_crop(shape, image_path, inherited_shape=inherited_shape)
    data = image_path.read_bytes()
    suffix = image_path.suffix.lower().lstrip(".")
    if placement and placement["exif_orientation"] != 1:
        # PowerPoint's handling of EXIF rotation varies by image format. Embed
        # the display-oriented pixels so that the crop and canvas agree.
        with Image.open(io.BytesIO(data)) as original:
            oriented = ImageOps.exif_transpose(original)
            if oriented.mode not in ("1", "L", "LA", "P", "RGB", "RGBA", "I", "I;16"):
                oriented = oriented.convert("RGBA" if "A" in oriented.getbands() else "RGB")
            stream = io.BytesIO()
            oriented.save(stream, format="PNG")
            data = stream.getvalue()
        suffix = "png"
    media = "ppt/generated/media/" + hashlib.sha256(data).hexdigest() + "." + suffix
    package[media] = data
    _override(ct, media, {"svg": "image/svg+xml", "png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg"}[suffix])
    rels = _relations(package, owner)
    rid = _append_rel(rels, "image", media)
    svg_rid = None
    if suffix == "svg":
        png = image_path.with_suffix(".png")
        if not png.exists():
            raise ValueError("Для вставки SVG в PowerPoint нужен PNG fallback из каталога иконок")
        png_data = png.read_bytes()
        png_media = "ppt/generated/media/" + hashlib.sha256(png_data).hexdigest() + ".png"
        package[png_media] = png_data
        _override(ct, png_media, "image/png")
        svg_rid = rid
        rid = _append_rel(rels, "image", png_media)
    if icon_mode:
        from .source_icons import fitted_icon_box
        placement = fitted_icon_box(shape, image_path, inherited_shape=inherited_shape)
        x, y, width, height = placement["slot_box_emu"]
        px, py, picture_width, picture_height = placement["content_box_emu"]
        source_properties = shape.xpath("./*/p:cNvPr", namespaces=NS)[0]
        next_id = 1 + max(int(node.get("id")) for node in shape.getroottree().getroot().xpath(".//p:cNvPr", namespaces=NS))
        group = E.Element(f"{{{P}}}grpSp")
        nv = E.SubElement(group, f"{{{P}}}nvGrpSpPr")
        properties = deepcopy(source_properties)
        properties.set("name", "Icon placeholder: " + source_properties.get("name", ""))
        properties.set("descr", "Квадратная область; иконка вписана с сохранением пропорций")
        nv.append(properties)
        locks = E.SubElement(nv, f"{{{P}}}cNvGrpSpPr")
        E.SubElement(locks, f"{{{A}}}grpSpLocks", noChangeAspect="1")
        E.SubElement(nv, f"{{{P}}}nvPr")
        group_properties = E.SubElement(group, f"{{{P}}}grpSpPr")
        transform = E.SubElement(group_properties, f"{{{A}}}xfrm", **placement.get("transform_attributes", {}))
        E.SubElement(transform, f"{{{A}}}off", x=str(x), y=str(y))
        E.SubElement(transform, f"{{{A}}}ext", cx=str(width), cy=str(height))
        E.SubElement(transform, f"{{{A}}}chOff", x="0", y="0")
        E.SubElement(transform, f"{{{A}}}chExt", cx=str(width), cy=str(height))
        picture = E.SubElement(group, f"{{{P}}}pic")
        picture_nv = E.SubElement(picture, f"{{{P}}}nvPicPr")
        E.SubElement(picture_nv, f"{{{P}}}cNvPr", id=str(next_id), name="Icon", descr=image_path.stem)
        picture_locks = E.SubElement(picture_nv, f"{{{P}}}cNvPicPr")
        E.SubElement(picture_locks, f"{{{A}}}picLocks", noChangeAspect="1")
        E.SubElement(picture_nv, f"{{{P}}}nvPr")
        fill = E.SubElement(picture, f"{{{P}}}blipFill")
        blip = E.SubElement(fill, f"{{{A}}}blip", **{f"{{{R}}}embed": rid})
        if svg_rid:
            ext_list = E.SubElement(blip, f"{{{A}}}extLst")
            ext = E.SubElement(ext_list, f"{{{A}}}ext", uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}")
            E.SubElement(ext, "{http://schemas.microsoft.com/office/drawing/2016/SVG/main}svgBlip", **{f"{{{R}}}embed": svg_rid})
        stretch = E.SubElement(fill, f"{{{A}}}stretch")
        E.SubElement(stretch, f"{{{A}}}fillRect")
        picture_properties = E.SubElement(picture, f"{{{P}}}spPr")
        picture_transform = E.SubElement(picture_properties, f"{{{A}}}xfrm")
        E.SubElement(picture_transform, f"{{{A}}}off", x=str(px), y=str(py))
        E.SubElement(picture_transform, f"{{{A}}}ext", cx=str(picture_width), cy=str(picture_height))
        geometry = E.SubElement(picture_properties, f"{{{A}}}prstGeom", prst="rect")
        E.SubElement(geometry, f"{{{A}}}avLst")
        shape.getparent().replace(shape, group)
        package[_relpath(owner)] = _bytes(rels)
        return placement
    if shape.tag == f"{{{P}}}pic":
        fill = shape.find(f"{{{P}}}blipFill")
        if fill is None:
            fill = E.Element(f"{{{P}}}blipFill")
            shape.insert(1, fill)
        lock_owner = shape.find(f"{{{P}}}nvPicPr/{{{P}}}cNvPicPr")
        lock_tag = f"{{{A}}}picLocks"
    elif shape.tag == f"{{{P}}}sp":
        properties = shape.find(f"{{{P}}}spPr")
        fill = properties.find(f"{{{A}}}blipFill")
        if fill is None:
            # Fill follows geometry and precedes lines/effects in DrawingML.
            fill = E.Element(f"{{{A}}}blipFill")
            for old_fill in properties.xpath("./a:noFill | ./a:solidFill | ./a:gradFill | ./a:pattFill | ./a:grpFill", namespaces=NS):
                properties.remove(old_fill)
            position = next((i for i, node in enumerate(properties)
                             if E.QName(node).localname not in ("xfrm", "prstGeom", "custGeom")), len(properties))
            properties.insert(position, fill)
        lock_owner = shape.find(f"{{{P}}}nvSpPr/{{{P}}}cNvSpPr")
        lock_tag = f"{{{A}}}spLocks"
    else:
        raise ValueError("Фотография должна быть изображением или заливкой фигуры")
    blip = fill.find(f"{{{A}}}blip")
    if blip is None:
        blip = E.Element(f"{{{A}}}blip")
        fill.insert(0, blip)
    blip.set(f"{{{R}}}embed", rid)
    blip.attrib.pop(f"{{{R}}}link", None)
    for ext in blip.findall(f"{{{A}}}extLst"):
        blip.remove(ext)
    if svg_rid:
        extlst = E.SubElement(blip, f"{{{A}}}extLst")
        ext = E.SubElement(extlst, f"{{{A}}}ext", uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}")
        E.SubElement(ext, "{http://schemas.microsoft.com/office/drawing/2016/SVG/main}svgBlip", **{f"{{{R}}}embed": svg_rid})
    for previous in fill.xpath("./a:srcRect | ./a:stretch | ./a:tile", namespaces=NS):
        fill.remove(previous)
    crop = E.Element(f"{{{A}}}srcRect", **{key: str(value) for key, value in placement["crop"].items()})
    fill.insert(list(fill).index(blip) + 1, crop)
    stretch = E.Element(f"{{{A}}}stretch")
    E.SubElement(stretch, f"{{{A}}}fillRect")
    fill.insert(list(fill).index(crop) + 1, stretch)
    if lock_owner is not None:
        locks = lock_owner.find(lock_tag)
        if locks is None:
            locks = E.Element(lock_tag)
            lock_owner.insert(0, locks)
        locks.set("noChangeAspect", "1")
    package[_relpath(owner)] = _bytes(rels)
    return placement


def build_source_presentation(spec, template_path, manifest_path, output_path, *, asset_catalog_path=None, icon_catalog_path=None, source_lookup=None):
    manifest = load_source_manifest(manifest_path)
    spec = validate_source_presentation(spec, manifest, set(source_lookup) if source_lookup is not None else None)
    with ZipFile(template_path) as archive:
        package = {n: archive.read(n) for n in archive.namelist()}
    presentation = _xml(package["ppt/presentation.xml"])
    rels = _relations(package, "ppt/presentation.xml")
    ct = _xml(package["[Content_Types].xml"])
    slide_list = presentation.find(f"{{{P}}}sldIdLst")
    slide_list.clear()
    for relation in list(rels):
        if relation.get("Type").endswith(("/slide", "/slideMaster")):
            rels.remove(relation)
    masters = presentation.find(f"{{{P}}}sldMasterIdLst")
    masters.clear()
    report = []
    registered_masters = set()
    date_text = _current_date_text()
    for index, slide in enumerate(spec.slides, 1):
        design = manifest["layouts"][slide.layout_id]["variants"][slide.variant_id]
        old_slide = design["package_part"]
        old_layout = _related(package, old_slide, "slideLayout")
        old_master = _related(package, old_layout, "slideMaster")
        updated_scopes = {design["fields"][u.field_id]["scope"] for u in slide.updates}
        inherited_edit = bool(updated_scopes & {"layout", "master"})
        chain = {old_slide: f"ppt/generated/slide{index}.xml"}
        if inherited_edit:
            chain.update({old_layout: f"ppt/generated/layout{index}.xml", old_master: f"ppt/generated/master{index}.xml"})
        for old, new in chain.items():
            _clone_part(package, ct, old, new, chain)
        # Unchanged inheritance is shared, as in the source deck. Clone it only
        # for an explicit inherited-field edit, isolating that occurrence.
        new_master = chain.get(old_master, old_master)
        if inherited_edit:
            master_root = _xml(package[new_master])
            master_rels = _relations(package, new_master)
            layout_rids = {r.get("Id") for r in master_rels if r.get("Type").endswith("/slideLayout") and _target(new_master, r.get("Target")) == chain[old_layout]}
            layout_list = master_root.find(f"{{{P}}}sldLayoutIdLst")
            if layout_list is not None:
                for item in list(layout_list):
                    if item.get(f"{{{R}}}id") not in layout_rids:
                        layout_list.remove(item)
            for relation in list(master_rels):
                if relation.get("Type").endswith("/slideLayout") and relation.get("Id") not in layout_rids:
                    master_rels.remove(relation)
            package[_relpath(new_master)] = _bytes(master_rels)
            package[new_master] = _bytes(master_root)
        if new_master not in registered_masters:
            master_rid = _append_rel(rels, "slideMaster", new_master)
            E.SubElement(masters, f"{{{P}}}sldMasterId", id=str(2147483648 + 1000 + index), **{f"{{{R}}}id": master_rid})
            registered_masters.add(new_master)
        owner_parts = {"slide": chain[old_slide], "layout": chain.get(old_layout, old_layout), "master": new_master}
        roots = {scope: _xml(package[part]) for scope, part in owner_parts.items()}
        title = source_title_metadata(design, roots["slide"])
        heading_fields = source_heading_fields(design, roots["slide"])
        for update in slide.updates:
            field = design["fields"][update.field_id]
            root = roots[field["scope"]]
            shape = _find_shape(root, field["shape_id"])
            if field["kind"] == "text":
                _write_body(shape.find(f"{{{P}}}txBody"), update.text.split("\n") if update.text is not None else update.paragraphs)
            elif field["kind"] == "table":
                for row, values in zip(shape.findall(f".//{{{A}}}tr"), update.rows):
                    for cell, value in zip(row.findall(f"{{{A}}}tc"), values):
                        body = cell.find(f"{{{A}}}txBody")
                        parts = value.split("\n")
                        if len(parts) > len(body.findall(f"{{{A}}}p")):
                            raise ValueError("Сохраните количество абзацев в ячейке исходной таблицы")
                        _write_body(body, parts)
            elif field["kind"] == "chart":
                chart_node = shape.xpath(".//*[local-name()='chart' and namespace-uri()='http://schemas.openxmlformats.org/drawingml/2006/chart']")[0]
                owner = owner_parts[field["scope"]]
                chart_rid = chart_node.get(f"{{{R}}}id")
                owner_rels = _relations(package, owner)
                chart_relation = next(r for r in owner_rels if r.get("Id") == chart_rid)
                old_chart = _target(owner, chart_relation.get("Target"))
                new_chart = f"ppt/generated/chart{index}_{field['shape_id']}.xml"
                _clone_part(package, ct, old_chart, new_chart, {})
                new_xml, workbook = replace_chart(package[old_chart], update.chart_data.model_dump())
                package[new_chart] = new_xml
                workbook_part = f"ppt/generated/workbook{index}_{field['shape_id']}.xlsx"
                package[workbook_part] = workbook
                _override(ct, workbook_part, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                chart_rels = _relations(package, new_chart)
                workbook_relation = next((r for r in chart_rels if r.get("Type").endswith("/package")), None)
                if workbook_relation is None:
                    workbook_rid = _append_rel(chart_rels, "package", workbook_part)
                    croot = _xml(package[new_chart])
                    external = E.SubElement(croot, "{http://schemas.openxmlformats.org/drawingml/2006/chart}externalData", **{f"{{{R}}}id": workbook_rid})
                    package[new_chart] = _bytes(croot)
                else:
                    workbook_relation.set("Target", "/" + workbook_part)
                package[_relpath(new_chart)] = _bytes(chart_rels)
                chart_relation.set("Target", "/" + new_chart)
                package[_relpath(owner)] = _bytes(owner_rels)
            else:
                asset = _asset(asset_catalog_path or SOURCE_PHOTO_CATALOG, update.photo_id) if update.photo_id else _asset(icon_catalog_path or SOURCE_ICON_CATALOG, update.icon_id)
                icon_mode = bool(update.icon_id) or field["kind"] == "icon"
                inherited_shape = None
                if field["scope"] != "master":
                    ph = shape.xpath("./*/p:nvPr/p:ph", namespaces=NS)
                    if ph:
                        idx, kind = ph[0].get("idx", "0"), ph[0].get("type", "obj")
                        scopes = ("layout", "master") if field["scope"] == "slide" else ("master",)
                        for scope in scopes:
                            matches = roots[scope].xpath(".//p:sp | .//p:pic", namespaces=NS)
                            for candidate in matches:
                                parent_ph = candidate.xpath("./*/p:nvPr/p:ph", namespaces=NS)
                                transforms = candidate.xpath("./p:spPr/a:xfrm[a:off and a:ext]", namespaces=NS)
                                if parent_ph and transforms and (parent_ph[0].get("idx", "0") == idx if scope == "layout" else parent_ph[0].get("type", "obj") == kind):
                                    inherited_shape = candidate
                                    break
                            if inherited_shape is not None:
                                break
                _replace_image(package, ct, owner_parts[field["scope"]], shape, asset,
                               icon_mode=icon_mode, inherited_shape=inherited_shape)
        for scope in roots:
            edited_fields = None if scope == "slide" else {u.field_id for u in slide.updates}
            if _normalize_text_weight(roots[scope], design, scope, title, heading_fields, edited_fields):
                updated_scopes.add(scope)
        if _fill_footer_date(roots["slide"], _xml(package[old_layout]), date_text):
            updated_scopes.add("slide")
        for scope, root in roots.items():
            # Source artwork is retained; generated content gets explicit
            # heading/body weight in addition to requested text and date edits.
            if scope in updated_scopes:
                package[owner_parts[scope]] = _bytes(root)
        # Notes have a back-reference to their owner. Clone them for every
        # occurrence, including unchanged slides, to keep dependencies isolated.
        old_notes = _related(package, chain[old_slide], "notesSlide")
        notes_source = old_notes
        if notes_source is None and (slide.notes or slide.source_refs):
            notes_source = next((part for part in package if "/notesSlides/" in part and part.endswith(".xml")), None)
        if notes_source:
            notes_part = f"ppt/generated/notes{index}.xml"
            previous_owner = _related(package, notes_source, "slide")
            _clone_part(package, ct, notes_source, notes_part, {previous_owner: chain[old_slide]})
            notes = _xml(package[notes_part])
            bodies = notes.xpath(".//p:sp[p:nvSpPr/p:nvPr/p:ph[@type='body']]/p:txBody", namespaces=NS)
            if bodies and not old_notes:
                for paragraph in bodies[0].findall(f"{{{A}}}p"):
                    bodies[0].remove(paragraph)
            if bodies and (slide.notes or slide.source_refs):
                for text in [slide.notes, *(["Источники: " + ", ".join(slide.source_refs)] if slide.source_refs else [])]:
                    para = E.SubElement(bodies[0], f"{{{A}}}p")
                    run = E.SubElement(para, f"{{{A}}}r")
                    E.SubElement(run, f"{{{A}}}t").text = text
            package[notes_part] = _bytes(notes)
            sr = _relations(package, chain[old_slide])
            relation = next((r for r in sr if r.get("Type").endswith("/notesSlide")), None)
            if relation is None:
                _append_rel(sr, "notesSlide", notes_part)
            else:
                relation.set("Target", "/" + notes_part)
            package[_relpath(chain[old_slide])] = _bytes(sr)
        rid = _append_rel(rels, "slide", chain[old_slide])
        E.SubElement(slide_list, f"{{{P}}}sldId", id=str(255 + index), **{f"{{{R}}}id": rid})
        report.append({"slide": index, "layout_id": slide.layout_id, "variant_id": slide.variant_id, "source_file": manifest["layouts"][slide.layout_id]["source_file"], "source_slide": design["source_slide"], "updates": len(slide.updates)})
    package["ppt/presentation.xml"] = _bytes(presentation)
    package[_relpath("ppt/presentation.xml")] = _bytes(rels)
    package["[Content_Types].xml"] = _bytes(ct)
    _set_presentation_title(package, spec.title)
    _fill_package_date_prompts(package, date_text)
    _prune_and_save(package, output_path)
    return {"template_version": "2.0", "render_mode": "source_native", "text_weight_policy": "headings_bold_body_regular", "slide_count": len(report), "footer_date": date_text, "slides": report, "output": str(Path(output_path).resolve())}


def _set_presentation_title(package, title):
    """Name the generated deck without changing other source core properties."""
    core_path = "docProps/core.xml"
    if core_path not in package:
        return
    core = _xml(package[core_path])
    title_tag = "{http://purl.org/dc/elements/1.1/}title"
    node = core.find(title_tag)
    if node is None:
        node = E.SubElement(core, title_tag)
    node.text = title
    package[core_path] = _bytes(core)


def _prune_and_save(package, output_path):
    """Keep only reachable OPC parts, preserving their bytes and native data."""
    _normalize_package_metadata(package)
    keep = {"[Content_Types].xml", "_rels/.rels"}
    pending = ["/"]
    while pending:
        owner = pending.pop()
        relpath = "_rels/.rels" if owner == "/" else _relpath(owner)
        if relpath not in package:
            continue
        keep.add(relpath)
        for relation in _xml(package[relpath]):
            if relation.get("TargetMode") == "External":
                continue
            part = _target("" if owner == "/" else owner, relation.get("Target"))
            if part not in package:
                raise ValueError(f"Повреждённая связь {owner} -> {part}")
            if part not in keep:
                keep.add(part); pending.append(part)
    ct = _xml(package["[Content_Types].xml"])
    for item in list(ct):
        if item.get("PartName") and item.get("PartName").lstrip("/") not in keep:
            ct.remove(item)
    package["[Content_Types].xml"] = _bytes(ct)
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        for part in sorted(keep):
            archive.writestr(part, package[part])


def _normalize_package_metadata(package):
    """PowerPoint requires global master/layout IDs and one notes master.

    These bookkeeping IDs are independent of object IDs, styles and artwork.
    Source decks share IDs because they were split from the same presentation.
    """
    presentation = _xml(package["ppt/presentation.xml"])
    rels = _relations(package, "ppt/presentation.xml")
    next_id = 2147483648
    for master_id in presentation.find(f"{{{P}}}sldMasterIdLst"):
        master_id.set("id", str(next_id)); next_id += 1
        relation = next(r for r in rels if r.get("Id") == master_id.get(f"{{{R}}}id"))
        part = _target("ppt/presentation.xml", relation.get("Target"))
        master = _xml(package[part])
        layouts = master.find(f"{{{P}}}sldLayoutIdLst")
        if layouts is not None:
            for layout in layouts:
                layout.set("id", str(next_id)); next_id += 1
        package[part] = _bytes(master)
    notes = presentation.find(f"{{{P}}}notesMasterIdLst")
    if notes is not None and len(notes):
        rid = notes[0].get(f"{{{R}}}id")
        canonical = _target("ppt/presentation.xml", next(r for r in rels if r.get("Id") == rid).get("Target"))
        for item in list(notes)[1:]:
            notes.remove(item)
        for relation in list(rels):
            if relation.get("Type").endswith("/notesMaster") and relation.get("Id") != rid:
                rels.remove(relation)
        for name in list(package):
            if name.endswith(".rels") and name != _relpath("ppt/presentation.xml"):
                relationships = _xml(package[name])
                changed = False
                for relation in relationships:
                    if relation.get("Type").endswith("/notesMaster"):
                        relation.set("Target", "/" + canonical); changed = True
                if changed:
                    package[name] = _bytes(relationships)
    package["ppt/presentation.xml"] = _bytes(presentation)
    package[_relpath("ppt/presentation.xml")] = _bytes(rels)


def check_source_text_fit(spec, template_path, manifest_path):
    validate_source_presentation(spec, load_source_manifest(manifest_path))


def inspect_source_template(template_path, manifest_path):
    manifest = load_source_manifest(manifest_path)
    with ZipFile(template_path) as archive:
        for layout in manifest["layouts"].values():
            for variant in layout["variants"].values():
                data = archive.read(variant["package_part"])
                if hashlib.sha256(data).hexdigest() != variant["slide_xml_sha256"]:
                    raise ValueError("Исходная XML-разметка слайда изменилась")
    return {"template": str(Path(template_path).resolve()), "template_version": "2.0", "render_mode": "source_native", "layout_count": len(manifest["layouts"]), "design_count": sum(len(l["variants"]) for l in manifest["layouts"].values()), "layouts": {k: {"name": v["name"], "source_file": v["source_file"], "variant_count": len(v["variants"])} for k, v in manifest["layouts"].items()}}
