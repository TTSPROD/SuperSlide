"""Conservative structural hints about repeated content in source variants.

These are capacities of the original artwork, not evidence that user data
contains that many facts. Derive hints from the manifest without loading PPTX.
"""

import re
import hashlib
from collections import defaultdict
from functools import lru_cache
import json
from pathlib import Path

from lxml import etree as E


_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS = {"a": _A}
_NUMBERED = re.compile(r"^(.*?)\s*#\s*(\d+)\b", re.IGNORECASE)
_NAMED = re.compile(
    r"\b(\d+|одна|один|две|два|три|четыре|пять|шесть|семь|восемь|девять|десять)\s+"
    r"(карточ\w*|колон\w*|диаграмм\w*|график\w*)\b", re.IGNORECASE,
)
_NUMBERS = {"одна": 1, "один": 1, "две": 2, "два": 2, "три": 3, "четыре": 4,
            "пять": 5, "шесть": 6, "семь": 7, "восемь": 8, "девять": 9, "десять": 10}
# The reviewed counts below are specific to these 211 source slide XML hashes.
_REVIEWED_SLIDES_SHA256 = "1500ecaf0c94208ee19c8ef2a0d7088d2f8cfc3bf2938fa6a4f981d9d6c79b65"


def verify_reviewed_slots(manifest):
    """Refuse stale manual counts if the underlying source template changes."""
    rows = [f"{layout_id}/{variant_id}:{variant.get('slide_xml_sha256', '')}"
            for layout_id, layout in sorted(manifest.get("layouts", {}).items())
            for variant_id, variant in sorted(layout.get("variants", {}).items())]
    digest = hashlib.sha256("\n".join(rows).encode("utf-8")).hexdigest()
    if digest != _REVIEWED_SLIDES_SHA256:
        raise ValueError("Исходные варианты шаблона изменились: пересчитайте и проверьте число блоков перед генерацией")
# Reviewed against the original 17 slides; values describe independent visual
# blocks per section, not runs, text shapes, or the numbering in a slide name.
_BULLET_SECTIONS = {
    1: (5,), 2: (8,), 3: (4,), 4: (4,), 5: (4,), 6: (5, 2),
    7: (5, 3), 8: (4, 4), 9: (4, 4), 10: (4, 4, 4),
    11: (4, 4, 4), 12: (3,), 13: (4,), 14: (4,),
    15: (3,), 16: (6,), 17: (2, 2, 2),
}
_CARD_SECTIONS = {
    1: (8,), 2: (8,), 3: (12,), 4: (12,), 5: (8,), 6: (3,), 7: (4,),
    8: (5,), 9: (3, 2, 1), 10: (2,), 11: (3,), 12: (3,), 13: (4,),
    14: (3,), 15: (4,), 16: (2,), 17: (2,),
}
_SCHEME_SECTIONS = {
    1: (5,), 2: (4,), 3: (4,), 4: (2,), 5: (2,), 6: (4,), 7: (4, 1),
    8: (3,), 9: (6,), 10: (8,), 11: (1, 4), 12: (4,), 13: (4,),
    14: (5,), 15: (5,), 16: (5,),
}
_REVIEWED_COUNTS = {
    "circle_schemes": (8, 8, 3, 6, 8, 3),
    "team": (9, 6, 4, 4, 4),
    "photo": (3, 3, 3, 3, 4, 6, 8, 4, 8),
}
_CONTENTS_COUNTS = (8, 8, 8, 8, 7)
_QUOTE_EXCEPTIONS = {
    3: {"kind": "content_blocks", "count": 3},
    13: {"kind": "quotes", "count": 3},
    14: {"kind": "quotes", "count": 7},
    19: {"kind": "mixed_content", "count": 4, "main_theses": 1, "supporting_blocks": 3},
    20: {"kind": "mixed_content", "count": 4, "main_theses": 1, "supporting_blocks": 3},
    21: {"kind": "indicators", "count": 3},
}
_DIAGRAM_COMPOSITES = {
    1: (5, 1), 11: (5, 3), 13: (3, 1), 17: (3, 1), 18: (2, 1),
    19: (4, 3), 20: (4, 3), 21: (4, 2),
    24: (4, 1), 25: (3, 3), 26: (3, 3), 27: (2, 6), 29: (2, 2),
    30: (7, 1),
}


@lru_cache(maxsize=1)
def _icon_page_counts():
    """Count extracted actual icons, not the underlying PowerPoint fragments."""
    path = Path(__file__).resolve().parents[2] / "assets/icons/severstal/catalog.json"
    counts = defaultdict(int)
    try:
        catalog = json.loads(path.read_text(encoding="utf-8-sig"))
        for record in catalog.get("icons", []):
            slide = record.get("source", {}).get("slide")
            if type(slide) is int:
                counts[slide] += 1
    except (OSError, ValueError, TypeError, AttributeError):
        return {}
    return dict(counts)


def drawn_table_metadata(fields):
    """Describe reviewed grids made from separate text, shapes and charts."""
    groups = {}
    roles = {"drawn_table_column_header": "headers", "drawn_table_cell": "cells",
             "drawn_table_section_header": "sections", "drawn_table_column_chart": "column_charts"}
    for field_id, field in fields.items():
        for key, destination in roles.items():
            role = field.get(key)
            if not isinstance(role, dict) or not role.get("table_id"):
                continue
            group = groups.setdefault(role["table_id"], {"table_id": role["table_id"],
                "header_placement": "external", "headers": [], "cells": [], "sections": [], "column_charts": []})
            group[destination].append({"field_id": field_id, "kind": field["kind"],
                                       **{name: value for name, value in role.items() if name != "table_id"}})
    for group in groups.values():
        group["headers"].sort(key=lambda item: item["column_index"])
        group["cells"].sort(key=lambda item: (item["row_index"], item["column_index"]))
        group["column_header_field_ids"] = [item["field_id"] for item in group["headers"]]
        group["column_count"] = len(group["headers"])
        group["data_row_count"] = max(
            [item["row_index"] + item.get("row_span", 1) - 1 for item in group["cells"]]
            + [item.get("row_span", 1) for item in group["column_charts"]], default=0)
    return list(groups.values())


def source_slot_metadata(variant):
    """Return only supported counts, with the source fields that justify them.

    Numbered samples like «Название #1» form a group until a number repeats.
    Separate groups can share the same label (e.g. rows with 3, 2 and 1 cards).
    Bullet positions are reported only when DrawingML declares them explicitly.
    """
    result = {"groups": []}
    # Count each kind independently: one text shape is not necessarily one
    # logical card, nor one bullet point. Include real IDs so the planner can
    # relate the quantities to the fields it later fills.
    fields = variant.get("fields", {})
    kinds = {kind: [] for kind in ("text", "picture", "icon", "table", "chart")}
    for field_id, field in fields.items():
        if (field.get("scope", field_id.split("_", 1)[0]) != "slide"
                or field.get("automatic") or field.get("kind") not in kinds):
            continue
        sample = field.get("sample_text", "").strip()
        if (sample in {"##.##.####", "‹#›", "<#>"}
                or " ".join(sample.split()).startswith("Любое несанкционированное копирование")):
            continue
        kinds[field["kind"]].append((field_id, field))
    inventory = {kind + "_fields": len(records) for kind, records in kinds.items()}
    inventory["text_paragraphs"] = sum(len(field.get("paragraphs", [])) for _, field in kinds["text"])
    inventory["tables"] = [{"field_id": fid, "rows": field.get("row_count"), "columns": field.get("column_count"),
                            **{key: field["table_structure"][key] for key in ("header_placement", "header_row_count", "body_row_count", "data_row_count", "section_header_row_indexes", "total_row_indexes")
                               if key in field.get("table_structure", {})}}
                           for fid, field in kinds["table"]]
    inventory["charts"] = [{"field_id": fid, "categories": field.get("chart", {}).get("category_count"),
                            "series": field.get("chart", {}).get("series_count")}
                           for fid, field in kinds["chart"]]
    result["block_inventory"] = inventory
    inventory["drawn_tables"] = [{"table_id": table["table_id"], "columns": table["column_count"],
                                   "data_rows": table["data_row_count"], "header_placement": "external"}
                                  for table in drawn_table_metadata(fields)]
    part = variant.get("package_part", "")
    source_type = part.split("/")[1] if part.startswith("sources/") and len(part.split("/")) > 2 else ""
    if source_type == "title" and variant.get("source_slide") == 1:
        result["primary_blocks"] = {"kind": "reference_instruction", "count": 4, "reference_only": True}
    elif source_type == "title" and variant.get("source_slide") in {9, 10}:
        result["primary_blocks"] = {"kind": "profile", "count": 1}
    elif source_type == "quotes" and variant.get("source_slide") in _QUOTE_EXCEPTIONS:
        result["primary_blocks"] = dict(_QUOTE_EXCEPTIONS[variant["source_slide"]])
    elif source_type in {"title", "separators", "big_text", "quotes", "last_slide"}:
        result["primary_blocks"] = {"kind": "single_composition", "count": 1}
    elif source_type == "contents" and type(variant.get("source_slide")) is int and 1 <= variant["source_slide"] <= len(_CONTENTS_COUNTS):
        result["primary_blocks"] = {"kind": "contents_entries", "count": _CONTENTS_COUNTS[variant["source_slide"] - 1]}
    elif source_type == "tables" and variant.get("source_slide") == 7:
        result["primary_blocks"] = {"kind": "drawn_table_rows", "count": 6,
                                    "section_counts": [3, 3], "data_columns": 4}
    elif source_type == "diagram" and variant.get("source_slide") == 28:
        result["primary_blocks"] = {"kind": "conclusion_blocks", "count": 4,
                                    "chart_points": 20, "first_block_bullets": 3}
    elif source_type == "diagram" and variant.get("source_slide") == 9:
        result["primary_blocks"] = {"kind": "comparison_categories", "count": 6,
                                    "charts": 2, "years": 2}
    elif source_type == "diagram" and variant.get("source_slide") == 10:
        result["primary_blocks"] = {"kind": "chart", "count": 3, "categories_per_chart": 8}
    elif source_type == "diagram" and variant.get("source_slide") in {15, 16}:
        count = 3 if variant["source_slide"] == 15 else 4
        result["primary_blocks"] = {"kind": "chart_text_cards", "count": count}
    elif source_type == "diagram" and variant.get("source_slide") in _DIAGRAM_COMPOSITES:
        texts, charts = _DIAGRAM_COMPOSITES[variant["source_slide"]]
        result["primary_blocks"] = {"kind": "mixed_charts_and_text", "count": texts + charts,
                                    "text_blocks": texts, "chart_blocks": charts}
        if variant["source_slide"] == 11:
            result["primary_blocks"]["categories_per_chart"] = 6
        if variant["source_slide"] == 17:
            result["primary_blocks"]["first_block_bullets"] = 3
    elif source_type == "diagram" and variant.get("source_slide") == 31:
        result["primary_blocks"] = {"kind": "chart_sections", "count": 6, "section_counts": [3, 3]}
    elif source_type == "diagram" and variant.get("source_slide") == 32:
        result["primary_blocks"] = {"kind": "chart_rows", "count": 8}
    elif source_type == "icons":
        count = _icon_page_counts().get(variant.get("source_slide"))
        if count is not None:
            result["primary_blocks"] = {"kind": "reference_icons", "count": count}
    elif source_type == "bullets" and variant.get("source_slide") in _BULLET_SECTIONS:
        sections = _BULLET_SECTIONS[variant["source_slide"]]
        result["primary_blocks"] = {"kind": "content_blocks", "count": sum(sections),
                                    "section_counts": list(sections)}
    elif source_type == "cards" and variant.get("source_slide") in _CARD_SECTIONS:
        sections = _CARD_SECTIONS[variant["source_slide"]]
        result["primary_blocks"] = {"kind": "cards", "count": sum(sections), "section_counts": list(sections)}
        if variant["source_slide"] == 10:
            result["primary_blocks"]["items_per_block"] = [5, 5]
    elif source_type == "schemes" and variant.get("source_slide") in _SCHEME_SECTIONS:
        sections = _SCHEME_SECTIONS[variant["source_slide"]]
        result["primary_blocks"] = {"kind": "scheme_nodes", "count": sum(sections), "section_counts": list(sections)}
    elif source_type == "photo" and variant.get("source_slide") == 9:
        result["primary_blocks"] = {"kind": "mixed_content", "count": 8, "photo_blocks": 5, "text_blocks": 3}
    elif source_type in _REVIEWED_COUNTS and type(variant.get("source_slide")) is int:
        number = variant["source_slide"]
        if 1 <= number <= len(_REVIEWED_COUNTS[source_type]):
            result["primary_blocks"] = {"kind": {"circle_schemes": "scheme_nodes", "team": "contacts", "photo": "content_blocks"}[source_type],
                                        "count": _REVIEWED_COUNTS[source_type][number - 1]}
    elif kinds["chart"]:
        result["primary_blocks"] = {"kind": "chart", "count": len(kinds["chart"])}
    elif kinds["table"]:
        result["primary_blocks"] = {"kind": "table", "count": len(kinds["table"])}
    match = _NAMED.search(variant.get("name", ""))
    if match:
        number, kind = match.groups()
        count = int(number) if number.isdecimal() else _NUMBERS[number.casefold()]
        if 1 <= count <= 40:
            result["named_item_count"] = count
            result["named_item_kind"] = kind.casefold()
    numbered = defaultdict(list)
    for field_id, field in variant.get("fields", {}).items():
        if field.get("kind") != "text" or field.get("scope", "slide") != "slide" or field.get("automatic"):
            continue
        first = field.get("sample_text", "").split("\n", 1)[0].strip()
        match = _NUMBERED.match(first)
        if match and match.group(1).strip():
            numbered[match.group(1).strip().casefold()].append((field_id, int(match.group(2))))
        positions = []
        for index, raw in enumerate(field.get("paragraph_styles", [])):
            try:
                paragraph = E.fromstring(raw.encode("utf-8"), E.XMLParser(resolve_entities=False, no_network=True))
            except (TypeError, AttributeError, E.XMLSyntaxError):
                continue
            if paragraph.xpath("./a:pPr/a:buChar | ./a:pPr/a:buAutoNum", namespaces=_NS):
                positions.append(index)
        if len(positions) >= 2:
            result["groups"].append({"kind": "bullet_paragraphs", "field_id": field_id,
                                     "count": len(positions), "paragraph_indexes": positions})
    for entries in numbered.values():
        if len(entries) < 2:
            continue
        run, seen = [], set()
        for field_id, number in entries:
            if number in seen:
                if run:
                    result["groups"].append({"kind": "numbered_text_fields", "count": len(run), "field_ids": run})
                run, seen = [], set()
            run.append(field_id)
            seen.add(number)
        if run:
            result["groups"].append({"kind": "numbered_text_fields", "count": len(run), "field_ids": run})
    if "named_item_count" in result and "primary_blocks" not in result:
        result["primary_blocks"] = {"kind": result["named_item_kind"], "count": result["named_item_count"]}
    return result
