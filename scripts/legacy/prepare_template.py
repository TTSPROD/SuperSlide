"""Prepare the small, empty corporate PowerPoint template from approved artwork.

Only this preparation step manipulates OOXML packages. Runtime presentations use
python-pptx's normal ``slides.add_slide(layout)`` API and the manifest's fields.
The source presentation is read only and never overwritten.
"""

from __future__ import annotations

import argparse
from copy import deepcopy
import hashlib
from io import BytesIO
import json
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

from lxml import etree as ET

try:
    from .prepare_variants import add_variant_specs, make_variant_layout, physical_designs
except ImportError:  # Direct ``python scripts/legacy/prepare_template.py`` invocation.
    from prepare_variants import add_variant_specs, make_variant_layout, physical_designs


ROOT = Path(__file__).resolve().parents[2]
SOURCE = ROOT / "resources" / "severstal" / "severstal-template-v29.pptx"
EMU_PER_INCH = 914400
NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}
COLORS = {
    "primary": "002F6C",
    "text": "002F6C",
    "muted": "909090",
    "background": "FFFFFF",
    "accent": "FF0000",
    "grid": "DEDEDE",
    "light_background": "F4F4F4",
}


def q(prefix: str, name: str) -> str:
    return f"{{{NS[prefix]}}}{name}"


def el(prefix: str, local_name: str, **attrs: object) -> ET._Element:
    return ET.Element(q(prefix, local_name), {key: str(value) for key, value in attrs.items()})


def xml_bytes(root: ET._Element) -> bytes:
    return ET.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)


def field(
    idx: int,
    box: tuple[float, float, float, float],
    size: float,
    minimum: float,
    *,
    lines: int | None = None,
    bold: bool = False,
    kind: str = "text",
    color: str = COLORS["text"],
) -> dict:
    x, y, width, height = (round(value * EMU_PER_INCH) for value in box)
    result = {
        "idx": idx,
        "type": kind,
        "x": x,
        "y": y,
        "width": width,
        "height": height,
        "font_size_pt": float(size),
        "min_font_size_pt": float(minimum),
        "color": color,
        "bold": bold,
    }
    if lines is not None:
        result["max_lines"] = lines
    return result


def layout_specs() -> dict:
    """Semantic fields are the stable public contract with the renderer."""
    title = field(0, (0.45, 0.34, 11.30, 0.90), 32, 26, lines=2, bold=True)
    specs = {
        "title_slide": {
            "layout_name": "title_slide",
            "fields": {
                "title": field(0, (0.465, 2.55, 7.10, 1.85), 36, 30, lines=3, bold=True),
                "subtitle": field(1, (0.465, 4.65, 7.00, 0.75), 22, 18, lines=2),
                "author": field(2, (0.78, 6.06, 6.50, 0.31), 16, 14, lines=1),
                "date": field(3, (0.78, 6.44, 6.50, 0.31), 16, 14, lines=1),
            },
        },
        "bullet_list": {
            "layout_name": "bullet_list",
            "fields": {
                "title": deepcopy(title),
                "bullets": field(1, (0.55, 1.65, 12.10, 4.95), 24, 20, lines=11),
            },
        },
        "two_columns": {
            "layout_name": "two_columns",
            "fields": {
                "title": deepcopy(title),
                "left_heading": field(1, (0.55, 1.65, 5.75, 0.55), 22, 20, lines=1, bold=True),
                "left_bullets": field(2, (0.55, 2.35, 5.75, 4.20), 22, 18, lines=10),
                "right_heading": field(3, (7.00, 1.65, 5.75, 0.55), 22, 20, lines=1, bold=True),
                "right_bullets": field(4, (7.00, 2.35, 5.75, 4.20), 22, 18, lines=10),
            },
        },
        "quote": {
            "layout_name": "quote",
            "fields": {
                "quote": field(0, (1.00, 1.85, 11.25, 3.10), 34, 28, lines=5),
                "author": field(1, (1.00, 5.35, 11.25, 0.65), 20, 18, lines=2, bold=True),
            },
        },
        "table": {
            "layout_name": "table",
            "fields": {
                "title": deepcopy(title),
                "table": field(1, (0.55, 1.65, 12.10, 4.75), 18, 14, kind="content"),
                "footnote": field(2, (0.55, 6.58, 12.10, 0.30), 11, 10, lines=1, color=COLORS["muted"]),
            },
        },
        "bar_chart": {
            "layout_name": "bar_chart",
            "fields": {
                "title": deepcopy(title),
                "chart": field(1, (0.55, 1.92, 8.35, 4.75), 16, 12, kind="content"),
                "unit": field(2, (0.55, 1.40, 8.35, 0.32), 14, 12, lines=1, color=COLORS["muted"]),
                "takeaway": field(3, (9.40, 2.15, 3.25, 3.90), 24, 20, lines=8, bold=True),
            },
        },
    }
    # The six version-1.0 layouts above retain their exact names and field boxes.
    specs.update({
        "section_divider": {
            "layout_name": "section_divider",
            "fields": {
                "number": field(0, (0.55, 1.95, 3.00, 1.55), 78, 72, lines=1, bold=True, color="FFFFFF"),
                "title": field(1, (0.55, 3.75, 7.10, 1.65), 46, 38, lines=2, bold=True, color="FFFFFF"),
                "subtitle": field(2, (0.55, 5.65, 7.10, 1.15), 22, 18, lines=3, color="FFFFFF"),
            },
        },
        "photo_text": {
            "layout_name": "photo_text",
            "fields": {
                "title": deepcopy(title),
                "image": field(1, (0.55, 1.65, 6.30, 4.80), 18, 16, kind="picture"),
                "heading": field(2, (7.30, 1.65, 5.45, 0.95), 24, 20, lines=2, bold=True),
                "bullets": field(3, (7.30, 2.75, 5.45, 3.65), 22, 18, lines=8),
                "caption": field(4, (0.55, 6.62, 6.30, 0.28), 11, 10, lines=1, color=COLORS["muted"]),
            },
        },
        "image_spotlight": {
            "layout_name": "image_spotlight",
            "fields": {
                "image": field(0, (0.00, 0.00, 13.3333333333, 4.80), 18, 16, kind="picture"),
                "title": field(1, (0.60, 5.08, 10.30, 1.25), 32, 28, lines=2, bold=True, color="FFFFFF"),
                "caption": field(2, (0.60, 6.52, 10.30, 0.62), 16, 14, lines=2, color="FFFFFF"),
            },
        },
        "gallery": {
            "layout_name": "gallery",
            "fields": {
                "title": deepcopy(title),
                **{f"image_{i}": field(i, (0.55 + (i - 1) * 4.20, 1.65, 3.90, 3.95), 18, 16, kind="picture") for i in range(1, 4)},
                **{f"caption_{i}": field(i + 3, (0.75 + (i - 1) * 4.20, 5.84, 3.50, 0.75), 18, 16, lines=2) for i in range(1, 4)},
            },
        },
        "kpi_cards": {
            "layout_name": "kpi_cards",
            "fields": {
                "title": field(0, (0.55, 0.40, 10.10, 1.10), 32, 26, lines=2, bold=True, color="FFFFFF"),
                **{f"value_{i}": field(i, (0.80 + (i - 1) * 4.19, 2.42, 3.35, 0.95), 46, 38, lines=1, bold=True) for i in range(1, 4)},
                **{f"label_{i}": field(i + 3, (0.80 + (i - 1) * 4.19, 3.62, 3.35, 0.90), 23, 20, lines=2, bold=True) for i in range(1, 4)},
                **{f"detail_{i}": field(i + 6, (0.80 + (i - 1) * 4.19, 4.82, 3.35, 1.40), 18, 16, lines=4) for i in range(1, 4)},
            },
        },
        "timeline": {
            "layout_name": "timeline",
            "fields": {
                "title": deepcopy(title),
                **{f"step_title_{i}": field(i, (0.55 + (i - 1) * 3.20, 3.05, 2.65, 0.92), 20, 18, lines=2, bold=True) for i in range(1, 5)},
                **{f"step_text_{i}": field(i + 4, (0.55 + (i - 1) * 3.20, 4.22, 2.65, 2.15), 18, 16, lines=6) for i in range(1, 5)},
            },
        },
        "comparison": {
            "layout_name": "comparison",
            "fields": {
                "title": deepcopy(title),
                "left_heading": field(1, (0.80, 1.91, 5.30, 0.46), 22, 20, lines=1, bold=True, color="FFFFFF"),
                "left_bullets": field(2, (0.80, 2.92, 5.15, 3.50), 22, 18, lines=8),
                "right_heading": field(3, (7.23, 1.91, 5.30, 0.46), 22, 20, lines=1, bold=True, color="FFFFFF"),
                "right_bullets": field(4, (7.23, 2.92, 5.15, 3.50), 22, 18, lines=8),
            },
        },
        "closing": {
            "layout_name": "closing",
            "fields": {
                "title": field(0, (0.55, 2.50, 7.10, 1.60), 46, 38, lines=2, bold=True, color="FFFFFF"),
                "subtitle": field(1, (0.55, 4.50, 7.10, 0.95), 24, 20, lines=2, color="FFFFFF"),
                "contact": field(2, (0.55, 6.00, 7.10, 0.80), 20, 18, lines=2, color="FFFFFF"),
            },
        },
    })
    for name in ("photo_text", "gallery", "timeline", "comparison"):
        specs[name]["fields"]["title"]["height"] = round(1.10 * EMU_PER_INCH)
    for name, spec in specs.items():
        spec["background_color"] = COLORS["primary"] if name in {"section_divider", "kpi_cards", "closing", "image_spotlight"} else COLORS["background"]
        spec["show_footer"] = name not in {"title_slide", "section_divider", "kpi_cards", "closing", "image_spotlight"}
    add_variant_specs(specs, field, COLORS, EMU_PER_INCH)
    return specs


def run_properties(size: float, color: str, bold: bool) -> ET._Element:
    result = el("a", "defRPr", sz=round(size * 100), b=int(bold), lang="ru-RU")
    fill = el("a", "solidFill")
    fill.append(el("a", "srgbClr", val=color))
    result.append(fill)
    result.append(el("a", "latin", typeface="Arial"))
    result.append(el("a", "ea", typeface="Arial"))
    result.append(el("a", "cs", typeface="Arial"))
    return result


def paragraph_style(name: str, spec: dict, *, level: int | None = None) -> ET._Element:
    tag = f"lvl{level}pPr" if level is not None else "pPr"
    attrs = {"algn": "l", "marL": 0, "indent": 0}
    bullets = "bullets" in name
    if bullets:
        attrs.update(marL=round(0.28 * EMU_PER_INCH), indent=-round(0.24 * EMU_PER_INCH))
    result = el("a", tag, **attrs)
    spacing = el("a", "lnSpc")
    spacing.append(el("a", "spcPct", val=110000))
    result.append(spacing)
    after = el("a", "spcAft")
    after.append(el("a", "spcPts", val=1100 if bullets else 0))
    result.append(after)
    if bullets:
        bullet_color = el("a", "buClr")
        bullet_color.append(el("a", "srgbClr", val=COLORS["primary"]))
        result.append(bullet_color)
        result.append(el("a", "buFont", typeface="Arial"))
        result.append(el("a", "buChar", char="•"))
    else:
        result.append(el("a", "buNone"))
    result.append(run_properties(spec["font_size_pt"], spec["color"], spec["bold"]))
    return result


def placeholder(name: str, spec: dict, shape_id: int) -> ET._Element:
    shape = el("p", "sp")
    nv = el("p", "nvSpPr")
    nv.append(el("p", "cNvPr", id=shape_id, name=name))
    visual = el("p", "cNvSpPr")
    visual.append(el("a", "spLocks", noGrp=1))
    nv.append(visual)
    application = el("p", "nvPr")
    ph_type = "pic" if spec["type"] == "picture" else "obj" if spec["type"] == "content" else "title" if name == "title" else "body"
    application.append(el("p", "ph", type=ph_type, idx=spec["idx"]))
    nv.append(application)
    shape.append(nv)
    props = el("p", "spPr")
    transform = el("a", "xfrm")
    transform.append(el("a", "off", x=spec["x"], y=spec["y"]))
    transform.append(el("a", "ext", cx=spec["width"], cy=spec["height"]))
    props.append(transform)
    props.append(el("a", "noFill"))
    shape.append(props)
    body = el("p", "txBody")
    # Runtime writes explicit typography. These styles also make the layouts usable
    # in PowerPoint's normal New Slide menu without the runtime.
    body_pr = el("a", "bodyPr", wrap="square", anchor="t", lIns=0, rIns=0, tIns=0, bIns=0)
    body_pr.append(el("a", "noAutofit"))
    body.append(body_pr)
    styles = el("a", "lstStyle")
    styles.append(paragraph_style(name, spec, level=1))
    body.append(styles)
    paragraph = el("a", "p")
    paragraph.append(paragraph_style(name, spec))
    body.append(paragraph)
    shape.append(body)
    return shape


def fixed_shape(name: str, box: tuple[float, float, float, float], color: str, shape_id: int, *, text: str | None = None, ellipse: bool = False) -> ET._Element:
    """Native layout decoration; never regenerated by the content renderer."""
    shape = el("p", "sp")
    nv = el("p", "nvSpPr")
    nv.append(el("p", "cNvPr", id=shape_id, name="Fixed: " + name))
    nv.append(el("p", "cNvSpPr"))
    nv.append(el("p", "nvPr"))
    shape.append(nv)
    props = el("p", "spPr")
    transform = el("a", "xfrm")
    x, y, width, height = (round(item * EMU_PER_INCH) for item in box)
    transform.append(el("a", "off", x=x, y=y))
    transform.append(el("a", "ext", cx=width, cy=height))
    props.append(transform)
    geometry = el("a", "prstGeom", prst="ellipse" if ellipse else "rect")
    geometry.append(el("a", "avLst"))
    props.append(geometry)
    fill = el("a", "solidFill")
    fill.append(el("a", "srgbClr", val=color))
    props.append(fill)
    line = el("a", "ln")
    line.append(el("a", "noFill"))
    props.append(line)
    shape.append(props)
    if text is not None:
        body = el("p", "txBody")
        body.append(el("a", "bodyPr", anchor="ctr", lIns=0, rIns=0, tIns=0, bIns=0))
        body.append(el("a", "lstStyle"))
        paragraph = el("a", "p")
        style = el("a", "pPr", algn="ctr")
        style.append(el("a", "buNone"))
        style.append(run_properties(18, "FFFFFF", True))
        paragraph.append(style)
        run = el("a", "r")
        run_props = run_properties(18, "FFFFFF", True)
        run_props.tag = q("a", "rPr")
        run.append(run_props)
        value = el("a", "t")
        value.text = text
        run.append(value)
        paragraph.append(run)
        body.append(paragraph)
        shape.append(body)
    return shape


def original_white_wordmark(blue_source: ET._Element, box: tuple[float, float, float, float]) -> ET._Element:
    """Reuse the original white logo variant; only its outer transform changes."""
    tree = blue_source.find("p:cSld/p:spTree", NS)
    logo = deepcopy(next(item for item in tree if item.tag == q("p", "grpSp")))
    transform = logo.find("p:grpSpPr/a:xfrm", NS)
    x, y, width, height = (round(item * EMU_PER_INCH) for item in box)
    transform.find("a:off", NS).set("x", str(x))
    transform.find("a:off", NS).set("y", str(y))
    transform.find("a:ext", NS).set("cx", str(width))
    transform.find("a:ext", NS).set("cy", str(height))
    return logo


def decorate(tree: ET._Element, name: str, blue_source: ET._Element) -> None:
    if name == "image_spotlight":
        tree.append(fixed_shape("photo bottom panel", (0, 4.80, 13.3333333333, 2.70), COLORS["primary"], 2000))
        tree.append(original_white_wordmark(blue_source, (11.30, 5.25, 1.55, 0.483)))
    elif name == "kpi_cards":
        tree.append(original_white_wordmark(blue_source, (11.30, 0.47, 1.55, 0.483)))
        for index in range(3):
            tree.append(fixed_shape(f"KPI card {index + 1}", (0.55 + index * 4.19, 1.90, 3.85, 4.60), "FFFFFF", 2000 + index))
    elif name == "gallery":
        for index in range(3):
            tree.append(fixed_shape(f"gallery card {index + 1}", (0.55 + index * 4.20, 1.65, 3.90, 4.95), COLORS["light_background"], 2000 + index))
    elif name == "timeline":
        tree.append(fixed_shape("timeline track", (1.875, 2.435, 9.60, 0.035), COLORS["grid"], 2000))
        for index in range(4):
            tree.append(fixed_shape(f"timeline marker {index + 1}", (1.625 + index * 3.20, 2.20, 0.50, 0.50), COLORS["primary"], 2001 + index, text=str(index + 1), ellipse=True))
    elif name == "comparison":
        for index, x in enumerate((0.55, 6.98)):
            tree.append(fixed_shape(f"comparison card {index + 1}", (x, 1.65, 5.80, 4.95), COLORS["light_background"], 2000 + index * 2))
            tree.append(fixed_shape(f"comparison header {index + 1}", (x, 1.65, 5.80, 0.95), COLORS["primary"], 2001 + index * 2))


def make_layout(original: ET._Element, name: str, spec: dict, blue_source: ET._Element) -> ET._Element:
    result = deepcopy(original)
    result.set("preserve", "1")
    result.set("userDrawn", "1")
    result.set("type", "cust")
    result.set("showMasterSp", "1" if spec["show_footer"] else "0")
    common = result.find("p:cSld", NS)
    common.set("name", spec["layout_name"])
    tree = common.find("p:spTree", NS)
    for shape in list(tree)[2:]:
        # On the title retain the ORIGINAL vector wordmark, author/date icons,
        # and diagonal supergraphic. Content layouts inherit original master art.
        if name not in {"title_slide", "section_divider", "closing"} or shape.xpath(".//p:ph", namespaces=NS):
            tree.remove(shape)
    if name in {"kpi_cards", "image_spotlight"}:
        old_background = common.find("p:bg", NS)
        if old_background is not None:
            common.remove(old_background)
        background = el("p", "bg")
        background_props = el("p", "bgPr")
        fill = el("a", "solidFill")
        fill.append(el("a", "srgbClr", val=COLORS["primary"]))
        background_props.append(fill)
        background_props.append(el("a", "effectLst"))
        background.append(background_props)
        common.insert(0, background)
    decorate(tree, name, blue_source)
    for offset, (field_name, field_spec) in enumerate(spec["fields"].items()):
        tree.append(placeholder(field_name, field_spec, 1000 + offset))
    # Discard source layout metadata, including old unique creation identifiers.
    for child in list(result):
        if child.tag not in {q("p", "cSld"), q("p", "clrMapOvr")}:
            result.remove(child)
    return result


def relationships(entries: list[tuple[str, str, str]]) -> ET._Element:
    result = ET.Element(q("rel", "Relationships"), nsmap={None: NS["rel"]})
    for identifier, kind, target in entries:
        result.append(el("rel", "Relationship", Id=identifier, Type=NS["r"] + "/" + kind, Target=target))
    return result


def build_parts(source: ZipFile, specs: dict) -> dict[str, bytes]:
    parts: dict[str, bytes] = {}
    designs = list(physical_designs(specs))
    master = ET.fromstring(source.read("ppt/slideMasters/slideMaster1.xml"))
    master_layout_ids = master.find("p:sldLayoutIdLst", NS)
    master_layout_ids.clear()
    for index in range(1, len(designs) + 1):
        item = el("p", "sldLayoutId", id=2147483648 + index)
        item.set(q("r", "id"), f"rIdLayout{index}")
        master_layout_ids.append(item)
    parts["ppt/slideMasters/slideMaster1.xml"] = xml_bytes(master)
    entries = [
        (f"rIdLayout{index}", "slideLayout", f"../slideLayouts/slideLayout{index}.xml")
        for index in range(1, len(designs) + 1)
    ]
    # Original ids match the unmodified logo shape in the source master.
    entries.extend([
        ("rId57", "theme", "../theme/theme1.xml"),
        ("rId58", "image", "../media/image1.png"),
        ("rId59", "image", "../media/image2.svg"),
    ])
    parts["ppt/slideMasters/_rels/slideMaster1.xml.rels"] = xml_bytes(relationships(entries))
    content_source = ET.fromstring(source.read("ppt/slideLayouts/slideLayout3.xml"))
    title_source = ET.fromstring(source.read("ppt/slideLayouts/slideLayout20.xml"))
    blue_source = ET.fromstring(source.read("ppt/slideLayouts/slideLayout48.xml"))
    for index, (name, variant, spec) in enumerate(designs, start=1):
        if variant == "default":
            original = title_source if name == "title_slide" else blue_source if name in {"section_divider", "closing"} else content_source
            layout = make_layout(original, name, spec, blue_source)
            layout_relationships = xml_bytes(relationships([
                ("rId1", "slideMaster", "../slideMasters/slideMaster1.xml"),
            ]))
        else:
            layout, layout_relationships, media = make_variant_layout(source, spec, globals())
            parts.update(media)
        parts[f"ppt/slideLayouts/slideLayout{index}.xml"] = xml_bytes(layout)
        parts[f"ppt/slideLayouts/_rels/slideLayout{index}.xml.rels"] = layout_relationships

    # Keep the original compatible corporate theme and exact original logo assets.
    for name in ["ppt/theme/theme1.xml", "ppt/media/image1.png", "ppt/media/image2.svg"]:
        parts[name] = source.read(name)

    presentation = ET.fromstring(source.read("ppt/presentation.xml"))
    for child in list(presentation):
        if child.tag not in {q("p", "sldSz"), q("p", "notesSz"), q("p", "defaultTextStyle")}:
            presentation.remove(child)
    masters = el("p", "sldMasterIdLst")
    master_id = el("p", "sldMasterId", id=2147483648)
    master_id.set(q("r", "id"), "rId1")
    masters.append(master_id)
    presentation.insert(0, masters)
    presentation.insert(1, el("p", "sldIdLst"))
    parts["ppt/presentation.xml"] = xml_bytes(presentation)
    parts["ppt/_rels/presentation.xml.rels"] = xml_bytes(relationships([
        ("rId1", "slideMaster", "slideMasters/slideMaster1.xml"),
    ]))
    root_rels = relationships([("rId1", "officeDocument", "ppt/presentation.xml")])
    root_rels.append(el("rel", "Relationship", Id="rId2", Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", Target="docProps/core.xml"))
    root_rels.append(el("rel", "Relationship", Id="rId3", Type=NS["r"] + "/extended-properties", Target="docProps/app.xml"))
    parts["_rels/.rels"] = xml_bytes(root_rels)
    parts["docProps/core.xml"] = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>SuperSlide corporate base</dc:title><dc:creator>SuperSlide</dc:creator><dc:description>Prepared from Severstal corporate template v29. Fourteen semantic types with thirty-one design compositions.</dc:description><cp:revision>3</cp:revision></cp:coreProperties>'''
    parts["docProps/app.xml"] = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>SuperSlide</Application><PresentationFormat>On-screen Show (16:9)</PresentationFormat><Slides>0</Slides><Notes>0</Notes><HiddenSlides>0</HiddenSlides><Company>Severstal</Company><AppVersion>1.2</AppVersion></Properties>'''

    types = ET.Element(q("ct", "Types"), nsmap={None: NS["ct"]})
    for extension, content_type in {"rels": "application/vnd.openxmlformats-package.relationships+xml", "xml": "application/xml", "png": "image/png", "svg": "image/svg+xml"}.items():
        types.append(el("ct", "Default", Extension=extension, ContentType=content_type))
    overrides = {
        "ppt/presentation.xml": "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
        "ppt/slideMasters/slideMaster1.xml": "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml",
        "ppt/theme/theme1.xml": "application/vnd.openxmlformats-officedocument.theme+xml",
        "docProps/core.xml": "application/vnd.openxmlformats-package.core-properties+xml",
        "docProps/app.xml": "application/vnd.openxmlformats-officedocument.extended-properties+xml",
    }
    overrides.update({f"ppt/slideLayouts/slideLayout{i}.xml": "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml" for i in range(1, len(designs) + 1)})
    for name, content_type in overrides.items():
        types.append(el("ct", "Override", PartName="/" + name, ContentType=content_type))
    parts["[Content_Types].xml"] = xml_bytes(types)
    return parts


def prepare(source_path: Path, output_dir: Path) -> dict:
    output_dir.mkdir(parents=True, exist_ok=True)
    target = output_dir / "corporate_base.pptx"
    if source_path.resolve() == target.resolve():
        raise ValueError("The source presentation must never be overwritten")
    specs = layout_specs()
    before = hashlib.sha256(source_path.read_bytes()).hexdigest()
    with ZipFile(source_path) as source:
        parts = build_parts(source, specs)
    # Fixed timestamps and stable part order make preparation reproducible.
    with ZipFile(target, "w", compression=ZIP_DEFLATED, compresslevel=9) as output:
        for name, data in sorted(parts.items()):
            info = ZipInfo(name, date_time=(2026, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            output.writestr(info, data)
    manifest = {
        "template_version": "1.2",
        "compatible_versions": ["1.0", "1.1", "1.2"],
        "template_file": target.name,
        "font_family": "Arial",
        "colors": COLORS,
        "layouts": specs,
    }
    (output_dir / "template_manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    if hashlib.sha256(source_path.read_bytes()).hexdigest() != before:
        raise RuntimeError("Source presentation changed during preparation")
    return {"template": str(target), "bytes": target.stat().st_size, "layouts": list(specs), "design_count": len(list(physical_designs(specs))), "source_sha256": before}


def prepare_photos(source_path: Path, image_dir: Path, preview_dir: Path | None = None) -> dict:
    """Re-extract the visually approved selection recorded in catalog.json.

    Keep original bytes when the source is below 2 MB. The hall photograph needs
    JPEG recompression at full resolution; never crop or resize stored assets.
    The small contact sheet is a separate preview, not an available photo asset.
    """
    from PIL import Image, ImageDraw, ImageFont, ImageOps

    catalog = json.loads((image_dir / "catalog.json").read_text(encoding="utf-8"))
    with ZipFile(source_path) as source:
        for asset in catalog["images"]:
            path = (image_dir / asset["file"]).resolve()
            if not path.is_relative_to(image_dir.resolve()):
                raise ValueError("Photo catalog path escapes its asset directory")
            data = source.read(asset["source_media"])
            if len(data) >= 2_000_000:
                with Image.open(BytesIO(data)) as picture:
                    encoded = BytesIO()
                    picture.convert("RGB").save(encoded, "JPEG", quality=90, optimize=True)
                    data = encoded.getvalue()
            if len(data) >= 2_000_000:
                raise ValueError(f"Full-resolution photo still exceeds 2 MB: {asset['id']}")
            path.write_bytes(data)
            with Image.open(path) as picture:
                if picture.size != (asset["width"], asset["height"]):
                    raise ValueError(f"Photo dimensions disagree with catalog: {asset['id']}")
    preview = Image.new("RGB", (1440, 900), "#F4F4F4")
    draw = ImageDraw.Draw(preview)
    try:
        font = ImageFont.truetype("C:/Windows/Fonts/arial.ttf", 22)
    except OSError:
        font = ImageFont.load_default(size=22)
    for index, asset in enumerate(catalog["images"]):
        x, y = (index % 3) * 480, (index // 3) * 450
        with Image.open(image_dir / asset["file"]) as picture:
            thumb = ImageOps.contain(picture.convert("RGB"), (450, 370))
            preview.paste(thumb, (x + (480 - thumb.width) // 2, y + 12 + (370 - thumb.height) // 2))
        draw.text((x + 15, y + 395), f"{asset['id']} | source slide {asset['source_slide']}", font=font, fill="#002F6C")
    preview_dir = preview_dir or ROOT / "output" / "assets" / "images"
    preview_dir.mkdir(parents=True, exist_ok=True)
    preview_path = preview_dir / "selection_preview.jpg"
    preview.save(preview_path, "JPEG", quality=88, optimize=True)
    return {"photos": len(catalog["images"]), "catalog": str(image_dir / "catalog.json"), "contact_sheet": str(preview_path)}


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=SOURCE)
    parser.add_argument("--output-dir", type=Path, default=ROOT / "templates" / "legacy")
    parser.add_argument("--extract-images", action="store_true", help="Re-extract the approved catalog photos and create a contact sheet")
    parser.add_argument("--image-dir", type=Path, default=ROOT / "assets" / "images")
    parser.add_argument("--preview-dir", type=Path, default=ROOT / "output" / "assets" / "images")
    args = parser.parse_args()
    report = prepare(args.source, args.output_dir)
    if args.extract_images:
        report.update(prepare_photos(args.source, args.image_dir, args.preview_dir))
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
