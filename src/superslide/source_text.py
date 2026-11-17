"""Shared source-text roles for the editor and generated PowerPoint styling.

Roles are inferred from the original source XML, before generated text or a
typography policy changes its appearance. Partial emphasis is not a heading.
"""

import re

from lxml import etree as E


A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS = {"a": A, "p": P}


def _xml(value):
    return E.fromstring(value, E.XMLParser(resolve_entities=False, no_network=True))


def _paragraph_style(paragraph):
    """Return representative size and whether all visible text is bold.

    A paragraph can inherit run properties through its default/end properties.
    Check each displayed run rather than letting one bold initial determine
    the role of the whole paragraph.
    """
    defaults = paragraph.xpath("./a:pPr/a:defRPr | ./a:endParaRPr", namespaces=NS)
    runs = paragraph.xpath("./a:r[a:t[string-length(normalize-space(.)) > 0]] | "
                           "./a:fld[a:t[string-length(normalize-space(.)) > 0]]", namespaces=NS)
    properties = [run.find(f"{{{A}}}rPr") for run in runs]
    if not properties:
        properties = [None]
    if not defaults and not any(item is not None for item in properties):
        return None, False

    def attribute(properties, name):
        for candidate in [properties, *defaults]:
            if candidate is not None and candidate.get(name) is not None:
                return candidate.get(name)
        return None

    size = attribute(properties[0], "sz")
    return (int(size) if size and size.isdigit() else None,
            all(attribute(item, "b") in {"1", "true"} for item in properties))


def paragraph_is_bold(paragraph):
    """True only when every visible run/field uses bold source formatting."""
    visible = paragraph.xpath("./a:r[a:t[string-length(normalize-space(.)) > 0]] | "
                              "./a:fld[a:t[string-length(normalize-space(.)) > 0]]", namespaces=NS)
    return bool(visible) and _paragraph_style(paragraph)[1]


def paragraph_roles(field, paragraphs):
    """Recognize an original heading/body pair without using inline emphasis."""
    styles = []
    for index, text in enumerate(paragraphs):
        if not text.strip():
            styles.append((None, False))
            continue
        raw = field.get("paragraph_styles", [])
        if index >= len(raw):
            return None
        try:
            paragraph = _xml(raw[index].encode("utf-8"))
        except (E.XMLSyntaxError, AttributeError, TypeError):
            return None
        styles.append(_paragraph_style(paragraph))
    filled = [index for index, text in enumerate(paragraphs) if text.strip()]
    if len(filled) < 2:
        return None
    first = filled[0]
    size, bold = styles[first]
    rest_sizes = [styles[index][0] for index in filled[1:] if styles[index][0] is not None]
    smaller_body = size is not None and rest_sizes and size > max(rest_sizes)
    plain_body = bold and any(not styles[index][1] for index in filled[1:])
    if not (smaller_body or plain_body):
        return None
    return ["separator" if not text.strip() else "heading" if index == first else "body"
            for index, text in enumerate(paragraphs)]


def source_heading_fields(variant, slide_root):
    """Confirm standalone subheadings by a smaller text companion below them.

    Original bold formatting alone can describe an emphasized body paragraph.
    Require a separate smaller text box in the same column and XML coordinate
    system, or an explicit source heading label/placeholder. Uncertain text
    remains regular body text.
    """
    shapes = {}
    for shape in slide_root.xpath(".//p:sp[p:txBody]", namespaces=NS):
        ids = shape.xpath("./p:nvSpPr/p:cNvPr/@id", namespaces=NS)
        if ids:
            shapes.setdefault(ids[0], shape)
    items = []
    headings = set()
    for field_id, field in variant.get("fields", {}).items():
        sample = field.get("sample_text", "")
        if (field.get("kind") != "text"
                or field.get("scope", field_id.split("_", 1)[0]) != "slide"
                or field.get("automatic")
                or not any(character.isalpha() for character in sample)
                or " ".join(sample.split()).startswith("Любое несанкционированное копирование")):
            continue
        shape = shapes.get(str(field.get("shape_id", field_id.rsplit("_", 1)[-1])))
        if shape is None:
            continue
        try:
            paragraphs = [_xml(raw.encode("utf-8")) for raw in field.get("paragraph_styles", [])]
        except (E.XMLSyntaxError, AttributeError, TypeError):
            continue
        visible = [paragraph for paragraph in paragraphs
                   if paragraph.xpath("./a:r/a:t[string-length(normalize-space(.)) > 0] | "
                                      "./a:fld/a:t[string-length(normalize-space(.)) > 0]", namespaces=NS)]
        if not visible:
            continue
        bold = all(paragraph_is_bold(paragraph) for paragraph in visible)
        descriptor = sample + " " + field.get("name", "")
        explicit_heading = (bool(re.search(r"\b(?:заголовок|title|heading)\b", descriptor, re.IGNORECASE))
                            or bool(shape.xpath("./p:nvSpPr/p:nvPr/p:ph[@type='title' or @type='ctrTitle']", namespaces=NS)))
        if explicit_heading:
            headings.add(field_id)
        transform = shape.find(f"{{{P}}}spPr/{{{A}}}xfrm")
        if transform is None:
            continue
        offset = transform.find(f"{{{A}}}off")
        extent = transform.find(f"{{{A}}}ext")
        if offset is None or extent is None:
            continue
        try:
            left, top = int(offset.get("x")), int(offset.get("y"))
            width, height = int(extent.get("cx")), int(extent.get("cy"))
        except (TypeError, ValueError):
            continue
        if width <= 0 or height <= 0:
            continue
        sizes = [_paragraph_style(paragraph)[0] for paragraph in visible]
        known_sizes = [size for size in sizes if size is not None]
        if not known_sizes:
            continue
        items.append({"field_id": field_id, "shape": shape, "left": left, "top": top,
                      "width": width, "size": min(known_sizes),
                      "bold": bold})
    for candidate in items:
        if not candidate["bold"]:
            continue
        for companion in items:
            if (candidate["size"] <= companion["size"]
                    or companion["top"] <= candidate["top"]
                    or candidate["shape"].getparent() is not companion["shape"].getparent()):
                continue
            overlap = min(candidate["left"] + candidate["width"], companion["left"] + companion["width"]) - max(candidate["left"], companion["left"])
            if overlap >= min(candidate["width"], companion["width"]) / 2:
                headings.add(candidate["field_id"])
                break
    return headings


def source_title_metadata(variant, slide_root, eligible_fields=None):
    """Identify the same visible title for editing and output typography.

    Native title placeholders win. Otherwise choose the largest actual text
    run, breaking ties by vertical position. Unused list-level sample fonts and
    purely numeric KPI boxes do not compete with visible headings.
    """
    fields = variant.get("fields", {})
    if eligible_fields is None:
        eligible_fields = {
            field_id for field_id, field in fields.items()
            if field.get("kind") == "text"
            and field.get("scope", field_id.split("_", 1)[0]) == "slide"
            and not field.get("automatic")
            and field.get("sample_text", "").strip() not in {"##.##.####", "‹#›", "<#>"}
            and not " ".join(field.get("sample_text", "").split()).startswith("Любое несанкционированное копирование")
        }
    text_shapes = {}
    native_titles = set()
    for shape in slide_root.xpath(".//p:sp[p:txBody]", namespaces=NS):
        ids = shape.xpath("./p:nvSpPr/p:cNvPr/@id", namespaces=NS)
        if not ids:
            continue
        field_id = "slide_text_" + ids[0]
        if field_id not in eligible_fields:
            continue
        text_shapes[field_id] = shape
        if shape.xpath("./p:nvSpPr/p:nvPr/p:ph[@type='title' or @type='ctrTitle']", namespaces=NS):
            native_titles.add(field_id)

    def heading_score(field_id):
        shape = text_shapes[field_id]
        sizes = shape.xpath("./p:txBody/a:p/a:r[a:t[string-length(.) > 0]]/a:rPr/@sz | "
                            "./p:txBody/a:p/a:fld[a:t[string-length(.) > 0]]/a:rPr/@sz", namespaces=NS)
        if not sizes:
            sizes = shape.xpath("./p:txBody/a:p/a:pPr/a:defRPr/@sz | "
                                "./p:txBody/a:p/a:endParaRPr/@sz", namespaces=NS)
        font_size = max((int(size) for size in sizes if size.isdigit()), default=0)
        positions = shape.xpath("./p:spPr/a:xfrm/a:off/@y", namespaces=NS)
        top = int(positions[0]) if positions else 0
        return (field_id in native_titles, font_size, -top)

    candidates = [field_id for field_id in text_shapes
                  if field_id in native_titles or any(c.isalpha() for c in fields[field_id].get("sample_text", ""))]
    title_id = max(candidates, key=heading_score) if candidates else None
    field = fields.get(title_id, {})
    paragraphs = field.get("paragraphs", [])
    roles = paragraph_roles(field, paragraphs) if paragraphs else None
    title_index = roles.index("heading") if roles and "heading" in roles else None
    return {"title_field_id": title_id, "title_paragraph_only": title_index is not None,
            "title_paragraph_index": title_index}
