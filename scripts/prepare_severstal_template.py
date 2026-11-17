"""Package the 15 supplied PPTX types without re-drawing a single object."""
from __future__ import annotations

import hashlib
import io
import json
import posixpath
from pathlib import Path
from zipfile import ZipFile

from lxml import etree as E
from PIL import Image
from pptx import Presentation

from superslide.source_template import (
    ROOT, P, A, R, REL, CT, NS, SOURCE_TYPES,
    _xml, _bytes, _relpath, _target, _related, _shape_id,
    _paragraph_text, _override, _append_rel, _prune_and_save,
    FOOTER_DATE_PLACEHOLDER,
)
from superslide.source_charts import inspect_chart
from superslide.source_icons import describe_icon_slot
from superslide.source_photos import describe_photo_frame


def _shape_placeholder(shape):
    values = shape.xpath("./p:nvSpPr/p:nvPr/p:ph | ./p:nvPicPr/p:nvPr/p:ph", namespaces=NS)
    return values[0] if values else None


def _has_geometry(shape):
    values = shape.xpath("./p:spPr/a:xfrm | ./p:grpSpPr/a:xfrm", namespaces=NS)
    return any(x.find(f"{{{A}}}off") is not None
               and (ext := x.find(f"{{{A}}}ext")) is not None
               and int(ext.get("cx", 0)) > 0 and int(ext.get("cy", 0)) > 0
               for x in values)


def _inherited_geometry_shape(package, part, shape):
    """Resolve placeholder geometry without modifying its source XML.

    Slide placeholders bind to their layout by idx. Layout placeholders bind
    to master placeholders by type; their local idx need not match the master.
    Some photos inherit their original frames, while icons currently have
    direct transforms. Both use the same source placeholder geometry.
    """
    if _has_geometry(shape):
        return None
    placeholder = _shape_placeholder(shape)
    if placeholder is None:
        return None
    owner = part
    for kind in ("slideLayout", "slideMaster"):
        target = _related(package, owner, kind)
        if target is None:
            continue
        root = _xml(package[target])
        matches = []
        for candidate in root.xpath(".//p:sp | .//p:pic", namespaces=NS):
            inherited = _shape_placeholder(candidate)
            if inherited is None:
                continue
            if kind == "slideLayout":
                matches_slot = inherited.get("idx", "0") == placeholder.get("idx", "0")
            else:
                matches_slot = inherited.get("type", "obj") == placeholder.get("type", "obj")
            if matches_slot:
                matches.append(candidate)
        if matches:
            if len(matches) != 1:
                raise ValueError(f"Ambiguous inherited image placeholder in {part}")
            candidate = matches[0]
            if _has_geometry(candidate):
                return candidate
            placeholder = _shape_placeholder(candidate)
        owner = target
    return None


_inherited_icon_shape = _inherited_geometry_shape


def _icon_placement(package, part, shape):
    inherited = _inherited_geometry_shape(package, part, shape)
    return describe_icon_slot(shape, inherited_shape=inherited)


def _photo_placement(package, part, shape):
    inherited = _inherited_geometry_shape(package, part, shape)
    return describe_photo_frame(shape, inherited_shape=inherited)


def _path_fill_fraction(path):
    """Measure filled contours, retaining opposite winding for transparent holes."""
    contours, points = [], []
    for command in path:
        kind = E.QName(command).localname
        values = [(float(p.get("x")), float(p.get("y"))) for p in command]
        if kind == "moveTo":
            if points:
                contours.append(points)
            points = values
        elif kind == "lnTo":
            points.extend(values)
        elif kind == "cubicBezTo" and points:
            p0, p1, p2, p3 = points[-1], *values
            for step in range(1, 17):
                t = step / 16
                points.append(tuple((1-t)**3*p0[d] + 3*(1-t)**2*t*p1[d]
                                    + 3*(1-t)*t*t*p2[d] + t**3*p3[d] for d in (0, 1)))
        elif kind == "close":
            if points:
                contours.append(points)
            points = []
        else:
            raise ValueError(f"Unsupported native icon contour: {kind}")
    if points:
        contours.append(points)
    signed_area = sum(sum(a[0]*b[1] - b[0]*a[1]
                          for a, b in zip(points, points[1:] + points[:1])) / 2
                      for points in contours)
    return abs(signed_area) / (float(path.get("w")) * float(path.get("h")))


def _native_icon_has_background(shape):
    """Find an opaque plate covering the icon's canvas, including grouped paths.

    Icon paths are already closed filled contours in these supplied sources.
    A large bounding box alone is insufficient: outline art has transparent
    holes and must remain in the group without a backing plate.
    """
    def inspect(node, canvas_fraction):
        if node.tag == f"{{{P}}}grpSp":
            transform = node.find(f"{{{P}}}grpSpPr/{{{A}}}xfrm")
            extent = transform.find(f"{{{A}}}chExt")
            area = int(extent.get("cx")) * int(extent.get("cy"))
            for child in node:
                if child.tag not in (f"{{{P}}}sp", f"{{{P}}}grpSp"):
                    continue
                transform = child.find(f"{{{P}}}spPr/{{{A}}}xfrm")
                if transform is None:
                    transform = child.find(f"{{{P}}}grpSpPr/{{{A}}}xfrm")
                extent = transform.find(f"{{{A}}}ext")
                fraction = canvas_fraction * int(extent.get("cx")) * int(extent.get("cy")) / area
                if inspect(child, fraction):
                    return True
            return False
        props = node.find(f"{{{P}}}spPr")
        if props.find(f"{{{A}}}solidFill") is None:
            return False
        preset = props.find(f"{{{A}}}prstGeom")
        if preset is not None and preset.get("prst") == "rect":
            return canvas_fraction >= .90
        paths = props.findall(f"{{{A}}}custGeom/{{{A}}}pathLst/{{{A}}}path")
        return any(path.get("fill") != "none" and canvas_fraction * _path_fill_fraction(path) >= .90
                   for path in paths)
    return inspect(shape, 1.0)


def _icon_background_metadata(package, part, shape):
    """Describe the retained original icon, without changing any source XML."""
    blips = shape.xpath(".//a:blip", namespaces=NS)
    if blips:
        rid = blips[0].get(f"{{{R}}}embed")
        relations = _xml(package[_relpath(part)])
        relation = next(r for r in relations if r.get("Id") == rid)
        data = package[_target(part, relation.get("Target"))]
        # SVG shapes carry an existing raster fallback in their a:blip. It has
        # exactly the same backing as the SVG; no additional renderer is needed.
        with Image.open(io.BytesIO(data)) as picture:
            histogram = picture.convert("RGBA").getchannel("A").histogram()
            backed = sum(histogram[241:]) / sum(histogram) >= .90
    else:
        backed = _native_icon_has_background(shape)
    return {"background_style": "with_background" if backed else "without_background",
            "has_background": backed}


def _fields_and_objects(prs, source_package, part, scope, prefix):
    """Record original object XML and bindings, including inherited artwork."""
    xml = _xml(source_package[part])
    objects = []
    fields = {}
    icon_ancestors = set()
    for z_order, shape in enumerate(xml.xpath(".//p:spTree/* | .//p:grpSp/*", namespaces=NS)):
        sid = _shape_id(shape)
        if sid is None:
            continue
        nv = shape.xpath("./*/p:cNvPr", namespaces=NS)
        transform = shape.find(f"{{{P}}}spPr/{{{A}}}xfrm")
        if transform is None:
            transform = shape.find(f"{{{P}}}grpSpPr/{{{A}}}xfrm")
        obj = {"shape_id": sid, "scope": scope, "name": nv[0].get("name") if nv else "", "tag": E.QName(shape).localname,
               "z_order": z_order, "transform_emu": {E.QName(c).localname: dict(c.attrib) for c in transform} if transform is not None else None}
        objects.append(obj)
        body = shape.find(f"{{{P}}}txBody")
        if body is not None:
            paragraphs = [_paragraph_text(p) for p in body.findall(f"{{{A}}}p")]
            text = "\n".join(paragraphs)
            ph = shape.find(f"{{{P}}}nvSpPr/{{{P}}}nvPr/{{{P}}}ph")
            if text.strip() and not (ph is not None and ph.get("type") in ("sldNum", "dt", "ftr")):
                # Dates / serial numbers are fixed only if marked by PPTX.
                kind = "text"
                fields[f"{scope}_{kind}_{sid}"] = {"kind": kind, "scope": scope, "shape_id": sid,
                    "name": obj["name"], "sample_text": text, "paragraphs": paragraphs,
                    "max_characters": max(len(text), 20),
                    "paragraph_styles": [E.tostring(p, encoding="unicode") for p in body.findall(f"{{{A}}}p")]}
                if text.strip() == FOOTER_DATE_PLACEHOLDER:
                    fields[f"{scope}_{kind}_{sid}"].update({"automatic": "current_date", "timezone": "Europe/Moscow", "format": "DD.MM.YYYY"})
        tables = shape.findall(f".//{{{A}}}tbl") if shape.tag == f"{{{P}}}graphicFrame" else []
        if tables:
            table = tables[0]
            rows = [["\n".join(_paragraph_text(p) for p in tc.findall(f"{{{A}}}txBody/{{{A}}}p")) for tc in tr.findall(f"{{{A}}}tc")] for tr in table.findall(f"{{{A}}}tr")]
            fields[f"{scope}_table_{sid}"] = {"kind": "table", "scope": scope, "shape_id": sid,
                "name": obj["name"], "row_count": len(rows), "column_count": len(rows[0]) if rows else 0, "rows": rows}
        charts = shape.xpath(".//*[local-name()='chart' and namespace-uri()='http://schemas.openxmlformats.org/drawingml/2006/chart']", namespaces=NS) if shape.tag == f"{{{P}}}graphicFrame" else []
        if charts:
            chart_rel = next(r for r in _xml(source_package[_relpath(part)]) if r.get("Id") == charts[0].get(f"{{{R}}}id"))
            chart_part = _target(part, chart_rel.get("Target"))
            fields[f"{scope}_chart_{sid}"] = {"kind": "chart", "scope": scope, "shape_id": sid, "name": obj["name"], "editable_in_powerpoint": True,
                "package_part": prefix + chart_part, "chart": inspect_chart(source_package[chart_part]),
                "update_policy": "Preserve source styling and data dimensions; replace values and embedded workbook together."}
        # The source also uses pictures as shape fills (portraits, circular
        # masks and inherited layout backgrounds), not only p:pic objects.
        if shape.tag in (f"{{{P}}}pic", f"{{{P}}}sp"):
            blips = shape.xpath(".//a:blip", namespaces=NS)
            if blips:
                rid = blips[0].get(f"{{{R}}}embed")
                rels = _xml(source_package[_relpath(part)])
                rel = next((r for r in rels if r.get("Id") == rid), None)
                if rel is not None:
                    media = _target(part, rel.get("Target"))
                    data = source_package[media]
                    is_icon = b"svgBlip" in E.tostring(shape) or media.endswith(".svg")
                    kind = "icon" if is_icon else "picture"
                    fields[f"{scope}_{kind}_{sid}"] = {"kind": kind, "scope": scope, "shape_id": sid,
                        "name": obj["name"], "media_part": prefix + media, "media_sha256": hashlib.sha256(data).hexdigest(),
                        "crop_and_transform_xml": E.tostring(shape, encoding="unicode")}
                    if kind == "icon":
                        fields[f"{scope}_{kind}_{sid}"]["placement"] = _icon_placement(source_package, part, shape)
                        fields[f"{scope}_{kind}_{sid}"].update(_icon_background_metadata(source_package, part, shape))
                    else:
                        fields[f"{scope}_{kind}_{sid}"]["placement"] = _photo_placement(source_package, part, shape)
        # Many source icons are native freeforms, rather than image objects.
        # Expose the original small Graphic slot, retaining its exact position.
        ext = transform.find(f"{{{A}}}ext") if transform is not None else None
        if (scope == "slide" and shape.tag in (f"{{{P}}}sp", f"{{{P}}}grpSp")
            and (obj["name"].startswith("Graphic")
                 or (obj["name"].startswith(("Рисунок", "Picture"))
                     and shape.find(f"{{{P}}}spPr/{{{A}}}custGeom") is not None))
            and ext is not None
            and 100000 <= int(ext.get("cx", 0)) <= 1400000
            and 100000 <= int(ext.get("cy", 0)) <= 1400000
            and not shape.xpath(".//a:t[text()] | .//a:blip", namespaces=NS)
            and not any(a in icon_ancestors for a in shape.iterancestors())):
            fields[f"{scope}_icon_{sid}"] = {"kind": "icon", "scope": scope, "shape_id": sid,
                "name": obj["name"], "native_shape": True, "transform_emu": obj["transform_emu"],
                "placement": _icon_placement(source_package, part, shape)}
            fields[f"{scope}_icon_{sid}"].update(_icon_background_metadata(source_package, part, shape))
            icon_ancestors.add(shape)
    return fields, objects


def prepare(root=ROOT):
    sources = root / "resources/severstal/templates"
    template = root / "templates/severstal_source.pptx"
    manifest_path = root / "templates/severstal_template.json"
    designs_dir = root / "templates/severstal_designs"
    designs_dir.mkdir(parents=True, exist_ok=True)
    blank = Presentation()
    blank.slide_width = 12192000; blank.slide_height = 6858000
    stream = io.BytesIO(); blank.save(stream)
    with ZipFile(stream) as archive:
        package = {n: archive.read(n) for n in archive.namelist()}
    ct = _xml(package["[Content_Types].xml"])
    presentation = _xml(package["ppt/presentation.xml"])
    rels = _xml(package[_relpath("ppt/presentation.xml")])
    for r in list(rels):
        if r.get("Type").endswith(("/slideMaster", "/notesMaster", "/slide")):
            rels.remove(r)
    masters = presentation.find(f"{{{P}}}sldMasterIdLst"); masters.clear()
    slides = E.Element(f"{{{P}}}sldIdLst")
    presentation.insert(list(presentation).index(presentation.find(f"{{{P}}}sldSz")), slides)
    notes_masters = E.Element(f"{{{P}}}notesMasterIdLst")
    presentation.insert(list(presentation).index(slides), notes_masters)
    manifest = {"template_version": "2.0", "render_mode": "source_native", "name": "Северсталь — исходные 15 типов",
        "slide_size_emu": [12192000, 6858000], "source_directory": "resources/severstal/templates",
        "policy": {"allowed_types": "Only the 15 supplied source files", "design": "Original native OOXML, no re-drawn layouts or font substitutions", "variants": "Every supplied source slide is a variant inside its original type", "text": "Uniform representative source style per replaced paragraph; preserve paragraph layout and object bounds; shorten content when necessary"},
        "photo_catalog": "assets/photos/severstal/catalog.json", "icon_catalog": "assets/icons/severstal/catalog.json",
        "icon_placement_policy": {"fit": "contain", "preserve_aspect_ratio": True, "alignment": "center", "slot_shape": "square", "square_side_policy": "min_original_extent", "coordinate_space": "parent"},
        "photo_placement_policy": {"fit": "cover", "preserve_aspect_ratio": True, "alignment": "center", "frame": "source", "preserve_source_mask": True, "coordinate_space": "parent"},
        "automatic_fields": {"footer_date": {"source_prompt": FOOTER_DATE_PLACEHOLDER, "value": "current_date_at_build", "timezone": "Europe/Moscow", "format": "DD.MM.YYYY"}}, "layouts": {}}
    global_slide = 0
    for layout_id, (stem, label) in SOURCE_TYPES.items():
        path = sources / f"Severstal_{stem}.pptx"
        with ZipFile(path) as archive:
            original = {n: archive.read(n) for n in archive.namelist()}
        if global_slide == 0:
            # Global theme/table styles affect objects even when a slide has its
            # own master. Carry the actual source globals, not python-pptx's.
            for global_part in ("ppt/theme/theme1.xml", "ppt/tableStyles.xml", "ppt/presProps.xml", "ppt/viewProps.xml"):
                if global_part in original:
                    package[global_part] = original[global_part]
            source_presentation = _xml(original["ppt/presentation.xml"])
            presentation.attrib.clear(); presentation.attrib.update(source_presentation.attrib)
            for tag in ("sldSz", "notesSz", "defaultTextStyle"):
                old = presentation.find(f"{{{P}}}{tag}")
                new = source_presentation.find(f"{{{P}}}{tag}")
                if old is not None and new is not None:
                    presentation.replace(old, deepcopy_element(new))
        prefix = f"sources/{layout_id}/"
        for part, data in original.items():
            if part == "[Content_Types].xml":
                continue
            package[prefix + part] = data
        for item in _xml(original["[Content_Types].xml"]):
            if item.tag == f"{{{CT}}}Override":
                _override(ct, prefix + item.get("PartName").lstrip("/"), item.get("ContentType"))
            elif not any(r.get("Extension") == item.get("Extension") for r in ct):
                ct.append(deepcopy_element(item))
        source_rels = _xml(original[_relpath("ppt/presentation.xml")])
        for relation in source_rels:
            kind = relation.get("Type").split("/")[-1]
            if kind in ("slideMaster", "notesMaster"):
                target = prefix + _target("ppt/presentation.xml", relation.get("Target"))
                rid = _append_rel(rels, kind, target)
                if kind == "slideMaster":
                    E.SubElement(masters, f"{{{P}}}sldMasterId", id=str(2147483648 + len(masters)), **{f"{{{R}}}id": rid})
                else:
                    E.SubElement(notes_masters, f"{{{P}}}notesMasterId", **{f"{{{R}}}id": rid})
        prs = Presentation(path)
        assert (prs.slide_width, prs.slide_height) == (12192000, 6858000)
        entry = {"name": label, "source_file": path.relative_to(root).as_posix(), "source_sha256": hashlib.sha256(path.read_bytes()).hexdigest(), "default_variant": "slide_002" if layout_id == "title" else "slide_001", "variants": {}}
        details = {"layout_id": layout_id, "source_file": entry["source_file"], "parts": {}, "variants": {}}
        for number, slide in enumerate(prs.slides, 1):
            global_slide += 1
            part = str(slide.part.partname).lstrip("/")
            layout_part = _related(original, part, "slideLayout")
            master_part = _related(original, layout_part, "slideMaster")
            fields = {}; layers = {}
            for scope, owner in (("slide", part), ("layout", layout_part), ("master", master_part)):
                f, objects = _fields_and_objects(prs, original, owner, scope, prefix)
                fields.update(f)
                layers[scope] = prefix + owner
                if prefix + owner not in details["parts"]:
                    details["parts"][prefix + owner] = {"scope": scope, "xml_sha256": hashlib.sha256(original[owner]).hexdigest(),
                        "drawingml": original[owner].decode("utf-8-sig"), "objects": objects}
            title_candidates = [f["sample_text"].strip().replace("\v", " ").replace("\n", " / ") for f in fields.values() if f["kind"] == "text" and f["scope"] == "slide" and not f["sample_text"].strip().isdigit()]
            source_xml = _xml(original[part])
            title_nodes = source_xml.xpath(".//p:sp[p:nvSpPr/p:nvPr/p:ph[@type='title' or @type='ctrTitle']]/p:txBody", namespaces=NS)
            title = " ".join(" ".join(_paragraph_text(p) for p in body.findall(f"{{{A}}}p")) for body in title_nodes).strip()
            if layout_id in ("title", "separators", "last_slide"):
                title = slide.slide_layout.name
            if not title:
                title = title_candidates[0] if title_candidates else f"{label} — вариант {number}"
            variant_id = f"slide_{number:03d}"
            variant = {"name": title[:160],
                "source_slide": number, "template_slide": global_slide, "package_part": prefix + part,
                "source_layout_part": prefix + layout_part, "source_master_part": prefix + master_part,
                "slide_xml_sha256": hashlib.sha256(original[part]).hexdigest(),
                "preview": f"output/severstal-source-previews/{path.stem}/slide-{number:02d}.png",
                "design_file": f"templates/severstal_designs/{layout_id}.json", "fields": fields}
            entry["variants"][variant_id] = variant
            details["variants"][variant_id] = layers
            rid = _append_rel(rels, "slide", prefix + part)
            E.SubElement(slides, f"{{{P}}}sldId", id=str(255 + global_slide), **{f"{{{R}}}id": rid})
        manifest["layouts"][layout_id] = entry
        (designs_dir / f"{layout_id}.json").write_text(json.dumps(details, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(f"{layout_id}: {len(prs.slides)} original variants")
    package["ppt/presentation.xml"] = _bytes(presentation)
    package[_relpath("ppt/presentation.xml")] = _bytes(rels)
    package["[Content_Types].xml"] = _bytes(ct)
    _prune_and_save(package, template)
    manifest["variant_count"] = global_slide
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    all_slides = {"template_version": "2.0", "title": "Северсталь — все исходные варианты", "slides": [{"layout_id": k, "variant_id": v, "updates": []} for k, l in manifest["layouts"].items() for v in l["variants"]]}
    (root / "examples/severstal_all_variants.json").write_text(json.dumps(all_slides, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    sample = {"template_version": "2.0", "title": "Северсталь — 15 типов", "slides": [{"layout_id": k, "variant_id": l["default_variant"], "updates": []} for k, l in manifest["layouts"].items()]}
    (root / "examples/severstal_15_types.json").write_text(json.dumps(sample, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"Prepared {len(manifest['layouts'])} types / {global_slide} original variants")


def deepcopy_element(node):
    return E.fromstring(E.tostring(node))


if __name__ == "__main__":
    prepare()
