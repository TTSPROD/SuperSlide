"""Check template structure, original artwork, and python-pptx layout round trips."""

from __future__ import annotations

import argparse
from copy import deepcopy
import hashlib
from io import BytesIO
import json
from pathlib import Path
import posixpath
import re
from zipfile import ZipFile

from lxml import etree as ET
from PIL import Image, ImageChops, ImageStat
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE


ROOT = Path(__file__).resolve().parents[1]
NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
EXPECTED = {"title_slide", "bullet_list", "two_columns", "quote", "table", "bar_chart", "section_divider", "photo_text", "image_spotlight", "gallery", "kpi_cards", "timeline", "comparison", "closing"}
# Version 1.1 defaults are intentionally immutable. Hashes cover every field,
# its geometry and text limits, plus footer/background settings; variants are
# checked separately and can never silently replace an existing composition.
DEFAULT_BASELINES = {
    "title_slide": "445f762b587dd94a96369818fac4d541eaf00ca9d94f417c85de8979242083e9",
    "bullet_list": "6770a3ee2d1244416cfb3d515193cdba819825758042b981e3012bf4b34f80a8",
    "two_columns": "93d3599344db9c8b8422b545a796a3d6a9705e439dd34fe82840112e4b36db20",
    "quote": "9ac41a5ff8ce43a8c2291c183bd9df4deb3487759bea4cf52513b4c2d05e81c7",
    "table": "1de9ec8f2979d1247f42540a948a8aed61f022588a5589950efee2e69fcf6e48",
    "bar_chart": "c50cf8aea87f6f88a52dfadf1d82b55874d6ab9a143e9c9eff64dac3a10071fe",
    "section_divider": "b31cebe8840bb1383660707bab7e692b3b8f27c3a685b60b42574546c0ca9ce9",
    "photo_text": "0645c6465e2b4a98a33f300d33446772970e82b21b70f66fb1395a6e17749d47",
    "image_spotlight": "14660dd0c68b45b4d0166e1e2f20b539879e0411933a04c3028934f41ed5d155",
    "gallery": "2aafcf24ec6e4e3da90072e53e1f23d4496e38d0e4d649c4695edb5d40830ed3",
    "kpi_cards": "b910a93cdd0ab2ec322d3adc00194b22e0af17ed5d9b633d076f438af523de2e",
    "timeline": "c3d7e262cfefc7e7779e366447598a847bb4f4f24633eb48f8768b74f474004a",
    "comparison": "e87a37280fa4c9ebf0fd9d3ca39f696010fe134fe1159b8a600385aaff4b58bd",
    "closing": "a586133a1b0bd6534b7725fe8572673acbf04683be167c22f93b4434ade323eb",
}


def design_specs(manifest: dict):
    """Yield semantic type, variant ID and physical layout, including defaults."""
    for semantic_id, default in manifest["layouts"].items():
        yield semantic_id, "default", default
        for variant_id, variant in default.get("variants", {}).items():
            assert variant_id != "default", "Default design must not be duplicated"
            assert re.fullmatch(r"[a-z][a-z0-9_]{0,63}", variant_id), variant_id
            assert not variant.get("variants"), "Variants cannot contain other variants"
            yield semantic_id, variant_id, variant


def _allowed_package_part(name: str) -> bool:
    """The empty template needs no embedded files, code, OLE or external links."""
    return name in {
        "[Content_Types].xml", "_rels/.rels", "docProps/app.xml",
        "docProps/core.xml", "ppt/presentation.xml", "ppt/_rels/presentation.xml.rels",
        "ppt/slideMasters/slideMaster1.xml", "ppt/slideMasters/_rels/slideMaster1.xml.rels",
        "ppt/theme/theme1.xml",
    } or bool(re.fullmatch(r"ppt/slideLayouts/(?:_rels/)?slideLayout[1-9][0-9]*\.xml(?:\.rels)?", name)) or bool(
        re.fullmatch(r"ppt/media/[^/]+\.(?:png|jpe?g|svg)", name, re.IGNORECASE)
    )


def _check_relationships(package: ZipFile, names: set[str]) -> None:
    office = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
    allowed = {
        office + suffix for suffix in (
            "officeDocument", "extended-properties", "slideMaster", "slideLayout", "theme", "image"
        )
    } | {"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"}
    for name in names:
        if not name.endswith(".rels"):
            continue
        owner_dir = "" if name == "_rels/.rels" else posixpath.dirname(posixpath.dirname(name))
        relationships = ET.fromstring(package.read(name))
        identifiers = [item.get("Id") for item in relationships]
        assert all(identifiers) and len(identifiers) == len(set(identifiers)), f"Duplicate relationship IDs: {name}"
        for relation in relationships:
            assert relation.get("Type") in allowed, f"Unexpected relationship type: {name}"
            assert relation.get("TargetMode") != "External", f"External relation: {name}"
            raw_target = relation.get("Target", "")
            assert raw_target and not raw_target.startswith(("/", "\\\\")) and ":" not in raw_target
            target = posixpath.normpath(posixpath.join(owner_dir, raw_target))
            assert target in names, f"Unresolved target: {name} -> {target}"


def _artwork_bytes(shape) -> bytes:
    """Shape IDs change when original groups are placed in a new layout."""
    clone = deepcopy(shape)
    for properties in clone.xpath(".//p:cNvPr", namespaces=NS):
        properties.attrib.pop("id", None)
    return ET.tostring(clone, method="c14n", exclusive=True)


def _check_variant_artwork(package: ZipFile, original: ZipFile, designs: list[tuple]) -> int:
    copied_shapes = 0
    by_name = {}
    for part in package.namelist():
        if re.fullmatch(r"ppt/slideLayouts/slideLayout[1-9][0-9]*\.xml", part):
            layout = ET.fromstring(package.read(part))
            by_name[layout.find("p:cSld", NS).get("name")] = layout
    for _, variant, spec in designs:
        if variant == "default":
            continue
        style = spec.get("artwork_style", "plain")
        if style not in {"original", "photo_cover", "quote_blue", "closing_minimal"}:
            continue
        source = ET.fromstring(original.read(spec["source_reference"]["source_layout_part"]))
        source_shapes = [node for node in source.find("p:cSld/p:spTree", NS)[2:] if not node.xpath(".//p:ph", namespaces=NS)]
        if style == "closing_minimal":
            source_shapes = [node for node in source_shapes if node.xpath(".//p:cNvPr/@name", namespaces=NS)[:1] == ["Рисунок 3"]]
        assert source_shapes, f"No original artwork found for {spec['layout_name']}"
        target_shapes = by_name[spec["layout_name"]].find("p:cSld/p:spTree", NS)[2:]
        target_bytes = [_artwork_bytes(node) for node in target_shapes if not node.xpath(".//p:ph", namespaces=NS)]
        for shape in source_shapes:
            assert _artwork_bytes(shape) in target_bytes, f"Original variant artwork changed: {spec['layout_name']}"
            copied_shapes += 1
    return copied_shapes


def audit(template: Path, manifest_path: Path, source_path: Path) -> dict:
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    assert set(manifest["layouts"]) == EXPECTED
    assert manifest["font_family"] == "Arial"
    assert manifest["template_file"] == template.name
    assert manifest["template_version"] == "1.2"
    assert manifest["compatible_versions"] == ["1.0", "1.1", "1.2"]
    designs = list(design_specs(manifest))
    assert len(designs) == 31, "Expected 14 original designs and 17 additional variants"
    names_in_manifest = [spec["layout_name"] for _, _, spec in designs]
    assert len(set(names_in_manifest)) == len(designs), "Duplicate physical layout name"
    assert len({(semantic, variant) for semantic, variant, _ in designs}) == len(designs)
    for semantic, variant, spec in designs:
        if variant == "default":
            assert spec["layout_name"] == semantic
            baseline = {key: spec[key] for key in ("layout_name", "fields", "background_color", "show_footer")}
            digest = hashlib.sha256(json.dumps(baseline, sort_keys=True, separators=(",", ":")).encode()).hexdigest()
            assert digest == DEFAULT_BASELINES[semantic], f"Existing default design changed: {semantic}"
            continue
        assert spec["layout_name"] == f"{semantic}__{variant}"
        assert isinstance(spec.get("description"), str) and spec["description"].strip()
        source = spec["source_reference"]
        assert source["brandbook_pages"] and all(isinstance(page, int) and 1 <= page <= 172 for page in source["brandbook_pages"])
        assert source["source_slides"] and all(isinstance(slide, int) and 1 <= slide <= 253 for slide in source["source_slides"])
    catalog_dir = ROOT / "assets" / "images"
    catalog = json.loads((catalog_dir / "catalog.json").read_text(encoding="utf-8"))
    assert 4 <= len(catalog["images"]) <= 6
    assert len({item["id"] for item in catalog["images"]}) == len(catalog["images"])
    source_hash = hashlib.sha256(source_path.read_bytes()).hexdigest()
    with ZipFile(template) as package, ZipFile(source_path) as original:
        names = set(package.namelist())
        assert len(names) == len(package.namelist()), "Duplicate package entries"
        assert all(_allowed_package_part(name) for name in names), "Unexpected package parts"
        assert not any(name.startswith(("ppt/slides/", "ppt/charts/", "ppt/embeddings/", "ppt/notesSlides/", "ppt/diagrams/")) for name in names)
        media_names = {name for name in names if name.startswith("ppt/media/")}
        assert {"ppt/media/image1.png", "ppt/media/image2.svg"} <= media_names
        original_media_hashes = {
            hashlib.sha256(original.read(name)).hexdigest()
            for name in original.namelist() if name.startswith("ppt/media/")
        }
        for name in media_names:
            assert hashlib.sha256(package.read(name)).hexdigest() in original_media_hashes, f"Media is not an unchanged original asset: {name}"
        for name in ("ppt/media/image1.png", "ppt/media/image2.svg", "ppt/theme/theme1.xml"):
            assert package.read(name) == original.read(name), f"Original asset changed: {name}"
        for _, variant, spec in designs:
            if variant != "default":
                source_part = spec["source_reference"].get("source_layout_part")
                if source_part:
                    assert source_part in original.namelist(), f"Missing source layout: {source_part}"
                    assert re.fullmatch(r"ppt/slideLayouts/slideLayout[1-9][0-9]*\.xml", source_part)
        for asset in catalog["images"]:
            assert set(asset) == {"id", "file", "description", "tags", "source_slide", "source_media", "width", "height"}
            path = (catalog_dir / asset["file"]).resolve()
            assert path.is_relative_to(catalog_dir.resolve())
            assert path.stat().st_size < 2_000_000
            original_bytes = original.read(asset["source_media"])
            if asset["id"] == "steelworks_hall":
                # This source JPEG exceeds 2 MB. It is recompressed at full
                # resolution; verify that content is unchanged within JPEG loss.
                with Image.open(BytesIO(original_bytes)) as source_photo, Image.open(path) as prepared_photo:
                    assert source_photo.size == prepared_photo.size
                    difference = ImageStat.Stat(ImageChops.difference(source_photo.convert("RGB"), prepared_photo.convert("RGB")))
                    assert max(difference.mean) < 4.0, "Recompressed photo differs from the source"
            else:
                assert path.read_bytes() == original_bytes, f"Extracted photo changed: {asset['id']}"
            with Image.open(path) as photo:
                assert photo.size == (asset["width"], asset["height"])
                photo.verify()
            source_rels = ET.fromstring(original.read(f"ppt/slides/_rels/slide{asset['source_slide']}.xml.rels"))
            targets = {posixpath.normpath(posixpath.join("ppt/slides", item.get("Target"))) for item in source_rels}
            assert asset["source_media"] in targets, f"Photo is not used on stated source slide: {asset['id']}"
        _check_relationships(package, names)
        referenced_media = set()
        for name in names:
            if not name.endswith(".xml"):
                continue
            xml = ET.fromstring(package.read(name))
            assert not xml.xpath("//p:oleObj | //a:videoFile | //a:audioFile", namespaces=NS), f"Unexpected embedded object: {name}"
            rels_name = posixpath.join(posixpath.dirname(name), "_rels", posixpath.basename(name) + ".rels")
            relations = ET.fromstring(package.read(rels_name)) if rels_name in names else []
            identifiers = {item.get("Id") for item in relations}
            used_ids = {value for node in xml.iter() for key, value in node.attrib.items() if key.startswith("{" + NS["r"] + "}")}
            assert used_ids <= identifiers, f"Unresolved relationship ID in XML: {name}: {used_ids - identifiers}"
            for item in relations:
                target = posixpath.normpath(posixpath.join(posixpath.dirname(name), item.get("Target")))
                if target.startswith("ppt/media/") and item.get("Id") in used_ids:
                    referenced_media.add(target)
        assert media_names == referenced_media, "Unreferenced media in generator template"
        old_tree = ET.fromstring(original.read("ppt/slideMasters/slideMaster1.xml")).find("p:cSld/p:spTree", NS)
        new_tree = ET.fromstring(package.read("ppt/slideMasters/slideMaster1.xml")).find("p:cSld/p:spTree", NS)
        # Every original master shape, including the logo, line, and disclaimer,
        # is preserved. Only its layout list and package dependencies are changed.
        assert len(old_tree) == len(new_tree)
        for old, new in zip(old_tree, new_tree):
            assert ET.tostring(old, method="c14n", exclusive=True) == ET.tostring(new, method="c14n", exclusive=True)
        title = ET.fromstring(package.read("ppt/slideLayouts/slideLayout1.xml"))
        original_title = ET.fromstring(original.read("ppt/slideLayouts/slideLayout20.xml"))
        old_art = [node for node in original_title.find("p:cSld/p:spTree", NS)[2:] if not node.xpath(".//p:ph", namespaces=NS)]
        new_art = [node for node in title.find("p:cSld/p:spTree", NS)[2:] if not node.xpath(".//p:ph", namespaces=NS)]
        assert len(old_art) == len(new_art)
        for old, new in zip(old_art, new_art):
            assert ET.tostring(old, method="c14n", exclusive=True) == ET.tostring(new, method="c14n", exclusive=True)
        blue_source = ET.fromstring(original.read("ppt/slideLayouts/slideLayout48.xml"))
        blue_art = [node for node in blue_source.find("p:cSld/p:spTree", NS)[2:] if not node.xpath(".//p:ph", namespaces=NS)]
        # Both blue separator and closing retain the genuine original white logo
        # variant and original blue-background supergraphic, without recoloring.
        for index in (7, 14):
            layout = ET.fromstring(package.read(f"ppt/slideLayouts/slideLayout{index}.xml"))
            art = [node for node in layout.find("p:cSld/p:spTree", NS)[2:] if not node.xpath(".//p:ph", namespaces=NS)]
            assert len(art) == len(blue_art)
            for old, new in zip(blue_art, art):
                assert ET.tostring(old, method="c14n", exclusive=True) == ET.tostring(new, method="c14n", exclusive=True)
        white_logo = next(item for item in blue_art if item.tag == "{" + NS["p"] + "}grpSp")
        for index in (9, 11):
            layout = ET.fromstring(package.read(f"ppt/slideLayouts/slideLayout{index}.xml"))
            logo = next(item for item in layout.find("p:cSld/p:spTree", NS) if item.tag == "{" + NS["p"] + "}grpSp")
            assert len(logo) == len(white_logo)
            for old, new in zip(white_logo, logo):
                # Repositioning changes only the outer group transform. All
                # vector logo paths, fills and inner coordinates stay intact.
                if old.tag.endswith("grpSpPr"):
                    continue
                assert ET.tostring(old, method="c14n", exclusive=True) == ET.tostring(new, method="c14n", exclusive=True)
        fixed_expectations = {10: 3, 11: 3, 12: 5, 13: 4}
        for index, expected_count in fixed_expectations.items():
            layout = ET.fromstring(package.read(f"ppt/slideLayouts/slideLayout{index}.xml"))
            static_names = layout.xpath('//p:cNvPr[starts-with(@name, "Fixed: ")]/@name', namespaces=NS)
            assert len(static_names) == expected_count, (index, static_names)
        for name in names:
            if name.startswith("ppt/slideLayouts/") and not name.endswith(".rels"):
                layout = ET.fromstring(package.read(name))
                shape_ids = layout.xpath("//p:cNvPr/@id", namespaces=NS)
                assert len(set(shape_ids)) == len(shape_ids), f"Duplicate shape IDs: {name}"
        verified_variant_shapes = _check_variant_artwork(package, original, designs)

    presentation = Presentation(template)
    assert len(presentation.slides) == 0
    assert len(presentation.slide_masters) == 1
    assert len(presentation.slide_layouts) == len(designs)
    assert (presentation.slide_width, presentation.slide_height) == (12192000, 6858000)
    layout_by_name = {item.name: item for item in presentation.slide_layouts}
    assert set(layout_by_name) == set(names_in_manifest)
    expected_tables = expected_charts = expected_pictures = 0
    for name, variant, layout_spec in designs:
        design_id = f"{name}/{variant}"
        layout = layout_by_name[layout_spec["layout_name"]]
        expected_indices = {spec["idx"] for spec in layout_spec["fields"].values()}
        assert len(expected_indices) == len(layout_spec["fields"]), f"Duplicate placeholder idx: {design_id}"
        assert {ph.placeholder_format.idx for ph in layout.placeholders} == expected_indices
        slide = presentation.slides.add_slide(layout)
        assert {ph.placeholder_format.idx for ph in slide.placeholders} == expected_indices
        for field_name, spec in layout_spec["fields"].items():
            ph = slide.placeholders[spec["idx"]]
            assert spec["type"] in {"text", "picture", "content"}
            assert (ph.left, ph.top, ph.width, ph.height) == tuple(spec[key] for key in ("x", "y", "width", "height"))
            assert ph.left >= 0 and ph.top >= 0 and ph.width > 0 and ph.height > 0
            assert ph.left + ph.width <= presentation.slide_width
            assert ph.top + ph.height <= presentation.slide_height
            if layout_spec["show_footer"]:
                assert ph.top + ph.height < 6489340, "Content overlaps the original footer rule"
            if spec["type"] == "text":
                assert ph.has_text_frame
                assert ph.placeholder_format.type in {1, 2}  # TITLE, BODY
                ph.text = f"{design_id}: {field_name}"
                assert spec["font_size_pt"] >= spec["min_font_size_pt"] > 0
                assert spec["max_lines"] > 0
                if variant == "default" and name not in {"title_slide", "bullet_list", "two_columns", "quote", "table", "bar_chart"}:
                    minimum_line_boxes = spec["max_lines"] * spec["min_font_size_pt"] * 1.12 * 1.2
                    assert minimum_line_boxes <= spec["height"] / 12700, f"Minimum text line boxes exceed new field height: {name}.{field_name}"
            elif spec["type"] == "picture":
                assert ph.placeholder_format.type == 18  # PP_PLACEHOLDER.PICTURE
                picture = ph.insert_picture(str(catalog_dir / catalog["images"][0]["file"]))
                assert (picture.left, picture.top, picture.width, picture.height) == tuple(spec[key] for key in ("x", "y", "width", "height"))
                expected_pictures += 1
            else:
                assert ph.placeholder_format.type == 7  # PP_PLACEHOLDER.OBJECT
                box = (ph.left, ph.top, ph.width, ph.height)
                if field_name == "table":
                    table = slide.shapes.add_table(3, 2, *box).table
                    table.cell(0, 0).text = "Test"
                    expected_tables += 1
                elif field_name == "chart":
                    data = CategoryChartData()
                    data.categories = ["A", "B", "C"]
                    data.add_series("Test", (1, 2, 3))
                    slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, *box, data)
                    expected_charts += 1
                else:
                    raise AssertionError(f"Unknown content field: {design_id}.{field_name}")
                ph._element.getparent().remove(ph._element)
    result = BytesIO()
    presentation.save(result)
    result.seek(0)
    reloaded = Presentation(result)
    assert len(reloaded.slides) == len(designs)
    assert sum(shape.has_table for slide in reloaded.slides for shape in slide.shapes) == expected_tables
    assert sum(shape.has_chart for slide in reloaded.slides for shape in slide.shapes) == expected_charts
    assert sum(shape.shape_type == 14 and shape.placeholder_format.type == 18 for slide in reloaded.slides for shape in slide.shapes) == expected_pictures
    assert {key for key in manifest["layouts"]["gallery"]["fields"] if key.startswith("image_")} == {"image_1", "image_2", "image_3"}
    assert {key for key in manifest["layouts"]["kpi_cards"]["fields"] if key.startswith("value_")} == {"value_1", "value_2", "value_3"}
    assert {key for key in manifest["layouts"]["timeline"]["fields"] if key.startswith("step_title_")} == {"step_title_1", "step_title_2", "step_title_3", "step_title_4"}
    spotlight = manifest["layouts"]["image_spotlight"]["fields"]
    assert spotlight["image"]["y"] + spotlight["image"]["height"] <= spotlight["title"]["y"]
    return {"ok": True, "template_version": manifest["template_version"], "slides_in_template": 0, "masters": 1, "semantic_types": len(manifest["layouts"]), "design_count": len(designs), "additional_variants": len(designs) - len(manifest["layouts"]), "layouts": list(layout_by_name), "template_bytes": template.stat().st_size, "native_table_chart_and_pictures_roundtrip": True, "roundtrip_tables": expected_tables, "roundtrip_charts": expected_charts, "roundtrip_pictures": expected_pictures, "original_master_and_title_art_preserved": True, "original_white_logo_preserved": True, "verified_original_variant_shapes": verified_variant_shapes, "verified_original_media": len(media_names), "verified_original_photos": len(catalog["images"]), "source_sha256": source_hash}


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--template", type=Path, default=ROOT / "templates" / "corporate_base.pptx")
    parser.add_argument("--manifest", type=Path, default=ROOT / "templates" / "template_manifest.json")
    parser.add_argument("--source", type=Path, default=ROOT / "SeverstalBrand" / "severstal-template-v29.pptx")
    args = parser.parse_args()
    print(json.dumps(audit(args.template, args.manifest, args.source), indent=2))


if __name__ == "__main__":
    main()
