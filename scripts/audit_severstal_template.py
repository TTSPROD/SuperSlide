"""Audit source fidelity, OPC relationships and extracted Severstal assets."""
from __future__ import annotations

import argparse
import hashlib
import json
import posixpath
from collections import Counter
from pathlib import Path
from urllib.parse import unquote, urlsplit
from zipfile import ZipFile

from lxml import etree as E


P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"p": P, "r": R}
EXPECTED = {
    "title": ("Title", 10), "contents": ("Contents", 5), "separators": ("Separators", 6),
    "big_text": ("BigText", 4), "bullets": ("Bullits", 17), "cards": ("Cards", 17),
    "tables": ("Tables", 20), "diagram": ("Diagram", 32), "schemes": ("Schemes", 16),
    "circle_schemes": ("CircleSchemes", 6), "team": ("Team", 5), "photo": ("Photo", 9),
    "quotes": ("Quotes", 24), "icons": ("Icons", 33), "last_slide": ("LastSlide", 7),
}


def xml(data):
    return E.fromstring(data, E.XMLParser(resolve_entities=False, no_network=True))


def sha(data):
    return hashlib.sha256(data).hexdigest()


def relpart(owner):
    return posixpath.join(posixpath.dirname(owner), "_rels", posixpath.basename(owner) + ".rels")


def target(owner, value):
    path = unquote(urlsplit(value).path)
    return path.lstrip("/") if path.startswith("/") else posixpath.normpath(posixpath.join(posixpath.dirname(owner), path))


def related(package, owner, kind):
    return next(target(owner, item.get("Target")) for item in xml(package[relpart(owner)])
                if item.get("Type", "").endswith("/" + kind))


def ordered_slides(package):
    relations = {item.get("Id"): item for item in xml(package[relpart("ppt/presentation.xml")])}
    return [target("ppt/presentation.xml", relations[item.get(f"{{{R}}}id")].get("Target"))
            for item in xml(package["ppt/presentation.xml"]).find(f"{{{P}}}sldIdLst")]


def master_design(data):
    root = xml(data)
    # This is the only permitted source-master metadata difference.
    for item in root.xpath("./p:sldLayoutIdLst/p:sldLayoutId", namespaces=NS):
        item.attrib.pop("id", None)
    return E.tostring(root, method="c14n")


def audit(root):
    errors = []
    checks = Counter()

    def require(condition, message, check=None):
        if not condition:
            errors.append(message)
        elif check:
            checks[check] += 1

    manifest_path = root / "templates/severstal_template.json"
    package_path = root / "templates/severstal_source.pptx"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8-sig"))
    with ZipFile(package_path) as archive:
        package = {part: archive.read(part) for part in archive.namelist()}
    source_dir = root / "SeverstalBrand/Template"
    expected_files = {f"Severstal_{stem}.pptx" for stem, _ in EXPECTED.values()}
    require({path.name for path in source_dir.glob("*.pptx")} == expected_files, "Папка источников должна содержать ровно 15 заданных PPTX")
    require(set(manifest["layouts"]) == set(EXPECTED), "Manifest должен содержать ровно 15 исходных типов")
    require(len(ordered_slides(package)) == 211, "В пакете должно быть ровно 211 исходных слайдов")
    require(sum(len(item["variants"]) for item in manifest["layouts"].values()) == 211, "Manifest должен содержать ровно 211 вариантов")
    originals = {}
    layout_reports = {}
    for layout_id, (stem, expected_count) in EXPECTED.items():
        source_path = source_dir / f"Severstal_{stem}.pptx"
        with ZipFile(source_path) as archive:
            source = {part: archive.read(part) for part in archive.namelist()}
        originals[source_path.name] = source
        slides = ordered_slides(source)
        layout = manifest["layouts"][layout_id]
        require(len(slides) == expected_count, f"{layout_id}: неожиданное число исходных слайдов")
        require(len(layout["variants"]) == expected_count, f"{layout_id}: потеряны или добавлены варианты")
        require(layout["source_sha256"] == sha(source_path.read_bytes()), f"{layout_id}: SHA-256 исходного PPTX изменился")
        before = len(errors)
        for number, source_slide in enumerate(slides, 1):
            variant_id = f"slide_{number:03}"
            variant = layout["variants"][variant_id]
            part = variant["package_part"]
            source_layout = related(source, source_slide, "slideLayout")
            source_master = related(source, source_layout, "slideMaster")
            require(package[part] == source[source_slide], f"{layout_id}/{variant_id}: исходная XML-разметка слайда изменена", "slide_xml_byte_equal")
            require(sha(package[part]) == variant["slide_xml_sha256"], f"{layout_id}/{variant_id}: SHA-256 slide XML не совпадает с manifest")
            require(package[variant["source_layout_part"]] == source[source_layout], f"{layout_id}/{variant_id}: XML исходного макета изменена", "layout_xml_byte_equal")
            require(master_design(package[variant["source_master_part"]]) == master_design(source[source_master]),
                    f"{layout_id}/{variant_id}: оформление master изменено за пределами sldLayoutId metadata", "master_design_equal")
        layout_reports[layout_id] = {"source_file": source_path.name, "variants": len(slides), "passed": len(errors) == before}

    # Validate every relationship in the serialized OPC package, including
    # chart workbooks, inherited pictures, layouts, masters, and OLE objects.
    for part, data in package.items():
        if not part.endswith(".rels"):
            continue
        owner = "" if part == "_rels/.rels" else posixpath.join(posixpath.dirname(posixpath.dirname(part)), posixpath.basename(part)[:-5])
        require(not owner or owner in package, f"{part}: отсутствует владелец relationships")
        relationships = list(xml(data))
        ids = [item.get("Id") for item in relationships]
        require(len(ids) == len(set(ids)), f"{part}: повторяющиеся relationship Id")
        for item in relationships:
            if item.get("TargetMode") == "External":
                checks["external_relationships"] += 1
                continue
            resolved = target(owner, item.get("Target"))
            require(resolved in package, f"{part}: не разрешается связь {item.get('Id')} → {resolved}", "internal_relationships_resolved")

    title = originals["Severstal_Title.pptx"]
    for part in ("ppt/theme/theme1.xml", "ppt/tableStyles.xml", "ppt/presProps.xml", "ppt/viewProps.xml"):
        require(package[part] == title[part], f"{part}: глобальное оформление отличается от исходного Title", "global_parts_byte_equal")
    presentation = xml(package["ppt/presentation.xml"])
    presentation_rels = list(xml(package[relpart("ppt/presentation.xml")]))
    by_id = {item.get("Id"): item for item in presentation_rels}
    global_ids = []
    for item in presentation.find(f"{{{P}}}sldMasterIdLst"):
        global_ids.append(item.get("id"))
        master_part = target("ppt/presentation.xml", by_id[item.get(f"{{{R}}}id")].get("Target"))
        master = xml(package[master_part])
        global_ids.extend(master.xpath("./p:sldLayoutIdLst/p:sldLayoutId/@id", namespaces=NS))
    require(len(global_ids) == len(set(global_ids)), "Master/layout metadata IDs должны быть глобально уникальны")
    checks["globally_unique_master_layout_ids"] = len(set(global_ids))
    notes_ids = presentation.find(f"{{{P}}}notesMasterIdLst")
    notes_rels = [item for item in presentation_rels if item.get("Type").endswith("/notesMaster")]
    require(notes_ids is not None and len(notes_ids) == 1 and len(notes_rels) == 1, "В корневой презентации должен быть ровно один notesMaster")
    canonical_notes = target("ppt/presentation.xml", notes_rels[0].get("Target"))
    for part, data in package.items():
        if not part.endswith(".rels") or part == relpart("ppt/presentation.xml"):
            continue
        owner = posixpath.join(posixpath.dirname(posixpath.dirname(part)), posixpath.basename(part)[:-5])
        for item in xml(data):
            if item.get("Type", "").endswith("/notesMaster"):
                require(target(owner, item.get("Target")) == canonical_notes, f"{part}: ссылка notesMaster не каноническая", "canonical_notes_master_references")

    def asset_path(catalog_path, record, key="file"):
        filename = record.get(key, record.get("path"))
        for candidate in (catalog_path.parent / filename, root / filename):
            if candidate.is_file():
                return candidate
        return catalog_path.parent / filename

    photo_path = root / "assets/photos/severstal/catalog.json"
    photos = json.loads(photo_path.read_text(encoding="utf-8-sig"))
    require(len(photos["images"]) == 99, "В каталоге должно быть ровно 99 фотографий")
    require(len({item["id"] for item in photos["images"]}) == 99, "id фотографий должны быть уникальны")
    for record in photos["images"] + photos.get("embedded_originals", []):
        path = asset_path(photo_path, record)
        require(path.is_file(), f"Нет фотографии: {path}")
        if path.is_file():
            require(sha(path.read_bytes()) == record["sha256"], f"Фото {record['id']}: SHA-256 не совпадает", "photo_files_sha256_equal")
        for provenance in record.get("sources", []):
            require(sha(originals[provenance["file"]][provenance["part"]]) == record["sha256"],
                    f"Фото {record['id']}: отличается от исходных байтов", "photo_source_bytes_equal")

    icon_path = root / "assets/icons/severstal/catalog.json"
    icons = json.loads(icon_path.read_text(encoding="utf-8-sig"))
    require(len(icons["groups"]) == 30, "В каталоге должно быть ровно 30 исходных групп иконок")
    require(len(icons["icons"]) == 1364, "В каталоге должно быть ровно 1364 исходных исполнения иконок")
    icon_ids = {item["id"] for item in icons["icons"]}
    require(len(icon_ids) == len(icons["icons"]), "id иконок должны быть уникальны")
    grouped = [icon_id for group in icons["groups"] for icon_id in group["icon_ids"]]
    require(set(grouped) == icon_ids and len(grouped) == len(icon_ids), "Каждая иконка должна относиться ровно к одной исходной группе")
    require({slide for group in icons["groups"] for slide in group["source_slides"]} == set(range(1, 34)), "Группы должны покрывать все 33 страницы библиотеки")
    for record in icons["icons"]:
        path = asset_path(icon_path, record)
        require(path.is_file(), f"Нет SVG иконки: {path}")
        if path.is_file():
            require(sha(path.read_bytes()) == record["sha256"], f"Иконка {record['id']}: SHA-256 не совпадает", "icon_svg_sha256_equal")
        require(asset_path(icon_path, record, "png_file").is_file(), f"Нет PNG fallback: {record['id']}", "icon_png_fallback_present")
    return {"status": "passed" if not errors else "failed", "template_version": manifest["template_version"],
            "template": package_path.relative_to(root).as_posix(), "template_sha256": sha(package_path.read_bytes()),
            "layout_count": len(manifest["layouts"]), "variant_count": sum(len(v["variants"]) for v in manifest["layouts"].values()),
            "package_parts": len(package), "photo_count": len(photos["images"]), "icon_groups": len(icons["groups"]),
            "icon_count": len(icons["icons"]), "selectable_icons": sum(item.get("selectable", True) for item in icons["icons"]),
            "checks": dict(checks), "layouts": layout_reports, "errors": errors}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    root = args.root.resolve()
    output = args.output or root / "output/severstal-audit.json"
    try:
        report = audit(root)
    except (OSError, ValueError, KeyError, StopIteration, E.XMLSyntaxError) as exc:
        report = {"status": "failed", "errors": [f"{type(exc).__name__}: {exc}"]}
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    compact = {key: report[key] for key in ("status", "layout_count", "variant_count", "photo_count", "icon_groups", "icon_count") if key in report}
    compact["error_count"] = len(report["errors"])
    compact["report"] = str(output.resolve())
    print(json.dumps(compact, ensure_ascii=False))
    if report["errors"]:
        print(json.dumps(report["errors"][:10], ensure_ascii=False))
    return 0 if report["status"] == "passed" else 1


if __name__ == "__main__":
    raise SystemExit(main())
