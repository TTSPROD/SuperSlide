"""Extract the visually classified photos from the 15 source decks, without re-encoding.

Metadata is keyed by SHA-256 rather than unstable PowerPoint image numbers.
Any new/changed source media fails closed and needs a visual classification.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import posixpath
from collections import Counter, defaultdict
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

from lxml import etree as ET
from PIL import Image, ImageDraw, ImageFont, ImageOps


ROOT = Path(__file__).resolve().parents[1]
NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
SHAPES = {f"{{{NS['p']}}}{name}" for name in ("pic", "sp", "graphicFrame", "cxnSp", "grpSp")}


def write_json(path: Path, value: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def slide_order(archive: ZipFile) -> dict[str, int]:
    rels = ET.fromstring(archive.read("ppt/_rels/presentation.xml.rels"))
    targets = {r.get("Id"): posixpath.normpath(posixpath.join("ppt", r.get("Target"))) for r in rels}
    document = ET.fromstring(archive.read("ppt/presentation.xml"))
    return {
        targets[node.get(f"{{{NS['r']}}}id")]: index
        for index, node in enumerate(document.findall("p:sldIdLst/p:sldId", NS), 1)
    }


def shape_info(node: ET._Element) -> dict:
    current = node
    while current is not None:
        if current.tag in SHAPES:
            props = current.find(".//p:cNvPr", NS)
            return {
                "shape_id": props.get("id") if props is not None else None,
                "shape_name": props.get("name", "") if props is not None else "",
            }
        current = current.getparent()
    return {"shape_id": None, "shape_name": ""}


def inventory(source: Path) -> tuple[dict, list]:
    blobs = {}
    decks = []
    for deck in sorted(source.glob("*.pptx")):
        with ZipFile(deck) as archive:
            order = slide_order(archive)
            decks.append({
                "file": deck.name,
                "sha256": hashlib.sha256(deck.read_bytes()).hexdigest(),
                "slides": len(order),
            })
            by_part = {}
            for part in sorted(archive.namelist()):
                if not part.startswith("ppt/media/"):
                    continue
                data = archive.read(part)
                digest = hashlib.sha256(data).hexdigest()
                if digest not in blobs:
                    blobs[digest] = {"data": data, "sources": [], "usages": []}
                blobs[digest]["sources"].append({"file": deck.name, "part": part})
                by_part[part] = digest
            for rel_part in sorted(archive.namelist()):
                if not rel_part.endswith(".rels") or "/_rels/" not in rel_part:
                    continue
                directory, filename = rel_part.rsplit("/_rels/", 1)
                owner = directory + "/" + filename.removesuffix(".rels")
                if owner not in archive.namelist():
                    continue
                relationships = ET.fromstring(archive.read(rel_part))
                media_rels = {}
                for rel in relationships:
                    if rel.get("TargetMode") == "External":
                        continue
                    target = posixpath.normpath(posixpath.join(posixpath.dirname(owner), rel.get("Target")))
                    if target in by_part:
                        media_rels[rel.get("Id")] = (target, rel.get("Type").rsplit("/", 1)[-1])
                if not media_rels:
                    continue
                document = ET.fromstring(archive.read(owner))
                for rel_id, (target, rel_type) in media_rels.items():
                    matches = document.xpath("//*[@r:embed=$rid or @r:link=$rid]", namespaces=NS, rid=rel_id)
                    if not matches:
                        matches = [None]
                    for node in matches:
                        usage = {
                            "file": deck.name, "part": owner, "media_part": target,
                            "slide": order.get(owner), "relationship_id": rel_id,
                            "relationship_type": rel_type,
                        }
                        usage.update(shape_info(node) if node is not None else {"shape_id": None, "shape_name": ""})
                        if node is not None:
                            usage["reference_element"] = ET.QName(node).localname
                        blobs[by_part[target]]["usages"].append(usage)
    return blobs, decks


def previews(destination: Path, photos: list[dict], output: Path) -> list[str]:
    """Build diagnostic sheets separately from the runtime photo library."""
    output.mkdir(parents=True, exist_ok=True)
    files = []
    try:
        font = ImageFont.truetype("arial.ttf", 15)
    except OSError:
        font = ImageFont.load_default()
    for page in range((len(photos) + 23) // 24):
        subset = photos[page * 24:(page + 1) * 24]
        sheet = Image.new("RGB", (1440, ((len(subset) + 3) // 4) * 280), "#edf0f3")
        draw = ImageDraw.Draw(sheet)
        for i, photo in enumerate(subset):
            x, y = i % 4 * 360, i // 4 * 280
            with Image.open(destination / photo["file"]) as original:
                rgba = original.convert("RGBA")
                thumb = ImageOps.contain(rgba, (340, 215))
                sheet.paste(thumb, (x + 10, y + 5), thumb)
            label = photo["name"]
            # Two bounded lines; filenames remain searchable in the JSON catalog.
            if len(label) > 42:
                cut = label.rfind(" ", 0, 42)
                label = label[:cut] + "\n" + label[cut + 1:]
            draw.text((x + 10, y + 225), f"{page * 24 + i + 1}. {label}", font=font, fill="#002f6c")
        filename = f"contact_sheet_{page + 1:02d}.jpg"
        path = output / filename
        sheet.save(path, quality=88)
        files.append(path.relative_to(ROOT).as_posix() if path.is_relative_to(ROOT) else str(path))
    return files


def extract(source: Path, destination: Path, metadata_path: Path, diagnostics: Path | None = None) -> dict:
    diagnostics = (diagnostics or ROOT / "output/assets/photos/severstal").resolve()
    metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
    blobs, decks = inventory(source)
    photos_by_hash = {item["sha256"]: item for item in metadata["photos"]}
    originals_by_hash = {item["sha256"]: item for item in metadata["embedded_originals"]}
    excluded = {
        digest: category for category, hashes in metadata["excluded_sha256_by_category"].items()
        for digest in hashes
    }
    classified = set(photos_by_hash) | set(originals_by_hash) | set(excluded)
    unknown = set(blobs) - classified
    missing = set(photos_by_hash) - set(blobs)
    if unknown or missing:
        raise ValueError(f"Media classification is stale: {len(unknown)} unknown, {len(missing)} missing photos")
    destination.mkdir(parents=True, exist_ok=True)
    photos = []
    for item in metadata["photos"]:
        digest = item["sha256"]
        blob = blobs[digest]
        extension = Path(item["reference"]["part"]).suffix.lower()
        filename = item["id"] + extension
        target = destination / filename
        target.write_bytes(blob["data"])
        with Image.open(BytesIO(blob["data"])) as image:
            width, height = image.size
            image_format = image.format
            image.verify()
        assert hashlib.sha256(target.read_bytes()).hexdigest() == digest
        photos.append({
            **{key: value for key, value in item.items() if key != "reference"},
            "file": filename, "width": width, "height": height, "format": image_format,
            "bytes": len(blob["data"]), "aspect_ratio": round(width / height, 6),
            "sources": blob["sources"], "usages": blob["usages"],
        })
    originals = []
    for digest, item in originals_by_hash.items():
        if digest not in blobs:
            continue
        blob = blobs[digest]
        filename = "originals/" + item["id"] + Path(item["reference"]["part"]).suffix.lower()
        target = destination / filename
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(blob["data"])
        assert hashlib.sha256(target.read_bytes()).hexdigest() == digest
        originals.append({
            **{key: value for key, value in item.items() if key != "reference"},
            "file": filename, "format": "JPEG XR", "bytes": len(blob["data"]),
            "sources": blob["sources"], "usages": blob["usages"],
            "use": "Original Office image-effect source; use associated PNG photo for insertion.",
        })
    sheets = previews(destination, photos, diagnostics)
    summary = {
        "source_decks": len(decks), "unique_media": len(blobs), "photos": len(photos),
        "embedded_originals": len(originals), "excluded_non_photographic_media": sum(d in excluded for d in blobs),
        "photo_occurrences": sum(len(p["usages"]) for p in photos),
        "all_photos_byte_identical": True,
        "unclassified_media": 0,
    }
    catalog = {
        "schema_version": "1.0", "brand": "Северсталь", "source_directory": "resources/severstal/templates",
        "description": "Все фотографии из 15 исходных PPTX. Дедупликация только по SHA-256 исходных байтов; разные кадры и размеры сохранены отдельно.",
        "summary": summary, "source_decks": decks, "images": photos,
        "embedded_originals": originals, "contact_sheets": sheets,
    }
    write_json(destination / "catalog.json", catalog)
    write_json(diagnostics / "extraction_report.json", {
        "summary": summary,
        "classification_method": metadata["classification_method"],
        "excluded_counts": dict(Counter(excluded[d] for d in blobs if d in excluded)),
        "excluded_media": [
            {"sha256": d, "category": excluded[d], "sources": blobs[d]["sources"]}
            for d in sorted(blobs) if d in excluded
        ],
    })
    (destination / "README.md").write_text(
        "# Фотографии Северстали\n\n"
        f"Из 15 файлов `resources/severstal/templates` извлечены **{len(photos)} фотографий**. "
        "PNG и JPEG сохранены побайтно, без пересжатия, изменения размеров и кадрирования. "
        "Одинаковые исходные байты объединены; разные исходные кадры, размеры и обработки одного сюжета оставлены отдельно.\n\n"
        "`catalog.json` содержит русские названия, описания, теги, размеры, SHA-256 и все исходные использования "
        "(PPTX, номер слайда, XML-часть, ID и имя фигуры). Модель выбирает `images[].id`; путь к файлу указан в `file`. "
        "Изображения не подтверждают личности изображенных сотрудников; портреты подписаны по внешнему виду.\n\n"
        "Два исходника JPEG XR из слоя эффектов PowerPoint сохранены в `originals/` и связаны с соответствующими PNG "
        "через `embedded_originals[].photo_id`. Для обычной вставки используйте PNG из `images`.\n\n"
        "В `output/assets/photos/severstal/` создаются `contact_sheet_*.jpg` — визуальный каталог — и "
        "`extraction_report.json`, который перечисляет исключенные нефотографические "
        "media: иконки, логотипы, цифры, фирменную графику, рамки устройств и образцы диаграмм. "
        "Эти объекты остаются в первоисточниках; библиотека иконок находится в `assets/icons/severstal`.\n\n"
        "Повторное извлечение из корня проекта: `python scripts/extract_severstal_photos.py`. "
        "Семантическая классификация хранится в `scripts/data/severstal_photo_metadata.json`; "
        "неизвестный SHA-256 останавливает скрипт, чтобы новые фото не были молча исключены.\n",
        encoding="utf-8",
    )
    return summary


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=ROOT / "resources/severstal/templates")
    parser.add_argument("--output", type=Path, default=ROOT / "assets/photos/severstal")
    parser.add_argument("--metadata", type=Path, default=ROOT / "scripts/data/severstal_photo_metadata.json")
    parser.add_argument("--diagnostics-output", type=Path, default=ROOT / "output/assets/photos/severstal",
                        help="Contact sheets and extraction report, separate from versioned assets")
    args = parser.parse_args()
    print(json.dumps(extract(args.source, args.output, args.metadata, args.diagnostics_output), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
