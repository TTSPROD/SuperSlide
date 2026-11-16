"""A bounded local photo bank; the model sees descriptions, never file paths."""

from __future__ import annotations

import copy
import json
import re
import warnings
from dataclasses import dataclass, field
from pathlib import Path, PureWindowsPath
from typing import Any
from urllib.parse import urlsplit

from PIL import Image


MAX_IMAGE_BYTES = 15 * 1024 * 1024
MAX_IMAGE_PIXELS = 40_000_000
MAX_CATALOG_BYTES = 2 * 1024 * 1024
_ASSET_ID = re.compile(r"[A-Za-z0-9][A-Za-z0-9_.-]{0,119}\Z")


class AssetError(ValueError):
    """The selected local catalog or image is invalid."""


def _resolve_local_file(root: Path, relative_file: str, asset_id: str) -> Path:
    try:
        path = Path(relative_file)
        windows_path = PureWindowsPath(relative_file)
        if (not relative_file or urlsplit(relative_file).scheme or path.is_absolute()
                or windows_path.drive or windows_path.root
                or ".." in path.parts or ".." in windows_path.parts):
            raise AssetError(f"Ассет {asset_id}: нужен локальный относительный путь внутри каталога.")
        if path.suffix.lower() not in {".png", ".jpg", ".jpeg"}:
            raise AssetError(f"Ассет {asset_id}: поддерживаются только PNG и JPEG.")
        resolved = (root / path).resolve(strict=True)
        if not resolved.is_relative_to(root) or not resolved.is_file():
            raise AssetError(f"Ассет {asset_id}: файл должен находиться внутри каталога изображений.")
        return resolved
    except AssetError:
        raise
    except (OSError, RuntimeError, ValueError):
        raise AssetError(f"Ассет {asset_id}: изображение недоступно или путь некорректен.") from None


def _verify_image(path: Path, asset_id: str) -> tuple[int, int]:
    try:
        if not 0 < path.stat().st_size <= MAX_IMAGE_BYTES:
            raise AssetError(f"Ассет {asset_id}: изображение должно быть не больше 15 МиБ.")
        with warnings.catch_warnings():
            warnings.simplefilter("error", Image.DecompressionBombWarning)
            with Image.open(path) as image:
                if image.format not in {"PNG", "JPEG"}:
                    raise AssetError(f"Ассет {asset_id}: содержимое файла должно быть PNG или JPEG.")
                width, height = image.size
                if width <= 0 or height <= 0 or width * height > MAX_IMAGE_PIXELS:
                    raise AssetError(f"Ассет {asset_id}: изображение превышает 40 миллионов пикселей.")
                image.verify()
        return width, height
    except AssetError:
        raise
    except (OSError, ValueError, SyntaxError, Image.DecompressionBombError, Image.DecompressionBombWarning):
        raise AssetError(f"Ассет {asset_id}: изображение повреждено или не прошло проверку.") from None


@dataclass
class AssetCatalog:
    records: list[dict[str, Any]]
    _root: Path = field(repr=False)
    _files: dict[str, str] = field(repr=False)
    _expected_sizes: dict[str, tuple[int | None, int | None]] = field(repr=False)
    _verified: dict[str, tuple[Path, int, int]] = field(default_factory=dict, repr=False)
    _extra: dict[str, Path] = field(default_factory=dict, repr=False)

    def resolve(self, asset_id: str) -> Path:
        if not isinstance(asset_id, str) or (asset_id not in self._files and asset_id not in self._extra):
            raise AssetError("Неизвестный asset_id: выберите изображение из локального каталога.")
        path = self._extra.get(asset_id)
        if path is None:
            path = _resolve_local_file(self._root, self._files[asset_id], asset_id)
        try:
            stat = path.stat()
            signature = (path, stat.st_size, stat.st_mtime_ns)
            if self._verified.get(asset_id) != signature:
                actual = _verify_image(path, asset_id)
                expected = self._expected_sizes[asset_id]
                if any(value is not None and value != actual[index] for index, value in enumerate(expected)):
                    raise AssetError(f"Ассет {asset_id}: размеры изображения не совпадают с каталогом.")
                self._verified[asset_id] = signature
            return path
        except AssetError:
            raise
        except OSError:
            raise AssetError(f"Ассет {asset_id}: изображение недоступно.") from None

    def register(self, asset_id: str, path: Path, description: str, tags: list[str] | None = None) -> tuple[int, int]:
        """Register a verified image stored outside the catalog directory (local web uploads)."""
        if not isinstance(asset_id, str) or not _ASSET_ID.fullmatch(asset_id) \
                or asset_id in self._files or asset_id in self._extra:
            raise AssetError("Недопустимый или повторяющийся id изображения.")
        if not isinstance(description, str) or not description or len(description) > 2000:
            raise AssetError("description должен быть строкой до 2000 символов.")
        path = Path(path)
        width, height = _verify_image(path, asset_id)
        record = {"id": asset_id, "file": path.name, "description": description, "tags": list(tags or []),
                  "width": width, "height": height, "uploaded": True}
        self.records.append(record)
        self._extra[asset_id] = path
        self._expected_sizes[asset_id] = (width, height)
        self._verified[asset_id] = (path, path.stat().st_size, path.stat().st_mtime_ns)
        return width, height

    def prompt_catalog(self) -> list[dict[str, Any]]:
        return [{"id": record["id"], "description": record["description"], "tags": list(record["tags"])}
                for record in self.records]

    def to_dict(self) -> dict[str, Any]:
        return {"images": copy.deepcopy(self.records)}


def load_asset_catalog(path: str | Path) -> AssetCatalog:
    """Validate the explicitly selected catalog and its images, without network access."""
    try:
        catalog_path = Path(path).expanduser().resolve(strict=True)
        if not catalog_path.is_file() or catalog_path.stat().st_size > MAX_CATALOG_BYTES:
            raise AssetError("Локальный каталог изображений должен быть JSON-файлом не больше 2 МиБ.")
        raw = catalog_path.read_bytes()
        if len(raw) > MAX_CATALOG_BYTES:
            raise AssetError("Каталог изображений превышает 2 МиБ.")
        payload = json.loads(raw.decode("utf-8-sig"))
    except AssetError:
        raise
    except (OSError, UnicodeError, json.JSONDecodeError, RuntimeError, ValueError):
        raise AssetError("Не удалось прочитать локальный JSON-каталог изображений.") from None
    images = payload.get("images") if isinstance(payload, dict) else None
    if not isinstance(images, list) or len(images) > 500:
        raise AssetError("Каталог должен содержать массив images, не больше 500 записей.")
    records, files, expected_sizes = [], {}, {}
    for raw_record in images:
        if not isinstance(raw_record, dict):
            raise AssetError("Запись каталога изображения должна быть JSON-объектом.")
        asset_id = raw_record.get("id")
        if not isinstance(asset_id, str) or not _ASSET_ID.fullmatch(asset_id):
            raise AssetError("id изображения должен содержать до 120 латинских букв, цифр, точек, дефисов или подчёркиваний.")
        if asset_id in files:
            raise AssetError(f"В каталоге повторяется id изображения {asset_id}.")
        relative_file = raw_record.get("file")
        description = raw_record.get("description", "")
        tags = raw_record.get("tags", [])
        if not isinstance(relative_file, str) or len(relative_file) > 1000:
            raise AssetError(f"Ассет {asset_id}: нужен относительный строковый file.")
        if not isinstance(description, str) or len(description) > 2000:
            raise AssetError(f"Ассет {asset_id}: description должен быть строкой до 2000 символов.")
        if not isinstance(tags, list) or len(tags) > 30 or any(not isinstance(tag, str) or len(tag) > 100 for tag in tags):
            raise AssetError(f"Ассет {asset_id}: tags должен быть массивом коротких строк.")
        dimensions = (raw_record.get("width"), raw_record.get("height"))
        if any(value is not None and (isinstance(value, bool) or not isinstance(value, int) or value <= 0) for value in dimensions):
            raise AssetError(f"Ассет {asset_id}: width и height должны быть положительными целыми числами.")
        records.append({**copy.deepcopy(raw_record), "description": description, "tags": list(tags)})
        files[asset_id] = relative_file
        expected_sizes[asset_id] = dimensions
    catalog = AssetCatalog(records, catalog_path.parent, files, expected_sizes)
    # Eager verification makes inspection and generation fail before an API call.
    for asset_id in files:
        catalog.resolve(asset_id)
    return catalog
