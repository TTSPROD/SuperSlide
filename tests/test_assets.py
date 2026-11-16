import json
import struct
import tempfile
import unittest
import zlib
from pathlib import Path
from unittest.mock import patch

from PIL import Image

from superslide.assets import AssetError, MAX_IMAGE_BYTES, load_asset_catalog


class AssetCatalogTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory(dir=Path(__file__).resolve().parents[1])
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.image_path = self.root / "photo.png"
        Image.new("RGB", (32, 20), "blue").save(self.image_path)
        self.catalog_path = self.root / "catalog.json"
        self.record = {
            "id": "photo_01", "file": "photo.png", "description": "Промышленная площадка",
            "tags": ["производство", "команда"], "source_slide": 12,
            "source_media": "image1.png", "width": 32, "height": 20,
        }

    def write_catalog(self, records=None):
        self.catalog_path.write_text(json.dumps({"images": records if records is not None else [self.record]},
                                               ensure_ascii=False), encoding="utf-8-sig")

    def test_local_image_resolution_and_prompt_only_metadata(self):
        self.write_catalog()
        catalog = load_asset_catalog(self.catalog_path)
        self.assertEqual(catalog.resolve("photo_01"), self.image_path.resolve())
        self.assertEqual(catalog.records[0]["source_slide"], 12)
        self.assertEqual(catalog.prompt_catalog(), [{
            "id": "photo_01", "description": "Промышленная площадка", "tags": ["производство", "команда"],
        }])
        self.assertNotIn("file", catalog.prompt_catalog()[0])
        snapshot = catalog.to_dict()
        snapshot["images"][0]["description"] = "changed"
        self.assertEqual(catalog.records[0]["description"], self.record["description"])

    def test_verification_is_cached_in_instance_and_changed_file_is_rechecked(self):
        import superslide.assets as assets
        self.write_catalog()
        original_verify = assets._verify_image
        with patch("superslide.assets._verify_image", wraps=original_verify) as verify:
            catalog = load_asset_catalog(self.catalog_path)
            catalog.resolve("photo_01")
            catalog.resolve("photo_01")
            self.assertEqual(verify.call_count, 1)
            Image.new("RGB", (50, 30), "green").save(self.image_path)
            with self.assertRaisesRegex(AssetError, "размеры"):
                catalog.resolve("photo_01")
            self.assertEqual(verify.call_count, 2)

    def test_unknown_ids_duplicates_and_empty_catalog(self):
        self.write_catalog()
        catalog = load_asset_catalog(self.catalog_path)
        for invalid in ("unknown", "../photo.png", None, []):
            with self.subTest(asset_id=invalid), self.assertRaises(AssetError):
                catalog.resolve(invalid)
        self.write_catalog([self.record, dict(self.record)])
        with self.assertRaisesRegex(AssetError, "повторяется"):
            load_asset_catalog(self.catalog_path)
        self.write_catalog([])
        empty = load_asset_catalog(self.catalog_path)
        self.assertEqual(empty.prompt_catalog(), [])

    def test_register_adds_verified_external_image(self):
        self.write_catalog()
        catalog = load_asset_catalog(self.catalog_path)
        external = self.root / "outside" / "user photo.png"
        external.parent.mkdir()
        Image.new("RGB", (64, 48), "green").save(external)
        dimensions = catalog.register("user_photo", external, "Фото пользователя", ["своё"])
        self.assertEqual(dimensions, (64, 48))
        self.assertEqual(catalog.resolve("user_photo"), external.resolve())
        self.assertEqual(catalog.prompt_catalog()[-1],
                         {"id": "user_photo", "description": "Фото пользователя", "tags": ["своё"]})
        with self.assertRaisesRegex(AssetError, "повторяющийся"):
            catalog.register("photo_01", external, "Повтор")
        broken = self.root / "outside" / "broken.png"
        broken.write_bytes(b"not an image")
        with self.assertRaisesRegex(AssetError, "повреждено"):
            catalog.register("broken_id", broken, "Битый файл")
        self.assertEqual(len(catalog.records), 2)

    def test_remote_absolute_and_parent_paths_are_rejected(self):
        invalid_paths = ["https://remote.invalid/photo.png", "file:///tmp/photo.png", "../photo.png",
                         "sub/../photo.png", "..\\photo.png", str(self.image_path.resolve()),
                         "C:\\photos\\photo.png", "\\\\server\\share\\photo.png"]
        for invalid in invalid_paths:
            with self.subTest(path=invalid):
                self.write_catalog([{**self.record, "file": invalid}])
                with self.assertRaises(AssetError):
                    load_asset_catalog(self.catalog_path)

    def test_missing_corrupt_or_non_png_jpeg_images_are_rejected(self):
        self.write_catalog([{**self.record, "file": "missing.jpg"}])
        with self.assertRaises(AssetError):
            load_asset_catalog(self.catalog_path)
        self.image_path.write_bytes(b"not an image")
        self.write_catalog()
        with self.assertRaises(AssetError):
            load_asset_catalog(self.catalog_path)
        Image.new("RGB", (32, 20)).save(self.image_path, format="GIF")
        with self.assertRaisesRegex(AssetError, "PNG"):
            load_asset_catalog(self.catalog_path)

    def test_jpeg_and_dimension_metadata(self):
        jpeg = self.root / "photo.jpeg"
        Image.new("RGB", (32, 20)).save(jpeg)
        self.write_catalog([{**self.record, "file": "photo.jpeg"}])
        self.assertEqual(load_asset_catalog(self.catalog_path).resolve("photo_01"), jpeg.resolve())
        self.write_catalog([{**self.record, "width": 33}])
        with self.assertRaisesRegex(AssetError, "размеры"):
            load_asset_catalog(self.catalog_path)

    def test_byte_and_pixel_limits_are_checked_without_decoding_large_images(self):
        self.write_catalog()
        with self.image_path.open("r+b") as handle:
            handle.truncate(MAX_IMAGE_BYTES + 1)
        with self.assertRaisesRegex(AssetError, "15 МиБ"):
            load_asset_catalog(self.catalog_path)

        def png_chunk(name, content):
            return struct.pack(">I", len(content)) + name + content + struct.pack(">I", zlib.crc32(name + content))

        # A real PNG header claims >40M pixels; no large bitmap is allocated.
        header = struct.pack(">IIBBBBB", 8000, 5001, 8, 2, 0, 0, 0)
        self.image_path.write_bytes(b"\x89PNG\r\n\x1a\n" + png_chunk(b"IHDR", header)
                                    + png_chunk(b"IDAT", zlib.compress(b"")) + png_chunk(b"IEND", b""))
        with self.assertRaisesRegex(AssetError, "40 миллионов"):
            load_asset_catalog(self.catalog_path)

    def test_invalid_catalog_shape_and_metadata_have_safe_errors(self):
        for records in ([{**self.record, "id": "../unsafe"}], [{**self.record, "tags": "wrong"}],
                        [{**self.record, "width": True}], [{**self.record, "description": []}]):
            with self.subTest(records=records):
                self.write_catalog(records)
                with self.assertRaises(AssetError):
                    load_asset_catalog(self.catalog_path)
        self.catalog_path.write_text("not json", encoding="utf-8")
        with self.assertRaises(AssetError):
            load_asset_catalog(self.catalog_path)


if __name__ == "__main__":
    unittest.main()
