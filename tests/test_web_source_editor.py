"""The compact source editor exposes titles and real photo slots only."""
import io
import json
import os
import unittest
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile

from PIL import Image

from support import temporary_directory
from superslide.web import WebState, _source_editor_metadata


class SourceEditorMetadataTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temporary = temporary_directory()
        cls.addClassCleanup(cls.temporary.cleanup)
        config = Path(cls.temporary.name) / "llm.json"
        config.write_text(json.dumps({"base_url": "https://editor-private.invalid/v1", "model": "test",
                                      "api_key_env": "SUPERSLIDE_EDITOR_TEST_KEY"}), encoding="utf-8")
        cls.state = WebState(config)
        cls.state.directory = Path(cls.temporary.name) / "web"
        cls.state.image_directory = cls.state.directory / "images"
        cls.addClassCleanup(lambda: cls.state.executor.shutdown(wait=True, cancel_futures=True))
        cls.addClassCleanup(lambda: cls.state.probe_executor.shutdown(wait=True, cancel_futures=True))
        cls.bootstrap = cls.state.bootstrap()
        cls.variants = {(layout["id"], variant["id"]): variant
                        for layout in cls.bootstrap["layouts"] for variant in layout["variants"]}

    def test_native_title_and_prominent_cover_text_are_selected_instead_of_dates_or_body(self):
        cover = self.variants[("title", "slide_002")]
        self.assertEqual(cover["title_field_id"], "slide_text_32")
        self.assertIn("Название презентации", cover["fields"][cover["title_field_id"]]["sample_text"])
        self.assertNotEqual(cover["title_field_id"], "slide_text_8")  # date is first in XML
        photo = self.variants[("photo", "slide_001")]
        self.assertEqual(photo["title_field_id"], "slide_text_13")
        self.assertTrue(all(variant["title_field_id"] is None or
                            (variant["title_field_id"] in variant["editable_field_ids"] and
                             variant["fields"][variant["title_field_id"]]["scope"] == "slide")
                            for variant in self.variants.values()))
        # Large KPI numerals are not a title when this composition has no
        # native heading placeholder.
        kpis = self.variants[("quotes", "slide_021")]
        self.assertNotIn(kpis["title_field_id"], {"slide_text_10", "slide_text_12", "slide_text_15"})

    def test_mixed_heading_body_binding_describes_just_the_heading_paragraph(self):
        mixed = self.variants[("title", "slide_010")]
        self.assertEqual(mixed["title_field_id"], "slide_text_15")
        self.assertTrue(mixed["title_paragraph_only"])
        self.assertEqual(mixed["title_paragraph_index"], 0)
        field = mixed["fields"][mixed["title_field_id"]]
        self.assertGreater(field["paragraph_count"], 1)
        self.assertEqual(field["paragraph_count"], len(field["paragraphs"]))
        plain = self.variants[("photo", "slide_001")]
        self.assertFalse(plain["title_paragraph_only"])
        self.assertIsNone(plain["title_paragraph_index"])

    def test_actual_photos_are_discovered_without_draft_updates_and_decorations_are_excluded(self):
        catalog = {record["sha256"]: record["id"] for record in self.state.catalog.records if record.get("sha256")}
        total = 0
        for (layout_id, variant_id), public in self.variants.items():
            original = self.state.manifest["layouts"][layout_id]["variants"][variant_id]
            for slot in public["photo_fields"]:
                field = original["fields"][slot["field_id"]]
                self.assertEqual(field["kind"], "picture")
                self.assertNotEqual(field["scope"], "master")
                self.assertEqual(slot["photo_id"], catalog[field["media_sha256"]])
                self.assertIn(slot["field_id"], public["editable_field_ids"])
                total += 1
        self.assertEqual(total, 156)
        photos = self.variants[("photo", "slide_001")]["photo_fields"]
        self.assertEqual(len(photos), 3)
        self.assertEqual({slot["field_id"] for slot in photos},
                         {"slide_picture_16", "slide_picture_18", "slide_picture_20"})
        self.assertFalse(self.variants[("icons", "slide_021")]["photo_fields"])
        closing = self.variants[("last_slide", "slide_001")]
        self.assertNotIn("layout_picture_35", [slot["field_id"] for slot in closing["photo_fields"]])

    def test_public_metadata_contains_friendly_names_and_no_private_package_paths(self):
        self.assertIsInstance(self.bootstrap["key_configured"], bool)
        raw = json.dumps(self.bootstrap, ensure_ascii=False)
        self.assertNotIn("editor-private.invalid", raw)
        self.assertNotIn("crop_and_transform_xml", raw)
        self.assertNotIn("media_part", raw)
        self.assertNotIn("sources/title/", raw)
        photo = next(item for item in self.state.image_list() if item["id"] == "worker_portrait_orange_helmet")
        self.assertTrue(photo["name"])
        self.assertNotEqual(photo["name"], photo["id"])
        self.assertEqual(photo["url"], "/api/images/worker_portrait_orange_helmet/file")
        self.assertNotIn("file", photo)

    def test_connection_flag_remains_boolean_with_populated_editor_fields(self):
        key = "SUPERSLIDE_EDITOR_TEST_KEY"
        with patch.dict(os.environ, {key: "editor-test-secret"}):
            public = self.state.bootstrap()
            self.assertIs(public["key_configured"], True)
            self.assertNotIn("editor-test-secret", json.dumps(public))
        with patch.dict(os.environ):
            os.environ.pop(key, None)
            self.assertIs(self.state.bootstrap()["key_configured"], False)

    def test_implicit_zero_placeholder_index_does_not_duplicate_inherited_photo(self):
        picture = '''<p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:nvPicPr><p:cNvPr id="2"/><p:cNvPicPr/><p:nvPr><p:ph type="pic"/></p:nvPr></p:nvPicPr>
        </p:pic>'''
        manifest = {"layouts": {"photo": {"name": "Фотографии", "variants": {"slide_001": {
            "name": "Фотография", "package_part": "ppt/slides/slide1.xml", "fields": {
                "slide_picture_2": {"kind": "picture", "scope": "slide", "media_sha256": "photo"},
                "layout_picture_2": {"kind": "picture", "scope": "layout", "media_sha256": "photo",
                                     "crop_and_transform_xml": picture},
            },
        }}}}}
        with temporary_directory() as directory:
            template = Path(directory) / "implicit-photo.pptx"
            with ZipFile(template, "w") as archive:
                archive.writestr("ppt/slides/slide1.xml", picture)
            metadata = _source_editor_metadata(manifest, template, [{"id": "photo_1", "sha256": "photo"}])
        slots = metadata[("photo", "slide_001")]["photo_fields"]
        self.assertEqual([slot["field_id"] for slot in slots], ["slide_picture_2"])

    def test_uploaded_image_has_a_readable_name_in_the_response_and_catalog(self):
        payload = io.BytesIO()
        Image.new("RGB", (80, 120), "navy").save(payload, format="PNG")
        uploaded = self.state.add_image("Фото цеха.png", payload.getvalue())
        self.assertEqual(uploaded["name"], "Фото цеха.png")
        public = next(item for item in self.state.image_list() if item["id"] == uploaded["id"])
        self.assertEqual(public["name"], uploaded["name"])
        self.assertEqual(public["url"], uploaded["url"])
        self.assertNotIn(str(Path(self.temporary.name)), json.dumps(uploaded))


if __name__ == "__main__":
    unittest.main()
