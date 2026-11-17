"""Verify image semantics, editable output and version compatibility together."""

import copy
import json
import unittest
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pydantic import ValidationError

from support import temporary_directory
from superslide.assets import load_asset_catalog
from superslide.models import iter_image_refs, validate_presentation
from superslide.render import build_presentation
from superslide.template import TemplateError, load_template


ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "templates/legacy/corporate_base.pptx"
MANIFEST = ROOT / "templates/legacy/template_manifest.json"
CATALOG = ROOT / "assets/images/catalog.json"


class VarietyTests(unittest.TestCase):
    def setUp(self):
        self.directory = temporary_directory()
        self.addCleanup(self.directory.cleanup)
        self.output = Path(self.directory.name) / "deck.pptx"
        self.payload = json.loads((ROOT / "examples/legacy/variety.json").read_text(encoding="utf-8"))
        self.catalog = load_asset_catalog(CATALOG)

    def build(self, payload=None):
        spec = validate_presentation(payload or self.payload)
        report = build_presentation(spec, TEMPLATE, MANIFEST, self.output, asset_catalog=self.catalog)
        return Presentation(self.output), report

    def test_fourteen_layouts_and_native_objects(self):
        prs, report = self.build()
        self.assertEqual(len(prs.slides), 14)
        self.assertEqual(len({s.slide_layout.name for s in prs.slides}), 14)
        self.assertEqual(len(report["images"]), 5)
        self.assertEqual(sum(s.has_chart for slide in prs.slides for s in slide.shapes), 1)
        self.assertEqual(sum(s.has_table for slide in prs.slides for s in slide.shapes), 1)
        self.assertIn("+18%", " ".join(s.text for s in prs.slides[3].shapes if s.has_text_frame))
        self.assertTrue(all(not warning for warning in report["warnings"]))
        gallery_pictures = [s for s in prs.slides[4].shapes if s.name.startswith("superslide_image_")]
        self.assertEqual(len(gallery_pictures), 3)

    def test_cover_is_native_crop_and_alt_text_roundtrips(self):
        prs, _ = self.build()
        shape = next(s for s in prs.slides[2].shapes if s.name.startswith("superslide_image_"))
        _, manifest, _ = load_template(TEMPLATE, MANIFEST)
        field = manifest["layouts"]["photo_text"]["fields"]["image"]
        self.assertEqual((shape.left, shape.top, shape.width, shape.height),
                         tuple(field[k] for k in ("x", "y", "width", "height")))
        self.assertGreater(sum((shape.crop_left, shape.crop_right, shape.crop_top, shape.crop_bottom)), 0)
        with Image.open(self.catalog.resolve("production_worker")) as source:
            expected_ratio = source.width / source.height
        visible_ratio = (shape.width / shape.height) * (1 - shape.crop_top - shape.crop_bottom) / (1 - shape.crop_left - shape.crop_right)
        self.assertAlmostEqual(visible_ratio, expected_ratio, delta=0.001)
        self.assertEqual(shape._element.nvPicPr.cNvPr.get("descr"), "Сотрудник на производственной площадке")

    def test_contain_preserves_complete_picture_and_stays_inside_frame(self):
        payload = copy.deepcopy(self.payload)
        payload["slides"][2]["image"]["fit"] = "contain"
        prs, _ = self.build(payload)
        shape = next(s for s in prs.slides[2].shapes if s.name.startswith("superslide_image_"))
        field = json.loads(MANIFEST.read_text(encoding="utf-8"))["layouts"]["photo_text"]["fields"]["image"]
        self.assertGreaterEqual(shape.left, field["x"])
        self.assertGreaterEqual(shape.top, field["y"])
        self.assertLessEqual(shape.left + shape.width, field["x"] + field["width"])
        self.assertLessEqual(shape.top + shape.height, field["y"] + field["height"])
        with Image.open(self.catalog.resolve("production_worker")) as source:
            self.assertAlmostEqual(shape.width / shape.height, source.width / source.height, places=5)
        self.assertEqual(sum((shape.crop_left, shape.crop_right, shape.crop_top, shape.crop_bottom)), 0)

    def test_missing_unknown_images_do_not_replace_output(self):
        spec = validate_presentation(self.payload)
        self.output.write_bytes(b"previous output")
        with self.assertRaisesRegex(ValueError, "каталог"):
            build_presentation(spec, TEMPLATE, MANIFEST, self.output)
        self.payload["slides"][2]["image"]["asset_id"] = "missing"
        with self.assertRaises(ValueError):
            self.build()
        self.assertEqual(self.output.read_bytes(), b"previous output")

    def test_new_layouts_require_version_11_and_exact_cardinality(self):
        for mutate in (
            lambda p: p.__setitem__("template_version", "1.0"),
            lambda p: p["slides"][4]["images"].pop(),
            lambda p: p["slides"][3]["cards"].pop(),
            lambda p: p["slides"][5]["steps"].pop(),
            lambda p: p["slides"][4]["images"][1].__setitem__("asset_id", p["slides"][4]["images"][0]["asset_id"]),
        ):
            payload = copy.deepcopy(self.payload)
            mutate(payload)
            with self.assertRaises(ValidationError):
                validate_presentation(payload)

    def test_manifest_picture_fields_cannot_be_retyped_to_text(self):
        manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
        manifest["layouts"]["photo_text"]["fields"]["image"]["type"] = "text"
        path = Path(self.directory.name) / "invalid-manifest.json"
        path.write_text(json.dumps(manifest), encoding="utf-8")
        with self.assertRaisesRegex(TemplateError, "picture"):
            load_template(TEMPLATE, path)

    def test_old_version_still_builds_without_asset_catalog(self):
        payload = json.loads((ROOT / "examples/legacy/presentation.json").read_text(encoding="utf-8"))
        spec = validate_presentation(payload)
        self.assertEqual(list(iter_image_refs(spec)), [])
        build_presentation(spec, TEMPLATE, MANIFEST, self.output)
        self.assertEqual(len(Presentation(self.output).slides), 6)


if __name__ == "__main__":
    unittest.main()
