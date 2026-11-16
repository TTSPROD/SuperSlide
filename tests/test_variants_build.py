"""Integration checks for native PowerPoint composition variants."""

import copy
import json
import tempfile
import unittest
from pathlib import Path

from pptx import Presentation

from superslide.assets import load_asset_catalog
from superslide.models import SUPPORTED_VARIANTS, validate_presentation
from superslide.render import build_presentation
from superslide.template import TemplateError, iter_layout_designs, load_template, select_layout


ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "templates/corporate_base.pptx"
MANIFEST = ROOT / "templates/template_manifest.json"


class VariantBuildTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory(dir=ROOT)
        self.addCleanup(self.directory.cleanup)
        self.output = Path(self.directory.name) / "variants.pptx"
        self.payload = json.loads((ROOT / "examples/variants.json").read_text(encoding="utf-8"))
        self.catalog = load_asset_catalog(ROOT / "assets/images/catalog.json")

    def build(self):
        spec = validate_presentation(self.payload)
        report = build_presentation(spec, TEMPLATE, MANIFEST, self.output, asset_catalog=self.catalog)
        return Presentation(self.output), report

    def test_all_31_compositions_have_distinct_bindings_and_build(self):
        _, manifest, layouts = load_template(TEMPLATE, MANIFEST)
        expected = {(layout_id, variant_id) for layout_id, variants in SUPPORTED_VARIANTS.items()
                    for variant_id in variants}
        designs = list(iter_layout_designs(manifest))
        self.assertEqual({(layout_id, variant_id) for layout_id, variant_id, _ in designs}, expected)
        self.assertEqual(len(layouts), 31)
        self.assertEqual(len({info["layout_name"] for _, _, info in designs}), 31)
        prs, report = self.build()
        self.assertEqual(len(prs.slides), 31)
        self.assertEqual(len(report["designs"]), 31)
        self.assertEqual(len({slide.slide_layout.name for slide in prs.slides}), 31)
        self.assertEqual(sum(s.has_table for slide in prs.slides for s in slide.shapes), 1)
        self.assertEqual(sum(s.has_chart for slide in prs.slides for s in slide.shapes), 1)

    def test_changing_photo_variant_preserves_content_and_changes_geometry(self):
        prs, _ = self.build()
        slides = [s for s in prs.slides if "_photo_text_" in s.name]
        self.assertEqual(len(slides), 3)
        photos = [next(shape for shape in slide.shapes if shape.name.startswith("superslide_image_"))
                  for slide in slides]
        self.assertEqual(len({photo.image.sha1 for photo in photos}), 1)
        self.assertEqual(len({(p.left, p.top, p.width, p.height) for p in photos}), 3)
        for slide in slides:
            text = " ".join(shape.text for shape in slide.shapes if shape.has_text_frame)
            self.assertIn("За каждым результатом", text)
            self.assertIn("Обратная связь сотрудников", text)

    def test_full_photo_cover_keeps_original_art_and_text_in_front(self):
        prs, _ = self.build()
        slide = next(slide for slide in prs.slides if slide.name.endswith("title_slide_photo"))
        shapes = list(slide.shapes)
        self.assertTrue(shapes[0].name.startswith("superslide_image_"))
        self.assertEqual((shapes[0].left, shapes[0].top, shapes[0].width, shapes[0].height),
                         (0, 0, prs.slide_width, prs.slide_height))
        self.assertGreater(sum(not shape.is_placeholder for shape in shapes), 2)
        ids = [node.get("id") for node in slide.shapes._spTree.xpath(".//p:cNvPr")]
        self.assertEqual(len(ids), len(set(ids)))
        text_positions = [i for i, shape in enumerate(shapes) if shape.is_placeholder and shape.has_text_frame]
        self.assertGreater(min(text_positions), 0)
        self.assertIn("От идеи к результату", " ".join(s.text for s in shapes if s.has_text_frame))

    def test_two_column_bullets_do_not_lose_or_duplicate_items(self):
        prs, _ = self.build()
        slide = next(slide for slide in prs.slides if slide.name.endswith("bullet_list_two_columns"))
        sample = next(item for item in self.payload["slides"]
                      if item["layout_id"] == "bullet_list" and item["variant_id"] == "two_columns")
        texts = [paragraph.text for shape in slide.shapes if shape.has_text_frame
                 for paragraph in shape.text_frame.paragraphs]
        for bullet in sample["bullets"]:
            self.assertEqual(texts.count(bullet), 1)

    def test_photo_band_keeps_the_workers_helmet_and_face_in_frame(self):
        prs, _ = self.build()
        slide = next(slide for slide in prs.slides if slide.name.endswith("photo_text_photo_band"))
        picture = next(shape for shape in slide.shapes if shape.name.startswith("superslide_image_"))
        # The chosen source photo's helmet and face occupy roughly y=70..435/1080.
        self.assertLess(picture.crop_top, 70 / 1080)
        self.assertGreater(1 - picture.crop_bottom, 435 / 1080)

    def test_missing_variant_is_a_template_error_before_output_replace(self):
        manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
        manifest["layouts"]["photo_text"]["variants"].pop("photo_right")
        path = Path(self.directory.name) / "subset.json"
        path.write_text(json.dumps(manifest), encoding="utf-8")
        self.output.write_bytes(b"previous")
        with self.assertRaisesRegex(TemplateError, "photo_right"):
            build_presentation(validate_presentation(self.payload), TEMPLATE, path, self.output,
                               asset_catalog=self.catalog)
        self.assertEqual(self.output.read_bytes(), b"previous")

    def test_invalid_variant_geometry_is_rejected_by_template_inspection(self):
        manifest = copy.deepcopy(json.loads(MANIFEST.read_text(encoding="utf-8")))
        manifest["layouts"]["photo_text"]["variants"]["photo_right"]["fields"]["image"]["x"] += 10
        path = Path(self.directory.name) / "bad-geometry.json"
        path.write_text(json.dumps(manifest), encoding="utf-8")
        with self.assertRaisesRegex(TemplateError, "геометрия"):
            load_template(TEMPLATE, path)


if __name__ == "__main__":
    unittest.main()
