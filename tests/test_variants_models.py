"""Design variants retain one content contract and explicit version compatibility."""

import copy
import json
import unittest
from pathlib import Path

from pydantic import ValidationError

from superslide.models import (
    SUPPORTED_VARIANTS, iter_image_refs, presentation_schema, validate_presentation,
)


ROOT = Path(__file__).resolve().parents[1]


class VariantModelTests(unittest.TestCase):
    def setUp(self):
        self.payload = json.loads((ROOT / "examples/variety.json").read_text(encoding="utf-8"))
        self.payload["template_version"] = "1.2"
        self.by_layout = {slide["layout_id"]: slide for slide in self.payload["slides"]}

    def single(self, layout_id, variant_id="default", *, version="1.2"):
        slide = copy.deepcopy(self.by_layout[layout_id])
        slide["variant_id"] = variant_id
        if layout_id == "title_slide" and variant_id == "photo":
            slide["image"] = {"asset_id": "production_worker", "fit": "cover", "alt_text": "Производство"}
        return {"template_version": version, "title": "Проверка композиции", "slides": [slide]}

    def test_all_31_known_designs_validate_without_changing_content_contract(self):
        self.assertEqual(len(SUPPORTED_VARIANTS), 14)
        self.assertEqual(sum(map(len, SUPPORTED_VARIANTS.values())), 31)
        for layout_id, variants in SUPPORTED_VARIANTS.items():
            self.assertEqual(variants[0], "default")
            for variant_id in variants:
                with self.subTest(layout=layout_id, variant=variant_id):
                    spec = validate_presentation(self.single(layout_id, variant_id))
                    self.assertEqual(spec.slides[0].layout_id, layout_id)
                    self.assertEqual(spec.slides[0].variant_id, variant_id)

    def test_old_presentations_keep_default_composition(self):
        for file_name in ("presentation.json", "variety.json"):
            with self.subTest(file=file_name):
                legacy = json.loads((ROOT / "examples" / file_name).read_text(encoding="utf-8"))
                for slide in legacy["slides"]:
                    slide.pop("variant_id", None)
                spec = validate_presentation(legacy)
                self.assertTrue(all(slide.variant_id == "default" for slide in spec.slides))

    def test_nondefault_designs_require_version_12(self):
        for version in ("1.0", "1.1"):
            with self.subTest(version=version), self.assertRaisesRegex(ValueError, "1.2"):
                validate_presentation(self.single("title_slide", "blue_minimal", version=version))

    def test_unknown_or_wrong_layout_variant_is_rejected(self):
        for layout_id, variant_id in (("title_slide", "accent"), ("quote", "blue_minimal"),
                                     ("table", "accent_panel"), ("timeline", "unknown"),
                                     ("bullet_list", "../layout"), ("bullet_list", ""),
                                     ("bullet_list", 12)):
            with self.subTest(layout=layout_id, variant=variant_id), self.assertRaises(ValidationError):
                validate_presentation(self.single(layout_id, variant_id))

    def test_title_photo_requires_image_and_image_is_never_silently_ignored(self):
        photo = self.single("title_slide", "photo")
        spec = validate_presentation(photo)
        self.assertEqual([image.asset_id for image in iter_image_refs(spec)], ["production_worker"])
        for absent in (None, "missing"):
            invalid = copy.deepcopy(photo)
            if absent == "missing":
                invalid["slides"][0].pop("image")
            else:
                invalid["slides"][0]["image"] = None
            with self.subTest(image=absent), self.assertRaisesRegex(ValueError, "требует image"):
                validate_presentation(invalid)
        for variant_id in ("default", "blue_minimal", "blue_graphic"):
            invalid = copy.deepcopy(photo)
            invalid["slides"][0]["variant_id"] = variant_id
            with self.subTest(variant=variant_id), self.assertRaisesRegex(ValueError, "только"):
                validate_presentation(invalid)

    def test_schema_enumerates_allowed_variants_per_layout(self):
        schema = presentation_schema()
        definitions = schema["$defs"].values()
        found = set()
        for definition in definitions:
            properties = definition.get("properties", {})
            layout_id = properties.get("layout_id", {}).get("const")
            if layout_id not in SUPPORTED_VARIANTS:
                continue
            variant = properties["variant_id"]
            expected = list(SUPPORTED_VARIANTS[layout_id])
            self.assertEqual(variant.get("enum", [variant.get("const")]), expected)
            self.assertEqual(variant["default"], "default")
            found.add(layout_id)
        self.assertEqual(found, set(SUPPORTED_VARIANTS))


if __name__ == "__main__":
    unittest.main()
