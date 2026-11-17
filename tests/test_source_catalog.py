"""Only actual content and useful source constraints belong in LLM prompts."""
import json
import unittest
from copy import deepcopy
from pathlib import Path

from superslide.source_template import (
    ROOT, SourceFieldError, load_source_manifest, source_prompt_catalog,
    validate_source_presentation,
)


class SourceCatalogTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.manifest = load_source_manifest(ROOT / "templates/severstal_template.json")

    def test_real_catalog_keeps_types_variants_and_only_model_editable_fields(self):
        before = json.dumps(self.manifest, ensure_ascii=False)
        catalog = source_prompt_catalog(self.manifest)
        self.assertEqual(set(catalog), set(self.manifest["layouts"]))
        self.assertEqual(sum(len(layout["variants"]) for layout in catalog.values()), 211)
        for lid, layout in catalog.items():
            for vid, variant in layout["variants"].items():
                self.assertEqual(set(variant["fields"]) - set(self.manifest["layouts"][lid]["variants"][vid]["fields"]), set())
                for fid, field in variant["fields"].items():
                    self.assertTrue(fid.startswith("slide_"), fid)
                    self.assertNotIn("drawingml", field)
                    self.assertNotIn("paragraph_styles", field)
                    self.assertNotIn("media_part", field)
                    if field["kind"] == "text":
                        original = self.manifest["layouts"][lid]["variants"][vid]["fields"][fid]
                        self.assertLessEqual(len(field["sample_text"]), 80)
                        self.assertNotIn("\v", field["sample_text"])
                        self.assertEqual(field["max_characters"], original["max_characters"])
                        self.assertEqual(field["paragraph_count"], len(original["paragraphs"]))
                        if field["paragraph_count"] > 1:
                            self.assertEqual(field["paragraph_characters"], [len(p) for p in original["paragraphs"]])
        self.assertEqual(before, json.dumps(self.manifest, ensure_ascii=False))
        self.assertFalse("Любое несанкционированное копирование" in json.dumps(catalog, ensure_ascii=False))

    def test_sample_omits_office_controls_and_fixed_brand_images(self):
        text = "Заголовок\vслайда\n\n" + "Описание " * 60
        base = {"kind": "text", "scope": "slide", "shape_id": 1, "sample_text": text,
                "paragraphs": text.split("\n"), "max_characters": len(text)}
        fixture = {"layouts": {"title": {"name": "Титульные", "variants": {"slide_001": {
            "name": "Титул", "fields": {
                "slide_text_1": base,
                "slide_icon_2": {"kind": "icon", "scope": "slide", "media_sha256": "brand"},
                "master_icon_3": {"kind": "icon", "scope": "master", "media_sha256": "brand"},
                "layout_text_4": {**base, "scope": "layout"},
                "slide_text_5": {**base, "sample_text": "##.##.####", "automatic": "current_date"},
            }}}}}}
        fields = source_prompt_catalog(fixture)["title"]["variants"]["slide_001"]["fields"]
        self.assertEqual(set(fields), {"slide_text_1"})
        self.assertEqual(fields["slide_text_1"]["paragraph_count"], 3)
        self.assertEqual(fields["slide_text_1"]["paragraph_samples"][1], "")
        self.assertNotIn("\v", fields["slide_text_1"]["sample_text"])

    def test_table_preview_is_explicitly_a_preview_and_chart_contract_is_named(self):
        catalog = source_prompt_catalog(self.manifest)
        table_found = chart_found = False
        for lid, layout in catalog.items():
            for vid, variant in layout["variants"].items():
                for fid, field in variant["fields"].items():
                    if field["kind"] == "table":
                        table_found = True
                        self.assertNotIn("rows", field)
                        self.assertLessEqual(len(field["sample_rows"]), 2)
                        self.assertEqual(len(field["cell_paragraph_counts"]), field["row_count"])
                        self.assertTrue(all(len(row) == field["column_count"] for row in field["cell_paragraph_counts"]))
                    if field["kind"] == "chart":
                        chart_found = True
                        self.assertTrue(field["chart"]["supported"])
                        self.assertEqual(field["operation"], "chart_data")
        self.assertTrue(table_found and chart_found)

    def test_text_failure_has_safe_slide_field_and_limit_metadata(self):
        variant = self.manifest["layouts"]["title"]["variants"]["slide_003"]
        field_id = "slide_text_11"
        limit = variant["fields"][field_id]["max_characters"]
        private_text = "СЕКРЕТ" * (limit + 1)
        spec = {"template_version": "2.0", "title": "x", "slides": [{
            "layout_id": "title", "variant_id": "slide_003",
            "updates": [{"field_id": field_id, "text": private_text}],
        }]}
        with self.assertRaises(SourceFieldError) as caught:
            validate_source_presentation(spec, self.manifest)
        details = caught.exception.validation_details
        self.assertEqual(details["slide_index"], 1)
        self.assertEqual(details["field_id"], field_id)
        self.assertEqual(details["reason"], "text_length")
        self.assertEqual(details["limit"], limit)
        self.assertNotIn("СЕКРЕТ", str(caught.exception))
        self.assertNotIn("СЕКРЕТ", json.dumps(details, ensure_ascii=False))

    def test_table_cell_paragraph_limit_is_checked_before_render(self):
        variant = self.manifest["layouts"]["tables"]["variants"]["slide_001"]
        fid, field = next((fid, f) for fid, f in variant["fields"].items() if f["kind"] == "table")
        rows = deepcopy(field["rows"])
        # Replace soft breaks with spaces, as the native input contract requires.
        rows = [[text.replace("\v", " ") for text in row] for row in rows]
        rows[0][0] = "\n".join(["x"] * (len(rows[0][0].split("\n")) + 1))
        spec = {"template_version": "2.0", "title": "x", "slides": [{
            "layout_id": "tables", "variant_id": "slide_001", "updates": [{"field_id": fid, "rows": rows}],
        }]}
        with self.assertRaises(SourceFieldError) as caught:
            validate_source_presentation(spec, self.manifest)
        self.assertEqual(caught.exception.validation_details["reason"], "cell_paragraph_count")
        self.assertEqual(caught.exception.validation_details["row_index"], 1)
        self.assertEqual(caught.exception.validation_details["column_index"], 1)

    def test_paragraph_character_counts_are_guidance_not_new_limits(self):
        field = {"kind": "text", "scope": "slide", "paragraphs": ["a", "longer"], "max_characters": 20}
        manifest = {"layouts": {"title": {"variants": {"slide_001": {"fields": {"slide_text_1": field}}}}}}
        spec = {"template_version": "2.0", "title": "x", "slides": [{
            "layout_id": "title", "variant_id": "slide_001", "updates": [{"field_id": "slide_text_1", "paragraphs": ["longer", "a"]}],
        }]}
        self.assertEqual(validate_source_presentation(spec, manifest).slides[0].updates[0].paragraphs, ["longer", "a"])


if __name__ == "__main__":
    unittest.main()
