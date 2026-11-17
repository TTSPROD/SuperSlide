"""Source-sized writing targets and complete overflow feedback reach the model."""
import json
import unittest

from superslide.source_template import (
    ROOT, SourceFieldError, load_source_manifest, source_prompt_catalog,
    validate_source_presentation,
)


class SourceTextBudgetTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.manifest = load_source_manifest(ROOT / "templates/severstal_template.json")

    def test_reported_fields_get_source_shaped_targets_before_generation(self):
        before = json.dumps(self.manifest, ensure_ascii=False)
        catalog = source_prompt_catalog(self.manifest)
        fields = catalog["bullets"]["variants"]["slide_012"]["fields"]
        for field_id, hard, target, paragraph_targets in (
            ("slide_text_22", 224, 190, [12, 176]),
            ("slide_text_27", 116, 98, [9, 87]),
        ):
            budget = fields[field_id]["text_budget"]
            self.assertEqual(budget["max_characters"], hard)
            self.assertEqual(budget["target_characters"], target)
            self.assertEqual(budget["max_paragraphs"], 2)
            self.assertEqual(budget["paragraph_targets"], paragraph_targets)
            self.assertEqual(budget["paragraph_roles"], ["heading", "body"])
            self.assertLessEqual(sum(budget["paragraph_targets"]), target)
        policy = catalog["bullets"]["text_budget_policy"]
        self.assertIn("spaces and punctuation", policy["counting_rule"])
        self.assertIn("paragraph separators do not count", policy["counting_rule"])
        self.assertIn("not a measured guarantee", policy["capacity_basis"])
        self.assertEqual(before, json.dumps(self.manifest, ensure_ascii=False))

    def test_reported_overflows_are_returned_together_without_their_text(self):
        spec = {"template_version": "2.0", "title": "Тест", "slides": [
            {"layout_id": "title", "variant_id": "slide_003", "updates": []},
            {"layout_id": "bullets", "variant_id": "slide_012", "updates": [
                {"field_id": "slide_text_22", "text": "С" * 284},
                {"field_id": "slide_text_27", "text": "К" * 146},
            ]},
        ]}
        with self.assertRaises(SourceFieldError) as caught:
            validate_source_presentation(spec, self.manifest)
        error = caught.exception
        self.assertEqual(error.validation_issues, [
            {"slide_index": 2, "reason": "text_length", "field_id": "slide_text_22", "limit": 224, "actual": 284, "target": 190},
            {"slide_index": 2, "reason": "text_length", "field_id": "slide_text_27", "limit": 116, "actual": 146, "target": 98},
        ])
        self.assertEqual(error.validation_details, error.validation_issues[0])
        self.assertIn("slide_text_22", str(error))
        self.assertIn("slide_text_27", str(error))
        self.assertNotIn("С" * 20, str(error))
        self.assertNotIn("К" * 20, json.dumps(error.validation_issues))

    def test_count_and_length_failures_are_aggregated_across_fields_and_slides(self):
        field = {"kind": "text", "paragraphs": ["a", "b"], "max_characters": 20}
        manifest = {"layouts": {"title": {"variants": {"slide_001": {"fields": {
            "slide_text_1": field, "slide_text_2": field,
        }}}}}}
        spec = {"template_version": "2.0", "title": "Тест", "slides": [
            {"layout_id": "title", "variant_id": "slide_001", "updates": [
                {"field_id": "slide_text_1", "paragraphs": ["a" * 21, "b", "c"]},
                {"field_id": "slide_text_2", "text": "d" * 22},
            ]},
            {"layout_id": "title", "variant_id": "slide_001", "updates": [
                {"field_id": "slide_text_2", "paragraphs": ["a", "b", "c"]},
            ]},
        ]}
        with self.assertRaises(SourceFieldError) as caught:
            validate_source_presentation(spec, manifest)
        issues = caught.exception.validation_issues
        self.assertEqual([(e["slide_index"], e["field_id"], e["reason"]) for e in issues], [
            (1, "slide_text_1", "paragraph_count"),
            (1, "slide_text_1", "text_length"),
            (1, "slide_text_2", "text_length"),
            (2, "slide_text_2", "paragraph_count"),
        ])

    def test_targets_are_editorial_guidance_not_new_hard_limits(self):
        # 200 characters is above the writing target190 but below the existing224.
        spec = {"template_version": "2.0", "title": "Тест", "slides": [{
            "layout_id": "bullets", "variant_id": "slide_012", "updates": [
                {"field_id": "slide_text_22", "paragraphs": ["a" * 80, "b" * 120]},
            ],
        }]}
        validated = validate_source_presentation(spec, self.manifest)
        self.assertEqual(sum(map(len, validated.slides[0].updates[0].paragraphs)), 200)

    def test_paragraph_targets_keep_empty_separators_and_do_not_invent_roles(self):
        manifest = {"layouts": {"title": {"name": "Титул", "variants": {"slide_001": {
            "name": "Титул", "fields": {"slide_text_1": {
                "kind": "text", "scope": "slide", "sample_text": "Header\n\nDescription",
                "paragraphs": ["Header", "", "Description"], "max_characters": 20,
            }},
        }}}}}
        budget = source_prompt_catalog(manifest)["title"]["variants"]["slide_001"]["fields"]["slide_text_1"]["text_budget"]
        self.assertEqual(budget["max_paragraphs"], 3)
        self.assertEqual(budget["paragraph_targets"][1], 0)
        self.assertNotIn("paragraph_roles", budget)
        self.assertLessEqual(sum(budget["paragraph_targets"]), budget["target_characters"])


if __name__ == "__main__":
    unittest.main()
