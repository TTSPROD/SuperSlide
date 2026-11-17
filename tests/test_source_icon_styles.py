"""A generated slide uses either backed or unbacked replacement icons."""
import copy
import json
import unittest
from pathlib import Path
from types import SimpleNamespace

from support import temporary_directory
from superslide.source_template import (
    ROOT, SourceFieldError, SourcePresentationSpec, build_source_presentation,
    load_source_manifest, validate_source_icon_styles,
)


ICONS = [
    {"id": "bare", "background_style": "without_background"},
    {"id": "bare-white", "background_style": "without_background"},
    {"id": "badge", "background_style": "with_background"},
    {"id": "badge-2", "has_background": True},
    {"id": "legacy"},
    {"id": "annotation", "background_style": "with_background", "selectable": False},
]


def payload(*slides):
    return {"template_version": "2.0", "title": "Проверка иконок", "slides": [
        {"layout_id": "cards", "variant_id": "slide_001", "updates": [
            {"field_id": f"slide_icon_{i}", "icon_id": icon} for i, icon in enumerate(ids, 1)
        ]} for ids in slides
    ]}


class SourceIconStyleTests(unittest.TestCase):
    def test_rejects_mixed_styles_for_both_json_and_validated_specs(self):
        draft = payload(["bare", "badge"])
        for spec in (draft, SourcePresentationSpec.model_validate(draft)):
            with self.subTest(spec=type(spec).__name__), self.assertRaises(SourceFieldError) as raised:
                validate_source_icon_styles(spec, ICONS)
            self.assertEqual(raised.exception.validation_details, {
                "slide_index": 1, "reason": "mixed_icon_background",
            })

    def test_style_is_per_slide_and_independent_of_icon_color(self):
        draft = payload(["bare", "bare-white"], ["badge", "badge-2"])
        before = copy.deepcopy(draft)
        self.assertIs(validate_source_icon_styles(draft, ICONS), draft)
        self.assertEqual(draft, before)

    def test_web_reports_the_style_rule_without_exposing_generated_content(self):
        from superslide.web import _safe_error
        error = SourceFieldError(1, None, "mixed_icon_background", "private generated content")
        message = _safe_error(error)
        self.assertIn("все с фоном или все без фона", message)
        self.assertNotIn("private", message)

    def test_single_icon_unchanged_slides_and_legacy_custom_catalogues(self):
        for draft in (payload([]), payload(["bare"]), payload(["legacy", "bare"])):
            self.assertIs(validate_source_icon_styles(draft, ICONS), draft)

    def test_unknown_and_nonselectable_icons_cannot_bypass_validation(self):
        for icon in ("missing", "annotation"):
            with self.subTest(icon=icon), self.assertRaises(SourceFieldError) as raised:
                validate_source_icon_styles(payload([icon]), ICONS)
            self.assertEqual(raised.exception.validation_details["reason"], "unknown_icon")

    def test_icons_in_picture_slots_are_included_and_photos_are_ignored(self):
        draft = payload(["bare", "badge"])
        draft["slides"][0]["updates"][1]["field_id"] = "slide_picture_8"
        draft["slides"][0]["updates"].append({"field_id": "slide_picture_9", "photo_id": "photo"})
        with self.assertRaisesRegex(SourceFieldError, "один тип"):
            validate_source_icon_styles(draft, ICONS)

    def test_partial_icon_edits_include_retained_source_icons(self):
        manifest = {"layouts": {"cards": {"variants": {"slide_001": {"fields": {
            "slide_icon_1": {"kind": "icon", "background_style": "without_background"},
            "slide_icon_2": {"kind": "icon", "background_style": "without_background"},
            "master_icon_3": {"kind": "icon", "background_style": "with_background"},
        }}}}}}
        with self.assertRaisesRegex(SourceFieldError, "один тип"):
            validate_source_icon_styles(payload(["badge"]), ICONS, manifest)
        # Same-style partial edits and a full conversion are both valid. Master
        # logos do not constrain the semantic slide icons' background choice.
        for draft in (payload(["bare"]), payload(["badge", "badge-2"]), payload([])):
            self.assertIs(validate_source_icon_styles(draft, ICONS, manifest), draft)

    def test_partial_real_table_edits_cannot_mix_with_original_icons(self):
        from superslide.paths import SOURCE_ICON_CATALOG
        manifest = load_source_manifest(ROOT / "templates/severstal_template.json")
        records = json.loads(SOURCE_ICON_CATALOG.read_text(encoding="utf-8"))["icons"]
        badge = next(item["id"] for item in records if item.get("selectable", True)
                     and item["background_style"] == "with_background")
        draft = {"template_version": "2.0", "title": "Частичная замена", "slides": [{
            "layout_id": "tables", "variant_id": "slide_019", "updates": [
                {"field_id": "slide_icon_15", "icon_id": badge},
            ],
        }]}
        with self.assertRaisesRegex(SourceFieldError, "один тип"):
            validate_source_icon_styles(draft, records, manifest)
        # Even replacing all seven previously exposed slots must account for
        # the renamed ship (shape 5), which was absent from the old catalogue.
        old_fields = ["slide_icon_15", "slide_icon_32", "slide_icon_33", "slide_icon_38",
                      "slide_icon_39", "slide_icon_40", "slide_icon_41"]
        draft["slides"][0]["updates"] = [{"field_id": fid, "icon_id": badge} for fid in old_fields]
        with self.assertRaisesRegex(SourceFieldError, "один тип"):
            validate_source_icon_styles(draft, records, manifest)

    def test_cli_web_and_direct_build_reject_real_catalogue_mixture(self):
        from superslide.cli import _validate_source_assets
        from superslide.paths import SOURCE_ICON_CATALOG, SOURCE_PHOTO_CATALOG
        from superslide.web import WebState
        manifest = load_source_manifest(ROOT / "templates/severstal_template.json")
        records = json.loads(SOURCE_ICON_CATALOG.read_text(encoding="utf-8"))["icons"]
        icons = [next(item["id"] for item in records
                      if item.get("selectable", True) and item["background_style"] == style)
                 for style in ("with_background", "without_background")]
        layout_id, variant_id, fields = next(
            (lid, vid, [fid for fid, field in variant["fields"].items()
                        if field["kind"] == "icon" and field["scope"] == "slide"])
            for lid, layout in manifest["layouts"].items()
            for vid, variant in layout["variants"].items()
            if lid != "icons" and sum(field["kind"] == "icon" and field["scope"] == "slide"
                                      for field in variant["fields"].values()) >= 2)
        draft = payload(icons)
        draft["slides"][0].update(layout_id=layout_id, variant_id=variant_id)
        for update, field_id in zip(draft["slides"][0]["updates"], fields):
            update["field_id"] = field_id
        spec = SourcePresentationSpec.model_validate(draft)
        with self.assertRaisesRegex(SourceFieldError, "один тип"):
            _validate_source_assets(spec, SOURCE_PHOTO_CATALOG, SOURCE_ICON_CATALOG)
        state = SimpleNamespace(source_mode=True, manifest=manifest, icons=records)
        with self.assertRaisesRegex(SourceFieldError, "один тип"):
            WebState.validate(state, draft)
        with temporary_directory() as temporary:
            output = Path(temporary) / "mixed.pptx"
            with self.assertRaisesRegex(SourceFieldError, "один тип"):
                build_source_presentation(spec, ROOT / "templates/severstal_source.pptx",
                                          ROOT / "templates/severstal_template.json", output)
            self.assertFalse(output.exists())


if __name__ == "__main__":
    unittest.main()
