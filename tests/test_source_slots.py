"""Slot hints describe the actual source artwork, not model-invented facts."""
import json
import hashlib
import unittest
from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile

from support import temporary_directory
from superslide.source_slots import source_slot_metadata
from superslide.source_template import A, NS, _paragraph_text, _xml, build_source_presentation, source_prompt_catalog


ROOT = Path(__file__).resolve().parents[1]


class SourceSlotMetadataTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.manifest = json.loads((ROOT / "templates/severstal_template.json").read_text(encoding="utf-8"))

    def test_every_source_variant_has_physical_block_inventory(self):
        variants = [variant for layout in self.manifest["layouts"].values() for variant in layout["variants"].values()]
        self.assertEqual(len(variants), 211)
        for variant in variants:
            with self.subTest(name=variant["name"], part=variant["package_part"]):
                inventory = source_slot_metadata(variant)["block_inventory"]
                primary = source_slot_metadata(variant).get("primary_blocks")
                self.assertIsInstance(primary, dict)
                self.assertGreater(primary["count"], 0)
                self.assertTrue(primary["kind"])
                fields = [field for field in variant["fields"].values() if field.get("scope") == "slide"
                          and not field.get("automatic") and field.get("sample_text", "").strip() not in {"##.##.####", "‹#›", "<#>"}
                          and not " ".join(field.get("sample_text", "").split()).startswith("Любое несанкционированное копирование")]
                for kind in ("text", "picture", "table", "chart", "icon"):
                    self.assertEqual(inventory[kind + "_fields"], sum(f["kind"] == kind for f in fields))

    def test_changed_source_slide_requires_a_new_block_count_review(self):
        manifest = deepcopy(self.manifest)
        variant = manifest["layouts"]["cards"]["variants"]["slide_009"]
        variant["slide_xml_sha256"] = "0" * 64
        with self.assertRaisesRegex(ValueError, "пересчитайте и проверьте число блоков"):
            source_prompt_catalog(manifest)

    def test_primary_sections_count_real_independent_bullets(self):
        variants = self.manifest["layouts"]["bullets"]["variants"]
        for vid, sections in {"slide_001": [5], "slide_007": [5, 3],
                              "slide_010": [4, 4, 4], "slide_017": [2, 2, 2]}.items():
            with self.subTest(variant=vid):
                self.assertEqual(source_slot_metadata(variants[vid])["primary_blocks"],
                                 {"kind": "content_blocks", "count": sum(sections), "section_counts": sections})

    def test_reviewed_multifield_cards_and_scheme_nodes_are_not_counted_as_text_shapes(self):
        cards = self.manifest["layouts"]["cards"]["variants"]
        self.assertEqual(source_slot_metadata(cards["slide_003"])["primary_blocks"]["count"], 12)
        self.assertEqual(source_slot_metadata(cards["slide_009"])["primary_blocks"]["section_counts"], [3, 2, 1])
        self.assertEqual(source_slot_metadata(cards["slide_010"])["primary_blocks"]["items_per_block"], [5, 5])
        schemes = self.manifest["layouts"]["schemes"]["variants"]
        self.assertEqual(source_slot_metadata(schemes["slide_007"])["primary_blocks"]["section_counts"], [4, 1])

    def test_source_objects_and_reference_icons_are_counted_separately(self):
        title = self.manifest["layouts"]["title"]["variants"]["slide_003"]
        self.assertEqual(source_slot_metadata(title)["primary_blocks"], {"kind": "single_composition", "count": 1})
        charts = self.manifest["layouts"]["diagram"]["variants"]["slide_004"]
        info = source_slot_metadata(charts)
        self.assertEqual(info["primary_blocks"], {"kind": "chart", "count": 4})
        self.assertEqual(len(info["block_inventory"]["charts"]), 4)
        self.assertTrue(all(item["categories"] == 4 for item in info["block_inventory"]["charts"]))
        icons = self.manifest["layouts"]["icons"]["variants"]["slide_001"]
        info = source_slot_metadata(icons)
        records = json.loads((ROOT / "assets/icons/severstal/catalog.json").read_text(encoding="utf-8"))["icons"]
        self.assertEqual(info["primary_blocks"]["count"], sum(item["source"]["slide"] == 1 for item in records))

    def test_reviewed_compound_contents_and_drawn_graph_are_counted_by_role(self):
        contents = self.manifest["layouts"]["contents"]["variants"]
        self.assertEqual([source_slot_metadata(contents[f"slide_{i:03}"])["primary_blocks"]["count"] for i in range(1, 6)], [8, 8, 8, 8, 7])
        drawn = self.manifest["layouts"]["tables"]["variants"]["slide_007"]
        self.assertEqual(source_slot_metadata(drawn)["primary_blocks"]["section_counts"], [3, 3])
        diagram = self.manifest["layouts"]["diagram"]["variants"]["slide_028"]
        self.assertEqual(source_slot_metadata(diagram)["primary_blocks"], {
            "kind": "conclusion_blocks", "count": 4, "chart_points": 20, "first_block_bullets": 3})
        diagram = self.manifest["layouts"]["diagram"]["variants"]["slide_019"]
        self.assertEqual(source_slot_metadata(diagram)["primary_blocks"], {
            "kind": "mixed_charts_and_text", "count": 7, "text_blocks": 4, "chart_blocks": 3})
        diagrams = self.manifest["layouts"]["diagram"]["variants"]
        self.assertEqual(source_slot_metadata(diagrams["slide_009"])["primary_blocks"], {
            "kind": "comparison_categories", "count": 6, "charts": 2, "years": 2})
        self.assertEqual(source_slot_metadata(diagrams["slide_010"])["primary_blocks"], {
            "kind": "chart", "count": 3, "categories_per_chart": 8})
        self.assertEqual(source_slot_metadata(diagrams["slide_011"])["primary_blocks"], {
            "kind": "mixed_charts_and_text", "count": 8, "text_blocks": 5, "chart_blocks": 3,
            "categories_per_chart": 6})
        self.assertEqual(source_slot_metadata(diagrams["slide_017"])["primary_blocks"], {
            "kind": "mixed_charts_and_text", "count": 4, "text_blocks": 3, "chart_blocks": 1,
            "first_block_bullets": 3})
        self.assertEqual(source_slot_metadata(diagrams["slide_018"])["primary_blocks"], {
            "kind": "mixed_charts_and_text", "count": 3, "text_blocks": 2, "chart_blocks": 1})
        quotes = self.manifest["layouts"]["quotes"]["variants"]
        self.assertEqual(source_slot_metadata(quotes["slide_003"])["primary_blocks"], {"kind": "content_blocks", "count": 3})
        for number, count in ((13, 3), (14, 7)):
            self.assertEqual(source_slot_metadata(quotes[f"slide_{number:03}"])["primary_blocks"], {"kind": "quotes", "count": count})
        for number in (19, 20):
            self.assertEqual(source_slot_metadata(quotes[f"slide_{number:03}"])["primary_blocks"],
                             {"kind": "mixed_content", "count": 4, "main_theses": 1, "supporting_blocks": 3})
        self.assertEqual(source_slot_metadata(quotes["slide_021"])["primary_blocks"], {"kind": "indicators", "count": 3})
        photo = self.manifest["layouts"]["photo"]["variants"]["slide_009"]
        self.assertEqual(source_slot_metadata(photo)["primary_blocks"],
                         {"kind": "mixed_content", "count": 8, "photo_blocks": 5, "text_blocks": 3})
        titles = self.manifest["layouts"]["title"]["variants"]
        self.assertEqual(source_slot_metadata(titles["slide_001"])["primary_blocks"],
                         {"kind": "reference_instruction", "count": 4, "reference_only": True})
        for number in (9, 10):
            self.assertEqual(source_slot_metadata(titles[f"slide_{number:03}"])["primary_blocks"], {"kind": "profile", "count": 1})

    def test_repeated_numbered_labels_are_grouped_with_restarts(self):
        variant = self.manifest["layouts"]["cards"]["variants"]["slide_009"]
        groups = source_slot_metadata(variant)["groups"]
        numbered = [group for group in groups if group["kind"] == "numbered_text_fields"]
        self.assertIn({"kind": "numbered_text_fields", "count": 3,
                       "field_ids": ["slide_text_38", "slide_text_43", "slide_text_44"]}, numbered)
        self.assertIn({"kind": "numbered_text_fields", "count": 2,
                       "field_ids": ["slide_text_50", "slide_text_51"]}, numbered)
        self.assertIn({"kind": "numbered_text_fields", "count": 1,
                       "field_ids": ["slide_text_53"]}, numbered)
        self.assertIn({"kind": "numbered_text_fields", "count": 3,
                       "field_ids": ["slide_text_14", "slide_text_37", "slide_text_72"]}, numbered)

    def test_explicit_bullet_paragraph_capacity_has_its_exact_field_and_positions(self):
        variant = self.manifest["layouts"]["bullets"]["variants"]["slide_003"]
        groups = source_slot_metadata(variant)["groups"]
        self.assertIn({"kind": "bullet_paragraphs", "field_id": "slide_text_7",
                       "count": 4, "paragraph_indexes": [1, 2, 3, 4]}, groups)

    def test_named_count_only_for_explicit_number_and_content_kind(self):
        variants = self.manifest["layouts"]["cards"]["variants"]
        self.assertEqual(source_slot_metadata(variants["slide_012"])["named_item_count"], 3)
        self.assertEqual(source_slot_metadata(variants["slide_013"])["named_item_count"], 4)
        self.assertNotIn("named_item_count", source_slot_metadata(variants["slide_006"]))

    def test_reordered_numbered_fields_still_form_one_group(self):
        variant = self.manifest["layouts"]["bullets"]["variants"]["slide_013"]
        groups = [g for g in source_slot_metadata(variant)["groups"] if g["kind"] == "numbered_text_fields"]
        self.assertTrue(any(g["count"] == 4 and set(g["field_ids"]) == {
            "slide_text_31", "slide_text_34", "slide_text_37", "slide_text_40"} for g in groups))

    def test_built_slide_removes_unused_bullets_for_short_and_explicit_empty_lists(self):
        template = ROOT / "templates/severstal_source.pptx"
        manifest = ROOT / "templates/severstal_template.json"
        original_digest = hashlib.sha256(template.read_bytes()).hexdigest()
        contents = ["Факт", "Первый пункт", "Второй пункт"]
        slides = [
            {"layout_id": "bullets", "variant_id": "slide_003", "updates": [
                {"field_id": "slide_text_7", "paragraphs": paragraphs},
            ]}
            for paragraphs in (contents, contents + ["", "  "])
        ]
        spec = {"template_version": "2.0", "title": "Проверка списка", "slides": slides}
        with temporary_directory() as directory:
            output = Path(directory) / "bullet-count.pptx"
            build_source_presentation(spec, template, manifest, output)
            with ZipFile(output) as archive:
                for slide_number in (1, 2):
                    with self.subTest(slide=slide_number):
                        root = _xml(archive.read(f"ppt/generated/slide{slide_number}.xml"))
                        paragraphs = root.xpath(
                            ".//p:sp[p:nvSpPr/p:cNvPr[@id='7']]/p:txBody/a:p", namespaces=NS,
                        )
                        self.assertEqual([_paragraph_text(p) for p in paragraphs], contents)
                        self.assertEqual(len(paragraphs), 3)
                        self.assertEqual(sum(bool(p.xpath("./a:pPr/a:buChar | ./a:pPr/a:buAutoNum", namespaces=NS))
                                             for p in paragraphs), 2)
        self.assertEqual(hashlib.sha256(template.read_bytes()).hexdigest(), original_digest)


if __name__ == "__main__":
    unittest.main()
