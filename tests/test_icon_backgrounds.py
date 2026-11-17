"""The icon catalogue separates actual backing plates from transparent art."""
import copy
import json
import unittest
from collections import Counter
from pathlib import Path

from PIL import Image

from scripts.extract_severstal_icons import annotate_icon_backgrounds


ROOT = Path(__file__).resolve().parents[1]


class IconBackgroundTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.catalog = json.loads((ROOT / "assets/icons/severstal/catalog.json").read_text(encoding="utf-8"))
        cls.by_source = {(r["source"]["slide"], r["source"]["index"]): r for r in cls.catalog["icons"]}

    def test_every_icon_has_explicit_consistent_background_metadata(self):
        for record in self.catalog["icons"]:
            with self.subTest(icon=record["id"]):
                self.assertIn(record["background_style"], {"with_background", "without_background"})
                self.assertIs(record["has_background"], record["background_style"] == "with_background")
                if record["has_background"]:
                    self.assertIn(record["background_color"], {"#002F6C", "#FFFFFF"})
                else:
                    self.assertIsNone(record["background_color"])
        self.assertEqual(self.catalog["stats"]["by_background_style"],
                         {"with_background": 683, "without_background": 679})

    def test_group_partitions_are_disjoint_complete_and_match_record_style(self):
        records = {r["id"]: r for r in self.catalog["icons"]}
        for group in self.catalog["groups"]:
            with self.subTest(group=group["id"]):
                partitions = group["icon_ids_by_background_style"]
                self.assertEqual(set(partitions), {"with_background", "without_background"})
                actual = Counter()
                for style, ids in partitions.items():
                    actual.update(ids)
                    for icon_id in ids:
                        self.assertEqual(records[icon_id]["background_style"], style)
                        self.assertEqual(records[icon_id]["group_id"], group["id"])
                self.assertEqual(actual, Counter(icon_id for icon_id in group["icon_ids"]
                                                 if records[icon_id]["selectable"]))

    def test_navigation_and_logistics_variants_match_the_user_example(self):
        # Ship, pin and map without a plate; their badge versions retain a plate.
        for unbacked, backed in [((15, 27), (15, 66)), ((15, 23), (15, 62)),
                                 ((15, 26), (15, 65)), ((28, 15), (27, 15))]:
            self.assertEqual(self.by_source[unbacked]["background_style"], "without_background")
            self.assertEqual(self.by_source[backed]["background_style"], "with_background")

    def test_white_transparent_icons_and_baked_white_rasters_are_distinguished(self):
        for key in [(4, 21), (5, 1)]:
            record = self.by_source[key]
            self.assertIn("белая на прозрачном фоне", record["name"])
            self.assertFalse(record["has_background"])
        for key in [(3, 7), (7, 33), (23, 16)]:
            record = self.by_source[key]
            self.assertFalse(record["is_vector"])
            self.assertEqual(record["background_color"], "#FFFFFF")
            with Image.open(ROOT / record["png_file"]) as picture:
                self.assertEqual(picture.getchannel("A").getextrema(), (255, 255))
                self.assertEqual(picture.getpixel((0, 0)), (255, 255, 255, 255))

    def test_all_backed_assets_actually_have_large_opaque_surfaces(self):
        # Check the rendered assets, independently of names and extraction rules.
        # Rounded/cut badge corners account for the remaining transparent area.
        for record in self.catalog["icons"]:
            if not record["has_background"]:
                continue
            with self.subTest(icon=record["id"]), Image.open(ROOT / record["png_file"]) as picture:
                histogram = picture.getchannel("A").histogram()
                self.assertGreater(sum(histogram[241:]) / sum(histogram), .90)

    def test_extraction_reproduces_metadata_without_changing_artwork(self):
        regenerated = copy.deepcopy(self.catalog)
        for record in regenerated["icons"]:
            for key in ("background_style", "has_background", "background_color"):
                record.pop(key)
        for group in regenerated["groups"]:
            group.pop("icon_ids_by_background_style")
        regenerated.pop("background_styles")
        regenerated["stats"].pop("by_background_style")
        annotate_icon_backgrounds(regenerated)
        self.assertEqual(regenerated, self.catalog)


if __name__ == "__main__":
    unittest.main()
