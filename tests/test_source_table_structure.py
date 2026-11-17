"""Table roles are extracted from source composition, not first-row styling."""
import copy
import json
import unittest
from pathlib import Path
from zipfile import ZipFile

from scripts.prepare_severstal_template import _fields_and_objects
from superslide.source_tables import annotate_table_structures
from superslide.source_template import A, _xml


ROOT = Path(__file__).resolve().parents[1]


class SourceTableStructureTests(unittest.TestCase):
    def test_all_native_table_structures_reproduce_from_the_untouched_sources(self):
        manifest = json.loads((ROOT / "templates/severstal_template.json").read_text(encoding="utf-8"))
        checked, placements = 0, []
        for layout_id, layout in manifest["layouts"].items():
            variants = [variant for variant in layout["variants"].values()
                        if any(field["kind"] == "table" and field["scope"] == "slide"
                               for field in variant["fields"].values())]
            if not variants:
                continue
            with ZipFile(ROOT / layout["source_file"]) as archive:
                package = {name: archive.read(name) for name in archive.namelist()}
            for variant in variants:
                part = f"ppt/slides/slide{variant['source_slide']}.xml"
                before = package[part]
                fields, _ = _fields_and_objects(None, package, part, "slide", f"sources/{layout_id}/")
                for field_id, field in fields.items():
                    if field["kind"] != "table":
                        continue
                    checked += 1
                    structure = field["table_structure"]
                    placements.append(structure["header_placement"])
                    self.assertNotEqual(structure["header_placement"], "unspecified")
                    self.assertEqual(structure, variant["fields"][field_id]["table_structure"])
                    self.assertEqual(structure["body_row_count"] + structure["header_row_count"], field["row_count"])
                    self.assertEqual(structure["data_row_count"] + structure["section_header_row_count"]
                                     + structure["total_row_count"], structure["body_row_count"])
                self.assertEqual(package[part], before)
        self.assertEqual(checked, 26)
        self.assertEqual(placements.count("external"), 1)
        self.assertEqual(placements.count("none"), 6)

    def test_unknown_component_table_does_not_get_a_header_from_first_row_style(self):
        with ZipFile(ROOT / "resources/severstal/templates/Severstal_Contents.pptx") as archive:
            package = {name: archive.read(name) for name in archive.namelist()}
        part = "ppt/slides/slide2.xml"
        fields, _ = _fields_and_objects(None, package, part, "slide", "sources/contents/")
        root = _xml(package[part])
        table = root.find(f".//{{{A}}}tbl")
        self.assertEqual(table.find(f"{{{A}}}tblPr").get("firstRow"), "1")
        table.find(f".//{{{A}}}t").text = "Changed sample"
        annotate_table_structures(root, fields, "slide")
        self.assertEqual(fields["slide_table_5"]["table_structure"], {"header_placement": "unspecified"})

    def test_partial_or_ambiguous_text_fields_do_not_define_an_external_header_row(self):
        with ZipFile(ROOT / "resources/severstal/templates/Severstal_Tables.pptx") as archive:
            package = {name: archive.read(name) for name in archive.namelist()}
        part = "ppt/slides/slide8.xml"
        fields, _ = _fields_and_objects(None, package, part, "slide", "sources/tables/")
        root = _xml(package[part])
        partial = copy.deepcopy(fields)
        del partial["slide_text_20"]
        annotate_table_structures(root, partial, "slide")
        self.assertNotEqual(partial["slide_table_18"]["table_structure"]["header_placement"], "external")
        ambiguous = copy.deepcopy(fields)
        ambiguous["slide_text_999"] = copy.deepcopy(ambiguous["slide_text_20"])
        annotate_table_structures(root, ambiguous, "slide")
        self.assertNotEqual(ambiguous["slide_table_18"]["table_structure"]["header_placement"], "external")


if __name__ == "__main__":
    unittest.main()
