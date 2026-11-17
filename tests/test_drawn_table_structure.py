"""Shared headers of drawn tables stay separate from text and chart cells."""
import copy
import json
import unittest
from pathlib import Path
from zipfile import ZipFile

from scripts.prepare_severstal_template import _fields_and_objects
from superslide.source_tables import (
    DRAWN_TABLE_ROLE_KEYS, annotate_drawn_table_structures,
)
from superslide.source_template import _xml


ROOT = Path(__file__).resolve().parents[1]
CASES = [("tables", "slide_007"), ("diagram", "slide_009"), ("diagram", "slide_032")]


class DrawnTableStructureTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.manifest = json.loads((ROOT / "templates/severstal_template.json").read_text(encoding="utf-8"))

    def test_reviewed_drawn_compositions_reproduce_from_unchanged_source_files(self):
        actual_fields = 0
        for layout_id, variant_id in CASES:
            layout = self.manifest["layouts"][layout_id]
            variant = layout["variants"][variant_id]
            with ZipFile(ROOT / layout["source_file"]) as archive:
                package = {name: archive.read(name) for name in archive.namelist()}
            part = f"ppt/slides/slide{variant['source_slide']}.xml"
            before = package[part]
            extracted, _ = _fields_and_objects(None, package, part, "slide", f"sources/{layout_id}/")
            for field_id, field in extracted.items():
                for role in DRAWN_TABLE_ROLE_KEYS:
                    if role in field:
                        actual_fields += 1
                        self.assertEqual(field[role], variant["fields"][field_id][role])
                        self.assertIn(field["kind"], {"text", "chart"})
                        self.assertEqual(field["scope"], "slide")
            self.assertEqual(package[part], before)
        self.assertEqual(actual_fields, 80)

    def test_drawn_table_preserves_section_stubs_and_sparse_merged_cells(self):
        fields = self.manifest["layouts"]["tables"]["variants"]["slide_007"]["fields"]
        for column, sid in enumerate((21, 24, 27, 30), 1):
            self.assertEqual(fields[f"slide_text_{sid}"]["drawn_table_column_header"],
                             {"table_id": "tables_007", "column_index": column})
        self.assertEqual(fields["slide_text_119"]["drawn_table_section_header"], {
            "table_id": "tables_007", "section_index": 0, "role": "section_column_header",
        })
        for row, sid in enumerate((7, 10, 115, 18, 16, 17), 1):
            role = fields[f"slide_text_{sid}"]["drawn_table_cell"]
            self.assertEqual((role["row_index"], role["column_index"]), (row, 0))
            self.assertEqual(role["section_index"], 1 if row <= 3 else 2)
        self.assertEqual(fields["slide_text_15"]["drawn_table_cell"]["row_span"], 2)
        body = [field["drawn_table_cell"] for field in fields.values() if "drawn_table_cell" in field]
        self.assertFalse(any(cell["column_index"] == 4 and cell["row_index"] >= 4 for cell in body))

    def test_chart_columns_and_individual_chart_rows_retain_native_operations(self):
        fields = self.manifest["layouts"]["diagram"]["variants"]["slide_009"]["fields"]
        for sid, column in ((5, 2), (39, 3)):
            self.assertEqual(fields[f"slide_chart_{sid}"]["drawn_table_column_chart"], {
                "table_id": "diagram_009", "column_index": column, "row_span": 6,
            })
            self.assertEqual(fields[f"slide_chart_{sid}"]["kind"], "chart")
        fields = self.manifest["layouts"]["diagram"]["variants"]["slide_032"]["fields"]
        for row, sid in enumerate((54, 57, 91, 69, 67, 75, 78, 93), 1):
            self.assertEqual(fields[f"slide_chart_{sid}"]["drawn_table_cell"], {
                "table_id": "diagram_032", "row_index": row, "column_index": 4,
            })
            self.assertEqual(fields[f"slide_chart_{sid}"]["kind"], "chart")

    def test_cards_process_schemes_and_native_tables_are_not_drawn_tables(self):
        tagged = set()
        for layout_id, layout in self.manifest["layouts"].items():
            for variant_id, variant in layout["variants"].items():
                if any(role in field for field in variant["fields"].values() for role in DRAWN_TABLE_ROLE_KEYS):
                    tagged.add((layout_id, variant_id))
        self.assertEqual(tagged, set(CASES))

    def test_changed_source_removes_stale_reviewed_field_roles(self):
        layout = self.manifest["layouts"]["tables"]
        variant = layout["variants"]["slide_007"]
        with ZipFile(ROOT / layout["source_file"]) as archive:
            root = _xml(archive.read("ppt/slides/slide7.xml"))
        fields = copy.deepcopy(variant["fields"])
        root.set("show", "0")
        annotate_drawn_table_structures(root, fields, "slide")
        self.assertFalse(any(role in field for field in fields.values() for role in DRAWN_TABLE_ROLE_KEYS))


if __name__ == "__main__":
    unittest.main()
