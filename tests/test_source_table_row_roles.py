"""Native section headings, totals and merged cells are not ordinary data slots."""
import copy
import json
import unittest
from pathlib import Path
from zipfile import ZipFile

from scripts.prepare_severstal_template import _fields_and_objects
from superslide.source_tables import annotate_table_structures
from superslide.source_template import A, NS, _xml


ROOT = Path(__file__).resolve().parents[1]


class SourceTableRowRoleTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.manifest = json.loads((ROOT / "templates/severstal_template.json").read_text(encoding="utf-8"))
        with ZipFile(ROOT / "resources/severstal/templates/Severstal_Tables.pptx") as archive:
            cls.package = {part: archive.read(part) for part in archive.namelist()}

    def table(self, slide):
        variant = self.manifest["layouts"]["tables"]["variants"][f"slide_{slide:03}"]
        field = next(field for field in variant["fields"].values() if field["kind"] == "table")
        source = _xml(self.package[f"ppt/slides/slide{slide}.xml"])
        return field["table_structure"], source.find(f".//{{{A}}}tbl")

    def test_category_headings_do_not_consume_factual_data_capacity(self):
        for slide, category_rows, data_count in ((9, [1, 6], 8), (10, [1, 5], 6)):
            with self.subTest(slide=slide):
                structure, table = self.table(slide)
                rows = table.findall(f"{{{A}}}tr")
                self.assertEqual(structure["section_header_row_indexes"], category_rows)
                self.assertEqual(structure["section_header_row_count"], 2)
                self.assertEqual(structure["data_row_count"], data_count)
                self.assertEqual(structure["body_row_count"], data_count + 2)
                self.assertEqual(structure["row_roles"][0], "column_header")
                for index in category_rows:
                    self.assertEqual(structure["row_roles"][index], "section_header")
                    self.assertNotIn(index, structure["data_row_indexes"])
                    # Independent source evidence: the category name spans all
                    # four columns, unlike an ordinary four-cell factual row.
                    cell = rows[index].find(f"{{{A}}}tc")
                    self.assertEqual(cell.get("gridSpan"), "4")
                    self.assertEqual("".join(cell.xpath(".//a:t/text()", namespaces=NS)), "Название категории")

    def test_source_total_rows_are_explicit_and_separate_from_data(self):
        for slide, total_row, total_label, data_count in ((6, 17, "Итого", 15), (17, 8, "TOTAL", 6)):
            with self.subTest(slide=slide):
                structure, table = self.table(slide)
                self.assertEqual(structure["header_row_count"], 2)
                self.assertEqual(structure["total_row_indexes"], [total_row])
                self.assertEqual(structure["total_row_count"], 1)
                self.assertEqual(structure["data_row_count"], data_count)
                self.assertEqual(structure["row_roles"][total_row], "total")
                self.assertNotIn(total_row, structure["data_row_indexes"])
                source_row = table.findall(f"{{{A}}}tr")[total_row]
                self.assertIn(total_label, source_row.xpath(".//a:t/text()", namespaces=NS))

    def test_merge_regions_exactly_cover_all_continuation_cells(self):
        seen = 0
        for layout in self.manifest["layouts"].values():
            for variant in layout["variants"].values():
                for field in variant["fields"].values():
                    if field["kind"] != "table":
                        continue
                    structure = field["table_structure"]
                    covered = set()
                    for merge in structure["merged_cells"]:
                        r, c = merge["row_index"], merge["column_index"]
                        self.assertLessEqual(r + merge["row_span"], field["row_count"])
                        self.assertLessEqual(c + merge["column_span"], field["column_count"])
                        cells = {(row, column) for row in range(r, r + merge["row_span"])
                                 for column in range(c, c + merge["column_span"])} - {(r, c)}
                        self.assertFalse(covered & cells)
                        covered.update(cells)
                    continuation = {tuple(cell) for cell in structure["merged_continuation_cells"]}
                    self.assertEqual(covered, continuation)
                    seen += len(continuation)
        self.assertEqual(seen, 42)

    def test_changed_unreviewed_table_does_not_inherit_old_section_indexes(self):
        part = "ppt/slides/slide9.xml"
        fields, _ = _fields_and_objects(None, self.package, part, "slide", "sources/tables/")
        before = copy.deepcopy(fields)
        root = _xml(self.package[part])
        root.find(f".//{{{A}}}tbl//{{{A}}}t").text = "Changed unknown header"
        annotate_table_structures(root, fields, "slide")
        self.assertEqual(before["slide_table_7"]["table_structure"]["section_header_row_indexes"], [1, 6])
        self.assertEqual(fields["slide_table_7"]["table_structure"], {"header_placement": "unspecified"})


if __name__ == "__main__":
    unittest.main()
