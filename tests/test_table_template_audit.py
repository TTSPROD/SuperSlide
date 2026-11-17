"""Whole-template table audit discovers grids from XML, including lost bindings."""
import copy
import unittest
from zipfile import ZipFile

from scripts.audit_severstal_template import audit_table_bindings, xml
from superslide.source_template import (
    NS, ROOT, SourceFieldError, _normalize_text_weight, _xml,
    load_source_manifest, source_prompt_catalog, validate_source_presentation,
)


class TableTemplateAuditTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.manifest = load_source_manifest(ROOT / "templates/severstal_template.json")
        with ZipFile(ROOT / cls.manifest["layouts"]["tables"]["source_file"]) as archive:
            cls.source8 = xml(archive.read("ppt/slides/slide8.xml"))

    def test_audit_detects_raw_table_even_when_its_manifest_binding_is_missing(self):
        fields = copy.deepcopy(self.manifest["layouts"]["tables"]["variants"]["slide_008"]["fields"])
        del fields["slide_table_18"]
        _, errors = audit_table_bindings(self.source8, fields, "slide")
        self.assertTrue(any("не описана" in error for error in errors))

    def test_audit_checks_external_header_links_against_the_grid(self):
        fields = copy.deepcopy(self.manifest["layouts"]["tables"]["variants"]["slide_008"]["fields"])
        self.assertEqual(audit_table_bindings(self.source8, fields, "slide")[1], [])
        fields["slide_text_10"]["table_column_header"]["column_index"] = 4
        _, errors = audit_table_bindings(self.source8, fields, "slide")
        self.assertTrue(any("неверная роль" in error for error in errors))

    def test_all_reviewed_drawn_tables_reach_model_catalog_with_true_row_counts(self):
        catalog = source_prompt_catalog(self.manifest)
        for lid, vid, columns, rows in (("tables", "slide_007", 4, 6),
                                         ("diagram", "slide_009", 3, 6),
                                         ("diagram", "slide_032", 4, 8)):
            with self.subTest(layout=lid, variant=vid):
                variant = catalog[lid]["variants"][vid]
                table, = variant["drawn_tables"]
                self.assertEqual(table["column_count"], columns)
                self.assertEqual(table["data_row_count"], rows)
                self.assertEqual(table["header_placement"], "external")
                self.assertEqual(len(table["column_header_field_ids"]), columns)
                for fid in table["column_header_field_ids"]:
                    self.assertIn("drawn_table_column_header", variant["fields"][fid])
                for cell in table["cells"]:
                    self.assertIn("drawn_table_cell", variant["fields"][cell["field_id"]])
        comparison = catalog["diagram"]["variants"]["slide_009"]["drawn_tables"][0]
        self.assertEqual(len(comparison["column_charts"]), 2)
        self.assertTrue(all(chart["row_span"] == 6 for chart in comparison["column_charts"]))

    def test_merged_continuations_cannot_accept_independent_values(self):
        variant = self.manifest["layouts"]["tables"]["variants"]["slide_003"]
        fid, field = next((fid, field) for fid, field in variant["fields"].items() if field["kind"] == "table")
        rows = copy.deepcopy(field["rows"])
        draft = {"template_version": "2.0", "title": "Проверка объединений", "slides": [{
            "layout_id": "tables", "variant_id": "slide_003", "updates": [{"field_id": fid, "rows": rows}],
        }]}
        validate_source_presentation(draft, self.manifest)
        row, column = field["table_structure"]["merged_continuation_cells"][0]
        rows[row][column] = "123"
        with self.assertRaises(SourceFieldError) as raised:
            validate_source_presentation(draft, self.manifest)
        self.assertEqual(raised.exception.validation_details["reason"], "table_merged_cell_content")
        self.assertEqual(raised.exception.validation_details["row_index"], row + 1)
        self.assertEqual(raised.exception.validation_details["column_index"], column + 1)

    def test_drawn_table_first_data_row_cannot_repeat_its_external_headers(self):
        catalog = source_prompt_catalog(self.manifest)
        for lid, vid in (("tables", "slide_007"), ("diagram", "slide_009"), ("diagram", "slide_032")):
            with self.subTest(layout=lid, variant=vid):
                fields = self.manifest["layouts"][lid]["variants"][vid]["fields"]
                table, = catalog[lid]["variants"][vid]["drawn_tables"]
                headers = {header["column_index"]: fields[header["field_id"]]["sample_text"] for header in table["headers"]}
                updates = [{"field_id": cell["field_id"], "text": headers[cell["column_index"]]}
                           for cell in table["cells"] if cell["kind"] == "text" and cell["row_index"] == 1
                           and cell["column_index"] in headers]
                draft = {"template_version": "2.0", "title": "Проверка заголовков", "slides": [{
                    "layout_id": lid, "variant_id": vid, "updates": updates,
                }]}
                with self.assertRaises(SourceFieldError) as raised:
                    validate_source_presentation(draft, self.manifest)
                self.assertEqual(raised.exception.validation_details["reason"], "table_header_in_body")
                self.assertEqual(len(raised.exception.validation_issues), len(updates))

    def test_external_column_titles_receive_heading_weight(self):
        with ZipFile(ROOT / "templates/severstal_source.pptx") as archive:
            for lid, vid in (("tables", "slide_007"), ("tables", "slide_008"),
                             ("diagram", "slide_009"), ("diagram", "slide_032")):
                with self.subTest(layout=lid, variant=vid):
                    variant = self.manifest["layouts"][lid]["variants"][vid]
                    root = _xml(archive.read(variant["package_part"]))
                    _normalize_text_weight(root, variant, "slide", {"title_field_id": None})
                    headers = [field for field in variant["fields"].values()
                               if field.get("table_column_header") or field.get("drawn_table_column_header")]
                    self.assertTrue(headers)
                    for field in headers:
                        props = root.xpath(f".//p:sp[p:nvSpPr/p:cNvPr[@id='{field['shape_id']}']]/p:txBody"
                                           "/a:p[1]/a:r/a:rPr", namespaces=NS)
                        self.assertTrue(props)
                        self.assertEqual({prop.get("b") for prop in props}, {"1"})


if __name__ == "__main__":
    unittest.main()
