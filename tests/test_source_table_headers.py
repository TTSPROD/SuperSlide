"""External column labels belong above the grid; its first row is data."""
import copy
import unittest
from pathlib import Path
from zipfile import ZipFile

from lxml import etree as E

from support import temporary_directory
from superslide.source_template import (
    A, NS, ROOT, SourceFieldError, _find_shape, _xml,
    build_source_presentation, load_source_manifest, source_prompt_catalog,
    validate_source_presentation,
)


HEADER_IDS = ["slide_text_10", "slide_text_5", "slide_text_20", "slide_text_7"]
HEADERS = ["Месяц", "Выручка", "Затраты", "Результат"]
MONTHS = [
    ["Январь", "12.5", "13.1", "-0.6"],
    ["Февраль", "14.75", "14.2", "0.55"],
    ["Март", "11.2", "12.05", "-0.85"],
    ["Апрель", "16.4", "15.6", "0.8"],
]


def external_slide(rows=None):
    return {"layout_id": "tables", "variant_id": "slide_008", "updates": [
        {"field_id": fid, "text": text} for fid, text in zip(HEADER_IDS, HEADERS)
    ] + [{"field_id": "slide_table_18", "rows": copy.deepcopy(rows if rows is not None else MONTHS + [[""] * 4])}]}


def presentation(slides):
    return {"template_version": "2.0", "title": "Учебные данные", "slides": slides}


class SourceTableHeaderTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.manifest_path = ROOT / "templates/severstal_template.json"
        cls.template = ROOT / "templates/severstal_source.pptx"
        cls.manifest = load_source_manifest(cls.manifest_path)
        cls.variant = cls.manifest["layouts"]["tables"]["variants"]["slide_008"]

    def test_external_header_contract_uses_visual_column_order(self):
        fields = self.variant["fields"]
        structure = fields["slide_table_18"]["table_structure"]
        self.assertEqual(structure["header_placement"], "external")
        self.assertEqual(structure["column_header_field_ids"], HEADER_IDS)
        self.assertEqual(structure["header_row_count"], 0)
        self.assertEqual(structure["data_row_count"], 5)
        for column, fid in enumerate(HEADER_IDS, 1):
            self.assertEqual(fields[fid]["table_column_header"], {
                "table_field_id": "slide_table_18", "column_index": column,
            })
        # firstRow is present in the native design for styling, but does not
        # imply a semantic header row in this particular source layout.
        with ZipFile(self.template) as archive:
            original = _xml(archive.read(self.variant["package_part"]))
        self.assertEqual(original.find(f".//{{{A}}}tbl/{{{A}}}tblPr").get("firstRow"), "1")

    def test_generation_catalog_exposes_header_binding_and_body_capacity(self):
        variant = source_prompt_catalog(self.manifest)["tables"]["variants"]["slide_008"]
        fields = variant["fields"]
        self.assertEqual(fields["slide_table_18"]["table_structure"],
                         self.variant["fields"]["slide_table_18"]["table_structure"])
        for fid in HEADER_IDS:
            self.assertEqual(fields[fid]["table_column_header"],
                             self.variant["fields"][fid]["table_column_header"])

    def test_duplicate_external_column_names_cannot_take_first_data_row(self):
        draft = presentation([external_slide([HEADERS] + MONTHS)])
        with self.assertRaises(SourceFieldError) as caught:
            validate_source_presentation(draft, self.manifest)
        self.assertEqual(caught.exception.validation_details["reason"], "table_header_in_body")
        self.assertEqual(caught.exception.validation_details["slide_index"], 1)
        self.assertEqual(caught.exception.validation_details["field_id"], "slide_table_18")

    def test_external_first_row_is_regular_data_and_native_grid_is_preserved(self):
        draft = presentation([external_slide()])
        validate_source_presentation(draft, self.manifest)
        with ZipFile(self.template) as archive:
            original = _xml(archive.read(self.variant["package_part"]))
        with temporary_directory() as temporary:
            output = Path(temporary) / "external-column-headers.pptx"
            build_source_presentation(draft, self.template, self.manifest_path, output)
            with ZipFile(output) as archive:
                generated = _xml(archive.read("ppt/generated/slide1.xml"))
        table = generated.find(f".//{{{A}}}tbl")
        source_table = original.find(f".//{{{A}}}tbl")
        actual = [["".join(cell.xpath(".//a:t/text()", namespaces=NS))
                   for cell in row.findall(f"{{{A}}}tc")]
                  for row in table.findall(f"{{{A}}}tr")]
        self.assertEqual(actual, MONTHS + [[""] * 4])
        self.assertEqual(table.xpath("./a:tr/@h", namespaces=NS),
                         source_table.xpath("./a:tr/@h", namespaces=NS))
        self.assertEqual(E.tostring(table.find(f"{{{A}}}tblGrid")),
                         E.tostring(source_table.find(f"{{{A}}}tblGrid")))
        generated_frame = _find_shape(generated, 18)
        original_frame = _find_shape(original, 18)
        self.assertEqual(E.tostring(generated_frame.find(f"{{{NS['p']}}}xfrm")),
                         E.tostring(original_frame.find(f"{{{NS['p']}}}xfrm")))
        properties = table.findall(f"{{{A}}}tr")[0].xpath(
            ".//a:rPr | .//a:defRPr | .//a:endParaRPr", namespaces=NS)
        self.assertTrue(properties)
        self.assertEqual({node.get("b") for node in properties}, {"0"})
        for fid, expected in zip(HEADER_IDS, HEADERS):
            sid = self.variant["fields"][fid]["shape_id"]
            shape = _find_shape(generated, sid)
            self.assertEqual("".join(shape.xpath("./p:txBody//a:t/text()", namespaces=NS)), expected)
            self.assertEqual(E.tostring(shape.find(f"{{{NS['p']}}}spPr")),
                             E.tostring(_find_shape(original, sid).find(f"{{{NS['p']}}}spPr")))

    def test_genuine_inside_headers_remain_inside_the_grid(self):
        variant = self.manifest["layouts"]["tables"]["variants"]["slide_001"]
        fid, field = next((fid, field) for fid, field in variant["fields"].items()
                          if field["kind"] == "table")
        self.assertEqual(field["table_structure"]["header_placement"], "inside")
        rows = copy.deepcopy(field["rows"])
        rows[0] = list(HEADERS)
        slide = {"layout_id": "tables", "variant_id": "slide_001", "updates": [
            {"field_id": fid, "rows": rows},
        ]}
        self.assertEqual(validate_source_presentation(presentation([slide]), self.manifest)
                         .slides[0].updates[0].rows[0], HEADERS)
        with temporary_directory() as temporary:
            output = Path(temporary) / "inside-column-headers.pptx"
            build_source_presentation(presentation([slide]), self.template, self.manifest_path, output)
            with ZipFile(output) as archive:
                generated = _xml(archive.read("ppt/generated/slide1.xml"))
        actual_rows = generated.find(f".//{{{A}}}tbl").findall(f"{{{A}}}tr")
        for index, expected in ((0, "1"), (1, "0")):
            properties = actual_rows[index].xpath(
                ".//a:rPr | .//a:defRPr | .//a:endParaRPr", namespaces=NS)
            self.assertTrue(properties)
            self.assertEqual({node.get("b") for node in properties}, {expected})


if __name__ == "__main__":
    unittest.main()
