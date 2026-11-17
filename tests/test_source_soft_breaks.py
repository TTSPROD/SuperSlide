"""Source text edits retain untouched soft breaks and write consistent new text."""
import unittest
from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile

from lxml import etree as E

from support import temporary_directory
from superslide.models import _xml_text
from superslide.source_template import (
    ROOT, A, NS, SourceUpdate, _paragraph_text, _write_paragraph, _xml,
    build_source_presentation, load_source_manifest, validate_source_presentation,
)


def paragraph():
    return _xml(f'''<a:p xmlns:a="{A}">
      <a:pPr algn="l"/>
      <a:r><a:rPr b="1" sz="2400"/><a:t>ABC</a:t></a:r>
      <a:r><a:rPr sz="1800"><a:solidFill><a:srgbClr val="003377"/></a:solidFill></a:rPr><a:t>defgh</a:t></a:r>
      <a:endParaRPr lang="ru-RU" sz="1800"/>
    </a:p>'''.encode())


def without_text_bold(node):
    """Compare original text and styling while allowing explicit font weight."""
    result = deepcopy(node)
    for properties in result.xpath(".//a:rPr | .//a:defRPr | .//a:endParaRPr", namespaces=NS):
        properties.attrib.pop("b", None)
        if not properties.attrib and not len(properties) and not properties.text:
            properties.getparent().remove(properties)
    return E.tostring(result, method="c14n")


class SourceSoftBreakTests(unittest.TestCase):
    def test_source_update_accepts_soft_breaks_without_relaxing_other_controls(self):
        for operation in (
            {"text": "Первая\vвторая строка"},
            {"paragraphs": ["Заголовок", "Первое\vвторое предложение"]},
            {"rows": [["Первая\vвторая строка"]]},
        ):
            with self.subTest(operation=next(iter(operation))):
                update = SourceUpdate(field_id="slide_text_15", **operation)
                self.assertEqual(getattr(update, next(iter(operation))), next(iter(operation.values())))
        for character in ("\x00", "\x01", "\x0c", "\x1f", "\ud800", "\ufffe", "\uffff"):
            with self.subTest(character=repr(character)), self.assertRaises(ValueError):
                SourceUpdate(field_id="slide_text_15", paragraphs=["Текст" + character])
        # Legacy contracts still disallow vertical tabs as literal XML text.
        with self.assertRaises(ValueError):
            _xml_text("Первая\vвторая строка")

    def test_editing_real_source_heading_preserves_body_except_explicit_text_weight(self):
        manifest_path = ROOT / "templates/severstal_template.json"
        template = ROOT / "templates/severstal_source.pptx"
        manifest = load_source_manifest(manifest_path)
        variant = manifest["layouts"]["title"]["variants"]["slide_010"]
        field = variant["fields"]["slide_text_15"]
        self.assertIn("\v", field["paragraphs"][1])
        spec = {"template_version": "2.0", "title": "Проверка заголовка", "slides": [{
            "layout_id": "title", "variant_id": "slide_010", "updates": [{
                "field_id": "slide_text_15", "paragraphs": ["Анна Соколова", *field["paragraphs"][1:]],
            }],
        }]}
        validate_source_presentation(spec, manifest)
        with ZipFile(template) as source:
            root = _xml(source.read(variant["package_part"]))
        xpath = ".//p:sp[p:nvSpPr/p:cNvPr[@id='15']]/p:txBody/a:p"
        original = root.xpath(xpath, namespaces=NS)
        before = without_text_bold(original[1])
        with temporary_directory() as directory:
            output = Path(directory) / "edited-heading.pptx"
            build_source_presentation(spec, template, manifest_path, output)
            with ZipFile(output) as archive:
                updated = _xml(archive.read("ppt/generated/slide1.xml")).xpath(xpath, namespaces=NS)
        self.assertEqual(_paragraph_text(updated[0]), "Анна Соколова")
        self.assertEqual(_paragraph_text(updated[1]), field["paragraphs"][1])
        self.assertEqual(without_text_bold(updated[1]), before)
        self.assertEqual(len(updated[1].findall(f"{{{A}}}br")), 1)
        for index, expected_bold in ((0, "1"), (1, "0")):
            for properties in updated[index].xpath("./a:r/a:rPr | ./a:br/a:rPr", namespaces=NS):
                self.assertEqual(properties.get("b"), expected_bold)

    def test_changed_soft_breaks_become_styled_native_breaks_before_end_properties(self):
        node = paragraph()
        text = "AB\vCD\vEF"
        _write_paragraph(node, text)
        serialized = E.tostring(node)
        self.assertNotIn(b"\x0b", serialized)
        self.assertEqual(_paragraph_text(_xml(serialized)), text)
        breaks = node.findall(f"{{{A}}}br")
        self.assertEqual(len(breaks), 2)
        for br in breaks:
            self.assertEqual(br.find(f"{{{A}}}rPr").get("sz"), "1800")
            self.assertEqual(br.find(f"{{{A}}}rPr/{{{A}}}solidFill/{{{A}}}srgbClr").get("val"), "003377")
        self.assertIs(node[-1], node.find(f"{{{A}}}endParaRPr"))
        # Reapplying the semantic text does not rebuild its styled runs/breaks.
        _write_paragraph(node, text)
        self.assertEqual(E.tostring(node), serialized)

    def test_new_soft_breaks_in_empty_paragraph_inherit_end_font_properties(self):
        node = _xml(f'<a:p xmlns:a="{A}"><a:pPr/><a:endParaRPr lang="ru-RU" sz="2200"/></a:p>'.encode())
        _write_paragraph(node, "Первая\v\vпоследняя\v")
        self.assertEqual(_paragraph_text(node), "Первая\v\vпоследняя\v")
        self.assertEqual(len(node.findall(f"{{{A}}}br")), 3)
        self.assertIs(node[-1], node.find(f"{{{A}}}endParaRPr"))
        self.assertTrue(all(run.find(f"{{{A}}}rPr").get("sz") == "2200"
                            for run in node.findall(f"{{{A}}}r")))
        _xml(E.tostring(node))

    def test_edit_after_soft_break_boundary_keeps_one_consistent_style(self):
        node = paragraph()
        _write_paragraph(node, "AB\vCD\vEF")
        self.assertEqual([run.findtext(f"{{{A}}}t", "") for run in node.findall(f"{{{A}}}r")], ["AB", "CD", "EF"])
        _write_paragraph(node, "123456789")
        self.assertEqual(_paragraph_text(node), "123456789")
        runs = node.findall(f"{{{A}}}r")
        self.assertEqual(len(runs), 1)
        self.assertEqual(runs[0].find(f"{{{A}}}rPr").get("sz"), "1800")
        self.assertFalse(node.findall(f"{{{A}}}br"))

    def test_replacing_existing_soft_breaks_removes_old_breaks_and_styles_new_ones(self):
        node = paragraph()
        old_break = E.Element(f"{{{A}}}br")
        E.SubElement(old_break, f"{{{A}}}rPr", sz="9900")
        node.insert(2, old_break)
        self.assertEqual(_paragraph_text(node), "ABC\vdefgh")
        _write_paragraph(node, "AB\vCD\vEF")
        self.assertEqual(_paragraph_text(node), "AB\vCD\vEF")
        self.assertNotIn(old_break, node)
        breaks = node.findall(f"{{{A}}}br")
        self.assertEqual([br.find(f"{{{A}}}rPr").get("sz") for br in breaks], ["1800", "1800"])
        _write_paragraph(node, "")
        self.assertEqual(_paragraph_text(node), "")
        self.assertFalse(node.findall(f"{{{A}}}br"))
        self.assertIs(node[-1], node.find(f"{{{A}}}endParaRPr"))

    def test_normal_changed_text_does_not_reuse_original_run_boundaries(self):
        node = paragraph()
        _write_paragraph(node, "1234567890")
        runs = node.findall(f"{{{A}}}r")
        self.assertEqual([run.findtext(f"{{{A}}}t") for run in runs], ["1234567890"])
        self.assertEqual(runs[0].find(f"{{{A}}}rPr").get("sz"), "1800")
        self.assertFalse(node.findall(f"{{{A}}}br"))
        self.assertIs(node[-1], node.find(f"{{{A}}}endParaRPr"))


if __name__ == "__main__":
    unittest.main()
