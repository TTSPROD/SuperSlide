"""Generated slides use consistent heading/body weight instead of run ratios."""
import hashlib
import json
import unittest
from pathlib import Path
from zipfile import ZipFile

from lxml import etree as E

from support import temporary_directory
from superslide.source_template import (
    A, NS, ROOT, _paragraph_text, _set_paragraph_bold, _xml,
    build_source_presentation,
)
from superslide.source_text import paragraph_is_bold, paragraph_roles, source_heading_fields


class SourceTypographyTests(unittest.TestCase):
    def test_explicit_heading_role_does_not_depend_on_original_bold(self):
        paragraph = f'<a:p xmlns:a="{A}"><a:r><a:rPr b="0" sz="1600"/><a:t>Заголовок блока</a:t></a:r></a:p>'
        field = {"kind": "text", "scope": "slide", "shape_id": 4,
                 "sample_text": "Заголовок блока", "paragraph_styles": [paragraph]}
        root = _xml(f'''<p:sld xmlns:p="{NS['p']}" xmlns:a="{A}"><p:cSld><p:spTree><p:sp>
          <p:nvSpPr><p:cNvPr id="4" name="TextBox"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:txBody>{paragraph}</p:txBody>
        </p:sp></p:spTree></p:cSld></p:sld>'''.encode())
        self.assertEqual(source_heading_fields({"fields": {"slide_text_4": field}}, root), {"slide_text_4"})

    def test_inline_bold_prefix_does_not_turn_body_into_a_heading(self):
        paragraphs = [
            f'<a:p xmlns:a="{A}"><a:r><a:rPr sz="1800" b="1"/><a:t>Важный факт: </a:t></a:r>'
            '<a:r><a:rPr sz="1800" b="0"/><a:t>описание продолжается обычным текстом</a:t></a:r></a:p>',
            f'<a:p xmlns:a="{A}"><a:r><a:rPr sz="1800" b="0"/><a:t>Второй абзац</a:t></a:r></a:p>',
        ]
        field = {"paragraph_styles": paragraphs}
        self.assertIsNone(paragraph_roles(field, ["Важный факт: описание", "Второй абзац"]))
        self.assertFalse(paragraph_is_bold(_xml(paragraphs[0].encode())))

    def test_explicit_weight_overrides_inherited_and_inline_bold_including_breaks(self):
        paragraph = _xml(f'''<a:p xmlns:a="{A}">
          <a:pPr><a:defRPr b="1" sz="1800"/></a:pPr>
          <a:r><a:rPr b="1" i="1"/><a:t>Первая</a:t></a:r>
          <a:br><a:rPr b="1"/></a:br>
          <a:r><a:t>вторая</a:t></a:r>
          <a:fld id="field-1" type="slidenum"><a:rPr b="1"/><a:t>2</a:t></a:fld>
          <a:endParaRPr b="1" lang="ru-RU"/>
        </a:p>'''.encode())
        original_text = _paragraph_text(paragraph)
        for bold in (False, True, False):
            _set_paragraph_bold(paragraph, bold)
            expected = "1" if bold else "0"
            properties = paragraph.xpath("./a:pPr/a:defRPr | ./a:r/a:rPr | ./a:br/a:rPr | ./a:fld/a:rPr | ./a:endParaRPr", namespaces=NS)
            self.assertEqual(len(properties), 6)
            self.assertTrue(all(prop.get("b") == expected for prop in properties))
            self.assertEqual(_paragraph_text(paragraph), original_text)
            self.assertEqual(paragraph.find(f"{{{A}}}r/{{{A}}}rPr").get("i"), "1")
            self.assertEqual(paragraph.find(f"{{{A}}}pPr/{{{A}}}defRPr").get("sz"), "1800")

    def test_real_card_headings_are_bold_and_all_body_runs_are_regular(self):
        template = ROOT / "templates/severstal_source.pptx"
        manifest_path = ROOT / "templates/severstal_template.json"
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        variant = manifest["layouts"]["cards"]["variants"]["slide_012"]
        before = hashlib.sha256(template.read_bytes()).hexdigest()
        # Both an edited and an unchanged instance must obey the same policy.
        spec = {"template_version": "2.0", "title": "Системы хранения", "slides": [
            {"layout_id": "cards", "variant_id": "slide_012", "updates": [
                {"field_id": "slide_text_15", "text": "Шаттлы 2D и 3D"},
                {"field_id": "slide_text_21", "paragraphs": ["Высокая скорость в каналах.", "Шаттлы перемещают паллеты."]},
            ]},
            {"layout_id": "cards", "variant_id": "slide_012", "updates": []},
            {"layout_id": "cards", "variant_id": "slide_017", "updates": []},
        ]}
        with temporary_directory() as directory:
            result = Path(directory) / "consistent-weight.pptx"
            report = build_source_presentation(spec, template, manifest_path, result)
            self.assertEqual(report["text_weight_policy"], "headings_bold_body_regular")
            with ZipFile(result) as archive:
                for number in (1, 2):
                    root = _xml(archive.read(f"ppt/generated/slide{number}.xml"))
                    for field_id, expected in {
                        "slide_text_4": "1", "slide_text_15": "1", "slide_text_53": "1", "slide_text_60": "1",
                        "slide_text_21": "0", "slide_text_54": "0", "slide_text_61": "0",
                    }.items():
                        with self.subTest(slide=number, field=field_id):
                            sid = variant["fields"][field_id]["shape_id"]
                            body = root.xpath(f".//p:sp[p:nvSpPr/p:cNvPr[@id='{sid}']]/p:txBody", namespaces=NS)[0]
                            properties = body.xpath("./a:p/a:r/a:rPr | ./a:p/a:fld/a:rPr | ./a:p/a:br/a:rPr | ./a:p/a:pPr/a:defRPr | ./a:p/a:endParaRPr", namespaces=NS)
                            self.assertTrue(properties)
                            self.assertTrue(all(node.get("b") == expected for node in properties))
                # Some source card descriptions are entirely bold; that alone
                # does not make them headings or permit bold in the result.
                root = _xml(archive.read("ppt/generated/slide3.xml"))
                for sid in (24, 36):
                    body = root.xpath(f".//p:sp[p:nvSpPr/p:cNvPr[@id='{sid}']]/p:txBody", namespaces=NS)[0]
                    properties = body.xpath("./a:p/a:r/a:rPr | ./a:p/a:endParaRPr", namespaces=NS)
                    self.assertTrue(properties)
                    self.assertTrue(all(node.get("b") == "0" for node in properties))
        self.assertEqual(hashlib.sha256(template.read_bytes()).hexdigest(), before)


if __name__ == "__main__":
    unittest.main()
