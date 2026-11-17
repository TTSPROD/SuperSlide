"""Generated text uses coherent paragraph styles and heading/body weights."""
import hashlib
import json
import unittest
from pathlib import Path
from zipfile import ZipFile

from lxml import etree as E

from support import temporary_directory
from superslide.source_template import (
    A, NS, ROOT, _paragraph_text, _set_paragraph_bold, _write_body, _write_paragraph, _xml,
    build_source_presentation,
)
from superslide.source_text import paragraph_is_bold, paragraph_roles, source_heading_fields


class SourceTypographyTests(unittest.TestCase):
    def test_replacement_chooses_main_text_instead_of_counter_or_dot_leader(self):
        for decoration in ("XX", ". " * 100):
            with self.subTest(decoration=decoration[:10]):
                paragraph = _xml(f'''<a:p xmlns:a="{A}">
                  <a:pPr algn="r" marL="100"><a:buNone/></a:pPr>
                  <a:r><a:rPr sz="2800"><a:latin typeface="Georgia"/>
                    <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
                  </a:rPr><a:t>{decoration}</a:t></a:r>
                  <a:r><a:rPr sz="1600" lang="en-US" err="1" dirty="1"><a:latin typeface="Arial"/>
                    <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
                  </a:rPr><a:t>Sample label</a:t></a:r>
                  <a:endParaRPr sz="2800"/>
                </a:p>'''.encode())
                layout = E.tostring(paragraph.find(f"{{{A}}}pPr"))
                text = "Руководитель\vнаправления"
                _write_paragraph(paragraph, text)
                self.assertEqual(_paragraph_text(paragraph), text)
                self.assertEqual(E.tostring(paragraph.find(f"{{{A}}}pPr")), layout)
                for props in paragraph.xpath("./a:r/a:rPr | ./a:br/a:rPr | ./a:endParaRPr", namespaces=NS):
                    self.assertEqual(props.get("sz"), "1600")
                    self.assertEqual(props.get("lang"), "ru-RU")
                    self.assertIsNone(props.get("err"))
                    self.assertIsNone(props.get("dirty"))
                    self.assertEqual(props.find(f"{{{A}}}latin").get("typeface"), "Arial")
                    self.assertEqual(props.find(f"{{{A}}}solidFill/{{{A}}}srgbClr").get("val"), "FFFFFF")

    def test_main_style_combines_fragments_split_only_by_spelling_or_language(self):
        paragraph = _xml(f'''<a:p xmlns:a="{A}">
          <a:r><a:rPr sz="2800"/><a:t>Counter</a:t></a:r>
          <a:r><a:rPr sz="1600" lang="ru-RU"/><a:t>First</a:t></a:r>
          <a:r><a:rPr sz="1600" lang="en-US" err="1"/><a:t>second</a:t></a:r>
        </a:p>'''.encode())
        _write_paragraph(paragraph, "Новая подпись")
        self.assertEqual(paragraph.xpath("./a:r/a:rPr/@sz", namespaces=NS), ["1600"])

    def test_replacement_preserves_different_sizes_between_heading_body_paragraphs(self):
        body = _xml(f'''<a:txBody xmlns:a="{A}">
          <a:bodyPr/><a:lstStyle/>
          <a:p><a:r><a:rPr sz="2400"/><a:t>Heading</a:t></a:r></a:p>
          <a:p><a:r><a:rPr sz="1600"/><a:t>Body text</a:t></a:r></a:p>
        </a:txBody>'''.encode())
        _write_body(body, ["Заголовок", "Основной текст"])
        self.assertEqual(body.xpath("./a:p/a:r/a:rPr/@sz", namespaces=NS), ["2400", "1600"])
        self.assertEqual([_paragraph_text(p) for p in body.findall(f"{{{A}}}p")], ["Заголовок", "Основной текст"])

    def test_main_style_can_inherit_paragraph_properties_without_using_end_font(self):
        paragraph = _xml(f'''<a:p xmlns:a="{A}">
          <a:pPr><a:defRPr sz="1600"><a:latin typeface="Arial"/></a:defRPr></a:pPr>
          <a:r><a:t>The actual label</a:t></a:r>
          <a:r><a:rPr sz="2800"/><a:t>XX</a:t></a:r>
          <a:endParaRPr sz="9900"/>
        </a:p>'''.encode())
        _write_paragraph(paragraph, "Новая подпись")
        self.assertEqual(paragraph.xpath("./a:pPr/a:defRPr/@sz", namespaces=NS), ["1600"])
        self.assertEqual(paragraph.xpath("./a:pPr/a:defRPr/a:latin/@typeface", namespaces=NS), ["Arial"])
        self.assertFalse(paragraph.xpath("./a:r/a:rPr/@sz | ./a:endParaRPr/@sz", namespaces=NS))
        self.assertEqual(_paragraph_text(paragraph), "Новая подпись")

    def test_real_contents_and_table_labels_do_not_change_style_mid_word(self):
        template = ROOT / "templates/severstal_source.pptx"
        manifest_path = ROOT / "templates/severstal_template.json"
        labels = ["Роль", "Автор", "Руководитель", "Диспетчер", "Исполнитель"]
        rows = [[label, "Этап процесса", "Действие", "Результат"] for label in labels]
        spec = {"template_version": "2.0", "title": "Роли и ответственность", "slides": [
            {"layout_id": "contents", "variant_id": "slide_005", "updates": [
                {"field_id": "slide_text_4", "text": "Роли и ответственность"},
                {"field_id": "slide_text_5", "text": "Учебный характер данных"},
            ]},
            {"layout_id": "tables", "variant_id": "slide_008", "updates": [
                {"field_id": "slide_table_18", "rows": rows},
            ]},
        ]}
        before = hashlib.sha256(template.read_bytes()).hexdigest()
        with temporary_directory() as directory:
            output = Path(directory) / "consistent-paragraphs.pptx"
            build_source_presentation(spec, template, manifest_path, output)
            with ZipFile(output) as archive:
                contents = _xml(archive.read("ppt/generated/slide1.xml"))
                for sid, text in ((4, "Роли и ответственность"), (5, "Учебный характер данных")):
                    paragraph = contents.xpath(f".//p:sp[p:nvSpPr/p:cNvPr[@id='{sid}']]/p:txBody/a:p", namespaces=NS)[0]
                    self.assertEqual(_paragraph_text(paragraph), text)
                    self.assertEqual(paragraph.xpath("./a:r/a:rPr/@sz", namespaces=NS), ["1800"])
                    self.assertEqual(paragraph.xpath("./a:r/a:rPr/a:solidFill/a:srgbClr/@val", namespaces=NS), ["FFFFFF"])
                table = _xml(archive.read("ppt/generated/slide2.xml")).find(f".//{{{A}}}tbl")
                for row, label in zip(table.findall(f"{{{A}}}tr"), labels):
                    paragraph = row.find(f"{{{A}}}tc/{{{A}}}txBody/{{{A}}}p")
                    self.assertEqual(_paragraph_text(paragraph), label)
                    self.assertEqual(paragraph.xpath("./a:r/a:rPr/@sz", namespaces=NS), ["1600"])
                    self.assertEqual(paragraph.xpath("./a:r/a:rPr/a:latin/@typeface", namespaces=NS), ["Arial"])
        self.assertEqual(hashlib.sha256(template.read_bytes()).hexdigest(), before)

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
