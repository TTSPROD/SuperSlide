import copy
import io
import json
import tempfile
import unittest
from xml.etree import ElementTree as ET
from pathlib import Path
from zipfile import ZipFile

from pptx import Presentation
from pydantic import ValidationError

from superslide.models import validate_presentation
from superslide.render import LayoutOverflowError, build_presentation, check_text_fit
from superslide.template import TemplateError, load_template


ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "templates/corporate_base.pptx"
MANIFEST = ROOT / "templates/template_manifest.json"


class BuildTests(unittest.TestCase):
    def setUp(self):
        self.workspace = tempfile.TemporaryDirectory(dir=ROOT)
        self.addCleanup(self.workspace.cleanup)
        self.work = Path(self.workspace.name)
        self.payload = json.loads((ROOT / "examples/presentation.json").read_text(encoding="utf-8"))

    def build(self, payload=None, template=TEMPLATE, manifest=MANIFEST):
        out = self.work / "test.pptx"
        report = build_presentation(validate_presentation(payload or self.payload), template, manifest, out)
        return Presentation(out), out, report

    def test_all_layouts_and_native_data_roundtrip(self):
        prs, out, _ = self.build()
        self.assertEqual(len(prs.slides), 6)
        self.assertEqual([s.slide_layout.name for s in prs.slides], [s["layout_id"] for s in self.payload["slides"]])
        table_shape = next(s for s in prs.slides[4].shapes if s.has_table)
        self.assertEqual(table_shape.table.cell(4, 1).text, "171")
        chart = next(s.chart for s in prs.slides[5].shapes if s.has_chart)
        self.assertEqual(list(chart.series[0].values), [120.0, 138.0, 156.0, 171.0])
        # This must be a real workbook, not a screenshot styled to look like a chart.
        with ZipFile(io.BytesIO(chart.part.chart_workbook.xlsx_part.blob)) as book:
            ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
            sheet = ET.fromstring(book.read("xl/worksheets/sheet1.xml"))
            strings = ET.fromstring(book.read("xl/sharedStrings.xml"))
            strings = ["".join(item.itertext()) for item in strings]
            cells = {c.attrib["r"]: c for c in sheet.findall(".//s:c", ns)}
            self.assertEqual(strings[int(cells["B1"].find("s:v", ns).text)], "Компания «Пример»")
            self.assertEqual([float(cells[f"B{r}"].find("s:v", ns).text) for r in range(2, 6)], [120, 138, 156, 171])
            self.assertFalse(sheet.findall(".//s:f", ns))
        with ZipFile(out) as archive:
            self.assertEqual(sum(n.startswith("ppt/charts/chart") and n.endswith(".xml") for n in archive.namelist()), 1)
            self.assertFalse(any("oleObject" in n or "thinkcell" in n.lower() for n in archive.namelist()))
        self.assertIn("ДЕМОНСТРАЦИОННЫЙ", prs.slides[5].notes_slide.notes_text_frame.text)

    def test_table_remains_within_placeholder_for_unequal_columns(self):
        self.payload["slides"][4]["columns"] = ["Очень длинное название показателя", "A", "B"]
        prs, _, _ = self.build()
        table = next(s for s in prs.slides[4].shapes if s.has_table)
        field = json.loads(MANIFEST.read_text(encoding="utf-8"))["layouts"]["table"]["fields"]["table"]
        self.assertEqual(table.width, field["width"])
        self.assertLessEqual(table.left + table.width, prs.slide_width)

    def test_sparse_placeholder_indices_are_identifiers(self):
        prs = Presentation(TEMPLATE)
        layout = next(l for l in prs.slide_layouts if l.name == "bullet_list")
        placeholder = next(p for p in layout.placeholders if p.placeholder_format.idx == 1)
        placeholder.element.xpath(".//p:ph")[0].set("idx", "10")
        template = self.work / "sparse.pptx"
        prs.save(template)
        manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
        manifest["layouts"]["bullet_list"]["fields"]["bullets"]["idx"] = 10
        manifest_path = self.work / "manifest.json"
        manifest_path.write_text(json.dumps(manifest), encoding="utf-8")
        built, _, _ = self.build(template=template, manifest=manifest_path)
        self.assertIn("За четыре", built.slides[1].placeholders[10].text)

    def test_latent_placeholders_rejected_before_build(self):
        prs = Presentation(TEMPLATE)
        placeholder = prs.slide_layouts[1].placeholders[1]
        placeholder.element.xpath(".//p:ph")[0].set("type", "ftr")
        template = self.work / "footer.pptx"
        prs.save(template)
        with self.assertRaisesRegex(TemplateError, "служебный"):
            load_template(template, MANIFEST)

    def test_unknown_sources_rejected(self):
        self.payload["slides"][1]["source_refs"] = ["invented"]
        with self.assertRaisesRegex(ValueError, "Неизвестные источники"):
            validate_presentation(self.payload, {"doc1:p1"})

    def test_invalid_table_chart_and_formula_labels_rejected(self):
        for mutate in (
            lambda p: p["slides"][4]["rows"][0].pop(),
            lambda p: p["slides"][5]["series"][0]["values"].pop(),
            lambda p: p["slides"][5]["categories"].__setitem__(0, "=1+1"),
            lambda p: p["slides"][5]["series"][0].__setitem__("name", "{=2+2}"),
            lambda p: p["slides"][5]["series"][0]["values"].__setitem__(0, float("nan")),
        ):
            payload = copy.deepcopy(self.payload)
            mutate(payload)
            with self.assertRaises(ValidationError):
                validate_presentation(payload)

    def test_failed_fit_does_not_replace_existing_output(self):
        payload = copy.deepcopy(self.payload)
        payload["slides"][2]["left"]["heading"] = "Очень длинное название " * 4
        out = self.work / "test.pptx"
        out.write_bytes(b"previous result")
        with self.assertRaises(LayoutOverflowError):
            build_presentation(validate_presentation(payload), TEMPLATE, MANIFEST, out)
        self.assertEqual(out.read_bytes(), b"previous result")

    def test_check_text_fit_is_a_dry_run_without_output_files(self):
        long_bullet = ("Подробное описание направления развития производства с указанием "
                       "конкретных мероприятий, сроков и ответственных подразделений " * 2)[:130]
        payload = copy.deepcopy(self.payload)
        payload["slides"][1]["bullets"] = [long_bullet] * 6
        with self.assertRaisesRegex(LayoutOverflowError, r"Слайд 2, поле bullets"):
            check_text_fit(validate_presentation(payload), TEMPLATE, MANIFEST)
        self.assertFalse(list(self.work.glob("*.pptx")))
        report = check_text_fit(validate_presentation(self.payload), TEMPLATE, MANIFEST)
        self.assertFalse(list(self.work.glob("*.pptx")))
        self.assertNotIn("output", report)


if __name__ == "__main__":
    unittest.main()
