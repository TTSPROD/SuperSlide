"""Native edits must retain artwork, isolate repeated slides and sync workbooks."""
import hashlib
import json
import tempfile
import unittest
from copy import deepcopy
from lxml import etree as E
from pathlib import Path
from zipfile import ZipFile

from superslide.source_template import (
    ROOT, P, A, R, load_source_manifest, validate_source_presentation,
    build_source_presentation, _xml, _relations, _target,
)
from superslide.source_charts import inspect_chart


class SourceNativeTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.manifest_path = ROOT / "templates/severstal_template.json"
        cls.template = ROOT / "templates/severstal_source.pptx"
        cls.manifest = load_source_manifest(cls.manifest_path)

    def test_actual_edits_and_duplicate_slide_isolation(self):
        spec = json.loads((ROOT / "examples/severstal_editable.json").read_text(encoding="utf-8"))
        spec["slides"][0]["notes"] = "Пояснение докладчика"
        spec["slides"][0]["source_refs"] = ["source-1"]
        original_variant = self.manifest["layouts"]["schemes"]["variants"]["slide_001"]
        spec["slides"].append({"layout_id": "schemes", "variant_id": "slide_001", "updates": []})
        inherited = next((lid, vid, v, fid) for lid, layout in self.manifest["layouts"].items()
                         for vid, v in layout["variants"].items() for fid, f in v["fields"].items()
                         if f["scope"] == "layout" and f["kind"] == "text")
        lid, vid, inherited_variant, fid = inherited
        spec["slides"].extend([
            {"layout_id": lid, "variant_id": vid, "updates": [{"field_id": fid, "text": "Тест"}]},
            {"layout_id": lid, "variant_id": vid, "updates": []},
        ])
        with tempfile.TemporaryDirectory() as temporary:
            output = Path(temporary) / "edited.pptx"
            build_source_presentation(spec, self.template, self.manifest_path, output)
            with ZipFile(output) as archive:
                package = {n: archive.read(n) for n in archive.namelist()}
            original = package["ppt/generated/slide6.xml"]
            # The automatic footer is the only change on this untouched slide.
            with ZipFile(self.template) as source:
                expected = _xml(source.read(original_variant["package_part"]))
            actual = _xml(original)
            for shape in expected.xpath(".//p:sp[p:nvSpPr/p:nvPr/p:ph]", namespaces={"p": P}):
                body = shape.find(f"{{{P}}}txBody")
                if body is not None and not "".join(body.itertext()).strip():
                    sid = shape.find(f"{{{P}}}nvSpPr/{{{P}}}cNvPr").get("id")
                    target = actual.xpath(f".//p:sp[p:nvSpPr/p:cNvPr[@id='{sid}']]", namespaces={"p": P})[0]
                    target.replace(target.find(f"{{{P}}}txBody"), deepcopy(body))
            self.assertEqual(E.tostring(actual, method="c14n"), E.tostring(expected, method="c14n"))
            changed = _xml(package["ppt/generated/slide5.xml"])
            self.assertTrue(changed.xpath(".//*[local-name()='svgBlip']"))
            self.assertFalse(_xml(original).xpath(".//*[local-name()='svgBlip']"))
            title_text = "".join(_xml(package["ppt/generated/slide1.xml"]).xpath(".//a:t/text()", namespaces={"a": A}))
            self.assertIn("Результаты команды", title_text)
            self.assertIn("Пояснение докладчика", package["ppt/generated/notes1.xml"].decode())
            self.assertIn("source-1", package["ppt/generated/notes1.xml"].decode())
            self.assertIn("Показатель", package["ppt/generated/slide3.xml"].decode())
            chart_update = spec["slides"][3]["updates"][0]
            chart_part = "ppt/generated/chart4_" + chart_update["field_id"].rsplit("_", 1)[1] + ".xml"
            chart = inspect_chart(package[chart_part])
            self.assertEqual(chart["categories"], chart_update["chart_data"]["categories"])
            self.assertEqual(chart["series"][0]["values"], chart_update["chart_data"]["series"][0]["values"])
            self.assertTrue(any(p.startswith("ppt/generated/workbook4_") for p in package))
            # An SVG is carried with a PNG fallback, sharing no old image link.
            self.assertTrue(any(p.endswith(".svg") and p.startswith("ppt/generated/media/") for p in package))
            self.assertTrue(any(p.endswith(".png") and p.startswith("ppt/generated/media/") for p in package))
            self.assertIn("ppt/generated/layout7.xml", package)
            self.assertNotEqual(package["ppt/generated/layout7.xml"], package[inherited_variant["source_layout_part"]])

    def test_unknown_type_and_source_geometry_constraints(self):
        with self.assertRaises(ValueError):
            validate_source_presentation({"template_version": "2.0", "title": "x", "slides": [{"layout_id": "timeline", "variant_id": "slide_001"}]}, self.manifest)
        v = self.manifest["layouts"]["title"]["variants"]["slide_003"]
        field = v["fields"]["slide_text_11"]
        with self.assertRaisesRegex(ValueError, "сократите"):
            validate_source_presentation({"template_version": "2.0", "title": "x", "slides": [{"layout_id": "title", "variant_id": "slide_003", "updates": [{"field_id": "slide_text_11", "text": "x" * (field["max_characters"] + 1)}]}]}, self.manifest)


if __name__ == "__main__":
    unittest.main()
