"""The source's footer prompt becomes a visible build date on actual slides."""
import json
import unittest
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile

from support import temporary_directory
from superslide.source_template import (
    ROOT, P, NS, FOOTER_DATE_PLACEHOLDER, build_source_presentation,
    load_source_manifest, source_prompt_catalog, _xml,
)


class SourceDateTests(unittest.TestCase):
    def test_all_source_footer_slots_get_build_date_without_changing_source(self):
        manifest = load_source_manifest(ROOT / "templates/severstal_template.json")
        spec = json.loads((ROOT / "examples/severstal_all_variants.json").read_text(encoding="utf-8"))
        with ZipFile(ROOT / "templates/severstal_source.pptx") as source:
            expected = {}
            for index, slide in enumerate(spec["slides"], 1):
                design = manifest["layouts"][slide["layout_id"]]["variants"][slide["variant_id"]]
                layout = _xml(source.read(design["source_layout_part"]))
                date_slots = []
                for shape in layout.xpath(".//p:sp", namespaces=NS):
                    if "".join(shape.xpath("./p:txBody//a:t/text()", namespaces=NS)).strip() == FOOTER_DATE_PLACEHOLDER:
                        date_slots.append(shape.find(f"{{{P}}}nvSpPr/{{{P}}}nvPr/{{{P}}}ph").get("idx", "0"))
                expected[index] = date_slots
            self.assertEqual(sum(bool(v) for v in expected.values()), 161)
            # The source archive retains its prompt; generated output resolves it.
            self.assertTrue(any(FOOTER_DATE_PLACEHOLDER in "".join(_xml(source.read(n)).xpath(".//a:t/text()", namespaces=NS))
                                for n in source.namelist() if "/slideLayouts/" in n and n.endswith(".xml")))
        with temporary_directory() as temp:
            output = Path(temp) / "dated.pptx"
            with patch("superslide.source_template._current_date_text", return_value="17.11.2026"):
                report = build_source_presentation(spec, ROOT / "templates/severstal_source.pptx", ROOT / "templates/severstal_template.json", output)
            self.assertEqual(report["footer_date"], "17.11.2026")
            with ZipFile(output) as archive:
                for index, slots in expected.items():
                    root = _xml(archive.read(f"ppt/generated/slide{index}.xml"))
                    for idx in slots:
                        shapes = root.xpath(f".//p:sp[p:nvSpPr/p:nvPr/p:ph[@idx='{idx}']]", namespaces=NS)
                        self.assertEqual(len(shapes), 1)
                        self.assertEqual("".join(shapes[0].xpath("./p:txBody//a:t/text()", namespaces=NS)), "17.11.2026")
                        # Office silently drops a new run placed after endParaRPr.
                        for paragraph in shapes[0].xpath("./p:txBody/a:p", namespaces=NS):
                            end = paragraph.find("{" + NS["a"] + "}endParaRPr")
                            if end is not None:
                                self.assertIs(paragraph[-1], end)
                for part in archive.namelist():
                    if part.endswith(".xml") and ("/slideLayouts/" in part or part.startswith("ppt/generated/slide")):
                        self.assertNotIn(FOOTER_DATE_PLACEHOLDER, "".join(_xml(archive.read(part)).xpath(".//a:t/text()", namespaces=NS)))
                # Numbered example captions must not be treated as dates.
                self.assertIn("#1", "".join(_xml(archive.read("ppt/generated/slide38.xml")).xpath(".//a:t/text()", namespaces=NS)))

    def test_next_build_gets_new_date_and_date_is_not_a_model_field(self):
        manifest = load_source_manifest(ROOT / "templates/severstal_template.json")
        catalog = source_prompt_catalog(manifest)
        self.assertNotIn(FOOTER_DATE_PLACEHOLDER, json.dumps(catalog, ensure_ascii=False))
        spec = {"template_version": "2.0", "title": "Дата", "slides": [{"layout_id": "tables", "variant_id": "slide_018", "updates": []}]}
        with temporary_directory() as temp:
            for date in ("31.12.2026", "01.01.2027"):
                output = Path(temp) / "dated.pptx"
                with patch("superslide.source_template._current_date_text", return_value=date):
                    build_source_presentation(spec, ROOT / "templates/severstal_source.pptx", ROOT / "templates/severstal_template.json", output)
                with ZipFile(output) as archive:
                    root = _xml(archive.read("ppt/generated/slide1.xml"))
                    self.assertIn(date, "".join(root.xpath(".//a:t/text()", namespaces=NS)))
                    ids = root.xpath(".//p:cNvPr/@id", namespaces=NS)
                    self.assertEqual(len(ids), len(set(ids)))


if __name__ == "__main__":
    unittest.main()
