"""The generated deck's document title follows its presentation content."""
import hashlib
import unittest
from pathlib import Path
from zipfile import ZipFile

from lxml import etree as E
from pptx import Presentation

from support import temporary_directory
from superslide.source_template import ROOT, _set_presentation_title, _xml, build_source_presentation


TITLE_TAG = "{http://purl.org/dc/elements/1.1/}title"


class PresentationTitleTests(unittest.TestCase):
    def test_source_deck_has_generated_russian_title_and_preserves_other_metadata(self):
        template = ROOT / "templates/severstal_source.pptx"
        manifest = ROOT / "templates/severstal_template.json"
        before_hash = hashlib.sha256(template.read_bytes()).hexdigest()
        with ZipFile(template) as archive:
            original_core = _xml(archive.read("docProps/core.xml"))
        title = "Шаттловые системы 2D и 3D: итоги & планы"
        spec = {"template_version": "2.0", "title": title, "slides": [
            {"layout_id": "title", "variant_id": "slide_002", "updates": []},
        ]}
        with temporary_directory() as directory:
            target = Path(directory) / "deck.pptx"
            build_source_presentation(spec, template, manifest, target)
            self.assertEqual(Presentation(target).core_properties.title, title)
            with ZipFile(target) as archive:
                generated_core = _xml(archive.read("docProps/core.xml"))
        self.assertEqual(generated_core.findtext(TITLE_TAG), title)
        for core in (original_core, generated_core):
            node = core.find(TITLE_TAG)
            if node is not None:
                core.remove(node)
        self.assertEqual(E.tostring(generated_core, method="c14n"),
                         E.tostring(original_core, method="c14n"))
        self.assertEqual(hashlib.sha256(template.read_bytes()).hexdigest(), before_hash)

    def test_missing_title_is_added_without_losing_author_or_fabricating_core_parts(self):
        package = {"docProps/core.xml": b'''<cp:coreProperties
          xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
          xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>Source author</dc:creator></cp:coreProperties>'''}
        _set_presentation_title(package, "Итоги проекта")
        core = _xml(package["docProps/core.xml"])
        self.assertEqual(core.findtext(TITLE_TAG), "Итоги проекта")
        self.assertEqual(core.findtext("{http://purl.org/dc/elements/1.1/}creator"), "Source author")
        empty = {}
        _set_presentation_title(empty, "Итоги проекта")
        self.assertEqual(empty, {})


if __name__ == "__main__":
    unittest.main()
