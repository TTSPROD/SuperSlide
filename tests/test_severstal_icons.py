"""Integrity checks against the supplied, unmodified icon PowerPoint."""
import hashlib
import json
import posixpath
import unittest
from collections import Counter
from pathlib import Path
from zipfile import ZipFile

from lxml import etree as ET
from PIL import Image


ROOT = Path(__file__).resolve().parents[1]
NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}


class SeverstalIconIntegrityTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.catalog=json.loads((ROOT/"assets/icons/severstal/catalog.json").read_text(encoding="utf-8"))

    def test_all_source_groups_and_pages_are_preserved(self):
        catalog=self.catalog
        self.assertEqual(len(catalog["groups"]),30)
        self.assertEqual(sorted(s for g in catalog["groups"] for s in g["source_slides"]),list(range(1,34)))
        self.assertEqual(len(catalog["icons"]),1364)
        self.assertEqual(len({r["id"] for r in catalog["icons"]}),1364)
        self.assertEqual(len({r["path"] for r in catalog["icons"]}),1364)
        members=Counter(i for g in catalog["groups"] for i in g["icon_ids"])
        self.assertEqual(members,Counter(r["id"] for r in catalog["icons"]))
        self.assertEqual(sum(r["selectable"] for r in catalog["icons"]),1362)
        self.assertEqual([(r["source"]["slide"],r["source"]["index"]) for r in catalog["icons"] if not r["selectable"]],[(7,35),(7,72)])

    def test_assets_are_named_complete_and_self_contained(self):
        groups={g["id"] for g in self.catalog["groups"]}
        for r in self.catalog["icons"]:
            with self.subTest(icon=r["id"]):
                self.assertIn(r["group_id"],groups)
                self.assertNotRegex(r["name"],r"^.+\s\d{2}$")
                self.assertTrue(r["description"])
                self.assertTrue(r["tags"])
                path=ROOT/r["path"]
                self.assertTrue(path.is_relative_to(ROOT/"assets/icons/severstal"))
                self.assertEqual(hashlib.sha256(path.read_bytes()).hexdigest(),r["sha256"])
                tree=ET.parse(str(path))
                self.assertEqual(ET.QName(tree.getroot()).localname,"svg")
                for node in tree.iter():
                    for key,value in node.attrib.items():
                        if ET.QName(key).localname=="href":
                            self.assertTrue(value.startswith(("data:","#")))
                with Image.open(ROOT/r["png_file"]) as image:
                    self.assertGreater(image.width,0)
                    self.assertGreater(image.height,0)
                    self.assertEqual(image.format,"PNG")
                    self.assertEqual(image.mode,"RGBA")

    def test_original_svg_artwork_is_copied_without_modification(self):
        with ZipFile(ROOT/self.catalog["source"]) as archive:
            self.assertEqual(hashlib.sha256((ROOT/self.catalog["source"]).read_bytes()).hexdigest(),self.catalog["source_sha256"])
            for r in self.catalog["icons"]:
                if r["extraction"]!="original_svg":
                    continue
                with self.subTest(icon=r["id"]):
                    slide=r["source"]["slide"]
                    tree=ET.fromstring(archive.read(f"ppt/slides/slide{slide}.xml"))
                    sid=str(r["source"]["shape_ids"][0])
                    node=tree.xpath(f'//p:cNvPr[@id="{sid}"]',namespaces=NS)[0].getparent().getparent()
                    rid=node.xpath('.//*[local-name()="svgBlip"]/@r:embed',namespaces=NS)[0]
                    rels=ET.fromstring(archive.read(f"ppt/slides/_rels/slide{slide}.xml.rels"))
                    target=next(x.get("Target") for x in rels if x.get("Id")==rid)
                    part=posixpath.normpath(posixpath.join("ppt/slides",target))
                    self.assertEqual((ROOT/r["path"]).read_bytes(),archive.read(part))


if __name__=="__main__":
    unittest.main()
