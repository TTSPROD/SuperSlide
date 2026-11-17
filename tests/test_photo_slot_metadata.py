"""Photo bindings retain their original frames and declare centered cover fit."""
import json
import unittest
from zipfile import ZipFile

from lxml import etree as E

from scripts.prepare_severstal_template import _fields_and_objects, _photo_placement
from superslide.source_template import ROOT, P, A, R, REL, _bytes, _relpath


def _picture(sid, idx, box=None, kind="pic"):
    shape = E.Element(f"{{{P}}}pic", nsmap={"p": P, "a": A})
    nv = E.SubElement(shape, f"{{{P}}}nvPicPr")
    E.SubElement(nv, f"{{{P}}}cNvPr", id=str(sid), name="Photo placeholder")
    E.SubElement(nv, f"{{{P}}}cNvPicPr")
    nvpr = E.SubElement(nv, f"{{{P}}}nvPr")
    E.SubElement(nvpr, f"{{{P}}}ph", idx=str(idx), type=kind)
    properties = E.SubElement(shape, f"{{{P}}}spPr")
    if box:
        x, y, width, height = box
        xfrm = E.SubElement(properties, f"{{{A}}}xfrm")
        E.SubElement(xfrm, f"{{{A}}}off", x=str(x), y=str(y))
        E.SubElement(xfrm, f"{{{A}}}ext", cx=str(width), cy=str(height))
    return shape


def _owner(shape, kind):
    root = E.Element(f"{{{P}}}{kind}", nsmap={"p": P, "a": A})
    tree = E.SubElement(E.SubElement(root, f"{{{P}}}cSld"), f"{{{P}}}spTree")
    tree.append(shape)
    return root


def _rels(kind, target):
    relations = E.Element(f"{{{REL}}}Relationships", nsmap={None: REL})
    E.SubElement(relations, f"{{{REL}}}Relationship", Id="rId1", Type=R + "/" + kind, Target=target)
    return _bytes(relations)


class PhotoSlotMetadataTests(unittest.TestCase):
    def assert_policy(self, placement):
        self.assertEqual(placement["fit"], "cover")
        self.assertTrue(placement["preserve_aspect_ratio"])
        self.assertEqual(placement["alignment"], "center")
        self.assertEqual(placement["frame"], "source")
        self.assertTrue(placement["preserve_source_mask"])
        self.assertEqual(placement["coordinate_space"], "parent")
        self.assertEqual(placement["frame_box_emu"], placement["source_box_emu"])
        _, _, width, height = placement["frame_box_emu"]
        self.assertGreater(width, 0)
        self.assertGreater(height, 0)
        sx, sy = placement["effective_scale"]
        self.assertAlmostEqual(placement["physical_aspect_ratio"], width * sx / (height * sy))

    def test_all_real_source_photos_resolve_original_frame_without_xml_changes(self):
        manifest = json.loads((ROOT / "templates/severstal_template.json").read_text(encoding="utf-8"))
        checked = inherited = 0
        masks = set()
        for layout_id, layout in manifest["layouts"].items():
            with ZipFile(ROOT / layout["source_file"]) as archive:
                package = {part: archive.read(part) for part in archive.namelist()}
            original_package = dict(package)
            prefix = f"sources/{layout_id}/"
            owners = {}
            for variant in layout["variants"].values():
                for field in variant["fields"].values():
                    if field["kind"] != "picture":
                        continue
                    key = {"slide": "package_part", "layout": "source_layout_part", "master": "source_master_part"}[field["scope"]]
                    owners[variant[key].removeprefix(prefix)] = field["scope"]
            for part, scope in owners.items():
                fields, _ = _fields_and_objects(None, package, part, scope, prefix)
                for field in fields.values():
                    if field["kind"] != "picture":
                        continue
                    self.assert_policy(field["placement"])
                    inherited += int(field["placement"]["geometry_inherited"])
                    source = E.fromstring(field["crop_and_transform_xml"].encode("utf-8"))
                    preset = source.find(f"{{{P}}}spPr/{{{A}}}prstGeom")
                    if preset is not None:
                        masks.add(preset.get("prst"))
                    if source.find(f"{{{P}}}spPr/{{{A}}}custGeom") is not None:
                        masks.add("custom")
                    checked += 1
            self.assertEqual(package, original_package)
        self.assertEqual(checked, 207)
        self.assertGreater(inherited, 0)
        self.assertTrue({"rect", "ellipse", "custom"}.issubset(masks))

    def test_slide_photo_inherits_original_layout_frame_by_idx(self):
        photo = _picture(4, 14)
        slide = _owner(photo, "sld")
        layout = _owner(_picture(99, 14, (100, 200, 600, 300)), "sldLayout")
        package = {
            "ppt/slides/slide1.xml": _bytes(slide),
            _relpath("ppt/slides/slide1.xml"): _rels("slideLayout", "../slideLayouts/slideLayout1.xml"),
            "ppt/slideLayouts/slideLayout1.xml": _bytes(layout),
        }
        before = E.tostring(photo)
        placement = _photo_placement(package, "ppt/slides/slide1.xml", photo)
        self.assert_policy(placement)
        self.assertTrue(placement["geometry_inherited"])
        self.assertEqual(placement["frame_box_emu"], [100, 200, 600, 300])
        self.assertEqual(E.tostring(photo), before)

    def test_layout_photo_inherits_master_frame_by_type(self):
        photo = _picture(4, 14)
        layout = _owner(photo, "sldLayout")
        master = _owner(_picture(99, 2, (100, 200, 300, 600)), "sldMaster")
        package = {
            "ppt/slideLayouts/slideLayout1.xml": _bytes(layout),
            _relpath("ppt/slideLayouts/slideLayout1.xml"): _rels("slideMaster", "../slideMasters/slideMaster1.xml"),
            "ppt/slideMasters/slideMaster1.xml": _bytes(master),
        }
        placement = _photo_placement(package, "ppt/slideLayouts/slideLayout1.xml", photo)
        self.assert_policy(placement)
        self.assertTrue(placement["geometry_inherited"])
        self.assertEqual(placement["frame_box_emu"], [100, 200, 300, 600])


if __name__ == "__main__":
    unittest.main()
