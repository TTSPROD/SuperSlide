"""Every original icon binding declares the same square containment policy."""
import json
import unittest
from pathlib import Path
from zipfile import ZipFile

from lxml import etree as E

from scripts.prepare_severstal_template import (
    _fields_and_objects, _inherited_icon_shape, _icon_placement,
)
from superslide.source_template import ROOT, P, A, R, REL, NS, _bytes, _relpath


def _shape(sid, idx, box=None, kind="body"):
    shape = E.Element(f"{{{P}}}sp", nsmap={"p": P, "a": A})
    nv = E.SubElement(shape, f"{{{P}}}nvSpPr")
    E.SubElement(nv, f"{{{P}}}cNvPr", id=str(sid), name="Icon placeholder")
    E.SubElement(nv, f"{{{P}}}cNvSpPr")
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


class IconSlotMetadataTests(unittest.TestCase):
    def assert_policy(self, placement):
        self.assertEqual(placement["fit"], "contain")
        self.assertTrue(placement["preserve_aspect_ratio"])
        self.assertEqual(placement["alignment"], "center")
        self.assertEqual(placement["slot_shape"], "square")
        self.assertEqual(placement["square_side_policy"], "min_original_extent")
        self.assertEqual(placement["coordinate_space"], "parent")
        x, y, width, height = placement["source_box_emu"]
        sx, sy, side_x, side_y = placement["slot_box_emu"]
        scale_x, scale_y = placement["effective_scale"]
        self.assertGreater(width, 0)
        self.assertGreater(height, 0)
        self.assertLessEqual(abs((sx + side_x / 2) - (x + width / 2)), 1)
        self.assertLessEqual(abs((sy + side_y / 2) - (y + height / 2)), 1)
        self.assertLessEqual(side_x, width + 1)
        self.assertLessEqual(side_y, height + 1)
        # An icon's parent coordinates can be anisotropically scaled; its
        # visible square is measured after that original group scaling.
        self.assertLessEqual(abs(side_x * scale_x - side_y * scale_y), max(scale_x, scale_y) + 0.01)

    def test_all_real_source_icon_bindings_have_resolved_square_policy(self):
        manifest = json.loads((ROOT / "templates/severstal_template.json").read_text(encoding="utf-8"))
        fields_checked = 0
        nested = 0
        for layout_id, layout in manifest["layouts"].items():
            with ZipFile(ROOT / layout["source_file"]) as archive:
                package = {part: archive.read(part) for part in archive.namelist()}
            prefix = f"sources/{layout_id}/"
            owners = {}
            for variant in layout["variants"].values():
                for field in variant["fields"].values():
                    if field["kind"] != "icon":
                        continue
                    key = {"slide": "package_part", "layout": "source_layout_part", "master": "source_master_part"}[field["scope"]]
                    part = variant[key].removeprefix(prefix)
                    owners[part] = field["scope"]
            for part, scope in owners.items():
                before = package[part]
                fields, _ = _fields_and_objects(None, package, part, scope, prefix)
                for field in fields.values():
                    if field["kind"] == "icon":
                        self.assert_policy(field["placement"])
                        fields_checked += 1
                        if field["placement"]["effective_scale"] != [1.0, 1.0]:
                            nested += 1
                self.assertEqual(package[part], before)
        self.assertGreater(fields_checked, 1400)
        self.assertGreater(nested, 0)

    def test_missing_slide_transform_inherits_layout_by_idx(self):
        slide_shape = _shape(4, 14)
        slide = _owner(slide_shape, "sld")
        layout_shape = _shape(99, 14, (100, 200, 600, 300))
        layout = _owner(layout_shape, "sldLayout")
        package = {
            "ppt/slides/slide1.xml": _bytes(slide),
            _relpath("ppt/slides/slide1.xml"): _rels("slideLayout", "../slideLayouts/slideLayout1.xml"),
            "ppt/slideLayouts/slideLayout1.xml": _bytes(layout),
        }
        inherited = _inherited_icon_shape(package, "ppt/slides/slide1.xml", slide_shape)
        self.assertEqual(inherited.find(f"{{{P}}}nvSpPr/{{{P}}}cNvPr").get("id"), "99")
        placement = _icon_placement(package, "ppt/slides/slide1.xml", slide_shape)
        self.assert_policy(placement)
        self.assertEqual(placement["source_box_emu"], [100, 200, 600, 300])
        self.assertEqual(placement["slot_box_emu"], [250, 200, 300, 300])

    def test_missing_layout_transform_inherits_master_by_type(self):
        layout_shape = _shape(4, 14)
        layout = _owner(layout_shape, "sldLayout")
        master_shape = _shape(99, 2, (100, 200, 300, 600))
        master = _owner(master_shape, "sldMaster")
        package = {
            "ppt/slideLayouts/slideLayout1.xml": _bytes(layout),
            _relpath("ppt/slideLayouts/slideLayout1.xml"): _rels("slideMaster", "../slideMasters/slideMaster1.xml"),
            "ppt/slideMasters/slideMaster1.xml": _bytes(master),
        }
        inherited = _inherited_icon_shape(package, "ppt/slideLayouts/slideLayout1.xml", layout_shape)
        self.assertEqual(inherited.find(f"{{{P}}}nvSpPr/{{{P}}}cNvPr").get("id"), "99")
        placement = _icon_placement(package, "ppt/slideLayouts/slideLayout1.xml", layout_shape)
        self.assert_policy(placement)
        self.assertEqual(placement["source_box_emu"], [100, 200, 300, 600])
        self.assertEqual(placement["slot_box_emu"], [100, 350, 300, 300])


if __name__ == "__main__":
    unittest.main()
