"""Retained source icon styles include native and renamed drawing objects."""
import io
import json
import unittest
from pathlib import Path
from zipfile import ZipFile

from lxml import etree as E
from PIL import Image

from scripts.prepare_severstal_template import (
    _fields_and_objects, _icon_background_metadata, _native_icon_has_background,
)
from superslide.source_template import A, P, R, REL, _bytes, _relpath


ROOT = Path(__file__).resolve().parents[1]


def native_shape(commands, *, sid=1, name="Graphic 1"):
    shape = E.fromstring(f'''<p:sp xmlns:p="{P}" xmlns:a="{A}">
      <p:nvSpPr><p:cNvPr id="{sid}" name="{name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="300000" cy="300000"/></a:xfrm>
      <a:custGeom><a:pathLst><a:path w="100" h="100">{commands}</a:path></a:pathLst></a:custGeom>
      <a:solidFill><a:srgbClr val="002F6C"/></a:solidFill></p:spPr></p:sp>''')
    return shape


def polygon(points):
    return ''.join(f'<a:{"moveTo" if i == 0 else "lnTo"}><a:pt x="{x}" y="{y}"/></a:{"moveTo" if i == 0 else "lnTo"}>'
                   for i, (x, y) in enumerate(points)) + '<a:close/>'


class SourceIconBackgroundTests(unittest.TestCase):
    def test_native_filled_plate_and_outline_have_different_styles(self):
        outer = polygon([(0, 0), (100, 0), (100, 100), (0, 100)])
        reverse_hole = polygon([(5, 5), (5, 95), (95, 95), (95, 5)])
        self.assertTrue(_native_icon_has_background(native_shape(outer)))
        # Same bounding box, but the opposite-winding inner contour is a hole.
        self.assertFalse(_native_icon_has_background(native_shape(outer + reverse_hole)))

    def test_a_small_filled_detail_inside_a_group_is_not_a_backing_plate(self):
        group = E.fromstring(f'''<p:grpSp xmlns:p="{P}" xmlns:a="{A}"><p:grpSpPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="300000" cy="300000"/>
          <a:chOff x="0" y="0"/><a:chExt cx="600000" cy="600000"/></a:xfrm>
          </p:grpSpPr></p:grpSp>''')
        group.append(native_shape(polygon([(0, 0), (100, 0), (100, 100), (0, 100)])))
        self.assertFalse(_native_icon_has_background(group))
        group.find(f"{{{P}}}grpSpPr/{{{A}}}xfrm/{{{A}}}chExt").set("cx", "300000")
        group.find(f"{{{P}}}grpSpPr/{{{A}}}xfrm/{{{A}}}chExt").set("cy", "300000")
        self.assertTrue(_native_icon_has_background(group))

    def test_picture_uses_existing_fallback_alpha_without_a_svg_renderer(self):
        shape = E.fromstring(f'''<p:pic xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}">
          <p:blipFill><a:blip r:embed="rId1"/></p:blipFill></p:pic>''')
        relations = E.Element(f"{{{REL}}}Relationships")
        E.SubElement(relations, f"{{{REL}}}Relationship", Id="rId1", Target="../media/icon.png")
        part = "ppt/slides/slide1.xml"
        package = {_relpath(part): _bytes(relations)}
        for opaque, expected in [(False, "without_background"), (True, "with_background")]:
            image = Image.new("RGBA", (100, 100), (0, 47, 108, 255 if opaque else 0))
            image.putpixel((50, 50), (0, 47, 108, 255))
            stream = io.BytesIO(); image.save(stream, format="PNG")
            package["ppt/media/icon.png"] = stream.getvalue()
            self.assertEqual(_icon_background_metadata(package, part, shape)["background_style"], expected)

    def test_table_source_exposes_all_eight_retained_row_icons(self):
        with ZipFile(ROOT / "resources/severstal/templates/Severstal_Tables.pptx") as archive:
            package = {name: archive.read(name) for name in archive.namelist()}
        before = package["ppt/slides/slide19.xml"]
        fields, _ = _fields_and_objects(None, package, "ppt/slides/slide19.xml", "slide", "sources/tables/")
        icons = {key: field for key, field in fields.items() if field["kind"] == "icon"}
        self.assertEqual(len(icons), 8)
        self.assertEqual(icons["slide_icon_5"]["name"], "Рисунок 2")
        self.assertEqual({r["background_style"] for r in icons.values()}, {"without_background"})
        self.assertEqual(package["ppt/slides/slide19.xml"], before)

    def test_manifest_carries_style_for_every_original_icon(self):
        manifest = json.loads((ROOT / "templates/severstal_template.json").read_text(encoding="utf-8"))
        for layout in manifest["layouts"].values():
            for variant in layout["variants"].values():
                for field in variant["fields"].values():
                    if field["kind"] == "icon":
                        self.assertIn(field["background_style"], {"with_background", "without_background"})
                        self.assertIs(field["has_background"], field["background_style"] == "with_background")


if __name__ == "__main__":
    unittest.main()
