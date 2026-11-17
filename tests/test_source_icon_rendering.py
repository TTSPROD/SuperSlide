"""The native renderer contains icons in a square, without inherited crop."""
import unittest
from pathlib import Path

from lxml import etree as E
from PIL import Image

from support import temporary_directory
from superslide.source_template import P, A, CT, NS, _xml, _replace_image


def picture():
    return _xml(f'''<p:pic xmlns:p="{P}" xmlns:a="{A}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:nvPicPr><p:cNvPr id="7" name="Source icon"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
      <p:blipFill><a:blip r:embed="rIdOld"/><a:srcRect l="25000" r="10000"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
      <p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="600" cy="300"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom></p:spPr>
    </p:pic>'''.encode())


def root_with_shape(nested=False):
    root = E.Element(f"{{{P}}}sld", nsmap={"p": P, "a": A})
    tree = E.SubElement(E.SubElement(root, f"{{{P}}}cSld"), f"{{{P}}}spTree")
    if nested:
        tree = E.SubElement(tree, f"{{{P}}}grpSp")
        properties = E.SubElement(tree, f"{{{P}}}grpSpPr")
        transform = E.SubElement(properties, f"{{{A}}}xfrm")
        E.SubElement(transform, f"{{{A}}}off", x="0", y="0")
        E.SubElement(transform, f"{{{A}}}ext", cx="2000", cy="1000")
        E.SubElement(transform, f"{{{A}}}chOff", x="0", y="0")
        E.SubElement(transform, f"{{{A}}}chExt", cx="1000", cy="1000")
    shape = picture(); tree.append(shape)
    return root, shape


class IconRenderingTests(unittest.TestCase):
    def render(self, size, *, nested=False, icon_mode=True):
        with temporary_directory() as temporary:
            path = Path(temporary) / "icon.png"
            Image.new("RGBA", size, (0, 47, 108, 255)).save(path)
            root, shape = root_with_shape(nested)
            package = {}
            types = E.Element(f"{{{CT}}}Types")
            _replace_image(package, types, "ppt/slides/slide1.xml", shape, path, icon_mode=icon_mode)
            return root

    def test_wide_tall_square_keep_ratio_and_have_an_invisible_square_frame(self):
        for size in ((200, 100), (100, 200), (100, 100)):
            with self.subTest(size=size):
                root = self.render(size)
                group = root.find(f"{{{P}}}cSld/{{{P}}}spTree/{{{P}}}grpSp")
                frame = group.find(f"{{{P}}}grpSpPr/{{{A}}}xfrm")
                off, ext = frame.find(f"{{{A}}}off"), frame.find(f"{{{A}}}ext")
                self.assertEqual((int(off.get("x")), int(off.get("y"))), (250, 200))
                self.assertEqual((int(ext.get("cx")), int(ext.get("cy"))), (300, 300))
                content = group.find(f"{{{P}}}pic/{{{P}}}spPr/{{{A}}}xfrm")
                p, s = content.find(f"{{{A}}}off"), content.find(f"{{{A}}}ext")
                width, height = int(s.get("cx")), int(s.get("cy"))
                self.assertAlmostEqual(width / height, size[0] / size[1])
                self.assertEqual(max(width, height), 300)
                self.assertLessEqual(abs(2 * int(p.get("x")) + width - 300), 1)
                self.assertLessEqual(abs(2 * int(p.get("y")) + height - 300), 1)
                self.assertFalse(root.xpath(".//a:srcRect", namespaces=NS))
                self.assertFalse(group.xpath("./p:grpSpPr/a:solidFill | ./p:grpSpPr/a:ln", namespaces=NS))
                self.assertEqual(root.xpath(".//p:cNvPr/@id", namespaces=NS), ["7", "8"])
                self.assertTrue(root.xpath(".//a:picLocks[@noChangeAspect='1']", namespaces=NS))

    def test_anisotropic_parent_group_does_not_stretch_the_icon(self):
        root = self.render((200, 100), nested=True)
        group = root.find(f"{{{P}}}cSld/{{{P}}}spTree/{{{P}}}grpSp/{{{P}}}grpSp")
        frame = group.find(f"{{{P}}}grpSpPr/{{{A}}}xfrm/{{{A}}}ext")
        self.assertEqual(2 * int(frame.get("cx")), int(frame.get("cy")))
        content = group.find(f"{{{P}}}pic/{{{P}}}spPr/{{{A}}}xfrm/{{{A}}}ext")
        self.assertAlmostEqual(2 * int(content.get("cx")) / int(content.get("cy")), 2)

    def test_photo_replacement_recalculates_crop_and_keeps_original_frame(self):
        root = self.render((200, 100), icon_mode=False)
        self.assertFalse(root.xpath(".//p:grpSp", namespaces=NS))
        crop = root.xpath(".//a:srcRect", namespaces=NS)[0]
        self.assertEqual(dict(crop.attrib), {"l": "0", "t": "0", "r": "0", "b": "0"})
        self.assertEqual(root.xpath(".//a:xfrm/a:ext/@cx", namespaces=NS), ["600"])
        self.assertEqual(root.xpath(".//a:xfrm/a:ext/@cy", namespaces=NS), ["300"])
        self.assertEqual(root.xpath(".//a:prstGeom/@prst", namespaces=NS), ["roundRect"])


if __name__ == "__main__":
    unittest.main()
