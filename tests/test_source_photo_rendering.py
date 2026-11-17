"""Native photo replacement covers frames without changing source artwork."""
import io
import tempfile
import unittest
from pathlib import Path

from lxml import etree as E
from PIL import Image

from superslide.source_template import (
    A, CT, NS, P, R, _relations, _replace_image, _target, _xml,
)


OWNER = "ppt/slides/slide1.xml"


def picture(*, mask="roundRect"):
    return _xml(f'''<p:pic xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}">
      <p:nvPicPr><p:cNvPr id="7" name="Source photo"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
      <p:blipFill><a:blip r:embed="rIdOld"/><a:srcRect l="25000" r="10000"/>
        <a:stretch><a:fillRect l="10000" t="7000"/></a:stretch></p:blipFill>
      <p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="600" cy="300"/></a:xfrm>
        <a:prstGeom prst="{mask}"><a:avLst/></a:prstGeom>
        <a:ln w="12700"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:ln>
      </p:spPr>
    </p:pic>'''.encode())


def group(child, scale_x, scale_y, *, rotation=0, flip_h=False):
    node = E.Element(f"{{{P}}}grpSp")
    properties = E.SubElement(node, f"{{{P}}}grpSpPr")
    transform = E.SubElement(properties, f"{{{A}}}xfrm", rot=str(rotation * 60000))
    if flip_h:
        transform.set("flipH", "1")
    E.SubElement(transform, f"{{{A}}}off", x="10", y="20")
    E.SubElement(transform, f"{{{A}}}ext", cx=str(1000 * scale_x), cy=str(1000 * scale_y))
    E.SubElement(transform, f"{{{A}}}chOff", x="30", y="40")
    E.SubElement(transform, f"{{{A}}}chExt", cx="1000", cy="1000")
    node.append(child)
    return node


class PhotoRenderingTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.directory = Path(self.temp.name)

    def image(self, size, *, filename="photo.png"):
        path = self.directory / filename
        with Image.new("RGB", size, (16, 80, 120)) as image:
            image.save(path)
        return path

    def replace(self, shape, image, *, inherited_shape=None):
        package = {}
        types = E.Element(f"{{{CT}}}Types")
        placement = _replace_image(package, types, OWNER, shape, image,
                                   inherited_shape=inherited_shape)
        return package, types, placement

    def crop(self, shape):
        rectangles = shape.xpath("./p:blipFill/a:srcRect | ./p:spPr/a:blipFill/a:srcRect", namespaces=NS)
        self.assertEqual(len(rectangles), 1)
        return {key: int(value) for key, value in rectangles[0].attrib.items()}

    def embedded(self, package, shape):
        rid = shape.xpath("./p:blipFill/a:blip/@r:embed | ./p:spPr/a:blipFill/a:blip/@r:embed", namespaces=NS)[0]
        relation = next(item for item in _relations(package, OWNER) if item.get("Id") == rid)
        target = _target(OWNER, relation.get("Target"))
        return target, package[target]

    def test_wide_tall_square_and_matching_assets_replace_old_crop(self):
        cases = (
            ((400, 100), {"l": 25000, "t": 0, "r": 25000, "b": 0}),
            ((100, 200), {"l": 0, "t": 37500, "r": 0, "b": 37500}),
            ((100, 100), {"l": 0, "t": 25000, "r": 0, "b": 25000}),
            ((200, 100), {"l": 0, "t": 0, "r": 0, "b": 0}),
        )
        for size, expected in cases:
            with self.subTest(size=size):
                shape = picture()
                image = self.image(size)
                package, _, placement = self.replace(shape, image)
                self.assertEqual(self.crop(shape), expected)
                self.assertEqual(placement["crop"], expected)
                self.assertEqual(self.embedded(package, shape)[1], image.read_bytes())
                self.assertEqual(shape.xpath("./p:nvPicPr/p:cNvPicPr/a:picLocks/@noChangeAspect", namespaces=NS), ["1"])

    def test_native_picture_frame_masks_and_line_remain_unchanged(self):
        for mask in ("roundRect", "ellipse"):
            with self.subTest(mask=mask):
                shape = picture(mask=mask)
                properties = shape.find(f"{{{P}}}spPr")
                properties.find(f"{{{A}}}xfrm").set("rot", "5400000")
                properties.find(f"{{{A}}}xfrm").set("flipV", "1")
                original_properties = E.tostring(properties)
                original_identity = E.tostring(shape.find(f"{{{P}}}nvPicPr/{{{P}}}cNvPr"))
                self.replace(shape, self.image((400, 100)))
                self.assertEqual(shape.tag, f"{{{P}}}pic")
                self.assertEqual(E.tostring(properties), original_properties)
                self.assertEqual(E.tostring(shape.find(f"{{{P}}}nvPicPr/{{{P}}}cNvPr")), original_identity)
                self.assertEqual(shape.xpath("./p:spPr/a:prstGeom/@prst", namespaces=NS), [mask])

    def test_nested_anisotropic_groups_yield_equal_physical_pixel_scales(self):
        shape = picture()
        shape.find(f"{{{P}}}spPr/{{{A}}}xfrm").set("rot", "5400000")
        inner = group(shape, 2, 1)
        outer = group(inner, 3, 3, rotation=90, flip_h=True)
        group_properties = [E.tostring(item) for item in outer.xpath(".//p:grpSpPr | ./p:grpSpPr", namespaces=NS)]
        frame = E.tostring(shape.find(f"{{{P}}}spPr"))
        _, _, placement = self.replace(shape, self.image((200, 100)))
        crop = self.crop(shape)
        # The photo axes see scales 3 and 6 after its own 90-degree rotation.
        # The 600x300 source frame is therefore physically square (1800x1800).
        visible_width = 200 * (1 - (crop["l"] + crop["r"]) / 100000)
        visible_height = 100 * (1 - (crop["t"] + crop["b"]) / 100000)
        self.assertEqual(crop, {"l": 25000, "t": 0, "r": 25000, "b": 0})
        self.assertAlmostEqual(600 * 3 / visible_width, 300 * 6 / visible_height)
        self.assertAlmostEqual(placement["effective_scale"][0], 3)
        self.assertAlmostEqual(placement["effective_scale"][1], 6)
        self.assertEqual(E.tostring(shape.find(f"{{{P}}}spPr")), frame)
        self.assertEqual([E.tostring(item) for item in outer.xpath(".//p:grpSpPr | ./p:grpSpPr", namespaces=NS)], group_properties)

    def test_inherited_frame_is_respected_without_materializing_local_geometry(self):
        shape = picture(mask="ellipse")
        properties = shape.find(f"{{{P}}}spPr")
        properties.remove(properties.find(f"{{{A}}}xfrm"))
        inherited = picture()
        original_properties, original_inherited = E.tostring(properties), E.tostring(inherited)
        _, _, placement = self.replace(shape, self.image((100, 100)), inherited_shape=inherited)
        self.assertTrue(placement["geometry_inherited"])
        self.assertEqual(placement["frame_box_emu"], [100, 200, 600, 300])
        self.assertEqual(self.crop(shape), {"l": 0, "t": 25000, "r": 0, "b": 25000})
        self.assertIsNone(properties.find(f"{{{A}}}xfrm"))
        self.assertEqual(E.tostring(properties), original_properties)
        self.assertEqual(E.tostring(inherited), original_inherited)

    def test_shape_picture_fill_preserves_custom_mask_and_line(self):
        for existing_picture_fill in (False, True):
            with self.subTest(existing_picture_fill=existing_picture_fill):
                shape = _xml(f'''<p:sp xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}">
                  <p:nvSpPr><p:cNvPr id="9" name="Custom photo frame"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
                  <p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="600" cy="300"/></a:xfrm>
                    <a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/>
                      <a:pathLst><a:path w="100" h="100"><a:moveTo><a:pt x="50" y="0"/></a:moveTo>
                        <a:lnTo><a:pt x="100" y="50"/></a:lnTo><a:lnTo><a:pt x="50" y="100"/></a:lnTo>
                        <a:lnTo><a:pt x="0" y="50"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>
                    <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
                    <a:ln w="19050"><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:ln>
                  </p:spPr>
                </p:sp>'''.encode())
                properties = shape.find(f"{{{P}}}spPr")
                if existing_picture_fill:
                    properties.remove(properties.find(f"{{{A}}}solidFill"))
                    fill = E.Element(f"{{{A}}}blipFill")
                    E.SubElement(fill, f"{{{A}}}blip", **{f"{{{R}}}embed": "rIdOld"})
                    E.SubElement(fill, f"{{{A}}}srcRect", l="40000")
                    E.SubElement(fill, f"{{{A}}}tile", sx="50000")
                    properties.insert(2, fill)
                protected = [E.tostring(properties.find(f"{{{A}}}{name}"))
                             for name in ("xfrm", "custGeom", "ln")]
                self.replace(shape, self.image((100, 100)))
                self.assertEqual(self.crop(shape), {"l": 0, "t": 25000, "r": 0, "b": 25000})
                self.assertEqual([E.tostring(properties.find(f"{{{A}}}{name}"))
                                  for name in ("xfrm", "custGeom", "ln")], protected)
                self.assertEqual([E.QName(item).localname for item in properties],
                                 ["xfrm", "custGeom", "blipFill", "ln"])
                self.assertEqual(shape.xpath("./p:nvSpPr/p:cNvSpPr/a:spLocks/@noChangeAspect", namespaces=NS), ["1"])

    def test_previous_tile_or_stretch_offsets_are_replaced_by_single_cover_mapping(self):
        for old_mapping in ("tile", "stretch"):
            with self.subTest(old_mapping=old_mapping):
                shape = picture()
                fill = shape.find(f"{{{P}}}blipFill")
                if old_mapping == "tile":
                    fill.remove(fill.find(f"{{{A}}}stretch"))
                    E.SubElement(fill, f"{{{A}}}tile", tx="9000", ty="12000", sx="50000", sy="80000")
                fill.find(f"{{{A}}}blip").set(f"{{{R}}}link", "rIdExternal")
                self.replace(shape, self.image((200, 100)))
                self.assertEqual(self.crop(shape), {"l": 0, "t": 0, "r": 0, "b": 0})
                self.assertEqual([E.QName(item).localname for item in fill], ["blip", "srcRect", "stretch"])
                self.assertFalse(fill.xpath("./a:tile | ./a:blip/@r:link", namespaces=NS))
                self.assertEqual(len(fill.xpath("./a:stretch/a:fillRect", namespaces=NS)), 1)
                self.assertEqual(dict(fill.find(f"{{{A}}}stretch/{{{A}}}fillRect").attrib), {})

    def test_exif_rotated_photo_embeds_display_oriented_png_and_preserves_source_file(self):
        path = self.directory / "rotated.jpg"
        with Image.new("RGB", (120, 80), (220, 20, 20)) as image:
            image.paste((20, 20, 220), (60, 0, 120, 80))
            exif = Image.Exif()
            exif[274] = 6
            image.save(path, exif=exif)
        original = path.read_bytes()
        shape = picture()
        extent = shape.find(f"{{{P}}}spPr/{{{A}}}xfrm/{{{A}}}ext")
        extent.set("cx", "80")
        extent.set("cy", "120")
        package, types, placement = self.replace(shape, path)
        target, payload = self.embedded(package, shape)
        self.assertTrue(target.endswith(".png"))
        self.assertEqual(placement["image_size"], [80, 120])
        self.assertEqual(self.crop(shape), {"l": 0, "t": 0, "r": 0, "b": 0})
        with Image.open(io.BytesIO(payload)) as embedded:
            self.assertEqual(embedded.format, "PNG")
            self.assertEqual(embedded.size, (80, 120))
            self.assertNotIn(274, embedded.getexif())
            upper, lower = embedded.getpixel((40, 20)), embedded.getpixel((40, 100))
            self.assertGreater(upper[0], upper[2])
            self.assertGreater(lower[2], lower[0])
        override = types.find(f"{{{CT}}}Override")
        self.assertEqual(override.get("PartName"), "/" + target)
        self.assertEqual(override.get("ContentType"), "image/png")
        self.assertEqual(path.read_bytes(), original)


if __name__ == "__main__":
    unittest.main()
