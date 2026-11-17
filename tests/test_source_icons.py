"""Wide and tall icons retain their aspect inside physical square source slots."""
import unittest
from pathlib import Path

from lxml import etree as E
from PIL import Image

from support import temporary_directory
from superslide.source_icons import (
    A, P, describe_icon_slot, fitted_icon_box, image_dimensions,
)


def shape(box=(100, 200, 600, 400), *, tag="sp", rotation=None, flip_h=False, sid=2):
    node = E.Element(f"{{{P}}}{tag}", nsmap={"p": P, "a": A})
    properties = E.SubElement(node, f"{{{P}}}{'grpSpPr' if tag == 'grpSp' else 'spPr'}")
    transform = E.SubElement(properties, f"{{{A}}}xfrm")
    if rotation is not None:
        transform.set("rot", str(rotation * 60000))
    if flip_h:
        transform.set("flipH", "1")
    x, y, width, height = box
    E.SubElement(transform, f"{{{A}}}off", x=str(x), y=str(y))
    E.SubElement(transform, f"{{{A}}}ext", cx=str(width), cy=str(height))
    if tag == "grpSp":
        E.SubElement(transform, f"{{{A}}}chOff", x="0", y="0")
        E.SubElement(transform, f"{{{A}}}chExt", cx=str(width), cy=str(height))
    return node


def group(child, scale_x, scale_y, *, rotation=None, flip_h=False):
    node = shape((0, 0, round(1000 * scale_x), round(1000 * scale_y)), tag="grpSp", rotation=rotation, flip_h=flip_h)
    extent = node.find(f"{{{P}}}grpSpPr/{{{A}}}xfrm/{{{A}}}chExt")
    extent.set("cx", "1000"); extent.set("cy", "1000")
    node.append(child)
    return node


class SourceIconTests(unittest.TestCase):
    def setUp(self):
        self.temp = temporary_directory()
        self.addCleanup(self.temp.cleanup)
        self.directory = Path(self.temp.name)

    def png(self, filename, size):
        path = self.directory / filename
        picture = Image.new("RGBA", size, (0, 0, 0, 0))
        picture.putpixel((size[0] // 2, size[1] // 2), (0, 47, 108, 255))
        picture.save(path)
        return path

    def svg(self, filename, attrs):
        path = self.directory / filename
        path.write_text(f'<svg xmlns="http://www.w3.org/2000/svg" {attrs}><path d="M0 0L1 1"/></svg>', encoding="utf-8")
        return path

    def test_wide_tall_and_square_assets_are_contained_and_centered(self):
        for size in ((200, 100), (50, 200), (100, 100)):
            with self.subTest(size=size):
                node = shape()
                result = fitted_icon_box(node, self.png(f"{size[0]}-{size[1]}.png", size))
                self.assertEqual(result["source_box_emu"], [100, 200, 600, 400])
                self.assertEqual(result["slot_box_emu"], [200, 200, 400, 400])
                x, y, width, height = result["content_box_emu"]
                self.assertAlmostEqual(width / height, size[0] / size[1])
                self.assertEqual(x * 2 + width, 400)
                self.assertEqual(y * 2 + height, 400)
                self.assertEqual(max(width, height), 400)
                # Canvas dimensions, not nontransparent pixel bounds, are used.
                self.assertEqual(result["asset_size"], list(size))

    def test_group_scale_is_compensated_in_frame_and_picture(self):
        node = shape((100, 100, 400, 300))
        ancestor = group(node, 2, 1)
        result = fitted_icon_box(node, self.png("wide.png", (200, 100)))
        self.assertEqual(result["effective_scale"], [2, 1])
        self.assertEqual(result["slot_box_emu"], [225, 100, 150, 300])
        self.assertEqual(result["content_box_emu"], [0, 75, 150, 150])
        sx, sy = result["effective_scale"]
        _, _, fw, fh = result["slot_box_emu"]
        _, _, width, height = result["content_box_emu"]
        self.assertEqual(fw * sx, fh * sy)
        self.assertEqual(width * sx / (height * sy), 2)
        self.assertIs(node.getparent(), ancestor)

    def test_nested_scale_rotation_and_flip_preserve_physical_proportions(self):
        node = shape((100, 100, 400, 300), rotation=90, flip_h=True)
        inner = group(node, 2, 1)
        group(inner, 3, 3, rotation=90, flip_h=True)
        result = fitted_icon_box(node, self.png("tall.png", (100, 300)))
        self.assertAlmostEqual(result["effective_scale"][0], 3)
        self.assertAlmostEqual(result["effective_scale"][1], 6)
        self.assertEqual(result["transform_attributes"], {"rot": "5400000", "flipH": "1"})
        sx, sy = result["effective_scale"]
        _, _, fw, fh = result["slot_box_emu"]
        _, _, width, height = result["content_box_emu"]
        self.assertAlmostEqual(fw * sx, fh * sy, delta=6)
        self.assertAlmostEqual(width * sx / (height * sy), 1 / 3, delta=.01)

    def test_inherited_placeholder_geometry_and_missing_geometry(self):
        empty = E.Element(f"{{{P}}}sp")
        before = E.tostring(empty)
        inherited = shape((50, 60, 120, 80))
        result = describe_icon_slot(empty, inherited_shape=inherited)
        self.assertEqual(result["slot_box_emu"], [70, 60, 80, 80])
        self.assertTrue(result["geometry_inherited"])
        self.assertEqual(before, E.tostring(empty))
        with self.assertRaisesRegex(ValueError, "геометрия"):
            describe_icon_slot(empty)

    def test_svg_explicit_canvas_wins_over_viewbox_and_units_are_resolved(self):
        mismatch = self.svg("square-canvas.svg", 'width="256" height="256" viewBox="0 0 250 100"')
        self.assertEqual(image_dimensions(mismatch), (256, 256))
        with_units = self.svg("units.svg", 'width="1in" height="36pt" viewBox="0 0 10 10"')
        self.assertEqual(image_dimensions(with_units), (96, 48))
        fallback = self.svg("viewbox.svg", 'width="100%" height="100%" viewBox="0 0 300 100"')
        self.assertEqual(image_dimensions(fallback), (300, 100))
        one_dimension = self.svg("one-dimension.svg", 'width="600" viewBox="0 0 300 100"')
        self.assertEqual(image_dimensions(one_dimension), (600, 200))

    def test_source_crop_does_not_affect_new_icon_and_source_is_not_mutated(self):
        node = shape(tag="pic")
        fill = E.SubElement(node, f"{{{P}}}blipFill")
        E.SubElement(fill, f"{{{A}}}srcRect", l="35000", r="35000")
        before = E.tostring(node)
        result = fitted_icon_box(node, self.svg("wide.svg", 'width="200" height="100"'))
        self.assertEqual(result["content_box_emu"], [0, 100, 400, 200])
        self.assertEqual(E.tostring(node), before)

    def test_future_sheared_groups_are_rejected_instead_of_distorting_icons(self):
        node = shape(rotation=45)
        group(node, 2, 1)
        with self.assertRaisesRegex(ValueError, "сдвиг осей"):
            describe_icon_slot(node)

    def test_invalid_sizes_are_rejected(self):
        with self.assertRaisesRegex(ValueError, "положительные"):
            describe_icon_slot(shape((0, 0, 0, 100)))
        bad_svg = self.svg("bad.svg", 'width="0" height="-2"')
        with self.assertRaisesRegex(ValueError, "размера холста"):
            image_dimensions(bad_svg)


if __name__ == "__main__":
    unittest.main()
