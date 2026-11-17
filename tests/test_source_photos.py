"""Photo cover crops preserve the displayed aspect and source frame geometry."""
import tempfile
import unittest
from pathlib import Path

from lxml import etree as E
from PIL import Image

from superslide.source_icons import A, P
from superslide.source_photos import describe_photo_frame, photo_crop


def shape(box=(100, 200, 600, 400), *, tag="pic", rotation=None, flip_h=False):
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
        E.SubElement(transform, f"{{{A}}}chExt", cx="1000", cy="1000")
    return node


def group(child, scale_x, scale_y, *, rotation=None, flip_h=False):
    node = shape((0, 0, round(1000 * scale_x), round(1000 * scale_y)),
                 tag="grpSp", rotation=rotation, flip_h=flip_h)
    node.append(child)
    return node


class SourcePhotoTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.directory = Path(self.temp.name)

    def image(self, filename, size, *, orientation=None):
        path = self.directory / filename
        with Image.new("RGB", size, (16, 80, 120)) as picture:
            if orientation is None:
                picture.save(path)
            else:
                exif = Image.Exif()
                exif[274] = orientation
                picture.save(path, exif=exif)
        return path

    def assert_proportional(self, result):
        width, height = result["image_size"]
        crop = result["crop"]
        kept_width = width * (1 - (crop["l"] + crop["r"]) / 100000)
        kept_height = height * (1 - (crop["t"] + crop["b"]) / 100000)
        # Integer srcRect precision can differ by one part in 100000 per edge.
        self.assertAlmostEqual(kept_width / kept_height,
                               result["physical_aspect_ratio"], delta=0.0001)
        self.assertEqual(crop["l"], crop["r"])
        self.assertEqual(crop["t"], crop["b"])

    def test_wide_tall_and_matching_images_fill_frame_with_centered_crop(self):
        cases = (
            ((200, 100), {"l": 12500, "t": 0, "r": 12500, "b": 0}),
            ((100, 200), {"l": 0, "t": 33333, "r": 0, "b": 33333}),
            ((300, 200), {"l": 0, "t": 0, "r": 0, "b": 0}),
        )
        for size, expected in cases:
            with self.subTest(size=size):
                result = photo_crop(shape(), self.image(f"{size[0]}-{size[1]}.png", size))
                self.assertEqual(result["crop"], expected)
                self.assertEqual(result["source_box_emu"], [100, 200, 600, 400])
                self.assertEqual(result["frame_box_emu"], [100, 200, 600, 400])
                self.assertEqual(result["fit"], "cover")
                self.assertTrue(result["preserve_aspect_ratio"])
                self.assertEqual(result["alignment"], "center")
                self.assert_proportional(result)

    def test_nonuniform_group_scale_controls_crop_without_changing_frame(self):
        node = shape((100, 100, 400, 300))
        ancestor = group(node, 2, 1)
        before = E.tostring(ancestor)
        result = photo_crop(node, self.image("landscape.png", (400, 300)))
        self.assertEqual(result["effective_scale"], [2, 1])
        self.assertEqual(result["frame_box_emu"], [100, 100, 400, 300])
        self.assertEqual(result["physical_size_emu"], [800, 300])
        self.assertEqual(result["crop"], {"l": 0, "t": 25000, "r": 0, "b": 25000})
        self.assert_proportional(result)
        self.assertEqual(E.tostring(ancestor), before)

    def test_nested_groups_rotation_and_flip_preserve_physical_aspect(self):
        node = shape((100, 100, 400, 300), rotation=90, flip_h=True)
        inner = group(node, 2, 1)
        ancestor = group(inner, 3, 3, rotation=90, flip_h=True)
        before = E.tostring(ancestor)
        result = photo_crop(node, self.image("portrait.png", (100, 300)))
        self.assertAlmostEqual(result["effective_scale"][0], 3)
        self.assertAlmostEqual(result["effective_scale"][1], 6)
        self.assertEqual(result["transform_attributes"], {"rot": "5400000", "flipH": "1"})
        self.assertEqual(result["crop"], {"l": 0, "t": 25000, "r": 0, "b": 25000})
        self.assert_proportional(result)
        self.assertEqual(E.tostring(ancestor), before)

    def test_inherited_placeholder_frame_and_invalid_geometry(self):
        empty = E.Element(f"{{{P}}}sp")
        inherited = shape((50, 60, 120, 80), tag="sp", rotation=90)
        before = E.tostring(empty), E.tostring(inherited)
        result = describe_photo_frame(empty, inherited_shape=inherited)
        self.assertEqual(result["frame_box_emu"], [50, 60, 120, 80])
        self.assertEqual(result["transform_attributes"], {"rot": "5400000"})
        self.assertTrue(result["geometry_inherited"])
        self.assertEqual((E.tostring(empty), E.tostring(inherited)), before)
        with self.assertRaisesRegex(ValueError, "геометрия"):
            describe_photo_frame(empty)
        with self.assertRaisesRegex(ValueError, "положительные"):
            describe_photo_frame(shape((0, 0, 0, 100)))

    def test_previous_crop_and_shape_mask_are_untouched_and_not_reused(self):
        image = self.image("new-square.png", (200, 200))
        original_bytes = image.read_bytes()
        for tag in ("pic", "sp"):
            with self.subTest(tag=tag):
                node = shape(tag=tag)
                properties = node.find(f"{{{P}}}spPr")
                E.SubElement(properties, f"{{{A}}}prstGeom", prst="ellipse")
                fill = E.SubElement(node if tag == "pic" else properties,
                                    f"{{{P if tag == 'pic' else A}}}blipFill")
                E.SubElement(fill, f"{{{A}}}srcRect", l="35000", r="35000")
                before = E.tostring(node)
                result = photo_crop(node, image)
                self.assertEqual(result["crop"], {"l": 0, "t": 16667, "r": 0, "b": 16667})
                self.assertEqual(E.tostring(node), before)
                self.assertEqual(image.read_bytes(), original_bytes)
                self.assert_proportional(result)

    def test_svg_intrinsic_canvas_and_viewbox_drive_crop(self):
        cases = (
            ('width="256" height="256" viewBox="0 0 250 100"', (256, 256), 16667),
            ('width="100%" height="100%" viewBox="0 0 300 100"', (300, 100), 25000),
        )
        for index, (attrs, size, edge) in enumerate(cases):
            with self.subTest(attrs=attrs):
                path = self.directory / f"photo-{index}.svg"
                path.write_text(f'<svg xmlns="http://www.w3.org/2000/svg" {attrs}/>', encoding="utf-8")
                result = photo_crop(shape(), path)
                self.assertEqual(result["image_size"], list(size))
                if size[0] == size[1]:
                    self.assertEqual(result["crop"], {"l": 0, "t": edge, "r": 0, "b": edge})
                else:
                    self.assertEqual(result["crop"], {"l": edge, "t": 0, "r": edge, "b": 0})
                self.assert_proportional(result)

    def test_exif_rotated_jpeg_uses_display_dimensions_without_rewriting_file(self):
        path = self.image("rotated.jpg", (120, 80), orientation=6)
        before = path.read_bytes()
        result = photo_crop(shape((0, 0, 80, 120)), path)
        self.assertEqual(result["image_size"], [80, 120])
        self.assertEqual(result["image_aspect_ratio"], 2 / 3)
        self.assertEqual(result["exif_orientation"], 6)
        self.assertEqual(result["crop"], {"l": 0, "t": 0, "r": 0, "b": 0})
        self.assertEqual(path.read_bytes(), before)
        with Image.open(path) as picture:
            self.assertEqual(picture.size, (120, 80))
            self.assertEqual(picture.getexif()[274], 6)

    def test_future_sheared_frames_are_rejected_instead_of_stretching(self):
        node = shape(rotation=45)
        group(node, 2, 1)
        with self.assertRaisesRegex(ValueError, "сдвиг осей"):
            describe_photo_frame(node)


if __name__ == "__main__":
    unittest.main()
