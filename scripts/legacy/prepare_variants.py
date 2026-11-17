"""Brandbook-backed design variants for the stable semantic slide contracts.

Original graphics are copied from the corporate PPTX, including their media
relationships. None of these layouts uses a rendered PDF page as artwork.
"""

from __future__ import annotations

from copy import deepcopy
import posixpath
from zipfile import ZipFile

from lxml import etree as ET


def add_variant_specs(specs: dict, field, colors: dict, emu: int) -> None:
    """Append nested variants while leaving the fourteen default designs intact."""
    blue, white, muted = colors["primary"], "FFFFFF", colors["muted"]
    title = field(0, (0.45, 0.34, 11.30, 1.10), 32, 26, lines=2, bold=True)

    def design(semantic: str, variant: str, description: str, fields: dict,
               pages: list[int], slides: list[int], source_part: str,
               *, background: str = white, footer: bool = True,
               foreground: bool = False, art: str | None = None) -> dict:
        result = {
            "layout_name": f"{semantic}__{variant}",
            "description": description,
            "fields": fields,
            "background_color": background,
            "show_footer": footer,
            "source_reference": {
                "brandbook_pages": pages,
                "source_slides": slides,
                "source_layout_part": source_part,
            },
        }
        if foreground:
            result["foreground_art"] = True
        if art:
            result["artwork_style"] = art
        specs[semantic].setdefault("variants", {})[variant] = result
        return result

    def cover_fields(*, wide: bool = False) -> dict:
        width = 11.40 if wide else 7.10
        return {
            "title": field(0, (0.465, 2.48, width, 1.55), 42, 32, lines=3, bold=True, color=white),
            "subtitle": field(1, (0.465, 4.10, width, 0.80), 22, 18, lines=2, color=white),
            "author": field(2, (0.78, 4.97, 6.50, 0.31), 16, 14, lines=1, color=white),
            "date": field(3, (0.78, 5.35, 6.50, 0.31), 16, 14, lines=1, color=white),
        }

    design("title_slide", "blue_minimal", "Синяя обложка без графики: вариант из брендбука, стр. 108.",
           cover_fields(wide=True), [108], [4], "ppt/slideLayouts/slideLayout24.xml",
           background=blue, footer=False, art="original")
    design("title_slide", "blue_graphic", "Синяя обложка с оригинальной диагональной суперграфикой, стр. 108.",
           cover_fields(), [108], [3], "ppt/slideLayouts/slideLayout23.xml",
           background=blue, footer=False, art="original")
    photo_cover = cover_fields()
    photo_cover["image"] = field(4, (0, 0, 13.3333333333, 7.5), 18, 16, kind="picture")
    design("title_slide", "photo", "Фотообложка с оригинальной диагональной графикой и полупрозрачным синим слоем для читаемости, стр. 108.",
           photo_cover, [108], [2, 8], "ppt/slideLayouts/slideLayout23.xml",
           background=blue, footer=False, foreground=True, art="photo_cover")

    for variant, part, slide, label in [
        ("circle_pattern", 30, 166, "круговым"),
        ("diagonal_pattern", 29, 165, "диагональным"),
    ]:
        design("section_divider", variant, f"Разделитель с оригинальным монохромным {label} паттерном, стр. 109.", {
            "number": field(0, (0.55, 1.35, 3.0, 1.0), 56, 48, lines=1, bold=True, color=white),
            "title": field(1, (0.55, 2.65, 10.6, 1.8), 44, 36, lines=3, bold=True, color=white),
            "subtitle": field(2, (0.55, 5.05, 9.4, 1.15), 22, 18, lines=3, color=white),
        }, [56, 109], [slide], f"ppt/slideLayouts/slideLayout{part}.xml",
               background=blue, footer=False, art="original")

    design("photo_text", "photo_right", "Текст слева и вертикальная фотография справа; композиция исходных слайдов 30–32 и стр. 110.", {
        "title": deepcopy(title),
        "image": field(1, (7.85, 1.65, 4.90, 4.92), 18, 16, kind="picture"),
        "heading": field(2, (0.55, 1.65, 6.80, 0.95), 24, 20, lines=2, bold=True),
        "bullets": field(3, (0.55, 2.85, 6.80, 3.65), 22, 18, lines=9),
        "caption": field(4, (7.85, 6.68, 4.90, 0.28), 11, 10, lines=1, color=muted),
    }, [110], [30, 31, 32], "ppt/slideLayouts/slideLayout32.xml", art="plain")
    design("photo_text", "photo_band", "Текст над широкой фотографией; адаптация допустимой фотокомпозиции со стр. 110 и нижней фотополосы слайда 177.", {
        "title": deepcopy(title),
        "image": {**field(1, (0.55, 4.10, 12.20, 2.50), 18, 16, kind="picture"), "crop_anchor_y": 0.08},
        "heading": field(2, (0.55, 1.65, 12.20, 0.65), 23, 20, lines=2, bold=True),
        "bullets": {**field(3, (0.55, 2.45, 12.20, 1.45), 18, 16, lines=5), "space_after_pt": 6},
        "caption": field(4, (0.55, 6.69, 12.20, 0.28), 11, 10, lines=1, color=muted),
    }, [110], [177], "ppt/slideLayouts/slideLayout2.xml", art="plain")

    design("image_spotlight", "white_caption", "Крупная фотография с белой нижней плашкой; адаптация фотослайдов и цветных плашек со стр. 110.", {
        "image": field(0, (0, 0, 13.3333333333, 5.20), 18, 16, kind="picture"),
        "title": field(1, (0.60, 5.42, 10.30, 1.10), 32, 28, lines=2, bold=True),
        "caption": field(2, (0.60, 6.68, 10.30, 0.48), 16, 14, lines=2, color=muted),
    }, [110], [177], "ppt/slideLayouts/slideLayout2.xml", footer=False, art="white_caption")
    design("image_spotlight", "side_panel", "Полноразмерное фото справа и синяя текстовая плашка слева; адаптация акцентных слайдов со стр. 112.", {
        "image": field(0, (5.15, 0, 8.1833333333, 7.50), 18, 16, kind="picture"),
        "title": field(1, (0.55, 2.50, 4.08, 2.20), 32, 26, lines=4, bold=True, color=white),
        "caption": field(2, (0.55, 5.45, 4.08, 1.15), 18, 16, lines=3, color=white),
    }, [110, 112], [5, 172], "ppt/slideLayouts/slideLayout18.xml",
           background=blue, footer=False, art="side_panel")

    design("bullet_list", "accent_panel", "Список на светлой плашке с фирменным синим акцентом; адаптация внутренних слайдов со стр. 110.", {
        "title": deepcopy(title),
        "bullets": field(1, (1.15, 1.90, 11.15, 4.55), 24, 20, lines=11),
    }, [110], [28], "ppt/slideLayouts/slideLayout3.xml", art="bullet_accent")
    design("bullet_list", "two_columns", "Список в двух колонках с красными акцентами; композиция внутренних слайдов со стр. 110 и исходного слайда 35.", {
        "title": deepcopy(title),
        "bullets": field(1, (0.75, 1.80, 5.55, 4.85), 22, 18, lines=11),
        "bullets_right": field(2, (7.00, 1.80, 5.55, 4.85), 22, 18, lines=11),
    }, [110], [35], "ppt/slideLayouts/slideLayout3.xml", art="bullet_columns")

    def kpi_fields(*, first_white: bool = False) -> dict:
        result = {"title": deepcopy(title)}
        for i in range(1, 4):
            color = white if i == 1 and first_white else blue
            x = 0.80 + (i - 1) * 4.19
            result[f"value_{i}"] = field(i, (x, 2.42, 3.35, 0.95), 46, 38, lines=1, bold=True, color=color)
            result[f"label_{i}"] = field(i + 3, (x, 3.62, 3.35, 0.90), 23, 20, lines=2, bold=True, color=color)
            result[f"detail_{i}"] = field(i + 6, (x, 4.82, 3.35, 1.40), 18, 16, lines=4, color=color)
        return result

    design("kpi_cards", "plain", "Три числовых акцента на белом фоне с разделителями; адаптация числовой инфографики и внутренних слайдов.",
           kpi_fields(), [110, 111], [44], "ppt/slideLayouts/slideLayout3.xml", art="kpi_plain")
    design("kpi_cards", "accent_panel", "Три карточки, первая на синей плашке; адаптация акцентов и цветных плашек со стр. 110.",
           kpi_fields(first_white=True), [110, 111], [44], "ppt/slideLayouts/slideLayout3.xml", art="kpi_accent")

    design("quote", "blue", "Белая цитата на оригинальном синем диагональном паттерне; исходный слайд 176, семейство паттернов со стр. 56.", {
        "quote": field(0, (1.0, 2.05, 11.25, 3.0), 34, 28, lines=5, color=white),
        "author": field(1, (1.0, 5.55, 11.25, 0.65), 20, 18, lines=2, bold=True, color=white),
    }, [56, 112], [165, 176], "ppt/slideLayouts/slideLayout29.xml",
           background=blue, footer=False, art="quote_blue")
    design("quote", "accent", "Крупная цитата на белом фоне с красным акцентом; адаптация цитатного слайда со стр. 112.", {
        "quote": field(0, (1.15, 1.85, 10.85, 3.35), 34, 28, lines=5),
        "author": field(1, (1.15, 5.65, 10.85, 0.65), 20, 18, lines=2, bold=True),
    }, [112], [176], "ppt/slideLayouts/slideLayout3.xml", art="quote_accent")

    design("closing", "blue_minimal", "Финальный слайд на синем фоне без суперграфики; адаптация минимальной обложки со стр. 108.", {
        "title": field(0, (0.55, 2.50, 11.40, 1.60), 46, 38, lines=2, bold=True, color=white),
        "subtitle": field(1, (0.55, 4.50, 11.40, 0.95), 24, 20, lines=2, color=white),
        "contact": field(2, (0.55, 6.00, 11.40, 0.80), 20, 18, lines=2, color=white),
    }, [108], [4, 14], "ppt/slideLayouts/slideLayout24.xml",
           background=blue, footer=False, art="closing_minimal")
    design("closing", "white_graphic", "Белый финальный слайд с оригинальным логотипом и диагональной суперграфикой исходного слайда 12.", {
        "title": field(0, (0.55, 2.50, 7.10, 1.60), 46, 38, lines=2, bold=True),
        "subtitle": field(1, (0.55, 4.50, 7.10, 0.95), 24, 20, lines=2),
        "contact": field(2, (0.55, 6.00, 7.10, 0.80), 20, 18, lines=2),
    }, [108], [12], "ppt/slideLayouts/slideLayout22.xml", footer=False, art="original")


def physical_designs(specs: dict):
    """Keep legacy indices stable: fourteen defaults, then all extra designs."""
    yield from ((name, "default", info) for name, info in specs.items())
    for name, info in specs.items():
        for variant, variant_info in info.get("variants", {}).items():
            yield name, variant, variant_info


def make_variant_layout(source: ZipFile, spec: dict, helpers: dict) -> tuple[ET._Element, bytes, dict[str, bytes]]:
    """Create native layout art and copy only the original media it references."""
    ns, q, el = helpers["NS"], helpers["q"], helpers["el"]
    field_ph, fixed, xml_bytes = helpers["placeholder"], helpers["fixed_shape"], helpers["xml_bytes"]
    blue, white = helpers["COLORS"]["primary"], "FFFFFF"
    source_part = spec["source_reference"]["source_layout_part"]
    original = ET.fromstring(source.read(source_part))
    result = deepcopy(original)
    result.set("preserve", "1")
    result.set("userDrawn", "1")
    result.set("type", "cust")
    result.set("showMasterSp", "1" if spec["show_footer"] else "0")
    common = result.find("p:cSld", ns)
    common.set("name", spec["layout_name"])
    tree = common.find("p:spTree", ns)
    style = spec.get("artwork_style", "plain")
    keep_original = style in {"original", "photo_cover", "quote_blue", "closing_minimal"}
    for shape in list(tree)[2:]:
        if not keep_original or shape.xpath(".//p:ph", namespaces=ns):
            tree.remove(shape)
    if style == "closing_minimal":
        # Source minimal cover contains author/date icons; closing uses only logo.
        for shape in list(tree)[2:]:
            names = shape.xpath(".//p:cNvPr/@name", namespaces=ns)
            if not names or names[0] != "Рисунок 3":
                tree.remove(shape)
    old_background = common.find("p:bg", ns)
    if old_background is not None:
        common.remove(old_background)
    background, props, fill = el("p", "bg"), el("p", "bgPr"), el("a", "solidFill")
    fill.append(el("a", "srgbClr", val=spec["background_color"]))
    props.extend([fill, el("a", "effectLst")])
    background.append(props)
    common.insert(0, background)

    def logo(*, white_variant: bool, box: tuple[float, float, float, float]):
        source_logo = ET.fromstring(source.read("ppt/slideLayouts/slideLayout48.xml" if white_variant else "ppt/slideLayouts/slideLayout20.xml"))
        source_tree = source_logo.find("p:cSld/p:spTree", ns)
        found = next(item for item in source_tree if item.tag == q("p", "grpSp") and item.xpath("./p:nvGrpSpPr/p:cNvPr/@name", namespaces=ns) == ["Рисунок 3"])
        shape = deepcopy(found)
        transform = shape.find("p:grpSpPr/a:xfrm", ns)
        x, y, width, height = (round(v * helpers["EMU_PER_INCH"]) for v in box)
        transform.find("a:off", ns).attrib.update({"x": str(x), "y": str(y)})
        transform.find("a:ext", ns).attrib.update({"cx": str(width), "cy": str(height)})
        return shape

    if style == "photo_cover":
        overlay = fixed("photo contrast layer", (0, 0, 13.3333333333, 7.5), blue, 3000)
        overlay.find("p:spPr/a:solidFill/a:srgbClr", ns).append(el("a", "alpha", val=42000))
        tree.insert(2, overlay)
    elif style == "white_caption":
        tree.append(fixed("white photo caption panel", (0, 5.2, 13.3333333333, 2.3), white, 3000))
        tree.append(logo(white_variant=False, box=(11.30, 5.60, 1.55, 0.483)))
    elif style == "side_panel":
        tree.append(fixed("blue side panel", (0, 0, 5.15, 7.5), blue, 3000))
        tree.append(logo(white_variant=True, box=(0.55, 0.80, 2.80, 0.872)))
    elif style == "bullet_accent":
        tree.append(fixed("bullet light panel", (0.55, 1.65, 12.20, 4.95), "F4F4F4", 3000))
        tree.append(fixed("bullet blue accent", (0.55, 1.65, 0.25, 4.95), blue, 3001))
    elif style == "bullet_columns":
        for i, x in enumerate((0.55, 6.80)):
            tree.append(fixed(f"column red accent {i + 1}", (x, 1.80, 0.045, 4.85), "FF0000", 3000 + i))
    elif style == "kpi_plain":
        for i, x in enumerate((4.52, 8.71)):
            tree.append(fixed(f"KPI separator {i + 1}", (x, 2.05, 0.018, 4.35), "DEDEDE", 3000 + i))
    elif style == "kpi_accent":
        for i in range(3):
            tree.append(fixed(f"KPI accent card {i + 1}", (0.55 + i * 4.19, 1.90, 3.85, 4.60), blue if i == 0 else "F4F4F4", 3000 + i))
    elif style == "quote_blue":
        tree.append(fixed("red quote accent", (1.0, 1.70, 0.80, 0.08), "FF0000", 3000))
    elif style == "quote_accent":
        tree.append(fixed("red quote rule", (0.65, 1.85, 0.065, 3.35), "FF0000", 3000))

    for i, (name, field_spec) in enumerate(spec["fields"].items()):
        tree.append(field_ph(name, field_spec, 5000 + i))
    for child in list(result):
        if child.tag not in {q("p", "cSld"), q("p", "clrMapOvr")}:
            result.remove(child)
    # Renumber ids throughout groups: source graphics and new shapes can overlap.
    for i, node in enumerate(result.xpath(".//p:cNvPr", namespaces=ns), 1):
        node.set("id", str(i))

    relationships = helpers["relationships"]([("rId1", "slideMaster", "../slideMasters/slideMaster1.xml")])
    media = {}
    needed_ids = set(result.xpath(".//@r:embed | .//@r:link", namespaces=ns))
    rel_path = posixpath.join(posixpath.dirname(source_part), "_rels", posixpath.basename(source_part) + ".rels")
    if rel_path in source.namelist():
        source_rels = ET.fromstring(source.read(rel_path))
        for relationship in source_rels:
            if relationship.get("Id") not in needed_ids:
                continue
            if not relationship.get("Type", "").endswith("/image") or relationship.get("TargetMode"):
                raise ValueError("Variant artwork may only reference embedded original images")
            relationships.append(deepcopy(relationship))
            part = posixpath.normpath(posixpath.join(posixpath.dirname(source_part), relationship.get("Target")))
            media[part] = source.read(part)
    found_ids = {r.get("Id") for r in relationships}
    if not needed_ids <= found_ids:
        raise ValueError("Variant artwork contains an unresolved image relationship")
    return result, xml_bytes(relationships), media
