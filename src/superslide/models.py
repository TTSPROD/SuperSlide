"""Versioned, bounded content contract. Layout geometry never comes from an LLM."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Annotated, Literal, Union

from pydantic import (
    AfterValidator,
    BaseModel,
    ConfigDict,
    Field,
    StringConstraints,
    model_validator,
)


def _xml_text(value: str) -> str:
    if any(
        ord(ch) < 32 and ch not in "\n\t\r"
        or 0xD800 <= ord(ch) <= 0xDFFF
        or ord(ch) in (0xFFFE, 0xFFFF)
        for ch in value
    ):
        raise ValueError("Текст содержит недопустимые для PowerPoint управляющие символы")
    return value


Title = Annotated[str, StringConstraints(strip_whitespace=True, min_length=1, max_length=100), AfterValidator(_xml_text)]
ShortText = Annotated[str, StringConstraints(strip_whitespace=True, max_length=180), AfterValidator(_xml_text)]
Bullet = Annotated[str, StringConstraints(strip_whitespace=True, min_length=1, max_length=130), AfterValidator(_xml_text)]
ColumnBullet = Annotated[str, StringConstraints(strip_whitespace=True, min_length=1, max_length=110), AfterValidator(_xml_text)]
Cell = Annotated[str, StringConstraints(strip_whitespace=True, max_length=80), AfterValidator(_xml_text)]
Label = Annotated[str, StringConstraints(strip_whitespace=True, min_length=1, max_length=45), AfterValidator(_xml_text)]
SourceId = Annotated[str, StringConstraints(strip_whitespace=True, min_length=1, max_length=160), AfterValidator(_xml_text)]
Notes = Annotated[str, StringConstraints(max_length=3000), AfterValidator(_xml_text)]
VariantId = Annotated[str, StringConstraints(pattern=r"^[a-z][a-z0-9_]{0,63}$")]

# The content type and its visual composition are independent. This catalog is
# shared with template validation; geometry remains solely in the manifest.
SUPPORTED_VARIANTS: dict[str, tuple[str, ...]] = {
    "title_slide": ("default", "blue_minimal", "blue_graphic", "photo"),
    "bullet_list": ("default", "accent_panel", "two_columns"),
    "two_columns": ("default",),
    "quote": ("default", "blue", "accent"),
    "table": ("default",),
    "bar_chart": ("default",),
    "section_divider": ("default", "circle_pattern", "diagonal_pattern"),
    "photo_text": ("default", "photo_right", "photo_band"),
    "image_spotlight": ("default", "white_caption", "side_panel"),
    "gallery": ("default",),
    "kpi_cards": ("default", "plain", "accent_panel"),
    "timeline": ("default",),
    "comparison": ("default",),
    "closing": ("default", "blue_minimal", "white_graphic"),
}


def _literal_chart_label(value: str) -> str:
    # The embedded workbook writer interprets leading '=' as an Excel formula.
    # Keep labels literal, matching the chart's cached values after data editing.
    if value.startswith(("=", "{=")):
        raise ValueError("Название категории или ряда диаграммы должно быть текстом, а не формулой Excel")
    return value


ChartLabel = Annotated[Label, AfterValidator(_literal_chart_label)]


class StrictModel(BaseModel):
    model_config = ConfigDict(extra="forbid", strict=True, allow_inf_nan=False)


class SlideBase(StrictModel):
    variant_id: VariantId = "default"
    source_refs: list[SourceId] = Field(default_factory=list, max_length=30)
    notes: Notes = ""

    @model_validator(mode="after")
    def unique_sources(self):
        if len(self.source_refs) != len(set(self.source_refs)):
            raise ValueError("source_refs содержит повторяющиеся идентификаторы")
        return self


class ImageRef(StrictModel):
    asset_id: SourceId
    fit: Literal["cover", "contain"] = "cover"
    alt_text: ShortText = ""


class TitleSlide(SlideBase):
    layout_id: Literal["title_slide"]
    variant_id: Literal[*SUPPORTED_VARIANTS["title_slide"]] = "default"
    title: Title
    subtitle: ShortText = ""
    author: ShortText = ""
    date: Annotated[str, StringConstraints(max_length=40), AfterValidator(_xml_text)] = ""
    image: ImageRef | None = None

    @model_validator(mode="after")
    def image_matches_variant(self):
        if self.variant_id == "photo" and self.image is None:
            raise ValueError("Вариант title_slide/photo требует image")
        if self.variant_id != "photo" and self.image is not None:
            raise ValueError("image разрешено только для варианта title_slide/photo")
        return self


class BulletSlide(SlideBase):
    layout_id: Literal["bullet_list"]
    variant_id: Literal[*SUPPORTED_VARIANTS["bullet_list"]] = "default"
    title: Title
    bullets: list[Bullet] = Field(min_length=1, max_length=6)


class Column(StrictModel):
    heading: Title
    bullets: list[ColumnBullet] = Field(min_length=1, max_length=4)


class TwoColumnsSlide(SlideBase):
    layout_id: Literal["two_columns"]
    variant_id: Literal[*SUPPORTED_VARIANTS["two_columns"]] = "default"
    title: Title
    left: Column
    right: Column


class QuoteSlide(SlideBase):
    layout_id: Literal["quote"]
    variant_id: Literal[*SUPPORTED_VARIANTS["quote"]] = "default"
    quote: Annotated[str, StringConstraints(strip_whitespace=True, min_length=1, max_length=220), AfterValidator(_xml_text)]
    author: ShortText


class TableSlide(SlideBase):
    layout_id: Literal["table"]
    variant_id: Literal[*SUPPORTED_VARIANTS["table"]] = "default"
    title: Title
    columns: list[Label] = Field(min_length=1, max_length=5)
    rows: list[list[Cell]] = Field(min_length=1, max_length=8)
    footnote: ShortText = ""

    @model_validator(mode="after")
    def rectangular(self):
        if any(len(row) != len(self.columns) for row in self.rows):
            raise ValueError("Число ячеек каждой строки должно совпадать с числом столбцов")
        if len(set(self.columns)) != len(self.columns):
            raise ValueError("Заголовки столбцов таблицы должны различаться")
        return self


class ChartSeries(StrictModel):
    name: ChartLabel
    values: list[float] = Field(min_length=1, max_length=8)
    role: Literal["company", "comparison", "neutral"] = "company"


class BarChartSlide(SlideBase):
    layout_id: Literal["bar_chart"]
    variant_id: Literal[*SUPPORTED_VARIANTS["bar_chart"]] = "default"
    title: Title
    categories: list[ChartLabel] = Field(min_length=1, max_length=8)
    series: list[ChartSeries] = Field(min_length=1, max_length=3)
    unit: Annotated[str, StringConstraints(max_length=60), AfterValidator(_xml_text)] = ""
    takeaway: ShortText = ""

    @model_validator(mode="after")
    def matched_data(self):
        if any(len(s.values) != len(self.categories) for s in self.series):
            raise ValueError("Число значений каждого ряда должно совпадать с числом категорий")
        if len(set(self.categories)) != len(self.categories):
            raise ValueError("Категории диаграммы должны различаться")
        if len({s.name for s in self.series}) != len(self.series):
            raise ValueError("Названия рядов диаграммы должны различаться")
        return self


class SectionDividerSlide(SlideBase):
    layout_id: Literal["section_divider"]
    variant_id: Literal[*SUPPORTED_VARIANTS["section_divider"]] = "default"
    number: Annotated[str, StringConstraints(max_length=12), AfterValidator(_xml_text)] = ""
    title: Title
    subtitle: ShortText = ""


class PhotoTextSlide(SlideBase):
    layout_id: Literal["photo_text"]
    variant_id: Literal[*SUPPORTED_VARIANTS["photo_text"]] = "default"
    title: Title
    image: ImageRef
    heading: Annotated[str, StringConstraints(strip_whitespace=True, max_length=80), AfterValidator(_xml_text)] = ""
    bullets: list[ColumnBullet] = Field(min_length=1, max_length=4)
    caption: Annotated[str, StringConstraints(max_length=120), AfterValidator(_xml_text)] = ""


class ImageSpotlightSlide(SlideBase):
    layout_id: Literal["image_spotlight"]
    variant_id: Literal[*SUPPORTED_VARIANTS["image_spotlight"]] = "default"
    image: ImageRef
    title: Title
    caption: Annotated[str, StringConstraints(max_length=140), AfterValidator(_xml_text)] = ""


class GalleryImage(ImageRef):
    caption: Annotated[str, StringConstraints(max_length=90), AfterValidator(_xml_text)] = ""


class GallerySlide(SlideBase):
    layout_id: Literal["gallery"]
    variant_id: Literal[*SUPPORTED_VARIANTS["gallery"]] = "default"
    title: Title
    images: list[GalleryImage] = Field(min_length=3, max_length=3)

    @model_validator(mode="after")
    def distinct_images(self):
        if len({image.asset_id for image in self.images}) != len(self.images):
            raise ValueError("Галерея должна содержать три разных изображения")
        return self


class KpiCard(StrictModel):
    value: Annotated[str, StringConstraints(strip_whitespace=True, min_length=1, max_length=20), AfterValidator(_xml_text)]
    label: Annotated[str, StringConstraints(strip_whitespace=True, min_length=1, max_length=65), AfterValidator(_xml_text)]
    detail: Annotated[str, StringConstraints(max_length=110), AfterValidator(_xml_text)] = ""


class KpiCardsSlide(SlideBase):
    layout_id: Literal["kpi_cards"]
    variant_id: Literal[*SUPPORTED_VARIANTS["kpi_cards"]] = "default"
    title: Title
    cards: list[KpiCard] = Field(min_length=3, max_length=3)


class TimelineStep(StrictModel):
    title: Label
    text: Annotated[str, StringConstraints(max_length=110), AfterValidator(_xml_text)] = ""


class TimelineSlide(SlideBase):
    layout_id: Literal["timeline"]
    variant_id: Literal[*SUPPORTED_VARIANTS["timeline"]] = "default"
    title: Title
    steps: list[TimelineStep] = Field(min_length=4, max_length=4)


class ComparisonSlide(SlideBase):
    layout_id: Literal["comparison"]
    variant_id: Literal[*SUPPORTED_VARIANTS["comparison"]] = "default"
    title: Title
    left: Column
    right: Column


class ClosingSlide(SlideBase):
    layout_id: Literal["closing"]
    variant_id: Literal[*SUPPORTED_VARIANTS["closing"]] = "default"
    title: Title
    subtitle: ShortText = ""
    contact: Annotated[str, StringConstraints(max_length=100), AfterValidator(_xml_text)] = ""


Slide = Annotated[
    Union[TitleSlide, BulletSlide, TwoColumnsSlide, QuoteSlide, TableSlide, BarChartSlide,
          SectionDividerSlide, PhotoTextSlide, ImageSpotlightSlide, GallerySlide,
          KpiCardsSlide, TimelineSlide, ComparisonSlide, ClosingSlide],
    Field(discriminator="layout_id"),
]


class PresentationSpec(StrictModel):
    template_version: Literal["1.0", "1.1", "1.2"]
    title: Annotated[str, StringConstraints(strip_whitespace=True, min_length=1, max_length=150), AfterValidator(_xml_text)]
    slides: list[Slide] = Field(min_length=1, max_length=40)

    @model_validator(mode="after")
    def versioned_layouts(self):
        original = {"title_slide", "bullet_list", "two_columns", "quote", "table", "bar_chart"}
        if self.template_version == "1.0" and any(slide.layout_id not in original for slide in self.slides):
            raise ValueError("Новые макеты требуют template_version 1.1")
        if self.template_version != "1.2" and any(slide.variant_id != "default" for slide in self.slides):
            raise ValueError("Варианты оформления требуют template_version 1.2")
        return self


def iter_image_refs(spec: PresentationSpec):
    for slide in spec.slides:
        if isinstance(slide, TitleSlide) and slide.image is not None:
            yield slide.image
        elif isinstance(slide, (PhotoTextSlide, ImageSpotlightSlide)):
            yield slide.image
        elif isinstance(slide, GallerySlide):
            yield from slide.images


def validate_presentation(payload: dict, known_source_ids: set[str] | None = None) -> PresentationSpec:
    spec = PresentationSpec.model_validate(payload)
    if known_source_ids is not None:
        unknown = {
            ref for slide in spec.slides for ref in slide.source_refs
            if ref not in known_source_ids
        }
        if unknown:
            raise ValueError("Неизвестные источники: " + ", ".join(sorted(unknown)))
    return spec


def load_presentation(path: str | Path, known_source_ids: set[str] | None = None) -> PresentationSpec:
    path = Path(path)
    if path.stat().st_size > 1_000_000:
        raise ValueError("JSON презентации превышает 1 МБ")
    with path.open(encoding="utf-8-sig") as stream:
        data = json.load(stream)
    return validate_presentation(data, known_source_ids)


def presentation_schema() -> dict:
    return PresentationSpec.model_json_schema()
