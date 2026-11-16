"""Create a reproducible catalog: identical content across composition variants."""

from copy import deepcopy
import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from superslide.models import SUPPORTED_VARIANTS, validate_presentation


def main():
    original = json.loads((ROOT / "examples/variety.json").read_text(encoding="utf-8"))
    samples = {slide["layout_id"]: slide for slide in original["slides"]}
    samples["title_slide"]["subtitle"] = "Коллекция композиций · все показатели и планы вымышлены"
    samples["photo_text"]["bullets"] = ["Единые цели для команды", "Обратная связь сотрудников", "Фокус на качестве"]
    samples["photo_text"]["heading"] = "Фокус проекта"
    samples["bullet_list"]["bullets"] = [
        "Одна главная мысль в заголовке", "Короткие тезисы", "Подходящая композиция",
        "Фотография по теме", "Проверяемые источники", "Понятный следующий шаг",
    ]
    samples["quote"]["author"] = "Демонстрационная цитата"
    samples["closing"]["subtitle"] = "Обсудим материалы для вашего сценария"
    order = ["title_slide", "section_divider", "bullet_list", "photo_text", "image_spotlight",
             "kpi_cards", "quote", "closing", "two_columns", "gallery", "timeline",
             "comparison", "table", "bar_chart"]
    slides = []
    for layout_id in order:
        for variant_id in SUPPORTED_VARIANTS[layout_id]:
            slide = deepcopy(samples[layout_id])
            slide["variant_id"] = variant_id
            slide["notes"] = (f"Композиция {layout_id}/{variant_id}. ДЕМОНСТРАЦИОННЫЙ ПРИМЕР: "
                              "все показатели, планы и цитаты вымышлены. Внутри одного типа "
                              "используется одинаковое содержание для сравнения композиций.")
            if layout_id == "title_slide" and variant_id == "photo":
                slide["image"] = {"asset_id": "steelworks_hall", "fit": "cover",
                                  "alt_text": "Производственный цех с оборудованием"}
            slides.append(slide)
    payload = {"template_version": "1.2", "title": "SuperSlide — фирменные варианты композиций", "slides": slides}
    validate_presentation(payload)
    output = ROOT / "examples/variants.json"
    output.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"{output}: {len(slides)} slides")


if __name__ == "__main__":
    main()
