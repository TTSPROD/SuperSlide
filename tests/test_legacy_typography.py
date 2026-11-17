"""Legacy slides enforce heading/body weights despite source-template emphasis."""

import json
import unittest
from pathlib import Path
from unittest.mock import patch

from pptx import Presentation

from support import temporary_directory
from superslide.models import validate_presentation
from superslide.render import _font, build_presentation


ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "templates/legacy/corporate_base.pptx"
MANIFEST = ROOT / "templates/legacy/template_manifest.json"


class LegacyTypographyTests(unittest.TestCase):
    def setUp(self):
        workspace = temporary_directory()
        self.addCleanup(workspace.cleanup)
        self.root = Path(workspace.name)

    def build(self, slides, weights):
        manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
        # Contradict the requested result in the source metadata to prove that
        # content roles, rather than inherited bold flags, control the output.
        for layout_id, fields in weights.items():
            for field_name, expected_bold in fields.items():
                manifest["layouts"][layout_id]["fields"][field_name]["bold"] = not expected_bold
        manifest_path = self.root / "manifest.json"
        manifest_path.write_text(json.dumps(manifest, ensure_ascii=False), encoding="utf-8")
        spec = validate_presentation({"template_version": "1.2", "title": "Типографика", "slides": slides})
        output = self.root / "typography.pptx"
        build_presentation(spec, TEMPLATE, manifest_path, output)
        return Presentation(output), manifest

    def assert_field_weight(self, slide, manifest, layout_id, field_name, expected):
        idx = manifest["layouts"][layout_id]["fields"][field_name]["idx"]
        paragraphs = slide.placeholders[idx].text_frame.paragraphs
        runs = [run for paragraph in paragraphs for run in paragraph.runs if run.text]
        self.assertTrue(runs, field_name)
        self.assertTrue(all(run.font.bold is expected for run in runs), field_name)
        self.assertTrue(all(paragraph.font.bold is expected for paragraph in paragraphs), field_name)

    def test_titles_and_body_ignore_opposite_source_bold_flags(self):
        weights = {
            "title_slide": {"title": True, "subtitle": False, "author": False, "date": False},
            "two_columns": {"title": True, "left_heading": True, "right_heading": True,
                            "left_bullets": False, "right_bullets": False},
            "kpi_cards": {"title": True, **{f"label_{i}": True for i in range(1, 4)},
                          **{f"value_{i}": False for i in range(1, 4)},
                          **{f"detail_{i}": False for i in range(1, 4)}},
            "timeline": {"title": True, **{f"step_title_{i}": True for i in range(1, 5)},
                         **{f"step_text_{i}": False for i in range(1, 5)}},
            "quote": {"quote": False, "author": False},
            "table": {"title": True, "footnote": False},
        }
        slides = [
            {"layout_id": "title_slide", "title": "Спутниковые системы 2D и 3D",
             "subtitle": "Сравнение решений", "author": "Команда", "date": "2026"},
            {"layout_id": "two_columns", "title": "Технологии хранения",
             "left": {"heading": "Система 2D", "bullets": ["Работа в одном тоннеле"]},
             "right": {"heading": "Система 3D", "bullets": ["Перемещение по глубине и высоте"]}},
            {"layout_id": "kpi_cards", "title": "Показатели",
             "cards": [{"value": str(i), "label": f"Направление {i}", "detail": "Основной текст"}
                       for i in range(1, 4)]},
            {"layout_id": "timeline", "title": "Этапы",
             "steps": [{"title": f"Этап {i}", "text": "Основной текст"} for i in range(1, 5)]},
            {"layout_id": "quote", "quote": "Описание результата", "author": "Автор"},
            {"layout_id": "table", "title": "Сравнение", "columns": ["Показатель", "Значение"],
             "rows": [["Глубина", "30 м"]], "footnote": "Примечание"},
        ]
        prs, manifest = self.build(slides, weights)
        for slide, payload in zip(prs.slides, slides):
            for field_name, expected in weights[payload["layout_id"]].items():
                with self.subTest(layout=payload["layout_id"], field=field_name):
                    self.assert_field_weight(slide, manifest, payload["layout_id"], field_name, expected)
        table = next(shape.table for shape in prs.slides[-1].shapes if shape.has_table)
        for r, row in enumerate(table.rows):
            for cell in row.cells:
                self.assertTrue(all(run.font.bold is (r == 0)
                                    for paragraph in cell.text_frame.paragraphs for run in paragraph.runs))

    def test_font_fit_uses_the_same_weight_as_the_saved_text(self):
        weights = {"title_slide": {"title": True, "subtitle": False, "author": False, "date": False}}
        with patch("superslide.render._font", wraps=_font) as metrics:
            self.build([{"layout_id": "title_slide", "title": "Заголовок",
                         "subtitle": "Основной текст", "author": "Автор", "date": "2026"}], weights)
        heading_metrics = [call for call in metrics.call_args_list if call.args[0] >= 30]
        body_metrics = [call for call in metrics.call_args_list if call.args[0] < 30]
        self.assertTrue(heading_metrics)
        self.assertTrue(body_metrics)
        self.assertTrue(all(call.args[1] is True for call in heading_metrics))
        self.assertTrue(all(call.args[1] is False for call in body_metrics))


if __name__ == "__main__":
    unittest.main()
