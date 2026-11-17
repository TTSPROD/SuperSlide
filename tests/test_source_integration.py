"""Exercise source-template CLI and model selection without contacting an API."""

import io
import json
import os
import socket
import unittest
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
from unittest.mock import MagicMock, patch

from support import temporary_directory
from superslide.cli import main
from superslide.config import LLMSettings
from superslide.documents import DocumentBundle, SourceChunk
from superslide.llm import LLMError, generate_draft


ROOT = Path(__file__).resolve().parents[1]


def response(payload):
    value = MagicMock()
    value.__enter__.return_value = value
    value.read.return_value = json.dumps({"choices": [{
        "message": {"content": json.dumps(payload, ensure_ascii=False)}, "finish_reason": "stop",
    }]}).encode("utf-8")
    return value


class SourcePromptTests(unittest.TestCase):
    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_planner_must_count_items_for_reviewed_repeating_blocks(self, request):
        from superslide.llm import _select_source_designs
        from superslide.source_template import load_source_manifest, source_prompt_catalog
        catalog = source_prompt_catalog(load_source_manifest(ROOT / "templates/severstal_template.json"))
        missing = {"title": "Решения", "slides": [
            {"layout_id": "cards", "variant_id": "slide_006", "item_count": None}], "icon_groups": []}
        corrected = json.loads(json.dumps(missing))
        corrected["slides"][0]["item_count"] = 3
        request.side_effect = [response(missing), response(corrected)]
        _, _, result = _select_source_designs("Три решения", DocumentBundle(),
                                              LLMSettings("https://mock.invalid/v1", "test"), catalog, [])
        self.assertEqual(result.draft["slides"][0]["item_count"], 3)
        self.assertEqual(request.call_count, 2)
        self.assertIn("item_count", json.loads(request.call_args.args[0].data)["messages"][-1]["content"])

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_planner_uses_reviewed_block_count_when_variant_name_has_no_count(self, request):
        from superslide.llm import _select_source_designs
        from superslide.source_template import load_source_manifest, source_prompt_catalog
        catalog = source_prompt_catalog(load_source_manifest(ROOT / "templates/severstal_template.json"))
        self.assertEqual(catalog["cards"]["variants"]["slide_006"]["item_slots"]["primary_blocks"]["count"], 3)
        self.assertEqual(catalog["cards"]["variants"]["slide_007"]["item_slots"]["primary_blocks"]["count"], 4)
        wrong = {"title": "Четыре решения", "slides": [
            {"layout_id": "cards", "variant_id": "slide_006", "item_count": 4}], "icon_groups": []}
        fixed = json.loads(json.dumps(wrong))
        fixed["slides"][0]["variant_id"] = "slide_007"
        request.side_effect = [response(wrong), response(fixed)]
        selected, _, result = _select_source_designs("Четыре решения", DocumentBundle(),
                                                     LLMSettings("https://mock.invalid/v1", "test"), catalog, [])
        self.assertEqual(result.draft["slides"][0]["variant_id"], "slide_007")
        self.assertEqual(request.call_count, 2)
        payload = json.loads(json.loads(request.call_args_list[0].args[0].data)["messages"][1]["content"])
        self.assertEqual(payload["layout_catalog"]["cards"]["variants"]["slide_006"]["item_slots"]["primary_blocks"]["count"], 3)
        self.assertIn("primary_blocks", json.loads(request.call_args_list[0].args[0].data)["messages"][0]["content"])
        self.assertEqual(set(selected["cards"]["variants"]), {"slide_007"})

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_planner_matches_real_cards_with_different_russian_inflections(self, request):
        from superslide.llm import _select_source_designs
        from superslide.source_template import load_source_manifest, source_prompt_catalog
        manifest = load_source_manifest(ROOT / "templates/severstal_template.json")
        catalog = source_prompt_catalog(manifest)
        self.assertEqual(catalog["cards"]["variants"]["slide_005"]["item_slots"]["named_item_kind"], "карточек")
        self.assertEqual(catalog["cards"]["variants"]["slide_016"]["item_slots"]["named_item_kind"], "карточки")
        incorrect = {"title": "Два преимущества", "slides": [
            {"layout_id": "cards", "variant_id": "slide_005", "item_count": 2}], "icon_groups": []}
        corrected = json.loads(json.dumps(incorrect))
        corrected["slides"][0]["variant_id"] = "slide_016"
        request.side_effect = [response(incorrect), response(corrected)]
        selected, _, result = _select_source_designs("Два преимущества", DocumentBundle(),
                                                     LLMSettings("https://mock.invalid/v1", "test"), catalog, [])
        self.assertEqual(request.call_count, 2)
        self.assertEqual(result.draft["slides"][0]["variant_id"], "slide_016")
        self.assertEqual(set(selected["cards"]["variants"]), {"slide_016"})

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_planner_does_not_match_unrelated_types_with_same_count(self, request):
        from superslide.llm import _select_source_designs
        variants = {
            "slide_001": {"item_slots": {"named_item_count": 4, "named_item_kind": "карточки", "groups": []}, "fields": {}},
            "slide_002": {"item_slots": {"named_item_count": 2, "named_item_kind": "диаграммы", "groups": []}, "fields": {}},
        }
        plan = {"title": "Два преимущества", "slides": [
            {"layout_id": "cards", "variant_id": "slide_001", "item_count": 2}], "icon_groups": []}
        request.return_value = response(plan)
        _, _, result = _select_source_designs("Два преимущества", DocumentBundle(),
                                              LLMSettings("https://mock.invalid/v1", "test"),
                                              {"cards": {"variants": variants}}, [])
        self.assertEqual(request.call_count, 1)
        self.assertEqual(result.draft["slides"][0]["variant_id"], "slide_001")

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_fact_count_does_not_force_a_different_number_of_charts(self, request):
        from superslide.llm import _select_source_designs
        from superslide.source_template import load_source_manifest, source_prompt_catalog
        manifest = load_source_manifest(ROOT / "templates/severstal_template.json")
        catalog = source_prompt_catalog(manifest)
        plan = {"title": "Два результата", "slides": [
            {"layout_id": "diagram", "variant_id": "slide_004", "item_count": 2}], "icon_groups": []}
        request.return_value = response(plan)
        selected, _, result = _select_source_designs("Два результата", DocumentBundle(),
                                                     LLMSettings("https://mock.invalid/v1", "test"), catalog, [])
        self.assertEqual(request.call_count, 1)
        self.assertEqual(result.draft["slides"][0]["variant_id"], "slide_004")
        self.assertEqual(set(selected["diagram"]["variants"]), {"slide_004"})

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_planner_repairs_wrong_number_of_cards_when_matching_variant_exists(self, request):
        from superslide.llm import _select_source_designs
        variants = {
            "slide_001": {"name": "Две карточки", "item_slots": {"named_item_count": 2, "groups": []}, "fields": {}},
            "slide_002": {"name": "Четыре карточки", "item_slots": {"named_item_count": 4, "groups": []}, "fields": {}},
        }
        catalog = {"cards": {"name": "Карточки", "variants": variants}}
        plan = {"title": "Преимущества", "slides": [{"layout_id": "cards", "variant_id": "slide_002", "item_count": 2}], "icon_groups": []}
        corrected = json.loads(json.dumps(plan))
        corrected["slides"][0]["variant_id"] = "slide_001"
        request.side_effect = [response(plan), response(corrected)]
        selected, _, result = _select_source_designs("Два подтверждённых преимущества", DocumentBundle(),
                                                     LLMSettings("https://mock.invalid/v1", "test"), catalog, [])
        self.assertEqual(request.call_count, 2)
        self.assertEqual(result.draft["slides"][0]["item_count"], 2)
        self.assertEqual(set(selected["cards"]["variants"]), {"slide_001"})
        request_payload = json.loads(request.call_args_list[0].args[0].data)
        catalog_payload = json.loads(request_payload["messages"][1]["content"])["layout_catalog"]
        self.assertEqual(catalog_payload["cards"]["variants"]["slide_001"]["item_slots"]["named_item_count"], 2)
        self.assertIn("самостоятельных фактов", request_payload["messages"][0]["content"])
        repair_payload = json.loads(request.call_args_list[1].args[0].data)
        self.assertIn("4 блоков", repair_payload["messages"][-1]["content"])

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_no_forced_matching_when_catalog_has_no_exact_variant(self, request):
        from superslide.llm import _select_source_designs
        variants = {"slide_001": {"item_slots": {"named_item_count": 4, "groups": []}, "fields": {}}}
        request.return_value = response({"title": "Преимущества", "slides": [
            {"layout_id": "cards", "variant_id": "slide_001", "item_count": 2}], "icon_groups": []})
        _, _, result = _select_source_designs("Два преимущества", DocumentBundle(),
                                             LLMSettings("https://mock.invalid/v1", "test"),
                                             {"cards": {"variants": variants}}, [])
        self.assertEqual(result.draft["slides"][0]["item_count"], 2)
        self.assertEqual(request.call_count, 1)

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_source_prompt_exposes_bindings_and_catalog_ids_without_asset_paths(self, request):
        request.return_value = response({"title": "Отчёт"})
        generate_draft("Отчёт", DocumentBundle(), LLMSettings("https://mock.invalid/v1", "test"),
                       lambda draft: None, {"type": "object"},
                       {"title": {"variants": {"slide_001": {"fields": {"slide_text_3": {
                           "kind": "text", "sample_text": "Заголовок", "max_characters": 30,
                       }}}}}}, template_version="2.0",
                       image_catalog=[{"id": "factory", "description": "Цех", "file": "private/photo.jpg"}],
                       icon_catalog=[{"id": "crane", "name": "Кран", "group_id": "metallurgy", "path": "private/icon.svg"}])
        body = json.loads(request.call_args.args[0].data)
        payload = json.loads(body["messages"][1]["content"])
        self.assertEqual(payload["photo_catalog"][0]["id"], "factory")
        self.assertEqual(payload["icon_catalog"][0]["id"], "crane")
        self.assertIn("slide_text_3", payload["layout_catalog"]["title"]["variants"]["slide_001"]["fields"])
        self.assertNotIn("private/", json.dumps(body))
        self.assertIn("updates", body["messages"][0]["content"])
        self.assertIn("В корневом поле title", body["messages"][0]["content"])
        self.assertIn("2–8 слов", body["messages"][0]["content"])
        self.assertEqual(request.call_count, 1)

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_large_catalog_selects_existing_designs_and_icon_groups_before_content(self, request):
        request.side_effect = [response({"title": "Технологии промышленной автоматизации",
                                          "slides": [{"layout_id": "title", "variant_id": "slide_002"}],
                                          "icon_groups": ["tech"]}), response({"template_version": "2.0", "title": "Готово",
                    "slides": [{"layout_id": "title", "variant_id": "slide_002", "updates": [
                        {"field_id": "slide_icon_2", "icon_id": "server"}]}]})]
        catalog = {"title": {"name": "Титульные", "variants": {
            f"slide_{number:03}": {"description": f"Вариант {number}", "fields": {"slide_text_1": {
                "kind": "text", "sample_text": "Длинный текст " * 6000,
            }, "slide_icon_2": {"kind": "icon"}}} for number in range(1, 5)}}}
        result = generate_draft("Отчёт", DocumentBundle(), LLMSettings("https://mock.invalid/v1", "test"),
                                lambda draft: None, {"type": "object"}, catalog, template_version="2.0",
                                icon_catalog=[{"id": "server", "group_id": "tech", "name": "Сервер"},
                                              {"id": "tree", "group_id": "ecology", "name": "Дерево"}])
        self.assertEqual(request.call_count, 2)
        self.assertEqual(result.attempts, 2)
        self.assertEqual(result.draft["title"], "Технологии промышленной автоматизации")
        planning = json.loads(request.call_args_list[0].args[0].data)
        self.assertNotIn("sample_text", planning["messages"][1]["content"])
        self.assertIn("В корневом поле title", planning["messages"][0]["content"])
        self.assertIn("title", json.loads(planning["messages"][1]["content"])["required_output_schema"]["required"])
        final = json.loads(request.call_args_list[1].args[0].data)
        payload = json.loads(final["messages"][1]["content"])
        self.assertEqual(set(payload["layout_catalog"]["title"]["variants"]), {"slide_002"})
        self.assertEqual([icon["id"] for icon in payload["icon_catalog"]], ["server"])


class SourceBatchTests(unittest.TestCase):
    @staticmethod
    def catalog():
        return {"title": {"name": "Титульные", "variants": {
            f"slide_{number:03}": {"name": f"Вариант {number}", "fields": {
                "slide_text_1": {"kind": "text", "sample_text": "Название", "max_characters": 40, "paragraph_count": 1}
            }} for number in range(1, 8)
        }}}

    @staticmethod
    def plan():
        return {"title": "Отчёт", "slides": [
            {"layout_id": "title", "variant_id": f"slide_{number:03}", "purpose": f"Часть {number}"}
            for number in range(1, 5)
        ], "icon_groups": []}

    @staticmethod
    def fragment(start, end):
        return {"template_version": "2.0", "title": "Отчёт", "slides": [
            {"layout_id": "title", "variant_id": f"slide_{number:03}", "updates": [
                {"field_id": "slide_text_1", "text": f"Слайд {number}"}]} for number in range(start, end + 1)
        ]}

    def generate(self, validate, progress, settings=None, **kwargs):
        return generate_draft("Подготовь ровно 4 слайда", DocumentBundle(),
                              settings or LLMSettings("https://mock.invalid/v1", "test", max_output_tokens=16000),
                              validate, {"type": "object"}, self.catalog(), template_version="2.0",
                              progress_callback=progress.append, **kwargs)

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_fills_two_small_slides_at_a_time_with_complete_presentation_validation(self, request):
        request.side_effect = [response(self.plan()), response(self.fragment(1, 2)), response(self.fragment(3, 4))]
        validated, progress = [], []

        def validate(value):
            self.assertEqual(len(value["slides"]), 4)
            validated.append(json.loads(json.dumps(value)))

        result = self.generate(validate, progress,
                               image_catalog=[{"id": "photo", "description": "Фото"}],
                               icon_catalog=[{"id": "icon", "group_id": "tech", "name": "Иконка"}])
        self.assertEqual(request.call_count, 3)
        self.assertEqual(result.attempts, 3)
        self.assertEqual(len(result.draft["slides"]), 4)
        self.assertTrue(all(s["updates"] for s in result.draft["slides"]))
        self.assertEqual([p["stage"] for p in progress], ["planning", "slides", "slides", "validating"])
        self.assertEqual([(p["slide_start"], p["slide_end"]) for p in progress if p["stage"] == "slides"], [(1, 2), (3, 4)])
        for call in request.call_args_list:
            self.assertLess(json.loads(call.args[0].data)["max_tokens"], 16000)
        for index, ids in ((1, {"slide_001", "slide_002"}), (2, {"slide_003", "slide_004"})):
            body = json.loads(request.call_args_list[index].args[0].data)
            payload = json.loads(body["messages"][1]["content"])
            self.assertEqual(set(payload["layout_catalog"]["title"]["variants"]), ids)
            self.assertEqual(payload["photo_catalog"], [])
            self.assertEqual(payload["icon_catalog"], [])
        self.assertGreaterEqual(len(validated), 4)

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_planned_presentation_title_is_preserved_across_all_content_fragments(self, request):
        plan = self.plan()
        plan["title"] = "Результаты автоматизации склада"
        first, final = self.fragment(1, 2), self.fragment(3, 4)
        first["title"], final["title"] = "Presentation", "Заключение"
        request.side_effect = [response(plan), response(first), response(final)]
        titles = []
        schema = {"type": "object", "properties": {"title": {"type": "string", "maxLength": 150}}}
        result = generate_draft("Подготовь ровно 4 слайда", DocumentBundle(),
                                LLMSettings("https://mock.invalid/v1", "test", max_output_tokens=16000),
                                lambda value: titles.append(value["title"]), schema,
                                self.catalog(), template_version="2.0")
        self.assertEqual(request.call_count, 3)
        self.assertEqual(result.draft["title"], plan["title"])
        self.assertTrue(titles)
        self.assertEqual(set(titles), {plan["title"]})
        for call in request.call_args_list[1:]:
            payload = json.loads(json.loads(call.args[0].data)["messages"][1]["content"])
            self.assertIn(plan["title"], payload["task"])
            self.assertEqual(payload["required_output_schema"]["properties"]["title"]["const"], plan["title"])

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_missing_plan_title_is_repaired_before_content_generation(self, request):
        missing = self.plan()
        del missing["title"]
        named = self.plan()
        named["title"] = "Результаты автоматизации склада"
        request.side_effect = [response(missing), response(named),
                               response(self.fragment(1, 2)), response(self.fragment(3, 4))]
        result = self.generate(lambda value: None, [])
        self.assertEqual(request.call_count, 4)
        self.assertEqual(result.draft["title"], named["title"])
        repair = json.loads(request.call_args_list[1].args[0].data)
        self.assertIn("title", repair["messages"][-1]["content"])

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_timeout_preserves_completed_slides_and_is_not_retried(self, request):
        request.side_effect = [response(self.plan()), response(self.fragment(1, 2)), socket.timeout("private endpoint")]
        progress = []
        with self.assertRaises(LLMError) as failure:
            self.generate(lambda value: None, progress)
        error = failure.exception
        self.assertEqual(error.code, "timeout")
        self.assertEqual(request.call_count, 3)
        self.assertEqual(error.context["stage"], "slides")
        self.assertEqual(error.context["slide_start"], 3)
        self.assertEqual(error.context["completed_slides"], 2)
        self.assertEqual([s["variant_id"] for s in error.partial_draft["slides"]], ["slide_001", "slide_002"])
        self.assertNotIn("private", json.dumps(error.context))

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_only_invalid_fragment_is_repaired_and_error_details_keep_global_slide_index(self, request):
        invalid = self.fragment(3, 4)
        invalid["slides"][0]["updates"][0]["field_id"] = "slide_text_999"
        request.side_effect = [response(self.plan()), response(self.fragment(1, 2)), response(invalid), response(self.fragment(3, 4))]
        progress = []
        result = self.generate(lambda value: None, progress)
        self.assertEqual(result.attempts, 4)
        repair = next(p for p in progress if p["stage"] == "repair")
        self.assertEqual((repair["slide_start"], repair["slide_end"], repair["attempt"]), (3, 4, 2))
        sent = json.loads(request.call_args.args[0].data)
        self.assertEqual(len(sent["messages"]), 4)
        self.assertIn("Слайд 3", sent["messages"][-1]["content"])
        self.assertNotIn("Слайд 1", json.dumps(sent["messages"], ensure_ascii=False))

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_failed_fragment_has_structured_validation_details(self, request):
        invalid = self.fragment(3, 4)
        invalid["slides"][0]["updates"][0]["field_id"] = "slide_text_999"
        request.side_effect = [response(self.plan()), response(self.fragment(1, 2)), response(invalid)]
        with self.assertRaises(LLMError) as failure:
            self.generate(lambda value: None, [], LLMSettings("https://mock.invalid/v1", "test", max_repair_attempts=0))
        details = failure.exception.diagnostics[0]["validation_details"]
        self.assertEqual(details, {"slide_index": 3, "reason": "unknown_field", "field_id": "slide_text_999"})


class SourceTextRepairTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        from superslide.source_template import load_source_manifest, source_prompt_catalog, source_presentation_schema
        cls.manifest = load_source_manifest(ROOT / "templates/severstal_template.json")
        cls.catalog = source_prompt_catalog(cls.manifest)
        cls.schema = source_presentation_schema()

    def fixture(self):
        plan = {"title": "Результаты", "icon_groups": [], "slides": [
            {"layout_id": "title", "variant_id": "slide_003"},
            {"layout_id": "bullets", "variant_id": "slide_012", "item_count": 3},
            {"layout_id": "last_slide", "variant_id": "slide_001"},
        ]}
        def long_paragraph(heading, body, actual):
            return [heading, body + "а" * (actual - len(heading) - len(body))]
        invalid = {"template_version": "2.0", "title": "Результаты", "slides": [
            {"layout_id": "title", "variant_id": "slide_003", "updates": [], "notes": "Обложка"},
            {"layout_id": "bullets", "variant_id": "slide_012", "source_refs": ["doc1"], "notes": "Пояснение докладчика",
             "updates": [{"field_id": "slide_text_2", "text": "Итоги"},
                         {"field_id": "slide_text_22", "paragraphs": long_paragraph("Результат", "Рост на 42%. ", 284)},
                         {"field_id": "slide_text_27", "paragraphs": long_paragraph("Эффект", "Экономия 12 млн руб. ", 146)},
                         {"field_id": "slide_text_25", "paragraphs": ["Показатель", "Рост производительности."]},
                         {"field_id": "slide_text_20", "paragraphs": ["Ресурсы", "Учет экономии затрат."]}]},
        ]}
        patch_value = {"field_repairs": [
            {"slide_index": 2, "field_id": "slide_text_22", "paragraphs": ["Результат", "Производительность выросла на 42%."]},
            {"slide_index": 2, "field_id": "slide_text_27", "paragraphs": ["Эффект", "Экономия — 12 млн руб."]},
        ]}
        closing = {"template_version": "2.0", "title": "Результаты", "slides": [
            {"layout_id": "last_slide", "variant_id": "slide_001", "updates": []}]}
        return plan, invalid, patch_value, closing

    def generate(self):
        from superslide.source_template import validate_source_presentation
        def validate(value):
            self.assertEqual(len(value["slides"]), 3)
            return validate_source_presentation(value, self.manifest, {"doc1"})
        return generate_draft("Подготовь ровно 3 слайда", DocumentBundle([
            SourceChunk("doc1", "facts.txt", "text", "Рост производительности 42%, экономия 12 млн руб.")]),
            LLMSettings("https://mock.invalid/v1", "test"), validate, self.schema, self.catalog, template_version="2.0")

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_all_overflow_fields_are_repaired_together_without_valid_field_drift(self, request):
        plan, invalid, patch_value, closing = self.fixture()
        request.side_effect = [response(plan), response(invalid), response(patch_value), response(closing)]
        result = self.generate()
        self.assertEqual(request.call_count, 4)
        slide = result.draft["slides"][1]
        self.assertEqual(slide["source_refs"], ["doc1"])
        self.assertEqual(slide["notes"], "Пояснение докладчика")
        self.assertEqual(slide["updates"][0], invalid["slides"][1]["updates"][0])
        self.assertEqual(slide["updates"][1]["paragraphs"], patch_value["field_repairs"][0]["paragraphs"])
        repair = json.loads(request.call_args_list[2].args[0].data)
        fields = json.loads(repair["messages"][-1]["content"])["field_repair_request"]["fields"]
        self.assertEqual([(f["slide_index"], f["field_id"], f["current_characters"], f["max_characters"], f["target_characters"])
                          for f in fields], [(2, "slide_text_22", 284, 224, 190), (2, "slide_text_27", 146, 116, 98)])
        self.assertEqual([f["characters_to_remove_for_target"] for f in fields], [94, 48])
        self.assertIn("Считаются все символы", repair["messages"][0]["content"])
        original = json.loads(request.call_args_list[1].args[0].data)
        payload = json.loads(original["messages"][1]["content"])
        schema = payload["required_output_schema"]["properties"]["slides"]["items"]["anyOf"]
        bullet = next(s for s in schema if s["properties"]["layout_id"]["const"] == "bullets")
        field = next(f for f in bullet["properties"]["updates"]["items"]["anyOf"] if f["properties"]["field_id"]["const"] == "slide_text_22")
        self.assertEqual(field["properties"]["text"], {"type": "null"})
        self.assertEqual(field["properties"]["paragraphs"]["minItems"], 2)
        self.assertEqual(field["properties"]["paragraphs"]["items"]["maxLength"], 224)

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_full_json_repair_cannot_rewrite_other_fields_or_source_refs(self, request):
        plan, invalid, patch_value, closing = self.fixture()
        full = json.loads(json.dumps(invalid))
        full["title"] = "Постороннее название"
        full["slides"][1]["source_refs"] = []
        full["slides"][1]["notes"] = "Потерянное пояснение"
        full["slides"][1]["updates"][0]["text"] = "Изменение"
        for index, patch_item in enumerate(patch_value["field_repairs"], 1):
            full["slides"][1]["updates"][index]["paragraphs"] = patch_item["paragraphs"]
        request.side_effect = [response(plan), response(invalid), response(full), response(closing)]
        result = self.generate()
        self.assertEqual(result.draft["title"], "Результаты")
        self.assertEqual(result.draft["slides"][1]["notes"], "Пояснение докладчика")
        self.assertEqual(result.draft["slides"][1]["source_refs"], ["doc1"])
        self.assertEqual(result.draft["slides"][1]["updates"][0]["text"], "Итоги")

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_shortening_cannot_drop_numeric_facts(self, request):
        plan, invalid, patch_value, closing = self.fixture()
        bad = json.loads(json.dumps(patch_value))
        bad["field_repairs"][0]["paragraphs"][1] = "Производительность выросла."
        fix_one = {"field_repairs": [patch_value["field_repairs"][0]]}
        request.side_effect = [response(plan), response(invalid), response(bad), response(fix_one), response(closing)]
        result = self.generate()
        self.assertEqual(request.call_count, 5)
        self.assertIn("42%", result.draft["slides"][1]["updates"][1]["paragraphs"][1])
        repair = json.loads(request.call_args_list[3].args[0].data)
        fields = json.loads(repair["messages"][-1]["content"])["field_repair_request"]["fields"]
        self.assertEqual([f["field_id"] for f in fields], ["slide_text_22"])
        self.assertEqual(fields[0]["preserve_numeric_literals"], ["42"])


class SourceCLIIntegrationTests(unittest.TestCase):
    def invoke(self, args):
        out, err = io.StringIO(), io.StringIO()
        with redirect_stdout(out), redirect_stderr(err):
            status = main(args)
        return status, out.getvalue(), err.getvalue()

    def test_default_inspect_and_schema_use_exactly_source_types(self):
        status, text, errors = self.invoke(["inspect-template"])
        self.assertEqual(status, 0, errors)
        inspected = json.loads(text)
        self.assertEqual(inspected["layout_count"], 15)
        self.assertEqual(inspected["design_count"], 211)
        self.assertEqual(inspected["render_mode"], "source_native")
        status, text, errors = self.invoke(["schema"])
        self.assertEqual(status, 0, errors)
        schema = json.loads(text)
        self.assertEqual(schema["properties"]["template_version"]["const"], "2.0")
        self.assertEqual(len(schema["$defs"]["SourceSlide"]["properties"]["layout_id"]["enum"]), 15)

    def test_source_json_validates_and_builds_through_cli(self):
        with temporary_directory() as directory:
            output = Path(directory) / "source.pptx"
            status, _, errors = self.invoke(["validate", str(ROOT / "examples/severstal_15_types.json")])
            self.assertEqual(status, 0, errors)
            status, _, errors = self.invoke(["build", str(ROOT / "examples/severstal_15_types.json"),
                                             "--output", str(output)])
            self.assertEqual(status, 0, errors)
            self.assertTrue(output.is_file())
            report = json.loads(output.with_suffix(".report.json").read_text(encoding="utf-8"))
            self.assertEqual(report["slide_count"], 15)
            self.assertEqual(report["render_mode"], "source_native")

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_default_draft_selects_source_design_then_repairs_unknown_binding(self, request):
        manifest = json.loads((ROOT / "templates/severstal_template.json").read_text(encoding="utf-8"))
        design = manifest["layouts"]["title"]["variants"]["slide_003"]
        field_id = next(key for key, value in design["fields"].items()
                        if value["kind"] == "text" and value["max_characters"] > 10)
        final = {"template_version": "2.0", "title": "Отчёт", "slides": [{
            "layout_id": "title", "variant_id": "slide_003", "source_refs": [],
            "updates": [{"field_id": field_id, "text": "Отчёт"}],
        }]}
        invalid = json.loads(json.dumps(final))
        invalid["slides"][0]["updates"][0]["field_id"] = "slide_text_999999"
        request.side_effect = [response({"title": "Отчёт", "slides": [{"layout_id": "title", "variant_id": "slide_003"}],
                                          "icon_groups": []}), response(invalid), response(final)]
        with temporary_directory() as directory:
            config = Path(directory) / "llm.json"
            config.write_text(json.dumps({"base_url": "https://mock.invalid/v1", "model": "test",
                                           "api_key_env": "SUPERSLIDE_SOURCE_TEST_KEY"}), encoding="utf-8")
            output = Path(directory) / "draft.json"
            with patch.dict(os.environ, {"SUPERSLIDE_SOURCE_TEST_KEY": "test-secret"}):
                status, _, errors = self.invoke(["draft", "--task", "Краткий отчёт", "--config", str(config),
                                                 "--output", str(output)])
            self.assertEqual(status, 0, errors)
            self.assertEqual(request.call_count, 3)
            self.assertEqual(json.loads(output.read_text(encoding="utf-8")), final)
            metadata = json.loads(output.with_suffix(".generation.json").read_text(encoding="utf-8"))
            self.assertEqual(metadata["attempts"], 3)
            self.assertEqual(len(metadata["validation_errors"]), 1)


if __name__ == "__main__":
    unittest.main()
