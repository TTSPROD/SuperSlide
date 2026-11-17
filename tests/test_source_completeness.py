"""Generated drafts must not leak omitted template sample text into slides."""
from copy import deepcopy
from pathlib import Path
from zipfile import ZipFile
import unittest
import json
from unittest.mock import MagicMock, patch

from support import temporary_directory
from superslide.source_completeness import (
    clear_unfilled_template_text, incomplete_slot_groups, incomplete_template_fields, is_template_placeholder,
)
from superslide.source_template import (
    A, ROOT, build_source_presentation, load_source_manifest, source_prompt_catalog, _xml,
)
from superslide.config import LLMSettings
from superslide.documents import DocumentBundle
from superslide.llm import LLMError, generate_draft


class SourceCompletenessTests(unittest.TestCase):
    @staticmethod
    def response(value):
        stream = MagicMock()
        stream.__enter__.return_value = stream
        stream.read.return_value = json.dumps({"choices": [{"message": {
            "content": json.dumps(value, ensure_ascii=False)}, "finish_reason": "stop"}]}).encode("utf-8")
        return stream

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_incomplete_model_group_is_repaired_before_publishing(self, request):
        fields = {f"slide_text_{i}": {"kind": "text", "sample_text": f"Название #{i}",
                                     "paragraph_count": 1, "max_characters": 80} for i in range(1, 4)}
        catalog = {"cards": {"variants": {"slide_001": {"fields": fields, "item_slots": {"groups": [
            {"kind": "numbered_text_fields", "field_ids": list(fields), "count": 3}]}}}}}
        draft = {"template_version": "2.0", "title": "Решения", "slides": [
            {"layout_id": "cards", "variant_id": "slide_001", "updates": [
                {"field_id": "slide_text_1", "text": "Продавец"}]}]}
        complete = deepcopy(draft)
        complete["slides"][0]["updates"].extend([
            {"field_id": "slide_text_2", "text": "Оператор"},
            {"field_id": "slide_text_3", "text": "Рекрутер"}])
        request.side_effect = [self.response(draft), self.response(complete)]
        result = generate_draft("Три решения", DocumentBundle(), LLMSettings("https://mock.invalid/v1", "test"),
                                lambda value: None, {"type": "object"}, catalog, template_version="2.0")
        self.assertEqual(result.draft, complete)
        self.assertEqual(request.call_count, 2)
        repair = json.loads(request.call_args.args[0].data)
        self.assertIn("slide_text_2", repair["messages"][-1]["content"])
        self.assertIn("не выдумывай", repair["messages"][-1]["content"])

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_incomplete_model_group_fails_instead_of_returning_placeholders(self, request):
        catalog = {"cards": {"variants": {"slide_001": {"fields": {}, "item_slots": {"groups": [
            {"kind": "numbered_text_fields", "field_ids": ["slide_text_1", "slide_text_2"], "count": 2}]}}}}}
        request.return_value = self.response({"title": "Отчёт", "slides": [
            {"layout_id": "cards", "variant_id": "slide_001", "updates": []}]})
        with self.assertRaises(LLMError) as caught:
            generate_draft("Одно решение", DocumentBundle(),
                           LLMSettings("https://mock.invalid/v1", "test", max_repair_attempts=0),
                           lambda value: None, {"type": "object"}, catalog, template_version="2.0")
        self.assertEqual(caught.exception.code, "validation")
        self.assertIn("slide_text_2", caught.exception.diagnostics[0]["validation_error"])

    def test_detects_only_explicit_template_samples(self):
        for sample in ("Lorem ipsum dolor sit amet", "Название #3", "Текстовый заголовок #2 И небольшой подзаголовок",
                       "Заголовок блока №4", "И небольшой подзаголовок"):
            with self.subTest(sample=sample):
                self.assertTrue(is_template_placeholder(sample))
        for sample in ("Итог", "Сложные сценарии", "Название продукта", "Услуги и технологии", "", None):
            with self.subTest(sample=sample):
                self.assertFalse(is_template_placeholder(sample))

    def test_adds_missing_updates_without_modifying_existing_data_or_input(self):
        catalog = {"cards": {"variants": {"slide_001": {"fields": {
            "slide_text_1": {"kind": "text", "sample_text": "Название #1"},
            "slide_text_2": {"kind": "text", "sample_text": "Название #2 Lorem ipsum", "safe_to_clear_if_unused": True},
            "slide_text_3": {"kind": "text", "sample_text": "Итог"},
            "layout_text_4": {"kind": "text", "sample_text": "Lorem ipsum"},
            "slide_icon_5": {"kind": "icon", "sample_text": "Lorem ipsum"},
            "slide_text_6": {"kind": "text", "sample_text": "Название #6 Lorem ipsum", "safe_to_clear_if_unused": True, "paragraph_count": 2,
                             "text_budget": {"paragraph_roles": ["heading", "body"], "max_paragraphs": 2}},
        }}}}}
        draft = {"title": "Отчёт", "slides": [{"layout_id": "cards", "variant_id": "slide_001",
                 "updates": [{"field_id": "slide_text_1", "text": "Системы 2D"}], "notes": "Сохранить"}]}
        original = deepcopy(draft)
        result = clear_unfilled_template_text(draft, catalog)
        self.assertEqual(draft, original)
        self.assertEqual(result["slides"][0]["notes"], "Сохранить")
        self.assertEqual(result["slides"][0]["updates"], [
            {"field_id": "slide_text_1", "text": "Системы 2D"},
            {"field_id": "slide_text_2", "text": ""},
            {"field_id": "slide_text_6", "paragraphs": ["", ""]},
        ])
        self.assertEqual(clear_unfilled_template_text(result, catalog), result)

    def test_incomplete_visual_group_rejects_missing_empty_and_template_values(self):
        catalog = {"cards": {"variants": {"slide_001": {"fields": {}, "item_slots": {
            "groups": [{"kind": "numbered_text_fields", "count": 3,
                        "field_ids": ["slide_text_1", "slide_text_2", "slide_text_3"]}],
        }}}}}
        draft = {"slides": [{"layout_id": "cards", "variant_id": "slide_001", "updates": [
            {"field_id": "slide_text_1", "text": "Продажи"},
            {"field_id": "slide_text_2", "paragraphs": [" ", "Lorem ipsum"]},
        ]}]}
        before = deepcopy(draft)
        self.assertEqual(incomplete_slot_groups(draft, catalog, [4]), [{"slide_index": 4, "reason": "unfilled_group",
                         "field_ids": ["slide_text_2", "slide_text_3"], "limit": 3, "actual": 1}])
        self.assertEqual(draft, before)
        draft["slides"][0]["updates"][1]["paragraphs"] = ["Оператор", "Обработка заявок"]
        draft["slides"][0]["updates"].append({"field_id": "slide_text_3", "text": "Рекрутер"})
        self.assertEqual(incomplete_slot_groups(draft, catalog), [])

    def test_unknown_geometry_reports_unfilled_template_without_inventing_or_removing(self):
        catalog = {"cards": {"variants": {"slide_001": {
            "fields": {"slide_text_1": {"kind": "text", "sample_text": "Название #1"},
                       "slide_text_2": {"kind": "text", "sample_text": "Lorem ipsum dolor sit amet"},
                       "slide_text_3": {"kind": "text", "sample_text": "Итог"}},
            "item_slots": {"groups": [{"kind": "numbered_text_fields", "field_ids": ["slide_text_1", "slide_text_4"]}]},
        }}}}
        draft = {"slides": [{"layout_id": "cards", "variant_id": "slide_001", "updates": []}]}
        before = deepcopy(draft)
        self.assertEqual(incomplete_template_fields(draft, catalog, [3]), [
            {"slide_index": 3, "field_id": "slide_text_2", "reason": "template_placeholder"}])
        self.assertEqual(clear_unfilled_template_text(draft, catalog), draft)
        self.assertEqual(draft, before)
        draft["slides"][0]["updates"].append({"field_id": "slide_text_2", "text": "Содержательный тезис"})
        self.assertEqual(incomplete_template_fields(draft, catalog), [])

    def test_singleton_numbered_group_is_still_checked_for_placeholder(self):
        catalog = {"cards": {"variants": {"slide_001": {
            "fields": {"slide_text_53": {"kind": "text", "sample_text": "Название #1 Lorem ipsum"}},
            "item_slots": {"groups": [{"kind": "numbered_text_fields", "field_ids": ["slide_text_53"], "count": 1}]},
        }}}}
        draft = {"slides": [{"layout_id": "cards", "variant_id": "slide_001", "updates": []}]}
        self.assertEqual(incomplete_template_fields(draft, catalog), [
            {"slide_index": 1, "field_id": "slide_text_53", "reason": "template_placeholder"}])

    def test_real_slide_omitted_placeholder_is_empty_in_generated_pptx(self):
        template = ROOT / "templates/severstal_source.pptx"
        manifest = ROOT / "templates/severstal_template.json"
        catalog = source_prompt_catalog(load_source_manifest(manifest))
        fields = catalog["bullets"]["variants"]["slide_012"]["fields"]
        self.assertIn("Lorem ipsum", fields["slide_text_20"]["sample_text"])
        original = {"template_version": "2.0", "title": "Проверка", "slides": [
            {"layout_id": "bullets", "variant_id": "slide_012", "updates": []},
        ]}
        # Mark a field safe only when separate structural metadata establishes
        # that its decoration is independently optional.
        fields["slide_text_20"]["safe_to_clear_if_unused"] = True
        cleared = clear_unfilled_template_text(original, catalog)
        self.assertTrue(any(u["field_id"] == "slide_text_20" for u in cleared["slides"][0]["updates"]))
        with temporary_directory() as temporary:
            path = Path(temporary) / "result.pptx"
            build_source_presentation(cleared, template, manifest, path)
            with ZipFile(path) as archive:
                root = _xml(archive.read("ppt/generated/slide1.xml"))
            body = root.xpath(".//p:sp[p:nvSpPr/p:cNvPr[@id='20']]/p:txBody",
                              namespaces={"a": A, "p": "http://schemas.openxmlformats.org/presentationml/2006/main"})[0]
            self.assertEqual("".join(body.xpath(".//a:t/text()", namespaces={"a": A})), "")


if __name__ == "__main__":
    unittest.main()
