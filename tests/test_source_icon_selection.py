"""Icon selection keeps both visual families available and rejects mixing."""

import copy
import json
import unittest
from collections import Counter
from unittest.mock import MagicMock, patch

from superslide.config import LLMSettings
from superslide.documents import DocumentBundle
from superslide.llm import LLMError, _rank_source_icons, _source_messages, generate_draft


def response(payload):
    value = MagicMock()
    value.__enter__.return_value = value
    value.read.return_value = json.dumps({"choices": [{
        "message": {"content": json.dumps(payload)}, "finish_reason": "stop",
    }]}).encode("utf-8")
    return value


class SourceIconSelectionTests(unittest.TestCase):
    @staticmethod
    def icons():
        return [{"id": "pin_badge", "name": "Геометка", "group_id": "navigation",
                 "background_style": "with_background", "has_background": True,
                 "path": "private/badge.svg"},
                {"id": "pin_outline", "name": "Геометка", "group_id": "navigation",
                 "background_style": "without_background", "has_background": False,
                 "path": "private/outline.svg"}]

    @staticmethod
    def catalog(variants=1):
        return {"contents": {"variants": {
            f"slide_{number:03}": {"name": "Пункты", "fields": {
                "slide_icon_1": {"kind": "icon"}, "slide_icon_2": {"kind": "icon"},
            }} for number in range(1, variants + 1)
        }}}

    @staticmethod
    def draft(*styles):
        return {"template_version": "2.0", "title": "Маршруты поставки", "slides": [
            {"layout_id": "contents", "variant_id": "slide_001", "source_refs": [],
             "updates": [{"field_id": f"slide_icon_{index}", "icon_id": icon_id}
                         for index, icon_id in enumerate(styles, 1)]},
        ]}

    def test_prompt_exposes_background_styles_without_asset_paths(self):
        records = self.icons()
        original = copy.deepcopy(records)
        messages = _source_messages("Пункты поставки", DocumentBundle(), {}, self.catalog(),
                                    icon_catalog=records)
        payload = json.loads(messages[1]["content"])
        self.assertEqual([item["background_style"] for item in payload["icon_catalog"]],
                         ["with_background", "without_background"])
        self.assertEqual([item["has_background"] for item in payload["icon_catalog"]], [True, False])
        self.assertEqual(payload["icon_style_policy"]["scope"], "slide")
        self.assertEqual(payload["icon_style_policy"]["max_background_styles"], 1)
        self.assertIn("На каждом слайде используй только один background_style", messages[0]["content"])
        self.assertNotIn("private/", json.dumps(messages))
        self.assertEqual(records, original)

    def test_shortlist_reserves_semantic_choices_for_each_style(self):
        badges = [{"id": f"badge_{index}", "name": "Москва", "background_style": "with_background"}
                  for index in range(70)]
        outlines = [{"id": f"outline_{index}", "name": "Регион", "background_style": "without_background"}
                    for index in range(40)]
        chosen = _rank_source_icons(badges + outlines, "Москва", 64)
        self.assertEqual(Counter(item["background_style"] for item in chosen),
                         {"with_background": 32, "without_background": 32})
        self.assertEqual([item["id"] for item in chosen[:32]], [f"badge_{i}" for i in range(32)])
        self.assertEqual([item["id"] for item in chosen[32:]], [f"outline_{i}" for i in range(32)])

    def test_shortlist_reuses_spare_capacity_and_excludes_nonselectable_art(self):
        badges = [{"id": f"badge_{index}", "background_style": "with_background"} for index in range(10)]
        outlines = [{"id": f"outline_{index}", "background_style": "without_background"} for index in range(70)]
        forbidden = {"id": "note", "name": "Москва", "selectable": False}
        chosen = _rank_source_icons([forbidden, *badges, *outlines], "Москва", 64)
        self.assertEqual(len(chosen), 64)
        self.assertNotIn(forbidden, chosen)
        self.assertEqual(Counter(item["background_style"] for item in chosen),
                         {"with_background": 10, "without_background": 54})
        self.assertEqual(_rank_source_icons(badges, "Москва", 0), [])

    def test_legacy_shortlist_without_background_metadata_keeps_semantic_order(self):
        records = [{"id": "factory", "name": "Завод"}, {"id": "pin", "name": "Москва"}]
        self.assertEqual([item["id"] for item in _rank_source_icons(records, "Москва", 2)], ["pin", "factory"])

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_direct_generation_repairs_mixed_styles_even_with_generic_validator(self, request):
        mixed = self.draft("pin_badge", "pin_outline")
        fixed = self.draft("pin_outline", "pin_outline")
        request.side_effect = [response(mixed), response(fixed)]
        result = generate_draft("Два пункта поставки", DocumentBundle(),
                                LLMSettings("https://mock.invalid/v1", "test"), lambda value: None,
                                {"type": "object"}, self.catalog(), template_version="2.0",
                                icon_catalog=self.icons())
        self.assertEqual(result.draft, fixed)
        self.assertEqual(request.call_count, 2)
        repair = json.loads(request.call_args_list[1].args[0].data)["messages"][-1]["content"]
        self.assertIn("with_background", repair)
        self.assertIn("without_background", repair)

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_batched_generation_repairs_mixed_styles(self, request):
        plan = {"title": "Маршруты поставки", "slides": [
            {"layout_id": "contents", "variant_id": "slide_001"}], "icon_groups": ["navigation"]}
        mixed = self.draft("pin_badge", "pin_outline")
        fixed = self.draft("pin_badge", "pin_badge")
        request.side_effect = [response(plan), response(mixed), response(fixed)]
        result = generate_draft("Два пункта поставки", DocumentBundle(),
                                LLMSettings("https://mock.invalid/v1", "test"), lambda value: None,
                                {"type": "object"}, self.catalog(variants=7), template_version="2.0",
                                icon_catalog=self.icons())
        self.assertEqual(result.draft, fixed)
        self.assertEqual(request.call_count, 3)
        planner = json.loads(json.loads(request.call_args_list[0].args[0].data)["messages"][1]["content"])
        self.assertEqual(planner["icon_groups"]["navigation"]["background_styles"],
                         ["with_background", "without_background"])
        self.assertIn("with_background", json.loads(request.call_args.args[0].data)["messages"][-1]["content"])

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_repeated_mixture_returns_structured_style_diagnostic(self, request):
        request.return_value = response(self.draft("pin_badge", "pin_outline"))
        with self.assertRaises(LLMError) as caught:
            generate_draft("Два пункта поставки", DocumentBundle(),
                           LLMSettings("https://mock.invalid/v1", "test", max_repair_attempts=0),
                           lambda value: None, {"type": "object"}, self.catalog(), template_version="2.0",
                           icon_catalog=self.icons())
        self.assertEqual(caught.exception.diagnostics[0]["validation_details"],
                         {"slide_index": 1, "reason": "mixed_icon_background"})

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_different_slides_can_choose_different_styles(self, request):
        draft = self.draft("pin_badge", "pin_badge")
        draft["slides"] += self.draft("pin_outline", "pin_outline")["slides"]
        request.return_value = response(draft)
        result = generate_draft("Два слайда", DocumentBundle(),
                                LLMSettings("https://mock.invalid/v1", "test"), lambda value: None,
                                {"type": "object"}, self.catalog(), template_version="2.0",
                                icon_catalog=self.icons())
        self.assertEqual(result.draft, draft)
        self.assertEqual(request.call_count, 1)


if __name__ == "__main__":
    unittest.main()
