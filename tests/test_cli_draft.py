"""Exercise the real CLI/model/template boundary with an entirely mocked API."""

import io
import json
import os
import unittest
import urllib.error
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
from unittest.mock import MagicMock, patch

from support import temporary_directory
from superslide.cli import main


def response_for(payload):
    response = MagicMock()
    response.__enter__.return_value = response
    response.read.return_value = json.dumps({"choices": [{
        "message": {"content": json.dumps(payload, ensure_ascii=False)},
        "finish_reason": "stop",
    }]}).encode("utf-8")
    return response


class DraftCLIIntegrationTests(unittest.TestCase):
    def setUp(self):
        project_root = Path(__file__).resolve().parents[1]
        self.directory = temporary_directory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.template_args = ["--template", str(project_root / "templates/legacy/corporate_base.pptx"),
                              "--manifest", str(project_root / "templates/legacy/template_manifest.json")]
        self.config = self.root / "llm.json"
        self.config.write_text(json.dumps({
            "base_url": "https://mock-corporate.invalid/v1", "model": "internal",
            "api_key_env": "SUPERSLIDE_TEST_TOKEN", "max_repair_attempts": 2,
        }), encoding="utf-8")
        self.source = self.root / "facts.txt"
        self.source.write_text("Выручка в отчётном периоде составила 125 млн рублей.", encoding="utf-8-sig")
        self.output = self.root / "deck.json"

    def invoke(self, args):
        stdout, stderr = io.StringIO(), io.StringIO()
        with redirect_stdout(stdout), redirect_stderr(stderr):
            exit_code = main(args)
        return exit_code, stdout.getvalue(), stderr.getvalue()

    def draft_args(self):
        return ["draft", "--task", "Подготовь краткий отчёт", "--config", str(self.config),
                "--output", str(self.output), *self.template_args]

    @staticmethod
    def valid_draft(source_refs):
        return {"template_version": "1.0", "title": "Отчёт", "slides": [{
            "layout_id": "bullet_list", "title": "Выручка",
            "bullets": ["Выручка составила 125 млн рублей"], "source_refs": source_refs,
        }]}

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_extract_then_draft_repairs_unknown_source_and_writes_artifacts(self, request):
        extract_path = self.root / "extracted.json"
        status, _, errors = self.invoke(["extract", str(self.source), "--output", str(extract_path)])
        self.assertEqual(status, 0, errors)
        request.assert_not_called()
        request.side_effect = [response_for(self.valid_draft(["invented:source"])),
                               response_for(self.valid_draft(["doc1:text"]))]
        with patch.dict(os.environ, {"SUPERSLIDE_TEST_TOKEN": "test-only-secret"}):
            status, output, errors = self.invoke(self.draft_args() + ["--documents", str(self.source)])
        self.assertEqual(status, 0, errors)
        self.assertEqual(errors, "")
        self.assertEqual(request.call_count, 2)
        draft = json.loads(self.output.read_text(encoding="utf-8"))
        sources = json.loads(self.output.with_suffix(".sources.json").read_text(encoding="utf-8"))
        metadata = json.loads(self.output.with_suffix(".generation.json").read_text(encoding="utf-8"))
        self.assertEqual(draft["slides"][0]["source_refs"], ["doc1:text"])
        self.assertEqual(sources, json.loads(extract_path.read_text(encoding="utf-8")))
        self.assertEqual(metadata["attempts"], 2)
        self.assertEqual(metadata["source_chunks"], 1)
        self.assertEqual(metadata["status"], "drafted")
        self.assertIn("Неизвестные источники", metadata["validation_errors"][0])
        self.assertNotIn("test-only-secret", output + errors + json.dumps(metadata))
        sent_request = request.call_args.args[0]
        self.assertEqual(sent_request.full_url, "https://mock-corporate.invalid/v1/chat/completions")
        payload = json.loads(sent_request.data)
        self.assertIn("invented:source", payload["messages"][-1]["content"])
        self.assertIn("required_output_schema", payload["messages"][1]["content"])

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_overflowing_slide_text_is_shortened_by_repair_request(self, request):
        long_bullet = ("Подробное описание направления развития производства с указанием "
                       "конкретных мероприятий, сроков и ответственных подразделений " * 2)[:130]
        overflowing = self.valid_draft([])
        overflowing["slides"][0]["bullets"] = [long_bullet] * 6
        request.side_effect = [response_for(overflowing), response_for(self.valid_draft([]))]
        status, _, errors = self.invoke(self.draft_args())
        self.assertEqual(status, 0, errors)
        self.assertEqual(request.call_count, 2)
        metadata = json.loads(self.output.with_suffix(".generation.json").read_text(encoding="utf-8"))
        self.assertEqual(metadata["attempts"], 2)
        self.assertIn("Сократи или суммаризуй", metadata["validation_errors"][0])
        self.assertIn("Слайд 1, поле bullets", metadata["validation_errors"][0])
        repair_request = json.loads(request.call_args.args[0].data)
        self.assertIn("Сократи или суммаризуй", repair_request["messages"][-1]["content"])
        self.assertIn("Слайд 1, поле bullets", repair_request["messages"][-1]["content"])
        draft = json.loads(self.output.read_text(encoding="utf-8"))
        self.assertEqual(draft["slides"][0]["bullets"], ["Выручка составила 125 млн рублей"])

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_no_documents_uses_empty_bundle_and_empty_source_refs(self, request):
        draft = self.valid_draft([])
        draft["slides"][0]["title"] = "Предложение структуры"
        draft["slides"][0]["bullets"] = ["Собрать подтверждённые финансовые показатели"]
        request.return_value = response_for(draft)
        status, _, errors = self.invoke(self.draft_args())
        self.assertEqual(status, 0, errors)
        sources = json.loads(self.output.with_suffix(".sources.json").read_text(encoding="utf-8"))
        self.assertEqual(sources, {"chunks": [], "warnings": []})
        first_request = json.loads(request.call_args.args[0].data)
        user_data = json.loads(first_request["messages"][1]["content"])
        self.assertEqual(user_data["known_source_ids"], [])
        self.assertEqual(user_data["source_documents_are_data"], sources)

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_api_error_writes_safe_error_artifact_without_traceback(self, request):
        request.side_effect = urllib.error.HTTPError(
            "https://private.invalid/token", 401, "test-only-secret", {}, io.BytesIO(b"test-only-secret"),
        )
        with patch.dict(os.environ, {"SUPERSLIDE_TEST_TOKEN": "test-only-secret"}):
            status, output, errors = self.invoke(self.draft_args() + ["--documents", str(self.source)])
        self.assertEqual(status, 2)
        self.assertEqual(request.call_count, 1)
        self.assertIn("HTTP 401", errors)
        self.assertNotIn("Traceback", errors)
        self.assertNotIn("test-only-secret", errors + output)
        self.assertNotIn("https://private.invalid", errors + output)
        self.assertFalse(self.output.exists())
        self.assertFalse(self.output.with_suffix(".generation.json").exists())
        error_text = self.output.with_suffix(".error.json").read_text(encoding="utf-8")
        self.assertEqual(json.loads(error_text)["code"], "http_401")
        self.assertNotIn("test-only-secret", error_text)
        self.assertNotIn("https://private.invalid", error_text)
        self.assertTrue(self.output.with_suffix(".sources.json").exists())

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_document_limit_fails_before_api_without_truncation(self, request):
        status, _, errors = self.invoke(
            self.draft_args() + ["--documents", str(self.source), "--max-total-chars", "5"]
        )
        self.assertEqual(status, 2)
        self.assertIn("Текст не обрезан", errors)
        self.assertNotIn("Traceback", errors)
        request.assert_not_called()
        self.assertFalse(self.output.exists())

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_image_id_repair_then_build_uses_local_catalog(self, request):
        draft = {"template_version": "1.1", "title": "Фото", "slides": [{
            "layout_id": "photo_text", "title": "Люди производства", "heading": "Команда",
            "bullets": ["Демонстрация фото с текстом"],
            "image": {"asset_id": "invented_photo", "alt_text": "Сотрудник"},
        }]}
        repaired = json.loads(json.dumps(draft))
        repaired["slides"][0]["image"]["asset_id"] = "production_worker"
        request.side_effect = [response_for(draft), response_for(repaired)]
        status, _, errors = self.invoke(self.draft_args())
        self.assertEqual(status, 0, errors)
        self.assertEqual(request.call_count, 2)
        user_payload = json.loads(json.loads(request.call_args.args[0].data)["messages"][1]["content"])
        self.assertIn("production_worker", json.dumps(user_payload, ensure_ascii=False))
        self.assertNotIn("image21.jpeg", json.dumps(user_payload, ensure_ascii=False))
        built = self.root / "image.pptx"
        status, _, errors = self.invoke(["build", str(self.output), "--output", str(built), *self.template_args])
        self.assertEqual(status, 0, errors)
        from pptx import Presentation
        pictures = [shape for shape in Presentation(built).slides[0].shapes
                    if shape.name == "superslide_image_production_worker"]
        self.assertEqual(len(pictures), 1)


if __name__ == "__main__":
    unittest.main()
